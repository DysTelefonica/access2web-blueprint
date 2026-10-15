#!/usr/bin/env python3
# ci-pattern asset — docker-ownership gate (DysTelefonica/team-skills#454)
"""Pure gate for docker resource ownership and protection (HR-*).

Un agente solo para o borra los contenedores, volúmenes y redes que
haya creado el mismo en la tarea en curso, con su prefijo, y nunca
toca recursos con la etiqueta de proteccion del consumer
(`apap.protect=true` por defecto, parametrizable), de la plataforma
(Coolify o su proxy) ni de los runners GitHub Actions self-hosted.
El incidente que motiva este gate (APAP_WEB, 2026-10-12, 17:15):
un agente tomo la base de produccion (`apap-pg-test`) por un
contenedor de test propio y la borro con `docker rm -fv`.

El gate no llama al demonio de Docker: opera sobre la lista de
recursos que el agente declara querer tocar (payload de un PR,
salida de `docker ps` capturada, etc.) y devuelve rc=0/1/2 segun
el contrato abajo. Asi es detectable en CI sin daemon y fail-closed
ante entradas invalidas (HR-3).

Invariantes:
- rc=0 -> el agente puede tocar todos los recursos declarados
  (cada uno tiene el prefijo del agente y ninguno lleva ninguna
   etiqueta de proteccion).
- rc=1 -> hay hallazgos; se imprime cada uno con su nombre y el
  motivo; el agente debe reducir el set antes de actuar.
- rc=2 -> entrada invalida (formato, agente vacio, etiquetas
  declaradas vacias, recursos vacios); falla cerrado (HR-3).

Uso (CLI):

    python3 check_docker_ownership.py \
        --agent-prefix skills1- \
        --protect-label apap.protect=true \
        --resource "container:apap-pg-test" \
        --resource "container:skills1-pg-454:labels:agent=skills1"
"""
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from typing import Iterable, Optional


SUPPORTED_KINDS = {"container", "volume", "network"}
LABEL_KEY_RE = re.compile(r"^[a-zA-Z0-9][a-zA-Z0-9_.-]{0,62}$")
LABEL_VALUE_RE = re.compile(r"^[^\n\r]{0,63}$")


class GateError(Exception):
    """Entrada invalida; rc=2, fail-closed (HR-3)."""


@dataclass(frozen=True)
class Resource:
    """Recurso Docker declarado para evaluacion.

    `name` es el nombre canonico del recurso. `labels` es un
    `dict` etiqueta -> valor, leido de `docker inspect`. Si
    `labels` es None se trata como ausente: cualquier
    `protect_label` que aplique se considera ausente.
    """

    kind: str
    name: str
    labels: Optional[dict] = None

    def label(self, key: str) -> Optional[str]:
        if not self.labels:
            return None
        return self.labels.get(key)


def _parse_label(spec: str) -> str:
    """`key=value` -> `key=value`. Solo valida, no parsea a dict.

    Lanza GateError si la sintaxis no encaja con label_key=label_value
    (las keys y valores siguen el formato `docker label`).
    """
    if "=" not in spec:
        raise GateError(
            f"protect-label invalido {spec!r}: formato esperado 'key=value'"
        )
    key, _, value = spec.partition("=")
    if not LABEL_KEY_RE.match(key):
        raise GateError(
            f"protect-label {spec!r}: key {key!r} no encaja con "
            "[a-zA-Z0-9][a-zA-Z0-9_.-]{0,62}"
        )
    if not LABEL_VALUE_RE.match(value):
        raise GateError(
            f"protect-label {spec!r}: value no encaja con "
            "longitud <= 63 sin saltos de linea"
        )
    return spec


def _parse_resource(spec: str) -> Resource:
    """`kind:name[:labels:k=v,k=v]`.

    Acepta dos formas:
      - `kind:name`          (sin etiquetas)
      - `kind:name:labels:k=v,k=v` (etiquetas embebidas)

    Los labels embebidos permiten que el payload del agente declare
    el resultado de `docker inspect` sin que el gate tenga que
    tocar el daemon.
    """
    parts = spec.split(":")
    if len(parts) == 2:
        kind, name = parts
        return Resource(kind=kind, name=name, labels=None)
    if len(parts) >= 4 and parts[2] == "labels":
        kind, name = parts[0], parts[1]
        labels: dict = {}
        for kv in ",".join(parts[3:]).split(","):
            if not kv:
                continue
            if "=" not in kv:
                raise GateError(
                    f"recurso {spec!r}: label {kv!r} sin '='"
                )
            k, _, v = kv.partition("=")
            if not LABEL_KEY_RE.match(k):
                raise GateError(
                    f"recurso {spec!r}: label key {k!r} invalida"
                )
            if not LABEL_VALUE_RE.match(v):
                raise GateError(
                    f"recurso {spec!r}: label value {v!r} invalido "
                    "(longitud <= 63 sin saltos de linea)"
                )
            labels[k] = v
        return Resource(kind=kind, name=name, labels=labels)
    raise GateError(
        f"recurso {spec!r}: formato esperado "
        "'kind:name' o 'kind:name:labels:k=v,k=v'"
    )


def _check_labels(resource: Resource) -> list[str]:
    return [
        f"label {k}={v!r}"
        for k, v in (resource.labels or {}).items()
        if LABEL_KEY_RE.match(k) and LABEL_VALUE_RE.match(v)
    ]


def evaluate(
    resources: Iterable[Resource],
    agent_prefix: str,
    protect_labels: Iterable[str],
) -> list[str]:
    """Devuelve la lista de hallazgos para los `resources`.

    Cada hallazgo es una frase accionable: nombra el recurso y la
    causa. Lista vacia significa que el agente puede tocar todos
    los recursos declarados.
    """
    out: list[str] = []
    protect_labels = list(protect_labels)
    if not agent_prefix:
        raise GateError("agent-prefix vacio; rc=2 (HR-3)")
    if not protect_labels:
        raise GateError(
            "no se declaro ninguna protect-label; el gate falla cerrado "
            "porque no hay forma de distinguir un recurso protegido"
        )
    for r in resources:
        if r.kind not in SUPPORTED_KINDS:
            raise GateError(
                f"recurso {r.name!r}: kind {r.kind!r} no soportado "
                f"(esperado uno de {sorted(SUPPORTED_KINDS)})"
            )
        if not r.name:
            raise GateError(f"recurso sin nombre (kind={r.kind!r})")
        if not r.name.startswith(agent_prefix):
            out.append(
                f"recurso {r.kind}/{r.name}: no lleva el prefijo del "
                f"agente {agent_prefix!r}; el agente solo borra sus "
                "propios recursos (Refs DysTelefonica/team-skills#454)"
            )
            continue
        seen_protect = [
            label
            for label in protect_labels
            if r.label(label.partition("=")[0]) == label.partition("=")[2]
        ]
        if seen_protect:
            out.append(
                f"recurso {r.kind}/{r.name}: lleva etiqueta(s) de "
                f"proteccion {seen_protect}; ni siquiera el agente "
                "propietario la borra (Refs #454)"
            )
            continue
        _ = _check_labels(r)
    return out


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(
        description="docker-ownership gate (Refs DysTelefonica/team-skills#454)"
    )
    ap.add_argument("--agent-prefix", required=True)
    ap.add_argument(
        "--protect-label",
        action="append",
        default=[],
        help="`key=value`; puede repetirse",
    )
    ap.add_argument(
        "--resource",
        action="append",
        default=[],
        help="`kind:name[:labels:k=v,k=v]`; puede repetirse",
    )
    args = ap.parse_args(argv)

    try:
        protect = [_parse_label(p) for p in args.protect_label]
        resources = [_parse_resource(r) for r in args.resource]
    except GateError as exc:
        print(f"DOCKER-OWNERSHIP INVALID: {exc}", file=sys.stderr)
        return 2

    try:
        findings = evaluate(resources, args.agent_prefix, protect)
    except GateError as exc:
        print(f"DOCKER-OWNERSHIP INVALID: {exc}", file=sys.stderr)
        return 2

    if findings:
        print(f"DOCKER-OWNERSHIP FAIL: {len(findings)} finding(s)")
        for f in findings:
            print(f"  - {f}")
        return 1
    print("DOCKER-OWNERSHIP OK: all declared resources are own and non-protected")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

# 02 · Topología del ecosistema — Matriz de dependencias

## Propósito

Vista relacional de las ocho aplicaciones: quién consume a quién, con qué intensidad y con qué excepciones. Sustituye a cualquier inferencia libre sobre la arquitectura actual; solo se documentan dependencias demostradas.

## Estados de dependencia

| Etiqueta | Significado |
|---|---|
| **Fuerte** | Evidencia estática o documental directa (código, doc, grafo). |
| **Probable** | Existe indicio, falta confirmación dirigida. |
| **Excepción** | Solo aplica a un flujo identificado; no generaliza. |
| **Desconocida** | No hay evidencia en esta pasada; queda como pendiente. |

## Matriz inicial (resumen de `exploration.md`)

Esta tabla es **provisional** y solo refleja la pasada de exploración. Se reescribirá al cierre del Lote 9 con la reconciliación transversal.

| Aplicación | Lanzadera | Expedientes | Observaciones |
|---|---|---|---|
| Lanzadera | Origen confirmado | No consumidora demostrada | Falta catálogo de las otras siete. |
| Gestion_Riesgos | Fuerte | Fuerte | Pendiente: naturaleza de `TbExpedientes1`. |
| No_Conformidades | Probable / fuerte parcial | Fuerte parcial + Excepción `Nemotecnico` | Separar flujos proyecto vs auditoría. |
| Condor | Probable | Fuerte (FK `tbSolicitudes.idExpediente`) | Falta backend y Dysflow. |
| HPS_Solicitudes | Desconocida en runtime | Fuerte documental | Sin checkout/binary localizado. |
| HPS | Fuerte parcial | Fuerte | Distinguir `HPST.accdb` actual de consolidaciones históricas. |
| Brass | Fuerte documental / parcial | Desconocida | Falta rastreo de `IDExpediente`/`CodExp`. |
| Expedientes | Fuerte parcial | Núcleo | Confirmar consumidores y ciclo de baja. |

## Reglas de la matriz

- Una celda **Fuerte** exige cita a código, doc o grafo; si no, baja a **Probable** o **Desconocida**.
- Una **Excepción** describe el flujo exacto (formulario, evento, módulo); no se eleva a regla general.
- APAP y APAP_WEB no aparecen en esta matriz.
- Cualquier afirmación de "tiempo real" exige `Verified-runtime`; el resto queda en `Verified-static` o `Likely`.

## Checklist

- [ ] Cada celda lleva una evidencia mínima (ruta a doc, módulo o símbolo).
- [ ] Las excepciones se separan visualmente del cuerpo principal.
- [ ] La versión inicial se marca como `provisional` hasta el Lote 9.

## Siguiente paso

Mantener esta matriz sincronizada con `01-inventario-aplicaciones.md` y con las fichas de `03-aplicaciones/` cada vez que se cierre un lote.

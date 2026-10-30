# Gates — línea base de fricción y job de veredicto único (#1179)

Fase 0 y fase 1 de la propuesta de un job de gates por repo con un veredicto
único (#1179, aprobado por el dueño para las fases 0 y 1 el 2026-10-24).

- `measure_gates.py` (**fase 0**): la línea base de fricción de un repo sobre un
  export de check-runs (JSON, sin red): rojos secundarios, gates rojos por
  commit, re-ejecuciones y tiempo hasta el verde.
- `run_gates.py` (**fase 1**): el job en sombra que ejecuta los gates declarados
  una vez y emite un veredicto único. En sombra: publica su check sin que nadie
  lo exija y no cambia required checks ni protección de rama.

```bash
# Fase 0: la línea base (el export lo produce el operador con la API).
python3 assets/gates/measure_gates.py --export <export.json> [--json]

# Fase 1: el veredicto único (los gates se declaran en el repo del consumer:
# `nombre: comando` por línea).
python3 assets/gates/run_gates.py --gates <fichero-de-gates> [--json]
```

El método para producir el export, las cifras medidas y los criterios del flip
están en `references/gates-baseline.md`; el fragmento que cablea el job en
sombra está en `references/porting-guide.md` (G0.5b).

Exit: fase 0, 0 métricas · 2 export inválido o sin commits (HR-3). Fase 1, 0
todos en verde · 1 algún gate en rojo · 2 declaración rota o un gate que no se
ejecutó: un gate que no midió nunca dice «OK».

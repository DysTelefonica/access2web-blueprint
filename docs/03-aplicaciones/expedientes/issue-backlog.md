# Backlog ejecutable de migración de Expedientes

Este índice enlaza las 110 unidades con sus issues reales. La [issue padre #170](https://github.com/DysTelefonica/access2web-blueprint/issues/170) gobierna el alcance completo; cada fila mantiene un límite inferior a 400 líneas cambiadas y conserva sus dependencias explícitas.

## Lectura rápida

- **Cobertura:** 110 unidades para 63 capacidades: 110 issues: #172–#264 y #267–#283; #171 quedó cerrada como variante superseded.
- **Entrega:** 68 unidades independientes y 42 encadenadas; estrategia base `stacked-to-main`.
- **Forecast:** 15.595 líneas cambiadas estimadas, con 110 PRs estrictamente menores de 400 líneas.
- **Gate CPV:** pertenece a Q01 / EXP-CAP-016; Q02 no lo gobierna.
- **Regla de revisión:** una unidad por PR, pruebas y documentación junto a la unidad, sin `size:exception`.
- **Fuente de detalle:** [tasks.md](../../../openspec/changes/expedientes-web-migration/tasks.md) define aceptación, rollback y dependencias de plataforma.

## Carriles

| Carril | Unidades | Issues | Alcance |
|---|---:|---:|---|
| Documentación y preparación | D01–D54 | 54 | Capacidades, diccionarios, mappings, contratos y cutover |
| Fundación | F01–F04 | 4 | Puertos, dominio, configuración, esquema y unidad de trabajo |
| Ciclo de vida | C01–C04 | 4 | Alta, edición, baja y transiciones |
| Datos relacionados | R01–R07 | 7 | Un vertical por capacidad relacionada |
| Consulta y catálogos | Q01–Q03 | 3 | Catálogos, búsquedas, bandeja, exportación y tareas |
| Escrituras resilientes | W01–W03 | 3 | Autosave, concurrencia, idempotencia y feedback |
| Intercambio E2E | E01–E07 | 7 | DTO, JSON canónico, hash, lotes, paquetes y sesiones |
| Seguridad | A01–A03 | 3 | Autorización, principal y auditoría |
| Runtime | T01 | 1 | Readiness, configuración, caché y binding |
| Integraciones | H01–H06 | 6 | HPS, AGEDYS, Riesgos, NC, notificaciones y documentos |
| E2E web | Z01–Z03 | 3 | REST y journeys Playwright |
| Migración de datos | M01–M04 | 4 | Extracción, promoción y reconciliación |
| Aceptación | U01–U10 | 10 | Un UAT por dominio de spec |
| Retirada | L01 | 1 | Puertas para la retirada verificable del legacy |

## Cadena de entrega

```text
Documentación D01–D54
  └─ Fundación F01–F04
      ├─ Ciclo de vida C01–C04 ─┬─ Relacionados R01–R07 ─┬─ Escrituras W01–W03
      │                          └─ Consultas Q01–Q03 ┘
      ├─ Intercambio E01–E07 ──→ REST y journeys Z01–Z03
      ├─ Seguridad A01–A03 ────→ Runtime T01
      └─ Integraciones H01–H06
Migración M01–M04 + UAT U01–U09 + cutover D50 ─→ U10 ─→ retirada L01
```

Las rutas `stacked` aterrizan después de su dependencia inmediata; las rutas `independent` pueden avanzar en paralelo cuando sus contratos de entrada estén aceptados. Las dependencias externas (`#21`, `#31`, etc.) remiten al backlog de plataforma y no se duplican.

## Issues por unidad de trabajo

| Unidad | Issue | Título | Ruta y dependencias | Estimación |
|---|---|---|---|---:|
| D01 | [#172](https://github.com/DysTelefonica/access2web-blueprint/issues/172) | docs(exp): document CAP-001 | `independent` · D0 ledger | +70/-10 |
| D02 | [#173](https://github.com/DysTelefonica/access2web-blueprint/issues/173) | docs(exp): document CAP-002 | `independent` · D0 | +70/-10 |
| D03 | [#174](https://github.com/DysTelefonica/access2web-blueprint/issues/174) | docs(exp): document CAP-003 | `independent` · D0 | +70/-10 |
| D04 | [#175](https://github.com/DysTelefonica/access2web-blueprint/issues/175) | docs(exp): document CAP-004 | `independent` · D0 | +70/-10 |
| D05 | [#176](https://github.com/DysTelefonica/access2web-blueprint/issues/176) | docs(exp): document CAP-005 | `independent` · D0 | +70/-10 |
| D06 | [#177](https://github.com/DysTelefonica/access2web-blueprint/issues/177) | docs(exp): document CAP-006 | `independent` · D0 | +70/-10 |
| D07 | [#178](https://github.com/DysTelefonica/access2web-blueprint/issues/178) | docs(exp): document CAP-007 | `independent` · D0 | +70/-10 |
| D08 | [#179](https://github.com/DysTelefonica/access2web-blueprint/issues/179) | docs(exp): document CAP-008 | `independent` · D0 | +70/-10 |
| D09 | [#180](https://github.com/DysTelefonica/access2web-blueprint/issues/180) | docs(exp): document CAP-009 | `independent` · D0 | +70/-10 |
| D10 | [#181](https://github.com/DysTelefonica/access2web-blueprint/issues/181) | docs(exp): document CAP-010 | `independent` · D0 | +70/-10 |
| D11 | [#182](https://github.com/DysTelefonica/access2web-blueprint/issues/182) | docs(exp): document CAP-011 | `independent` · D0 | +70/-10 |
| D12 | [#183](https://github.com/DysTelefonica/access2web-blueprint/issues/183) | docs(exp): document CAP-012 | `independent` · D0 | +70/-10 |
| D13 | [#184](https://github.com/DysTelefonica/access2web-blueprint/issues/184) | docs(exp): document CAP-013 | `independent` · D0 | +70/-10 |
| D14 | [#185](https://github.com/DysTelefonica/access2web-blueprint/issues/185) | docs(exp): document CAP-014 | `independent` · D0 | +70/-10 |
| D15 | [#186](https://github.com/DysTelefonica/access2web-blueprint/issues/186) | docs(exp): document CAP-015 | `independent` · D0 | +70/-10 |
| D16 | [#187](https://github.com/DysTelefonica/access2web-blueprint/issues/187) | docs(exp): document CAP-016 | `independent` · D0 | +70/-10 |
| D17 | [#188](https://github.com/DysTelefonica/access2web-blueprint/issues/188) | docs(exp): document CAP-017 | `independent` · D0 | +70/-10 |
| D18 | [#189](https://github.com/DysTelefonica/access2web-blueprint/issues/189) | docs(exp): document CAP-018 | `independent` · D0 | +70/-10 |
| D19 | [#190](https://github.com/DysTelefonica/access2web-blueprint/issues/190) | docs(exp): document CAP-019 | `independent` · D0 | +70/-10 |
| D20 | [#191](https://github.com/DysTelefonica/access2web-blueprint/issues/191) | docs(exp): document CAP-020 | `independent` · D0 | +70/-10 |
| D21 | [#192](https://github.com/DysTelefonica/access2web-blueprint/issues/192) | docs(exp): document CAP-021 | `independent` · D0 | +70/-10 |
| D22 | [#193](https://github.com/DysTelefonica/access2web-blueprint/issues/193) | docs(exp): document CAP-022 | `independent` · D0 | +70/-10 |
| D23 | [#194](https://github.com/DysTelefonica/access2web-blueprint/issues/194) | docs(exp): document CAP-023 | `independent` · D0 | +70/-10 |
| D24 | [#195](https://github.com/DysTelefonica/access2web-blueprint/issues/195) | docs(exp): document CAP-024 | `independent` · D0 | +70/-10 |
| D25 | [#196](https://github.com/DysTelefonica/access2web-blueprint/issues/196) | docs(exp): document CAP-025 | `independent` · D0 | +70/-10 |
| D26 | [#197](https://github.com/DysTelefonica/access2web-blueprint/issues/197) | docs(exp): document CAP-026 | `independent` · D0 | +70/-10 |
| D27 | [#198](https://github.com/DysTelefonica/access2web-blueprint/issues/198) | docs(exp): document CAP-027 | `independent` · D0 | +70/-10 |
| D28 | [#199](https://github.com/DysTelefonica/access2web-blueprint/issues/199) | docs(exp): document CAP-028 | `independent` · D0 | +70/-10 |
| D29 | [#200](https://github.com/DysTelefonica/access2web-blueprint/issues/200) | docs(exp): document CAP-029 | `independent` · D0 | +70/-10 |
| D30 | [#201](https://github.com/DysTelefonica/access2web-blueprint/issues/201) | docs(exp): document CAP-030 | `independent` · D0 | +70/-10 |
| D31 | [#202](https://github.com/DysTelefonica/access2web-blueprint/issues/202) | docs(exp): document CAP-031 | `independent` · D0 | +70/-10 |
| D32 | [#203](https://github.com/DysTelefonica/access2web-blueprint/issues/203) | docs(exp): document CAP-032 | `independent` · D0 | +70/-10 |
| D33 | [#204](https://github.com/DysTelefonica/access2web-blueprint/issues/204) | docs(exp): document CAP-033 | `independent` · D0 | +70/-10 |
| D34 | [#205](https://github.com/DysTelefonica/access2web-blueprint/issues/205) | docs(exp): document CAP-034 | `independent` · D0 | +70/-10 |
| D35 | [#206](https://github.com/DysTelefonica/access2web-blueprint/issues/206) | docs(exp): document CAP-035 | `independent` · D0 | +70/-10 |
| D36 | [#207](https://github.com/DysTelefonica/access2web-blueprint/issues/207) | docs(exp): document CAP-036 | `independent` · D0 | +70/-10 |
| D37 | [#208](https://github.com/DysTelefonica/access2web-blueprint/issues/208) | docs(exp): document CAP-037 | `independent` · D0 | +70/-10 |
| D38 | [#209](https://github.com/DysTelefonica/access2web-blueprint/issues/209) | docs(exp): document CAP-038 | `independent` · D0 | +70/-10 |
| D39 | [#210](https://github.com/DysTelefonica/access2web-blueprint/issues/210) | docs(exp): document E2E export capabilities | `independent` · D0 | +110/-15 |
| D51 | [#267](https://github.com/DysTelefonica/access2web-blueprint/issues/267) | docs(exp): document access-control capabilities | `independent` · D0 | +110/-15 |
| D52 | [#268](https://github.com/DysTelefonica/access2web-blueprint/issues/268) | docs(exp): document runtime capabilities | `independent` · D0 | +110/-15 |
| D53 | [#269](https://github.com/DysTelefonica/access2web-blueprint/issues/269) | docs(exp): document integration capabilities | `independent` · D0 | +110/-15 |
| D54 | [#270](https://github.com/DysTelefonica/access2web-blueprint/issues/270) | docs(exp): document legacy-retirement capabilities | `independent` · D0 | +110/-15 |
| D40 | [#211](https://github.com/DysTelefonica/access2web-blueprint/issues/211) | docs(exp): dictionary tranche 01 | `independent` · #44,#45 | +120/-15 |
| D41 | [#212](https://github.com/DysTelefonica/access2web-blueprint/issues/212) | docs(exp): dictionary tranche 02 | `stacked` · D40 | +120/-15 |
| D42 | [#213](https://github.com/DysTelefonica/access2web-blueprint/issues/213) | docs(exp): dictionary tranche 03 | `stacked` · D41 | +120/-15 |
| D43 | [#214](https://github.com/DysTelefonica/access2web-blueprint/issues/214) | docs(exp): dictionary tranche 04 | `stacked` · D42 | +120/-15 |
| D44 | [#215](https://github.com/DysTelefonica/access2web-blueprint/issues/215) | docs(exp): dictionary tranche 05 | `stacked` · D43 | +120/-15 |
| D45 | [#216](https://github.com/DysTelefonica/access2web-blueprint/issues/216) | docs(exp): mapping tranche 01 | `stacked` · D40 | +160/-20 |
| D46 | [#217](https://github.com/DysTelefonica/access2web-blueprint/issues/217) | docs(exp): mapping tranche 02 | `stacked` · D41,D45 | +160/-20 |
| D47 | [#218](https://github.com/DysTelefonica/access2web-blueprint/issues/218) | docs(exp): mapping tranche 03 | `stacked` · D44,D46 | +180/-20 |
| D48 | [#219](https://github.com/DysTelefonica/access2web-blueprint/issues/219) | docs(exp): readiness contracts | `independent` · D39,D51–D54 | +180/-25 |
| D49 | [#220](https://github.com/DysTelefonica/access2web-blueprint/issues/220) | docs(exp): reconciliation runbook | `stacked` · D47 | +150/-20 |
| D50 | [#221](https://github.com/DysTelefonica/access2web-blueprint/issues/221) | docs(exp): cutover pack | `stacked` · D48,D49,#54,#55,#107 | +150/-20 |
| F01 | [#222](https://github.com/DysTelefonica/access2web-blueprint/issues/222) | feat(exp): ports | `independent` · D48,#21 | +120/-30 |
| F02 | [#223](https://github.com/DysTelefonica/access2web-blueprint/issues/223) | feat(exp): lifecycle domain | `stacked` · F01,D01 | +150/-40 |
| F03 | [#224](https://github.com/DysTelefonica/access2web-blueprint/issues/224) | feat(exp): DI readiness | `independent` · F01,#31,#32 | +120/-30 |
| F04 | [#225](https://github.com/DysTelefonica/access2web-blueprint/issues/225) | feat(exp): schema UoW | `stacked` · F01,D44 | +160/-45 |
| C01 | [#226](https://github.com/DysTelefonica/access2web-blueprint/issues/226) | feat(exp): create expediente | `stacked` · F02,F04 | +130/-35 |
| C02 | [#227](https://github.com/DysTelefonica/access2web-blueprint/issues/227) | feat(exp): edit expediente | `stacked` · C01 | +130/-35 |
| C03 | [#228](https://github.com/DysTelefonica/access2web-blueprint/issues/228) | feat(exp): conditional delete | `stacked` · C01 | +120/-30 |
| C04 | [#229](https://github.com/DysTelefonica/access2web-blueprint/issues/229) | feat(exp): lifecycle transitions | `stacked` · C01,C02 | +180/-45 |
| R01 | [#230](https://github.com/DysTelefonica/access2web-blueprint/issues/230) | feat(exp): milestone vertical | `stacked` · F02,F04,D08 | +120/-30 |
| R02 | [#271](https://github.com/DysTelefonica/access2web-blueprint/issues/271) | feat(exp): modification-history vertical | `stacked` · F02,F04,D09 | +120/-30 |
| R03 | [#272](https://github.com/DysTelefonica/access2web-blueprint/issues/272) | feat(exp): annex vertical | `stacked` · F02,F04,D10 | +120/-30 |
| R04 | [#273](https://github.com/DysTelefonica/access2web-blueprint/issues/273) | feat(exp): annuity vertical | `stacked` · F02,F04,D11 | +120/-30 |
| R05 | [#274](https://github.com/DysTelefonica/access2web-blueprint/issues/274) | feat(exp): responsibility vertical | `stacked` · F02,F04,D12 | +120/-30 |
| R06 | [#275](https://github.com/DysTelefonica/access2web-blueprint/issues/275) | feat(exp): entity-relations vertical | `stacked` · F02,F04,D13 | +120/-30 |
| R07 | [#276](https://github.com/DysTelefonica/access2web-blueprint/issues/276) | feat(exp): supplier-tree vertical | `stacked` · F02,F04,D14 | +120/-30 |
| Q01 | [#231](https://github.com/DysTelefonica/access2web-blueprint/issues/231) | feat(exp): catalog group A | `stacked` · F01,F04,D15–D19 | +170/-40 |
| Q02 | [#232](https://github.com/DysTelefonica/access2web-blueprint/issues/232) | feat(exp): catalog group B | `independent` · Q01,D20–D24 | +170/-40 |
| Q03 | [#233](https://github.com/DysTelefonica/access2web-blueprint/issues/233) | feat(exp): query services | `independent` · F04,Q01 | +180/-45 |
| W01 | [#234](https://github.com/DysTelefonica/access2web-blueprint/issues/234) | feat(exp): autosave core | `stacked` · C01 | +140/-35 |
| W02 | [#235](https://github.com/DysTelefonica/access2web-blueprint/issues/235) | feat(exp): autosave related | `stacked` · R01–R07,W01 | +140/-35 |
| W03 | [#236](https://github.com/DysTelefonica/access2web-blueprint/issues/236) | feat(exp): resilient feedback | `stacked` · W01,W02 | +170/-40 |
| E01 | [#237](https://github.com/DysTelefonica/access2web-blueprint/issues/237) | feat(exp): E2E DTO | `independent` · F01 | +120/-30 |
| E02 | [#238](https://github.com/DysTelefonica/access2web-blueprint/issues/238) | feat(exp): canonical JSON | `stacked` · E01 | +150/-35 |
| E03 | [#239](https://github.com/DysTelefonica/access2web-blueprint/issues/239) | feat(exp): E2E hash | `stacked` · E02 | +120/-30 |
| E04 | [#240](https://github.com/DysTelefonica/access2web-blueprint/issues/240) | feat(exp): E2E ordinal families | `stacked` · E02 | +150/-35 |
| E05 | [#241](https://github.com/DysTelefonica/access2web-blueprint/issues/241) | feat(exp): E2E batch | `stacked` · E02,E03 | +160/-40 |
| E06 | [#242](https://github.com/DysTelefonica/access2web-blueprint/issues/242) | feat(exp): export package sink | `stacked` · E05 | +150/-35 |
| E07 | [#243](https://github.com/DysTelefonica/access2web-blueprint/issues/243) | feat(exp): export sessions | `stacked` · E05,E06 | +170/-40 |
| A01 | [#244](https://github.com/DysTelefonica/access2web-blueprint/issues/244) | feat(exp): authorization policy | `independent` · F01,F03 | +140/-35 |
| A02 | [#245](https://github.com/DysTelefonica/access2web-blueprint/issues/245) | feat(exp): CurrentPrincipal adapter | `stacked` · A01,#31 | +120/-30 |
| A03 | [#246](https://github.com/DysTelefonica/access2web-blueprint/issues/246) | feat(exp): session audit | `stacked` · A02,W03 | +170/-40 |
| T01 | [#247](https://github.com/DysTelefonica/access2web-blueprint/issues/247) | feat(exp): runtime policies | `stacked` · F03,A02 | +170/-40 |
| H01 | [#248](https://github.com/DysTelefonica/access2web-blueprint/issues/248) | feat(exp): HPS adapter | `independent` · F01,A02,#49 | +140/-35 |
| H02 | [#249](https://github.com/DysTelefonica/access2web-blueprint/issues/249) | feat(exp): AGEDYS adapter | `stacked` · H01 | +140/-35 |
| H03 | [#250](https://github.com/DysTelefonica/access2web-blueprint/issues/250) | feat(exp): Risks adapter | `independent` · F01 | +120/-30 |
| H04 | [#251](https://github.com/DysTelefonica/access2web-blueprint/issues/251) | feat(exp): NC adapter | `independent` · F01 | +120/-30 |
| H05 | [#252](https://github.com/DysTelefonica/access2web-blueprint/issues/252) | feat(exp): notification adapter | `independent` · F01 | +120/-30 |
| H06 | [#253](https://github.com/DysTelefonica/access2web-blueprint/issues/253) | feat(exp): document storage adapter | `independent` · D10,A01,#50,#51 | +170/-40 |
| Z01 | [#254](https://github.com/DysTelefonica/access2web-blueprint/issues/254) | feat(exp): REST E2E | `stacked` · E01–E07,A01 | +150/-35 |
| Z02 | [#255](https://github.com/DysTelefonica/access2web-blueprint/issues/255) | test(exp): lifecycle journeys | `independent` · C01–Q03 | +180/-45 |
| Z03 | [#256](https://github.com/DysTelefonica/access2web-blueprint/issues/256) | test(exp): resilience journeys | `independent` · W01–H06,Z01 | +190/-45 |
| M01 | [#257](https://github.com/DysTelefonica/access2web-blueprint/issues/257) | feat(exp): extractor tranche 01 | `independent` · F04,D40,D45 | +180/-45 |
| M02 | [#258](https://github.com/DysTelefonica/access2web-blueprint/issues/258) | feat(exp): extractor tranche 02 | `stacked` · M01,D41,D46 | +180/-45 |
| M03 | [#259](https://github.com/DysTelefonica/access2web-blueprint/issues/259) | feat(exp): extractor tranche 03 | `stacked` · M02,D44,D47 | +180/-45 |
| M04 | [#260](https://github.com/DysTelefonica/access2web-blueprint/issues/260) | feat(exp): promotion reconciliation | `stacked` · M01–M03,D49 | +190/-45 |
| U01 | [#261](https://github.com/DysTelefonica/access2web-blueprint/issues/261) | test(exp): lifecycle UAT | `independent` · C01–C04,Z02 | +150/-35 |
| U02 | [#262](https://github.com/DysTelefonica/access2web-blueprint/issues/262) | test(exp): related-data UAT | `independent` · R01–R07,W02,Z02 | +150/-35 |
| U03 | [#263](https://github.com/DysTelefonica/access2web-blueprint/issues/263) | test(exp): catalogs UAT | `independent` · Q01,Q02,Z02 | +140/-30 |
| U04 | [#277](https://github.com/DysTelefonica/access2web-blueprint/issues/277) | test(exp): query-and-tasks UAT | `independent` · Q03,Z02 | +130/-30 |
| U05 | [#278](https://github.com/DysTelefonica/access2web-blueprint/issues/278) | test(exp): write-resilience UAT | `independent` · W01–W03,Z03 | +130/-30 |
| U06 | [#279](https://github.com/DysTelefonica/access2web-blueprint/issues/279) | test(exp): E2E export UAT | `independent` · E01–E07,Z01,Z03 | +160/-35 |
| U07 | [#280](https://github.com/DysTelefonica/access2web-blueprint/issues/280) | test(exp): access-control UAT | `independent` · A01–A03,Z02 | +120/-30 |
| U08 | [#281](https://github.com/DysTelefonica/access2web-blueprint/issues/281) | test(exp): runtime UAT | `independent` · F03,T01,Z03 | +110/-25 |
| U09 | [#282](https://github.com/DysTelefonica/access2web-blueprint/issues/282) | test(exp): integrations UAT | `independent` · H01–H06,Z03 | +160/-35 |
| U10 | [#283](https://github.com/DysTelefonica/access2web-blueprint/issues/283) | test(exp): retirement and cutover UAT | `stacked` · U01–U09,M01–M04,D50 | +150/-35 |
| L01 | [#264](https://github.com/DysTelefonica/access2web-blueprint/issues/264) | chore(exp): legacy retirement gates | `stacked` · U10,D50,#107 | +160/-40 |

## Criterio de uso

1. Seleccionar una unidad cuyas dependencias estén cerradas o tengan contrato aceptado.
2. Crear la rama desde la base exigida por su ruta.
3. Implementar únicamente el outcome y los criterios de `tasks.md`.
4. Verificar que `additions + deletions ≤ 400`; si deja de caber, dividir antes de abrir el PR.
5. Mantener la issue padre #170 abierta hasta cerrar U10 y L01.

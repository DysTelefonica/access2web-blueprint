# Tasks: Migración web de Expedientes

## Review Workload Forecast

| Campo | Valor |
|---|---|
| Work units | 110 unidades con issue: #172–#264 y #267–#283, más la issue padre #170 |
| Capability coverage | 63 capacidades (EXP-CAP-001..063) |
| Estimated changed lines | 15.595 (suma de adiciones+eliminaciones estimadas por unidad) |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | D01–D54 documentación; F01–T01 foundation y producto; H01–Z03 integraciones/E2E; M01–M04 migración; U01–U10 UAT; L01 retirada |
| Delivery strategy | auto-chain |
| Chain strategy | stacked-to-main (feature-branch-chain solo si una unidad funcional no puede aterrizar independiente) |

Decision needed before apply: No
Chained PRs recommended: Yes
Chain strategy: stacked-to-main
400-line budget risk: High

Cada PR queda en ≤400 líneas (adiciones+eliminaciones); no se autoriza `size:exception`. Los PR documentales son independientes salvo que una matriz consuma el documento anterior. Las unidades de producto incluyen sus tests.

## Decisiones bloqueantes (separadas de implementación)

- Runtime de las 49 tablas: diccionario, ownership, volumen y uso deben tener evidencia runtime antes de DDL final.
- CPV: dígitos/normalización y compatibilidad de códigos.
- Anexos: límite (la spec deja 25 MB como pregunta), retención y proveedor.
- FNV-1a: golden de continuidad antes de aprobar el algoritmo.
- AGEDYS: ownership, errores, reintentos e idempotencia.
- Batch E2E: atomicidad, ciclo, reanudación y semántica de fallo.

## Diagrama de carriles

```text
D01…D39,D51…D54 ─┬─→ D40→D41→D42→D43→D44 ─→ F04 ─→ M01→M02→M03→M04
                  └─→ D48 ─→ F01 ─┬─→ F02 ─→ C01→C02→C04 ─→ Z02
                                   ├─→ R01…R07; Q01→Q02→Q03
                                   ├─→ E01→E02→E03/E04→E05→E06→E07→Z01
                                   ├─→ A01→A02→A03; T01
                                   └─→ H01…H06 ───────────────→ Z03
D45→D46→D47→D49; D48+D49→D50; dominios→U01…U09→U10→L01
```

## Work units / PRs

Todas las unidades siguientes son tickets independientes o encadenados solo por dependencia real. Cada estimación es adiciones+eliminaciones y queda por debajo de 400 líneas. `VERIF-DOC` valida estructura, trazabilidad y ausencia de PII; las unidades de código siguen RED→GREEN→REFACTOR.

| ID | Outcome / rutas / caps-spec / deps / estimación | Issue / ruta | Acceptance / rollback |
|---|---|---|---|
| D01 | Capability doc EXP-001 en `docs/05-capacidades/expedientes/EXP-CAP-001.md`; lifecycle; +70/-10; D0 ledger | [#172](https://github.com/DysTelefonica/access2web-blueprint/issues/172) · docs(exp): document CAP-001 / independent / documentation | §2/3/6/8 + UAT; revertir archivo. |
| D02 | EXP-002; +70/-10; D0 | [#173](https://github.com/DysTelefonica/access2web-blueprint/issues/173) · docs(exp): document CAP-002 / independent / documentation | invariantes/concurrencia trazables; revertir archivo. |
| D03 | EXP-003; +70/-10; D0 | [#174](https://github.com/DysTelefonica/access2web-blueprint/issues/174) · docs(exp): document CAP-003 / independent / documentation | baja/relacionados/rollback; revertir archivo. |
| D04 | EXP-004; +70/-10; D0 | [#175](https://github.com/DysTelefonica/access2web-blueprint/issues/175) · docs(exp): document CAP-004 / independent / documentation | elegibilidad y efectos; revertir archivo. |
| D05 | EXP-005; +70/-10; D0 | [#176](https://github.com/DysTelefonica/access2web-blueprint/issues/176) · docs(exp): document CAP-005 / independent / documentation | transiciones/motivos; revertir archivo. |
| D06 | EXP-006; +70/-10; D0 | [#177](https://github.com/DysTelefonica/access2web-blueprint/issues/177) · docs(exp): document CAP-006 / independent / documentation | jerarquía padre-hijo; revertir archivo. |
| D07 | EXP-007; +70/-10; D0 | [#178](https://github.com/DysTelefonica/access2web-blueprint/issues/178) · docs(exp): document CAP-007 / independent / documentation | ordinal; revertir archivo. |
| D08 | EXP-008; +70/-10; D0 | [#179](https://github.com/DysTelefonica/access2web-blueprint/issues/179) · docs(exp): document CAP-008 / independent / documentation | hitos temporales; revertir archivo. |
| D09 | EXP-009; +70/-10; D0 | [#180](https://github.com/DysTelefonica/access2web-blueprint/issues/180) · docs(exp): document CAP-009 / independent / documentation | historial/modificados; revertir archivo. |
| D10 | EXP-010; +70/-10; D0 | [#181](https://github.com/DysTelefonica/access2web-blueprint/issues/181) · docs(exp): document CAP-010 / independent / documentation | anexos/retención y decisión 25 MB; revertir archivo. |
| D11 | EXP-011; +70/-10; D0 | [#182](https://github.com/DysTelefonica/access2web-blueprint/issues/182) · docs(exp): document CAP-011 / independent / documentation | anualidades; revertir archivo. |
| D12 | EXP-012; +70/-10; D0 | [#183](https://github.com/DysTelefonica/access2web-blueprint/issues/183) · docs(exp): document CAP-012 / independent / documentation | responsables; revertir archivo. |
| D13 | EXP-013; +70/-10; D0 | [#184](https://github.com/DysTelefonica/access2web-blueprint/issues/184) · docs(exp): document CAP-013 / independent / documentation | entidades/jurídicas; revertir archivo. |
| D14 | EXP-014; +70/-10; D0 | [#185](https://github.com/DysTelefonica/access2web-blueprint/issues/185) · docs(exp): document CAP-014 / independent / documentation | UTE/árbol; revertir archivo. |
| D15 | EXP-015; +70/-10; D0 | [#186](https://github.com/DysTelefonica/access2web-blueprint/issues/186) · docs(exp): document CAP-015 / independent / documentation | comerciales; revertir archivo. |
| D16 | EXP-016; +70/-10; D0 | [#187](https://github.com/DysTelefonica/access2web-blueprint/issues/187) · docs(exp): document CAP-016 / independent / documentation | CPV y decisión dígitos; revertir archivo. |
| D17 | EXP-017; +70/-10; D0 | [#188](https://github.com/DysTelefonica/access2web-blueprint/issues/188) · docs(exp): document CAP-017 / independent / documentation | ejércitos; revertir archivo. |
| D18 | EXP-018; +70/-10; D0 | [#189](https://github.com/DysTelefonica/access2web-blueprint/issues/189) · docs(exp): document CAP-018 / independent / documentation | suministradores; revertir archivo. |
| D19 | EXP-019; +70/-10; D0 | [#190](https://github.com/DysTelefonica/access2web-blueprint/issues/190) · docs(exp): document CAP-019 / independent / documentation | lugares; revertir archivo. |
| D20 | EXP-020; +70/-10; D0 | [#191](https://github.com/DysTelefonica/access2web-blueprint/issues/191) · docs(exp): document CAP-020 / independent / documentation | PECAL; revertir archivo. |
| D21 | EXP-021; +70/-10; D0 | [#192](https://github.com/DysTelefonica/access2web-blueprint/issues/192) · docs(exp): document CAP-021 / independent / documentation | RAC; revertir archivo. |
| D22 | EXP-022; +70/-10; D0 | [#193](https://github.com/DysTelefonica/access2web-blueprint/issues/193) · docs(exp): document CAP-022 / independent / documentation | grados; revertir archivo. |
| D23 | EXP-023; +70/-10; D0 | [#194](https://github.com/DysTelefonica/access2web-blueprint/issues/194) · docs(exp): document CAP-023 / independent / documentation | órganos; revertir archivo. |
| D24 | EXP-024; +70/-10; D0 | [#195](https://github.com/DysTelefonica/access2web-blueprint/issues/195) · docs(exp): document CAP-024 / independent / documentation | oficinas; revertir archivo. |
| D25 | EXP-025; +70/-10; D0 | [#196](https://github.com/DysTelefonica/access2web-blueprint/issues/196) · docs(exp): document CAP-025 / independent / documentation | bandeja; revertir archivo. |
| D26 | EXP-026; +70/-10; D0 | [#197](https://github.com/DysTelefonica/access2web-blueprint/issues/197) · docs(exp): document CAP-026 / independent / documentation | búsqueda; revertir archivo. |
| D27 | EXP-027; +70/-10; D0 | [#198](https://github.com/DysTelefonica/access2web-blueprint/issues/198) · docs(exp): document CAP-027 / independent / documentation | búsqueda técnica; revertir archivo. |
| D28 | EXP-028; +70/-10; D0 | [#199](https://github.com/DysTelefonica/access2web-blueprint/issues/199) · docs(exp): document CAP-028 / independent / documentation | export; revertir archivo. |
| D29 | EXP-029; +70/-10; D0 | [#200](https://github.com/DysTelefonica/access2web-blueprint/issues/200) · docs(exp): document CAP-029 / independent / documentation | tareas calculadas; revertir archivo. |
| D30 | EXP-030; +70/-10; D0 | [#201](https://github.com/DysTelefonica/access2web-blueprint/issues/201) · docs(exp): document CAP-030 / independent / documentation | autosave general; revertir archivo. |
| D31 | EXP-031; +70/-10; D0 | [#202](https://github.com/DysTelefonica/access2web-blueprint/issues/202) · docs(exp): document CAP-031 / independent / documentation | autosave relacionados; revertir archivo. |
| D32 | EXP-032; +70/-10; D0 | [#203](https://github.com/DysTelefonica/access2web-blueprint/issues/203) · docs(exp): document CAP-032 / independent / documentation | idempotencia/feedback; revertir archivo. |
| D33 | EXP-033; +70/-10; D0 | [#204](https://github.com/DysTelefonica/access2web-blueprint/issues/204) · docs(exp): document CAP-033 / independent / documentation | DTO; revertir archivo. |
| D34 | EXP-034; +70/-10; D0 | [#205](https://github.com/DysTelefonica/access2web-blueprint/issues/205) · docs(exp): document CAP-034 / independent / documentation | JSON; revertir archivo. |
| D35 | EXP-035; +70/-10; D0 | [#206](https://github.com/DysTelefonica/access2web-blueprint/issues/206) · docs(exp): document CAP-035 / independent / documentation | batch; revertir archivo. |
| D36 | EXP-036; +70/-10; D0 | [#207](https://github.com/DysTelefonica/access2web-blueprint/issues/207) · docs(exp): document CAP-036 / independent / documentation | hash/FNV gate; revertir archivo. |
| D37 | EXP-037; +70/-10; D0 | [#208](https://github.com/DysTelefonica/access2web-blueprint/issues/208) · docs(exp): document CAP-037 / independent / documentation | paquete; revertir archivo. |
| D38 | EXP-038; +70/-10; D0 | [#209](https://github.com/DysTelefonica/access2web-blueprint/issues/209) · docs(exp): document CAP-038 / independent / documentation | trazabilidad; revertir archivo. |
| D39 | EXP-039..042 en `docs/05-capacidades/expedientes/e2e-export.md`; +110/-15; D0 | [#210](https://github.com/DysTelefonica/access2web-blueprint/issues/210) · docs(exp): document E2E export capabilities / independent / documentation | selección, destino, sesión y ordinal E2E; revertir archivo. |
| D51 | EXP-043..046 en `docs/05-capacidades/expedientes/access-control.md`; +110/-15; D0 | [#267](https://github.com/DysTelefonica/access2web-blueprint/issues/267) · docs(exp): document access-control capabilities / independent / documentation | deny-by-default, principal, sesión y auditoría; revertir archivo. |
| D52 | EXP-047..050 en `docs/05-capacidades/expedientes/runtime.md`; +110/-15; D0 | [#268](https://github.com/DysTelefonica/access2web-blueprint/issues/268) · docs(exp): document runtime capabilities / independent / documentation | readiness, configuración, caché y binding; revertir archivo. |
| D53 | EXP-051..056 en `docs/05-capacidades/expedientes/integrations.md`; +110/-15; D0 | [#269](https://github.com/DysTelefonica/access2web-blueprint/issues/269) · docs(exp): document integration capabilities / independent / documentation | HPS, AGEDYS, Riesgos, NC, correo y documentos; revertir archivo. |
| D54 | EXP-057..063 en `docs/05-capacidades/expedientes/legacy-retirement.md`; +110/-15; D0 | [#270](https://github.com/DysTelefonica/access2web-blueprint/issues/270) · docs(exp): document legacy-retirement capabilities / independent / documentation | sustituciones y puertas de retirada; revertir archivo. |
| D40 | Diccionario tablas 01–10 `migration/source-dictionary-01.md`; EXP-047..050; +120/-15; #44,#45 | [#211](https://github.com/DysTelefonica/access2web-blueprint/issues/211) · docs(exp): dictionary tranche 01 / independent / documentation | tablas/campos/ownership/volumen; revertir archivo. |
| D41 | Diccionario 11–20; +120/-15; D40 | [#212](https://github.com/DysTelefonica/access2web-blueprint/issues/212) · docs(exp): dictionary tranche 02 / stacked / documentation | evidencia runtime; revertir archivo. |
| D42 | Diccionario 21–30; +120/-15; D41 | [#213](https://github.com/DysTelefonica/access2web-blueprint/issues/213) · docs(exp): dictionary tranche 03 / stacked / documentation | evidencia runtime; revertir archivo. |
| D43 | Diccionario 31–40; +120/-15; D42 | [#214](https://github.com/DysTelefonica/access2web-blueprint/issues/214) · docs(exp): dictionary tranche 04 / stacked / documentation | evidencia runtime; revertir archivo. |
| D44 | Diccionario 41–49; +120/-15; D43 | [#215](https://github.com/DysTelefonica/access2web-blueprint/issues/215) · docs(exp): dictionary tranche 05 / stacked / documentation | 49/49 cerradas o bloqueadas explícitamente; revertir archivo. |
| D45 | Field mapping 01–15; +160/-20; D40 | [#216](https://github.com/DysTelefonica/access2web-blueprint/issues/216) · docs(exp): mapping tranche 01 / stacked / documentation | mapping campo-a-campo; revertir archivo. |
| D46 | Field mapping 16–30; +160/-20; D41,D45 | [#217](https://github.com/DysTelefonica/access2web-blueprint/issues/217) · docs(exp): mapping tranche 02 / stacked / documentation | transformaciones/claves; revertir archivo. |
| D47 | Field mapping 31–49 + rechazos; +180/-20; D44,D46 | [#218](https://github.com/DysTelefonica/access2web-blueprint/issues/218) · docs(exp): mapping tranche 03 / stacked / documentation | cuarentena y motivos; revertir archivo. |
| D48 | Authz/integraciones/UI workflows; +180/-25; D39,D51–D54 | [#219](https://github.com/DysTelefonica/access2web-blueprint/issues/219) · docs(exp): readiness contracts / independent / documentation | perfiles, adapters y flujos; revertir archivo. |
| D49 | Reconciliation/rejects/rollback; +150/-20; D47 | [#220](https://github.com/DysTelefonica/access2web-blueprint/issues/220) · docs(exp): reconciliation runbook / stacked / documentation | conteos/muestras/hash/rollback; revertir archivo. |
| D50 | UAT/cutover/retirement; +150/-20; D48,D49,#54,#55,#107 | [#221](https://github.com/DysTelefonica/access2web-blueprint/issues/221) · docs(exp): cutover pack / stacked / documentation | go/no-go y CAP-057..063; revertir archivo. |
| F01 | Puertos y tipos `modules/expedientes/ports`; +120/-30; D48,#21 | [#222](https://github.com/DysTelefonica/access2web-blueprint/issues/222) · feat(exp): ports / independent / feat | contratos y fakes deterministas; revertir ports. |
| F02 | Dominio expediente/lifecycle; +150/-40; F01,D01 | [#223](https://github.com/DysTelefonica/access2web-blueprint/issues/223) · feat(exp): lifecycle domain / stacked / feat | RED invariants→GREEN; revertir dominio. |
| F03 | DI/config/readiness; +120/-30; F01,#31,#32 | [#224](https://github.com/DysTelefonica/access2web-blueprint/issues/224) · feat(exp): DI readiness / independent / feat | deny/no fake prod; revertir DI. |
| F04 | Schema/migración base/UoW; +160/-45; F01,D44 | [#225](https://github.com/DysTelefonica/access2web-blueprint/issues/225) · feat(exp): schema UoW / stacked / feat | reversible expand; revertir migration. |
| C01 | CAP-001 alta; +130/-35; F02,F04 | [#226](https://github.com/DysTelefonica/access2web-blueprint/issues/226) · feat(exp): create expediente / stacked / feat | UAT alta transaccional; revertir slice. |
| C02 | CAP-002 edición/concurrencia; +130/-35; C01 | [#227](https://github.com/DysTelefonica/access2web-blueprint/issues/227) · feat(exp): edit expediente / stacked / feat | 409 sin commit parcial; revertir slice. |
| C03 | CAP-003 baja; +120/-30; C01 | [#228](https://github.com/DysTelefonica/access2web-blueprint/issues/228) · feat(exp): conditional delete / stacked / feat | auditoría/relacionados; revertir slice. |
| C04 | CAP-004..007 tipo/estado/jerarquía/ordinal; +180/-45; C01,C02 | [#229](https://github.com/DysTelefonica/access2web-blueprint/issues/229) · feat(exp): lifecycle transitions / stacked / feat | reglas y ordinal; revertir slice. |
| R01 | CAP-008 hitos; +120/-30; F02,F04,D08 | [#230](https://github.com/DysTelefonica/access2web-blueprint/issues/230) · feat(exp): milestone vertical / stacked / feat | reglas temporales; revertir vertical. |
| R02 | CAP-009 modificados e historial; +120/-30; F02,F04,D09 | [#271](https://github.com/DysTelefonica/access2web-blueprint/issues/271) · feat(exp): modification-history vertical / stacked / feat | cambios e historial reciente; revertir vertical. |
| R03 | CAP-010 anexos; +120/-30; F02,F04,D10 | [#272](https://github.com/DysTelefonica/access2web-blueprint/issues/272) · feat(exp): annex vertical / stacked / feat | eliminación, retención y autorización; revertir vertical. |
| R04 | CAP-011 anualidades; +120/-30; F02,F04,D11 | [#273](https://github.com/DysTelefonica/access2web-blueprint/issues/273) · feat(exp): annuity vertical / stacked / feat | anualidades vinculadas; revertir vertical. |
| R05 | CAP-012 responsables y jefaturas; +120/-30; F02,F04,D12 | [#274](https://github.com/DysTelefonica/access2web-blueprint/issues/274) · feat(exp): responsibility vertical / stacked / feat | roles y jefaturas; revertir vertical. |
| R06 | CAP-013 entidades y jurídicas; +120/-30; F02,F04,D13 | [#275](https://github.com/DysTelefonica/access2web-blueprint/issues/275) · feat(exp): entity-relations vertical / stacked / feat | cadena de entidades y jurídicas; revertir vertical. |
| R07 | CAP-014 suministradores y UTE; +120/-30; F02,F04,D14 | [#276](https://github.com/DysTelefonica/access2web-blueprint/issues/276) · feat(exp): supplier-tree vertical / stacked / feat | contratistas, UTE y jerarquía; revertir vertical. |
| Q01 | Catálogos CAP-015..019; +170/-40; F01,F04,D15–D19 | [#231](https://github.com/DysTelefonica/access2web-blueprint/issues/231) · feat(exp): catalog group A / stacked / feat | filtros/permiso y gate de decisión CPV (CAP-016); revertir grupo. |
| Q02 | Catálogos CAP-020..024; +170/-40; Q01,D20–D24 | [#232](https://github.com/DysTelefonica/access2web-blueprint/issues/232) · feat(exp): catalog group B / independent / feat | filtros/vigencia/permisos; revertir grupo. |
| Q03 | Query services CAP-025..029; +180/-45; F04,Q01 | [#233](https://github.com/DysTelefonica/access2web-blueprint/issues/233) · feat(exp): query services / independent / feat | paginación/export/tareas; revertir read models. |
| W01 | CAP-030 autosave generales; +140/-35; C01 | [#234](https://github.com/DysTelefonica/access2web-blueprint/issues/234) · feat(exp): autosave core / stacked / feat | idempotencia básica; revertir endpoint. |
| W02 | CAP-031 relacionados; +140/-35; R01–R07,W01 | [#235](https://github.com/DysTelefonica/access2web-blueprint/issues/235) · feat(exp): autosave related / stacked / feat | concurrencia; revertir endpoint. |
| W03 | CAP-032 UoW/idempotencia/feedback; +170/-40; W01,W02 | [#236](https://github.com/DysTelefonica/access2web-blueprint/issues/236) · feat(exp): resilient feedback / stacked / feat | aria-busy/double submit; revertir delivery. |
| E01 | CAP-033 DTO canonical; +120/-30; F01 | [#237](https://github.com/DysTelefonica/access2web-blueprint/issues/237) · feat(exp): E2E DTO / independent / feat | schema contract; revertir serializer. |
| E02 | CAP-034 canonical JSON/order/null; +150/-35; E01 | [#238](https://github.com/DysTelefonica/access2web-blueprint/issues/238) · feat(exp): canonical JSON / stacked / feat | golden; revertir serializer. |
| E03 | CAP-036 FNV/hash; +120/-30; E02 | [#239](https://github.com/DysTelefonica/access2web-blueprint/issues/239) · feat(exp): E2E hash / stacked / feat | solo tras golden aprobado; revertir hash. |
| E04 | CAP-042 ordinal/family cycle expansion; +150/-35; E02 | [#240](https://github.com/DysTelefonica/access2web-blueprint/issues/240) · feat(exp): E2E ordinal families / stacked / feat | ciclos cortan seguro; revertir expansion. |
| E05 | CAP-035 batch atomicity; +160/-40; E02,E03 | [#241](https://github.com/DysTelefonica/access2web-blueprint/issues/241) · feat(exp): E2E batch / stacked / feat | decisión atomicidad; revertir batch. |
| E06 | CAP-037 paquete/sink; +150/-35; E05 | [#242](https://github.com/DysTelefonica/access2web-blueprint/issues/242) · feat(exp): export package sink / stacked / feat | destino opaco; revertir sink. |
| E07 | CAP-038..041 session/selection/destination; +170/-40; E05,E06 | [#243](https://github.com/DysTelefonica/access2web-blueprint/issues/243) · feat(exp): export sessions / stacked / feat | aislamiento/reanudación; revertir session. |
| A01 | CAP-043 authorization; +140/-35; F01,F03 | [#244](https://github.com/DysTelefonica/access2web-blueprint/issues/244) · feat(exp): authorization policy / independent / feat | deny-by-default; revertir policy. |
| A02 | CAP-044 principal Lanzadera; +120/-30; A01,#31 | [#245](https://github.com/DysTelefonica/access2web-blueprint/issues/245) · feat(exp): CurrentPrincipal adapter / stacked / feat | errores deniegan; revertir adapter. |
| A03 | CAP-045..046 session/audit; +170/-40; A02,W03 | [#246](https://github.com/DysTelefonica/access2web-blueprint/issues/246) · feat(exp): session audit / stacked / feat | auditoría completa; revertir audit. |
| T01 | CAP-047..050 runtime/cache/binding; +170/-40; F03,A02 | [#247](https://github.com/DysTelefonica/access2web-blueprint/issues/247) · feat(exp): runtime policies / stacked / feat | readiness/config/cache; revertir config. |
| H01 | CAP-051 HPS adapter; +140/-35; F01,A02,#49 | [#248](https://github.com/DysTelefonica/access2web-blueprint/issues/248) · feat(exp): HPS adapter / independent / feat | contract/idempotencia; revertir adapter. |
| H02 | CAP-052 AGEDYS adapter; +140/-35; H01 | [#249](https://github.com/DysTelefonica/access2web-blueprint/issues/249) · feat(exp): AGEDYS adapter / stacked / feat | ownership decision; revertir adapter. |
| H03 | CAP-053 Riesgos adapter; +120/-30; F01 | [#250](https://github.com/DysTelefonica/access2web-blueprint/issues/250) · feat(exp): Risks adapter / independent / feat | ID estable; revertir adapter. |
| H04 | CAP-054 NC adapter; +120/-30; F01 | [#251](https://github.com/DysTelefonica/access2web-blueprint/issues/251) · feat(exp): NC adapter / independent / feat | estado asociado; revertir adapter. |
| H05 | CAP-055 notification adapter; +120/-30; F01 | [#252](https://github.com/DysTelefonica/access2web-blueprint/issues/252) · feat(exp): notification adapter / independent / feat | puerto/trazabilidad; revertir adapter. |
| H06 | CAP-056 document storage; +170/-40; D10,A01,#50,#51 | [#253](https://github.com/DysTelefonica/access2web-blueprint/issues/253) · feat(exp): document storage adapter / independent / feat | retención/autz; annex gate; revertir adapter. |
| Z01 | REST E2E adapter; +150/-35; E01–E07,A01 | [#254](https://github.com/DysTelefonica/access2web-blueprint/issues/254) · feat(exp): REST E2E / stacked / feat | transporte separado; revertir routes. |
| Z02 | Playwright lifecycle/query UAT; +180/-45; C01–Q03 | [#255](https://github.com/DysTelefonica/access2web-blueprint/issues/255) · test(exp): lifecycle journeys / independent / test | perfiles/errores; revertir tests. |
| Z03 | Playwright writes/E2E/integrations; +190/-45; W01–H06,Z01 | [#256](https://github.com/DysTelefonica/access2web-blueprint/issues/256) · test(exp): resilience journeys / independent / test | concurrencia/dependencias; revertir tests. |
| M01 | Extractor staging grupos tablas 01–15; +180/-45; F04,D40,D45 | [#257](https://github.com/DysTelefonica/access2web-blueprint/issues/257) · feat(exp): extractor tranche 01 / independent / feat | watermark/rechazo; revertir extractor. |
| M02 | Extractor grupos 16–30; +180/-45; M01,D41,D46 | [#258](https://github.com/DysTelefonica/access2web-blueprint/issues/258) · feat(exp): extractor tranche 02 / stacked / feat | repeat-safe; revertir tranche. |
| M03 | Extractor grupos 31–49; +180/-45; M02,D44,D47 | [#259](https://github.com/DysTelefonica/access2web-blueprint/issues/259) · feat(exp): extractor tranche 03 / stacked / feat | 49 tablas; revertir tranche. |
| M04 | Promotion/reconciliation; +190/-45; M01–M03,D49 | [#260](https://github.com/DysTelefonica/access2web-blueprint/issues/260) · feat(exp): promotion reconciliation / stacked / feat | conteos/hash/muestras; revertir promotion. |
| U01 | UAT lifecycle CAP-001..007; +150/-35; C01–C04,Z02 | [#261](https://github.com/DysTelefonica/access2web-blueprint/issues/261) · test(exp): lifecycle UAT / independent / test | alta, edición, baja, estado y jerarquía; revertir tests. |
| U02 | UAT related-data CAP-008..014; +150/-35; R01–R07,W02,Z02 | [#262](https://github.com/DysTelefonica/access2web-blueprint/issues/262) · test(exp): related-data UAT / independent / test | siete verticales y errores; revertir tests. |
| U03 | UAT catalogs CAP-015..024; +140/-30; Q01,Q02,Z02 | [#263](https://github.com/DysTelefonica/access2web-blueprint/issues/263) · test(exp): catalogs UAT / independent / test | diez catálogos, vigencia, permisos y CPV; revertir tests. |
| U04 | UAT query-and-tasks CAP-025..029; +130/-30; Q03,Z02 | [#277](https://github.com/DysTelefonica/access2web-blueprint/issues/277) · test(exp): query-and-tasks UAT / independent / test | bandejas, filtros, exportación y tareas; revertir tests. |
| U05 | UAT write-resilience CAP-030..032; +130/-30; W01–W03,Z03 | [#278](https://github.com/DysTelefonica/access2web-blueprint/issues/278) · test(exp): write-resilience UAT / independent / test | concurrencia, idempotencia y feedback; revertir tests. |
| U06 | UAT E2E CAP-033..042; +160/-35; E01–E07,Z01,Z03 | [#279](https://github.com/DysTelefonica/access2web-blueprint/issues/279) · test(exp): E2E export UAT / independent / test | JSON, hash, batch, paquete y sesiones; revertir tests. |
| U07 | UAT access-control CAP-043..046; +120/-30; A01–A03,Z02 | [#280](https://github.com/DysTelefonica/access2web-blueprint/issues/280) · test(exp): access-control UAT / independent / test | deny-by-default, principal, sesión y auditoría; revertir tests. |
| U08 | UAT runtime CAP-047..050; +110/-25; F03,T01,Z03 | [#281](https://github.com/DysTelefonica/access2web-blueprint/issues/281) · test(exp): runtime UAT / independent / test | readiness, configuración, caché y binding; revertir tests. |
| U09 | UAT integrations CAP-051..056; +160/-35; H01–H06,Z03 | [#282](https://github.com/DysTelefonica/access2web-blueprint/issues/282) · test(exp): integrations UAT / independent / test | contratos, fallos, reintentos y trazabilidad; revertir tests. |
| U10 | UAT retirement/cutover CAP-057..063; +150/-35; U01–U09,M01–M04,D50 | [#283](https://github.com/DysTelefonica/access2web-blueprint/issues/283) · test(exp): retirement and cutover UAT / stacked / test | go/no-go, rollback y sustituciones verificadas; revertir runbook/checks. |
| L01 | CAP-057..063 retirada verificable; +160/-40; U10,D50,#107 | [#264](https://github.com/DysTelefonica/access2web-blueprint/issues/264) · chore(exp): legacy retirement gates / stacked / chore | sustituciones probadas, legacy intacto hasta aprobación; revertir flags. |

## Matriz capability → task → spec → UAT

| Capacidades | Task | Spec | UAT |
|---|---|---|---|
| 001–007 | C01–C04 | lifecycle | U01: alta/edición/baja/tipo/estado/jerarquía/ordinal por perfil autorizado |
| 008–014 | R01–R07 | related-data | U02: hitos, modificados, anexos, anualidades, responsables, entidades, UTE |
| 015–024 | Q01–Q02 | catalogs | U03: diez catálogos, filtros, vigencia, permisos y CPV |
| 025–029 | Q03 | query-and-tasks | U04: bandeja, búsqueda avanzada/técnica, export y tareas calculadas |
| 030–032 | W01–W03 | write-resilience | U05: autosave, doble envío, concurrencia, feedback accesible |
| 033–042 | E01–E07 | e2e | U06: DTO, JSON, batch, hash, paquete, trazabilidad, selección, destino, sesión, ordinal |
| 043–046 | A01–A03 | access-control | U07: deny-by-default, principal Lanzadera, sesión aislada, auditoría |
| 047–050 | T01 | runtime | U08: readiness, config, caché reconstruible, binding por despliegue |
| 051–056 | H01–H06 | integrations | U09: HPS, AGEDYS, Riesgos, NC, correo, documentos; fallos/reintentos |
| 057–063 | U10,L01 | uat-cutover-legacy-retirement | U10: sustitución Win32/OLE/globals/backend/popup/menu; cutover y rollback |

## Dependencias GitHub existentes

El plan consume, sin duplicar alcance, #21 (monorepo/hexagonal), #31–#32 (identidad y configuración), #34–#36 (catálogos/consultas), #38 (UI SSR), #43 (auditoría), #44–#45 (migración/runtime), #49–#51 (integraciones/documentos), #54–#55 (E2E/UAT) y #107 (retirada legacy). Cada issue de work unit debe enlazar los números aplicables y bloquearse si su contrato aún no está aceptado.

## Verificación transversal

- RED primero para invariantes, autorización, concurrencia, idempotencia, canonicalización, ciclos, cuarentena, reconciliación y fallos de dependencias; después GREEN y REFACTOR en la misma unidad.
- `pytest tests/expedientes -q` y suites contract/Playwright son checks focalizados; staging/reconciliación y runbooks usan escenarios sin PII. La evidencia runtime de Access es solo lectura.
- No copiar PII, rutas locales ni secretos; no ejecutar retirada antes de reconciliación, UAT aprobada y rollback ensayado.

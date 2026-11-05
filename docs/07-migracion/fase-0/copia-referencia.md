# Copia de producción de referencia de la fase 0

Las comprobaciones de datos de la fase 0 (tablas vacías antes de una cuarentena, duplicados, recuentos) se hacen contra una sola foto fija de los backends de producción. Este documento la identifica. Resuelve la decisión 8 del [catálogo de conflictos](catalogo-accionables-conflictos.md#qué-necesita-decidir-el-propietario).

## Ubicación y origen

- **Carpeta:** `C:\00repos\access2web-datos`, de uso exclusivo para este trabajo; nadie más la toca.
- **Origen:** copia de `C:\00repos\datos` hecha por el propietario el 2026-11-05.
- **Fijada:** 2026-11-05T12:23+01:00.

`C:\00repos\datos` no sirve de referencia: otras pruebas leen y escriben filas en ella. Su estructura sí es fiable, porque esas pruebas no crean tablas ni cambian su forma; las comprobaciones de estructura pueden seguir haciéndose allí.

## Preparación

1. Se comprobó que las 18 bases abren y que todas sus tablas locales se leen. `Gestion_Riesgos_Datos.accdb` se copió mientras estaba abierta en origen; abre y sus 64 tablas se leen. Se borró el `.laccdb` huérfano que llegó con la copia.
2. Los 127 vínculos apuntaban a `C:\00repos\datos`. Se reapuntaron a `C:\00repos\access2web-datos` y `RefreshLink` validó cada uno contra su tabla destino, para que ninguna lectura a través de un vínculo salga de la foto.
3. Los hashes se tomaron después del reapuntado.

Las copias `*.bak-*` y `*.BACKUP-*` que llegaron con la carpeta no forman parte de la referencia.

## Manifiesto

| Backend | Bytes | SHA-256 |
|---|---|---|
| `AGEDO20_Datos.accdb` | 27832320 | `58a1d6da97fb406e1cadcd91a7f2d0ed07521762ebd3c93a05a30295a2ee56a9` |
| `AGEDYS_DATOS.accdb` | 71340032 | `79bfa91645d278302ec3d042db48868466e06253d99e9928041b9dd9ad44f84c` |
| `Control_Cambios_datos.accdb` | 704512 | `fb150ecf13ec4c565a88756ca6c8d11b88d897125b55cfe55d803a74b904c730` |
| `Correos_datos.accdb` | 417603584 | `d854760d995407c94a65788d520febd976b29d28242361020d9cdaba26ad0121` |
| `Expedientes_datos.accdb` | 5767168 | `2090e0807ce9ecd71a0d858d396c3a8b2667509187180f5a41524e5dd1ed2f38` |
| `GestionContratos_datos.accdb` | 1904640 | `182da060affb13f5cfdc4bf1e77ced21b58f89f46f78d8ef399a503c710b3a57` |
| `Gestion_Brass_Gestion_Datos.accdb` | 62525440 | `40848de0681aa5d548f0ce1701168ab93bb32c40b0dcfc053df89baa65f099e6` |
| `Gestion_Riesgos_Datos.accdb` | 23040000 | `c2376e134333aa1c7bdc9875159fdd2d3290fb49de9fa71b5f461a96a3f34bfb` |
| `HPST.accdb` | 7938048 | `4f088948ff195cc1b0f92235383002c59574f15293b7d3dfcdfc6e885e56fcb5` |
| `Lanzadera_Datos.accdb` | 5902336 | `dfc9f090baa925fa21ccd57f4143e0cc29b0b4c9130119a7c206579e4aaf06b0` |
| `NoConformidades_Datos.accdb` | 10973184 | `427caa2c513e89f2f80a915e79730da1e00b58889bbfe70c3c463cb058bcf257` |
| `Registro_Datos.accdb` | 5193728 | `ad82d343930528306a49d0b4d41240ca0cf1f5eb9a08936b00cc3d8a5e5bf7b2` |
| `Registro_Ent_Salida_Datos.accdb` | 11329536 | `6a955e7dc0110936431846c6e0c084e3cfeb56596bb40570c36ded696c011cff` |
| `SICA_datos.accdb` | 1482752 | `4d34ddbc562b453e6ee7a804df85eae060fd680fab11fdd30a3efba7084828dc` |
| `Seguridad_datos.accdb` | 2613248 | `0f74aa2d163ce55246ef3bdc7f69cffa361048c102bab291fb5bc95c5cbd3746` |
| `Solicitudes_HPS_datos.accdb` | 13762560 | `28750af7af0bacc1255e399f2b6cb80d6ae78c5f5294a4a10af6148d46cfea3a` |
| `Tareas_datos1.accdb` | 25509888 | `d7941c1ff54c259780b7bfe188478f110cb8309cbc83de0301bec19118a880ae` |
| `condor_datos.accdb` | 1331200 | `4055d0df1232020b2aeb9ca70292a6adc09a641260dcf65f2c22fea646b4dd92` |

## Regla de uso

- Nunca se abre un fichero de la referencia: se copia a un directorio temporal y se abre la copia con DAO en solo lectura, como hace `scripts/dysflow-erd/extract-erd.ps1`. Access puede escribir en un `.accdb` al abrirlo.
- Antes de cada tanda de comprobaciones se recalcula el SHA-256 y se compara con el manifiesto. Si un hash no coincide, la comprobación no vale: se para y se decide si se fija una foto nueva.
- Toda evidencia de datos de la fase 0 cita este documento como origen.
- La contraseña de los backends llega por la variable de entorno `ACCDB_PASSWORD` y no se escribe en ningún artefacto.

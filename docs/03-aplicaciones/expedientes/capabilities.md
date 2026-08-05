# Expedientes — capacidades observadas

## Resultado

La aplicación cubre un agregado de contratación pública, no solo una pantalla CRUD. La paridad futura debe incluir como mínimo las capacidades siguientes; ninguna se marca como retirada.

| Dominio | Capacidades evidenciadas | Evidencia principal |
|---|---|---|
| Arranque e identidad | splash, configuración de backend, usuario de red, roles Administrador/Calidad/Técnico, preferencias | `Variables Globales.EVE`, `Form_frmSplash`, `Usuario`, `UsuarioAplicacionPermisos` |
| Expediente | alta, edición, eliminación condicionada, cambio de tipo, alta HPS, lectura técnica | formularios `Form_FormExpediente*`, `ExpedienteOperaciones.Registrar` |
| Ciclo de vida | Preoferta, Oferta, Adjudicada, EnEjecucion, EnGarantia, Cerrado, Desestimado, Perdido, NoAPlica, Desconocido; fechas calculadas y garantía | `Expediente.ESTADOCalculado*`, `MotivoNoOK`, `EnumEstados` |
| Identificación | `IDExpediente`, `IDExpedientePadre`, `CodExp`, `CodExpLargo`, `Nemotecnico`, `CodS4H`, proyecto/actividad/pedido | `TbExpedientes`, `ExpedienteJsonExporter`, búsquedas |
| Tipos y clasificación | AM, Lote, Basado de AM, Basado de Lote, individual, HPS; ámbito, grado, órgano, oficina, ejército, PECAL/CPV | `EnumTipoExpediente`, catálogos y subformularios |
| Personas y entidades | responsables por rol, JP, Calidad, Seguridad, RAC, jurídica, comerciales | tablas de relación y `Form_FormExpedienteEntidades` |
| Suministradores | catálogo, relación jerárquica, contratista/subcontratista, UTE, drag & drop y regla «El Árbol Manda» | `Form_FormExpedienteSuministradores`, `ExpedienteSuministradorServicio` |
| Seguimiento | anualidades, hitos, modificados, observaciones, último cambio y cambios de campo | subformularios y `TbCambios`/`TbUltimoCambio` |
| Documentación | anexos y referencias SharePoint; eliminación condicionada por fichero/registro según helper | `Form_FormExpedienteDocumentacion`, `TbExpedientesAnexos` |
| Consulta | gestión general, vista técnica, búsqueda avanzada, filtros por estado/código/JP/jurídica/suministrador, detalle solo lectura | `Form_FormExpedientesGestion*`, helpers de consulta |
| Tareas | estado desconocido, recepción completa/hito, adjudicado sin contrato, TSOL sin S4H, oferta prolongada | `PintarTareas`, `Form_FormTareas`, colecciones de `Entorno` |
| Catálogos | alta/edición/baja master-detail de comerciales, CPV, ejército, grado, lugar, oficina, órgano, PECAL, RAC y suministrador | gestor de entidades y pares de formularios |
| Salidas | Excel, JSON completo, exportación familiar/batch E2E, historial y destino por usuario | `Form_FormEleccionTipoConsulta`, `ExpedienteJsonExporter`, E2E |
| Integraciones | Lanzadera, AGEDYS, HPS, riesgos, no conformidades y correos | `constructor`, `Variables Globales`, clases de integración |

## Reglas de conservación

- La escritura del agregado es transaccional y actualiza cabecera, hijos, caché y último cambio.
- `TbExpedientesConEntidades` es un read-model persistido; no debe perderse aunque se regenere en el futuro.
- La ausencia de pruebas runtime no equivale a ausencia de capacidad: la matriz OpenSpec registra reglas `Verified-static` y diferidas.
- El inventario de capacidades incluye capacidades raras/técnicas: E2E, configuración local, cachés, anti-spam y selección temporal.

## Evidencia previa

Se han cosechado PRD, Discovery Map, Architecture Overview, ERD, OpenSpec CAP-001..055, UAT y releases antes de inspeccionar staging. Las afirmaciones divergentes entre esos documentos y el código quedan abiertas, no resueltas por intención.

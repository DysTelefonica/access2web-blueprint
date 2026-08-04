# 02 · Topología del ecosistema — Identificadores y glosario

## Propósito

Diccionario controlado de los identificadores transversales y de los términos recurrentes del ecosistema. Evita mezclar conceptos que en Access son parecidos pero contractualmente distintos.

## Identificadores a documentar

| Identificador | Origen | Significado provisional | Estado |
|---|---|---|---|
| `IDExpediente` | `TbExpedientes` y FKs consumidoras | Identificador numérico del expediente. | Confirmado como campo; contrato de equivalencia por aplicación, pendiente. |
| `CodExp` | `TbExpedientes` / consumidores | Código de expediente (formato a verificar). | Pendiente de formato y cardinalidad. |
| `CodExpLargo` | `TbExpedientes` | Variante extendida; uso por confirmar. | Pendiente. |
| `Nemotecnico` | `TbExpedientes` y consumidores | Nemotécnico del expediente; base de la excepción en No Conformidades. | Pendiente de alcance. |
| `UsuarioRed` | Lanzadera / consumidores | Cuenta de dominio del usuario. | Confirmado como familia de patrones. |
| `CorreoUsuario` | Lanzadera / consumidores | Dirección SMTP del usuario. | Confirmado como familia de patrones. |
| `IDAplicacion` | `TbAplicaciones` | Clave de la aplicación en el catálogo. | Confirmado; falta extraer las ocho filas vigentes. |
| `NombreCorto` | `TbAplicaciones` | Etiqueta corta de la aplicación. | Confirmado; valores por extraer. |

## Glosario (mínimo, provisional)

| Término | Significado |
|---|---|
| Backend | Fichero `.accdb` de datos, separado del frontend. |
| Frontend | Fichero `.accdb` de interfaz y código. |
| `TbConfiguracionBackends` | Tabla del frontend con rutas a backends. |
| `getdb()` / `getdbLanzadera` / `getdbExpedientes` | Helpers de acceso al backend. |
| `TbTablasAVincular` | Catálogo Lanzadera de tablas que se enlazan por aplicación. |
| `Verified-runtime` | Hecho verificado con prueba ejecutada. |
| `Verified-static` | Hecho leído en código o doc, sin prueba aún (deuda a cerrar). |
| `Divergent` | Discrepancia entre SDD/intención y código real. |

## Reglas

- Los identificadores **no** se asumen intercambiables: cualquier normalización web exige tabla de correspondencias.
- El glosario crece solo cuando aparece un término nuevo y relevante; no se duplican entradas de cada aplicación.
- APAP y APAP_WEB no entran ni como identificador ni como glosario.

## Checklist

- [ ] Cada identificador lleva su tabla origen y, si la hay, su equivalencia verificada.
- [ ] Cada entrada del glosario lleva una cita o se retira.
- [ ] Los términos nuevos se añaden al cerrar un lote, no a mitad de pasada.

## Siguiente paso

Reutilizar este glosario en `04-integraciones-y-operacion/identidad-arranque-y-permisos.md` y en cada ficha de `03-aplicaciones/`.

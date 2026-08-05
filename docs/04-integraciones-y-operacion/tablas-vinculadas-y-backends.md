# 04 · Integraciones y operación — Tablas vinculadas y backends

## Propósito

Inventario de cómo se enlazan frontend y backend en el legado y la dirección objetivo de la plataforma web. Sirve de base para la sustitución por adapters hexagonales y para que `07-migracion/` pueda documentar cada paso sin reintroducir secretos en la administración.

## Estado del contenido

- **Legacy documentado (parcial)**: patrón `getdb()` / `getdbLanzadera` / `getdbExpedientes`; tabla frontend `TbConfiguracionBackends`; tablas backend por aplicación.
- **Objetivo registrado**: persistencia por adapter (PostgreSQL por esquema), rutas y credenciales externalizadas al despliegue.

## Patrones legacy observados

- **`TbConfiguracionBackends`** (frontend Lanzadera): columnas relevantes `BackendActivo`, `BackendProduccion`, `BackendSandbox`, `PasswordBackend`, `IDAplicacion`, rutas PROD/LOCAL y `EnPruebas`. La configuración actual declara `BackendActivo=PROD`, `IDAplicacion=12` y rutas internas; los valores sensibles (rutas absolutas, contraseñas de backend) **no se reproducen en este documento**.
- **`Variables Globales.getdb()`** y variantes `getdbLanzadera` / `getdbExpedientes`: eligen backend activo y abren DAO. El code-path tiene 76 callers internos en Lanzadera, lo que confirma que no es un detalle local aislado.
- **`Config_BackendHelper.bas`**: rutas y constantes de entorno hardcoded coexisten con la configuración persistida. Duplicidad a eliminar antes de migrar.
- **`TbTablasAVincular`** (documentada): modela tablas que cada aplicación debe enlazar. No se inspeccionó contenido en esta pasada.
- **`TbConexiones` / `TbConexionesRegistro`** (backend Lanzadera): registran eventos de conexión y, en algunos campos, telemetría de SSID, oficina y coordenadas.

## Decisión: retirada de la administración técnica en Lanzadera (APROBADO)

- La gestión de **rutas de backend**, **contraseñas de backend**, **ficheros Access por entorno** y **configuración de tablas vinculadas** se **retira** como capacidad de administración de Lanzadera.
- Los valores se migran a **configuración de despliegue** (variables de entorno o adapter de secretos) accesible solo por mecanismos de plataforma, no por UI de aplicación.
- Esta retirada elimina la duplicidad entre `TbConfiguracionBackends` y `Config_BackendHelper.bas`.

## Objetivo: configuración técnica externalizada (APROBADO dirección)

- Cada backend se invoca a través de un **port**. Los adapters (PostgreSQL por esquema; object storage; colas; scheduler; proveedores de secretos) son reemplazables.
- **Secretos**: nunca en código, configuración versionada o logs. Entrada por variables de entorno o por un adapter de proveedor de secretos.
- **Persistencia objetivo preferida**: una base de datos PostgreSQL compartida con **esquemas por módulo**. Una BD física compartida no implica propiedad compartida: los límites de esquema y las reglas de acceso deben preservar el aislamiento hexagonal y modular.

## Decisiones aún no tomadas (ABIERTO)

- Tecnología concreta del adapter de caché.
- Topología de despliegue (on-premise, nube corporativa, OCP).
- Stack exacto de implementación.
- Estrategia de migración de datos desde los `.accdb` a PostgreSQL.

## Fuentes de autoridad

1. `TbConfiguracionBackends` y `TbTablasAVincular` en cada frontend.
2. `C:\00repos\codigo\<app>\00_main` por aplicación.
3. Inspección Dysflow solo lectura.
4. CodeGraph-VBA para trazado de símbolos.
5. Engram como contexto histórico.

## Reglas de evidencia

- Una tabla se declara `vinculada` solo si la fuente lo afirma; un enlace manual no es una FK.
- Las rutas se citan textualmente, **sin secretos**: se mencionan campos cuando el contexto lo requiere, pero no se exponen valores internos.
- Las configuraciones Dysflow no se modifican aquí.

## Checklist

- [ ] Cada tabla vinculada lleva su backend origen y su helper de apertura (cuando aplique).
- [ ] Las rutas observadas se distinguen de las rutas resueltas en runtime.
- [ ] Las contraseñas y rutas internas no aparecen en este documento.

## Siguiente paso

Cruzar con `02-topologia-ecosistema/topologia-frontends-backends.md` y con `04-integraciones-y-operacion/rutas-entornos-y-contingencia.md` para mantener una única fuente de verdad de rutas.

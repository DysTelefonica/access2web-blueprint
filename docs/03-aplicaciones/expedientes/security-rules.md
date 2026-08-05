# Expedientes — seguridad y reglas

## Autorización

La matriz común se mantiene en [06-autorizacion-legacy-matriz.md](../../06-autorizacion-legacy-matriz.md). Aquí solo se conserva el comportamiento específico: Administrador tiene prioridad; Calidad se evalúa después; cualquier usuario restante se clasifica como Técnico. Las pantallas técnicas son de solo lectura en el flujo observado.

## Reglas de negocio críticas

- Alta de lote/basado requiere padre válido.
- Eliminación queda condicionada por hijos/relaciones.
- El árbol de suministradores impide auto-padre y borrar nodos con hijos; las raíces propias dependen de `ConsorcioPropio`.
- `Registrar` valida, inicia transacción, persiste cabecera/agregados, actualiza caché y hace commit o rollback completo.
- Estado y garantía se calculan a partir de fechas, flags y estado manual; deben conservarse tanto el valor fuente como el calculado.
- Auditoría registra usuario y fecha de último cambio; cambios de campos se modelan aparte.
- Errores inesperados se propagan y pueden notificar al administrador; el fallo de correo no debe borrar evidencia.

## Riesgos de privacidad/migración

No se han incluido nombres, correos, contraseñas, cadenas de conexión, hashes, hosts ni filas personales. El ERD histórico contiene referencias sensibles saneadas solo conceptualmente; no se reproducen en estos artefactos. La migración deberá separar identidad/autorización externa de los datos de negocio y preservar las claves de enlace de forma controlada.

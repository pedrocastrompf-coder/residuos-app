# Integrar Hallazgos en el Apps Script (3 pasos)

1. En Apps Script: **+ > Script** → nómbralo `Hallazgos` → pega todo `Hallazgos.gs`.
2. En `Code.gs`: borra todo y pega el `Code.gs` nuevo (ya incluye el cierre de `data` que aplicaste).
3. Ejecuta `crearEstructura` (crea las pestañas **Hallazgos** y **Áreas** y la columna `areas` en Usuarios).
4. Ejecuta `pruebaHallazgos` y **acepta los permisos de Drive** (esto va antes de publicar). En el registro debe verse: crear, «repetido», asignar, cerrar, «estado_invalido» y la carpeta `Hallazgos Ambientales/2026/11`. Después borra la fila «PRUEBA» de Hallazgos y las áreas «Prueba …» de Áreas.
5. Recién ahora: **Implementar > Administrar implementaciones > lápiz > Nueva versión**.

## Qué cambió en Code.gs (para revisarlo)
- `ROLES` suma `"Responsable de área"`; `PERMISOS` suma `hallazgo_reportar/asignar/cerrar/reabrir` y `areas_admin`.
- `USU_COLS` suma `areas`; se leen, guardan y devuelven (`usuariosTodos_`, `escribirUsuario_`, `pub_`, `usuarioCrear_`, `usuarioEditar_`).
- `doPost`: 8 rutas nuevas y una regla: el **Responsable de área solo puede usar hallazgos** (no pesa ni despacha ni lee pesajes, aunque fuerce la llamada: también en `doGet`).
- `SIN_TURNO` suma `hallazgos_listar`, `hallazgo_foto`, `areas_listar` (solo lectura).
- `crearEstructura()` y `plantaCrear_()` crean las dos pestañas nuevas.
- No toca `CLAVE`.

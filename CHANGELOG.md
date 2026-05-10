Changelog — Reporte Diario de Mantenimiento Ferring
v2.3 — 2026-05-08
Mejoras de Dashboard
Solo OTs del turno actual: el Dashboard ahora muestra únicamente las OTs correctivas que fueron creadas o modificadas en el turno actual. Las OTs heredadas del turno anterior (carry-over) que nadie tocó en este turno NO aparecen en el Dashboard, pero siguen visibles en "Cargar Reporte" para que el responsable pueda actualizarlas. Esto evita que el Dashboard quede saturado de OTs viejas y refleja realmente el trabajo del turno.
Detalle por técnico más compacto: la sección "Por técnico" del card Preventivos del Turno ahora muestra los técnicos en grid de 2 columnas en vez de uno debajo del otro. Reduce ~60% el espacio vertical ocupado.
Título del card Correctivos: cambia de "Correctivos (N)" a "Correctivos del turno (N)" para reflejar que es solo del turno actual.
Cambios críticos en Estadísticas (BUG FIX)
⚠️ Aviso de impacto: los números en la pestaña Estadísticas van a cambiar (algunos bajar) respecto a las versiones anteriores. Esto NO es una pérdida de datos: es la corrección de un bug que estaba inflando los conteos.
El bug
En V1.0 a V2.2, una OT correctiva pendiente se contaba una vez por cada turno que aparecía en el reporte. Si una OT estaba "En Curso" durante 5 turnos, se contaba 5 veces en las estadísticas. Lo mismo con "Equipos con más correctivos" (un equipo con 1 OT pendiente 5 días sumaba 5 puntos).
El fix (V2.3)
Cada OT correctiva (identificada por su número) ahora cuenta una sola vez en todas las estadísticas. Su estado es el del reporte más reciente del rango analizado.
Lugares afectados:
KPI "Correctivos" (total)
KPI "Realizados"
KPI "Pendientes"
KPI "% Cumpl."
Pie chart "Estado de Correctivos"
Bar chart "Equipos con más correctivos"
Bar chart "Carga por técnico" (parte de correctivos)
Bar chart "Distribución por turno" (parte de correctivos)
Bar chart "Trabajos en el período" (parte de correctivos)
Cards "Último Día" y "Fin de Semana" (correctivos generados/realizados)
Sin cambios en los números de:
Preventivos (cada turno hace su propio trabajo preventivo, no se duplican)
Comentarios urgentes
Reportes guardados (sigue siendo el conteo de turnos)
Casos borde
Las OTs sin número de OT (`OT-XXXX` vacío) NO se pueden deduplicar; cada aparición sigue contándose. En operación normal todas las OTs deberían tener número.
Para el bar chart "Trabajos en el período", cada OT se muestra en la fecha de su primera aparición (cuando se "abrió" la OT en el sistema), no en la fecha en que se cerró.
Detalles técnicos
Se agrega un campo `createdInShift` (formato `YYYY-MM-DD-Turno`) al objeto OT cuando se crea con el botón "Agregar OT". Reportes viejos sin este campo siguen funcionando (compatibilidad hacia atrás: se asume que pertenecen al turno actual).
Se agrega `lastModifiedInShift` que se setea automáticamente cada vez que se edita un campo de la OT.
El Dashboard filtra usando estos dos campos.
El filtro de Cargar Reporte (carry-over de OTs pendientes) funciona igual que antes — las OTs heredadas siguen apareciendo ahí para que se las pueda actualizar.
Sin cambios en
Esquema de Supabase (los nuevos campos van en el JSONB existente)
Catálogos
Lógica del formulario, validaciones, exports a Excel
Dashboard layout (sigue con opción A: correctivos 50% izq con sub-columnas)
Botones de export a PNG/PDF (V2.2 sigue funcionando igual)
---
v2.2 — 2026-05-08
Nuevas funcionalidades
Exportar Dashboard a PNG: nuevo botón en la pestaña Dashboard que descarga una imagen PNG con todo el contenido del Dashboard expandido.
Exportar Dashboard a PDF: nuevo botón que genera un PDF A4 horizontal con el Dashboard completo.
Las librerías `html2canvas` y `jspdf` se cargan lazy desde CDN solo cuando el usuario hace click.
Antes de capturar, la app expande automáticamente todos los contenedores con scroll interno.
Resolución de la imagen: 2x (alta calidad para impresión).
Nombre de archivo: `Dashboard_YYYY-MM-DD_Turno.png` o `.pdf`.
---
v2.1 — 2026-05-08 (fixes post-testing V2.0)
Mejoras de UX
Reordenado del formulario.
Dashboard con nuevo layout (correctivos 50% izq con sub-columnas Realizadas | Pendientes).
Texto del Último Día corregido a "Turnos del día (N): ...".
Bug fixes
Mensaje de error persistente al hacer Limpiar.
---
v2.0 — 2026-05-08
Nuevas funcionalidades
Logo Biomas en el header.
Estadísticas: apartados "Fin de Semana" y "Último Día".
Resumen Preventivos del Turno con validación cruzada.
Exportación a Excel: Solo Comentarios, Solo Proveedores.
Versión visible en el header.
Cambios de comportamiento
Eliminar OT correctiva removido.
Limpiar conserva pendientes y borra realizadas.
Validación de técnicos obligatorios en correctivos "Realizada".
Validación cruzada de preventivos.
Fechas en formato `dd/mmm/aa`.
Equipo en Dashboard con wrap multi-línea.
Cambios de schema (Planta de Efluentes y Caldera)
Nuevos campos: PTEL (Caudal m³/h, Vacío, ΔT, %TK1, %TK2, %TK7), Caldera (Conductividad mS, pH), Ablandadores (Conductividad mS, pH).
Eliminados: Ablandador (campo único), TK Emergencia, TK4.
---
v1.0 — 2026-05-04 (versión inicial en producción)
Dashboard con tanques/cisternas y gauges semicirculares
Estadísticas extendidas con filtro personalizado
Histórico no editable
Exportación Excel matching exacto del template
Persistencia en Supabase (sin login)

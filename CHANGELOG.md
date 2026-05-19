Changelog — Reporte Diario de Mantenimiento Ferring
## [v2.8] - 2026-05-19
  ### Bug crítico resuelto: carry-over stale (OTs Realizadas que reaparecían)
  - **Detección de conflictos al guardar**: antes de cada save, la app verifica contra Supabase si las OTs del reporte ya fueron cerradas por otro turno. Si hay conflicto, abre modal para decidir.
  - **Modal de resolución**: por cada OT en conflicto, el usuario debe elegir entre quitarla o reabrirla. Reapertura solo en modo admin y con motivo obligatorio registrado en timeline.
  - **Migración SQL del histórico**: script one-time corrigió todas las apariciones contradictorias previas, agregando entradas auditables al timeline.

  ### Notas técnicas
  - La verificación al guardar agrega ~1s al guardado, dependiente de la latencia de Supabase. Si falla por red, el guardado procede sin verificación (no bloquea al usuario).
## [v2.7] - 2026-05-17
  ### Cambios en pestaña Estadísticas
  - **Visibilidad admin-only** para tarjetas "Distribución por turno" y "Carga por técnico" (antes visibles para todos)
  - **Nueva tarjeta admin**: "OTs dejadas pendientes por turno de origen" — ranking M/T/N + barras
  - **Nueva tarjeta admin**: "OTs heredadas cerradas por turno" — ranking M/T/N + barras
  - Excluye OTs legacy (sin formato XXX-YYYYY) en los rankings nuevos
  - Turnos sin datos en la ventana se muestran como "—"
v2.6 — 2026-05-17
Cambios principales
1. Modo administrador con password
Nuevo botón "Admin" (candado) en el header, debajo del badge "Supabase conectado". Al clickearlo se abre un modal de login con password hardcoded (FerringBiomas2026). Soporta Enter para confirmar y Escape para cancelar.
Una vez activado, aparece un badge rojo "MODO ADMIN" junto con un botón "Salir" para desactivarlo.
Sin trazabilidad / audit trail: decisión deliberada por tratarse de control personal del responsable del proyecto, no flujo GMP regulado.
2. Edición y eliminación de OTs correctivas en modo admin
Botón 🗑️ en la esquina superior derecha de cada OT correctiva para eliminarla individualmente del reporte.
Botones ✏️ y 🗑️ por cada entrada del timeline de Estado de avance, permitiendo edición inline y eliminación de entradas históricas.
3. Eliminación de reportes completos en modo admin
Desde FormView: nuevo botón "Eliminar reporte" en la action bar (solo aparece si el reporte ya está guardado en Supabase).
Desde Histórico: columna nueva "Acciones" con botón 🗑️ rojo en cada fila de reportes guardados.
Ambos disparan un modal de confirmación antes del borrado. El delete es hard delete (sin papelera ni recuperación).
4. Edición desde Dashboard en modo admin
Click en una OT correctiva del Dashboard → redirige a "Cargar Reporte" con ese reporte abierto para edición.
Click en la card de Preventivos del Turno → mismo comportamiento.
Las cards muestran hover azulado y ring sky para indicar que son clickeables.
5. Las ediciones en modo admin no marcan lastModifiedInShift
Las correcciones retroactivas no aparecen como trabajo del turno actual en filtros del Dashboard. Esto preserva la coherencia del filtro V2.3 (Dashboard muestra solo OTs trabajadas en el turno).
v2.5 — 2026-05-17
Cambios principales
1. Dashboard: correctivos divididos en 4 sub-secciones
La columna "Realizadas" ahora se divide en "Del turno" (creadas y cerradas en este turno) y "Heredados realizados" (creadas en turno previo, cerradas acá).
La columna "Pendientes" ahora se divide en "Del turno" (creadas acá, aún sin cerrar) y "Heredados" (creadas en turno previo, aún sin cerrar).
El criterio es por creación de la OT (createdInShift), no por trabajo realizado.
El total del header (Correctivos del turno (N)) se mantiene combinado.
Si una categoría está vacía, se muestra "Sin novedades" debajo del subtítulo.
2. Marca de "Avance hoy" en pendientes heredados
Cuando una OT heredada tiene una entrada de timeline cargada en el turno actual, aparece un badge verde "Avance hoy" al lado del N° OT.
Adicionalmente, debajo de la descripción, se muestra una línea destacada en verde con el texto del último avance del turno: "↳ Avance del turno: ...".
3. Técnico obligatorio en correctivos y preventivos
Antes: el técnico era obligatorio solo en correctivos con estado "Realizada".
Ahora: es obligatorio en cualquier estado (Sin Iniciar, En Curso, Realizada) y también en todos los preventivos cargados.
Las filas sin técnico se muestran con borde rojo y mensaje inline antes de apretar Guardar.
4. Banner permanente en sección Correctivos del formulario
Arriba de la lista de OTs aparece un banner rojo recordando que, para OTs heredadas de turnos previos, solo cargue "Estado de avance" si hay novedades.
5. Modal de confirmación al guardar con entradas vacías
Si al apretar Guardar hay correctivos o preventivos completamente vacíos (sin N° OT/equipo, sin descripción y sin técnicos), se abre un modal indicando cuántas filas se van a eliminar y pidiendo confirmación.
Al confirmar: se eliminan las filas vacías y se guarda el resto.
Al cancelar: vuelve al formulario sin cambios.
6. Bug fix: texto cortado al exportar Dashboard a PNG/PDF
Los campos largos (códigos de equipo, descripciones) que en pantalla quedan con truncate ahora se des-truncan durante la captura.
La vista normal en pantalla sigue con truncate para mantener el layout compacto.
Sin cambios en
Esquema de Supabase (todo sigue en el JSONB existente).
Lógica de carry-over de OTs pendientes al siguiente turno.
Catálogos (responsables, técnicos, equipos auxiliares).
Exports a Excel (mismas hojas y columnas que V2.4).
Estadísticas y KPIs.
Filtrado del Dashboard de reportes nuevos vs guardados.
v2.4 — 2026-05-10
Cambios principales
1. N° OT con formato estandarizado XXX-YYYYY (según SOP 10.3.2)
El campo "N° OT" ahora es un input compuesto: dropdown de sector + input numérico de 5 dígitos.
Sectores disponibles (lista cerrada según SOP):
`FOA1`, `FB2`, `RO`, `BIO`, `DEP`, `PP`, `PAD`, `LIM`, `EHS`, `MAN`, `FAC`
El correlativo se completa automáticamente con ceros a la izquierda al perder foco (ej: escribís `1395` → guarda `01395`).
Al guardar reporte, las OTs nuevas (creadas en el turno actual) deben tener formato válido. Si no, no se puede guardar.
Las OTs heredadas del carry-over con formato legacy se muestran con un input plano y badge "L" (Legacy). NO se valida el formato de las legacy. Si el responsable las cambia al formato nuevo, dejan de ser legacy.
2. Estado de avance (timeline acumulable)
Para OTs en "En Curso" se agrega una sección "Estado de avance" con un timeline acumulable.
Cada vez que un responsable carga texto y hace click en "Guardar avance", se agrega una entrada con: fecha, turno, autor (responsable del turno), y el texto.
Las entradas anteriores no son editables ni borrables (auditoría GMP).
Al guardar reporte, si una OT está en "En Curso", es obligatorio que tenga al menos una entrada del turno actual en el timeline. Si no, no se puede guardar.
Si la OT pasa a "Realizada" o "Sin Iniciar", el timeline queda visible (read-only) pero no exige entradas nuevas.
El timeline se exporta a Excel en una columna nueva `EstadoAvance` con todas las entradas concatenadas.
3. Multi-select de técnicos en Resumen Preventivos (opción C)
El "Detalle por técnico" del Resumen de Preventivos ahora permite cargar grupos de técnicos.
Si Juan + Pedro hicieron 4 preventivos juntos, se carga UNA fila con ambos seleccionados y cantidad 4.
La cantidad cuenta una sola vez para validación cruzada con "Preventivos realizados". Es decir, el grupo de arriba suma 4, no 8.
En Estadísticas (gráfico "Carga por técnico"), tanto Juan como Pedro reciben +4 cada uno por ese grupo.
En el Excel exportado, el grupo se desglosa en filas individuales (Juan con cantidad 4, Pedro con cantidad 4) más una columna nueva `EnGrupoCon` que indica con quién más estuvo en el grupo.
4. Orden de OTs correctivas
Las OTs nuevas (creadas en el turno actual) ahora se insertan arriba del listado al hacer click en "Agregar OT". Antes iban al final.
En el carry-over, las OTs se ordenan por última aparición descendente: las más recientemente vistas arriba, las que llevan días sin movimiento abajo.
5. "Carga por técnico" muestra los 21 técnicos del catálogo
El gráfico de Estadísticas ahora muestra siempre los 21 técnicos del catálogo, aunque no tengan OTs en el período.
Orden descendente por carga total (los más cargados arriba, los que están en cero al final).
Útil para detectar técnicos sub-utilizados o no asignados.
El alto del gráfico se ajusta automáticamente para que entren todas las barras.
6. Bug fix: KPI "Pendientes" ahora cuenta Sin Iniciar + En Curso
Antes, el KPI "Pendientes" en Estadísticas solo contaba OTs en estado "Sin Iniciar".
Ahora cuenta correctamente: Sin Iniciar + En Curso.
Compatibilidad hacia atrás
Todos los cambios son retro-compatibles con reportes guardados en V1.0 a V2.3:
Reportes viejos sin `timeline` → al cargarlos, se inicializa como array vacío.
Resumen preventivos viejo con schema `{tecnico, cantidad}` → se migra automáticamente a `{tecnicos: [tecnico], cantidad}` al cargar (la migración es solo en memoria, no se reescribe la base).
OTs con formato legacy (sin `XXX-YYYYY`) → se muestran tal cual con badge "L". No se exige migrar.
OTs sin `createdInShift` (V2.2 o anteriores) → se asume formato legacy y se permite editarlas sin validar.
Sin cambios en
Esquema de Supabase (todos los nuevos campos van en el JSONB existente)
Catálogos de técnicos, foguistas, responsables, compresores, grupos electrógenos
Lógica de Servicios (Planta, Cisternas, Agua de Pozo, Proveedores)
Comentarios del turno
Botones de export PNG/PDF del Dashboard (V2.2)
Filtro Dashboard solo turno actual (V2.3)
Deduplicación de stats (V2.3)
Bug conocido (pendiente)
El input nativo `<input type="date">` muestra el formato según el idioma del SO/navegador. En sistemas en inglés muestra `mm/dd/aaaa`. Solución temporal: configurar idioma del navegador a "Español (Argentina)".
---
v2.3 — 2026-05-10
Dashboard solo OTs creadas o modificadas en el turno actual
Detalle por técnico en preventivos: grid 2-col compacto
BUG FIX en Estadísticas: cada OT correctiva cuenta una sola vez
Datos limpios en Supabase (8 OTs sin número eliminadas + 5 con espacios normalizadas)
v2.2 — 2026-05-10
Botones de export Dashboard a PNG y PDF (Dashboard)
Carga lazy de html2canvas y jspdf desde CDN
v2.1 — 2026-05-08
Reordenado del formulario
Dashboard layout opción A (correctivos 50% izq con sub-columnas Realizadas | Pendientes)
Texto Último Día corregido a "Turnos del día (N): ..."
Bug fix mensaje de error persistente al hacer Limpiar
v2.0 — 2026-05-08
Logo Biomas en header
Estadísticas: apartados "Fin de Semana" y "Último Día"
Resumen Preventivos del Turno con validación cruzada
Exportación a Excel: Solo Comentarios, Solo Proveedores
Versión visible en header
Eliminar OT correctiva removido
Limpiar conserva pendientes y borra realizadas
Validación de técnicos obligatorios en correctivos "Realizada"
Fechas en formato dd/mmm/aa
Equipo en Dashboard con wrap multi-línea
Schema Planta de Efluentes y Caldera nuevo (PTEL + Caldera + Ablandadores)
v1.0 — 2026-05-04 (versión inicial en producción)
Dashboard con tanques/cisternas y gauges semicirculares
Estadísticas extendidas con filtro personalizado
Histórico no editable
Exportación Excel matching exacto del template
Persistencia en Supabase (sin login)

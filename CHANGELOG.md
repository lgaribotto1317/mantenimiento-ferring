Changelog — Reporte Diario de Mantenimiento Ferring
v2.0 — 2026-05-08
Nuevas funcionalidades
Logo Biomas en el header: reemplaza el icono del casco. Fondo blanco según diseño corporativo.
Estadísticas — Apartado "Fin de Semana": muestra el último FDS cerrado (viernes Noche + sábado completo + domingo completo) con 4 métricas: correctivos generados, correctivos realizados, preventivos asignados, preventivos realizados.
Estadísticas — Apartado "Último Día": muestra los últimos 3 turnos del día más reciente con datos cargados (típicamente Mañana + Tarde + Noche del día previo cuando se carga el reporte por la mañana). Mismas 4 métricas que el apartado de Fin de Semana.
Resumen Preventivos del Turno: nueva sección al final del formulario con campos numéricos "Asignados" y "Realizados", más detalle por técnico (técnico del turno + cantidad). Validación cruzada: la suma del detalle debe coincidir con "Realizados".
Exportación a Excel — Solo Comentarios: botón nuevo en Histórico & Excel, exporta solo la hoja de comentarios con fecha y turno.
Exportación a Excel — Solo Proveedores: botón nuevo en Histórico & Excel, exporta solo la hoja de proveedores con fecha y turno.
Versión visible en el header: badge `v2.0` al lado del título.
Cambios de comportamiento
Eliminar OT correctiva removido: ya no se puede eliminar una OT correctiva una vez creada. Las OTs en estado "Sin Iniciar" o "En Curso" persisten siempre entre turnos. Solo se "limpian" del formulario al pasar a "Realizada".
Limpiar formulario: el botón "Limpiar" ahora conserva todas las correctivas en estado "Sin Iniciar" o "En Curso". Solo se quitan las que están en "Realizada". El resto del formulario (preventivos, servicios, comentarios, resumen, etc.) se vacía como antes.
Validación al guardar: las OTs correctivas en estado "Realizada" deben tener al menos un técnico asignado. Si falta, el guardado se bloquea con mensaje de error.
Validación cruzada de preventivos: si "Preventivos realizados" > 0, la suma de cantidades en el detalle por técnico debe coincidir. Si no, el guardado se bloquea.
Fechas en formato `dd/mmm/aa`: en Dashboard, Histórico, y todas las pestañas. Ej: `08/may/26`. El input de fecha sigue usando el formato del navegador/SO (limitación del HTML nativo).
Equipo en Dashboard con wrap multi-línea: todos los técnicos del turno se muestran como badges, sin truncamiento. Cuando el equipo es grande, ocupa más renglones.
Cambios de schema (Planta de Efluentes y Caldera)
⚠️ Aviso: a partir de V2.0 cambia la estructura de los parámetros de Planta de Efluentes y Caldera. Esto afecta también el archivo Excel exportado en la hoja "Planta de efluentes" (las columnas viejas se reemplazan por las nuevas). El usuario confirmó que esta hoja no constituye registro GMP.
Campos nuevos:
PTEL: Caudal (m³/h), Vacío del equipo, ΔT entre torres (°C), % Nivel TK1, % Nivel TK2, % Nivel TK7
Caldera: Conductividad (mS), pH
Agua Ablandadores: Conductividad (mS), pH
Campos eliminados (solo del formulario y del Excel exportado; los datos viejos quedan archivados en Supabase JSONB pero no se muestran):
Ablandador (campo único anterior)
TK Emergencia
TK4
Migración aplicada: opción A — los reportes viejos quedan tal cual están en Supabase. Cuando se abren con el nuevo schema, los campos viejos no aparecen y los nuevos están vacíos. No se migran datos automáticamente.
Excel — hojas nuevas
Resumen Preventivos Turno: `ResumenID | Fecha | Turno | Asignados | Realizados`
Preventivos por Tecnico: `RegistroID | Fecha | Turno | Tecnico | TecnicoID | Cantidad`
Excel — hoja modificada
Planta de efluentes: las columnas `Ablandador`, `TKEmergencia`, `TK4` se reemplazan por `PTEL_Caudal_m3h`, `PTEL_Vacio`, `PTEL_DeltaT_C`, `PTEL_TK1_pct`, `PTEL_TK2_pct`, `PTEL_TK7_pct`, `Caldera_Conductividad_mS`, `Caldera_pH`, `Ablandador_Conductividad_mS`, `Ablandador_pH`.
UI
Nueva sección "Resumen Preventivos del Turno" al final del formulario, después de Comentarios.
Dashboard: la columna "Preventivos" ahora muestra solo el resumen (asignados, realizados, detalle por técnico). El detalle individual de tareas preventivas sigue cargándose en el formulario y se exporta al Excel, pero no se visualiza en Dashboard.
Hidratación automática de reportes viejos: la app ahora completa la estructura de objetos al leer reportes guardados con schema viejo, evitando errores de undefined.
Sin cambios en la base de datos
Toda la modificación de datos se hace en el campo `data` JSONB de la tabla `reportes`. No se requiere ALTER TABLE ni ningún SQL en Supabase.
---
v1.0 — 2026-05-04 (versión inicial en producción)
Dashboard con tanques/cisternas y gauges semicirculares
Estadísticas extendidas con filtro personalizado
Histórico no editable
Exportación Excel matching exacto del template
Persistencia en Supabase (sin login)
Compresor PTEL agregado
Grupo Electrógeno Depósito 10 agregado
Niveles cisternas con estados
Dashboard reorganizado a una pantalla sin scroll
Foguistas con multi-select, persistencia de correctivos al limpiar formulario
Sección "Agua de Pozo" con cloro de Pozo 3 y Pozo 6 numéricos

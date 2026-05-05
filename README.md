# Reporte Diario de Mantenimiento

App web para registrar y consultar reportes de mantenimiento por turno, con dashboard, estadísticas y exportación a Excel.

## Despliegue rápido en Vercel

Ver el archivo `INSTRUCCIONES_DEPLOY.txt` adjunto para los pasos detallados.

Variables de entorno necesarias (configurar en el panel de Vercel):
- `VITE_SUPABASE_URL` — URL del proyecto Supabase
- `VITE_SUPABASE_ANON_KEY` — anon public key de Supabase

## Stack

- React 18 + Vite
- Tailwind CSS
- Recharts (gráficos)
- SheetJS (Excel)
- Lucide React (iconos)
- Supabase (base de datos compartida)
<!-- trigger build -->

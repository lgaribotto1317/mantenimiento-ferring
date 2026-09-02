import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import {
  ClipboardList, BarChart3, Download, Plus, Trash2, Save, Calendar, Users,
  Wrench, Activity, FileSpreadsheet, CheckCircle2, AlertTriangle, Building2,
  HardHat, Beaker, ListChecks, ChevronDown, ChevronLeft, ChevronRight, X, FileText, TrendingUp, Flame,
  Cog, Zap, Filter, Search, Cloud, CloudOff, RefreshCw, Settings, MessageSquare,
  CalendarDays, Clock, Image as ImageIcon, FileDown,
  Lock, LogOut, Edit3, Shield, RotateCcw, Inbox, Ban, Timer
} from 'lucide-react';
// #66 — xlsx-js-style en vez de xlsx: mismo motor SheetJS 0.18.5 por dentro
// (la Community Edition que ya usaba el proyecto no escribe estilos de
// celda), agrega la propiedad `.s` por celda para poder poner negrita y
// color de fondo en el Excel de Horas Extras. Mismo import, misma API para
// todo lo demás — el resto de los exports (Correctivos/Preventivos/etc.) no
// cambia de comportamiento.
import * as XLSX from 'xlsx-js-style';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, RadialBarChart, RadialBar, Area, AreaChart
} from 'recharts';

// ═══════════════════════════════════════════════════════════════════
// SUPABASE CONFIG
// Set in Vercel: Settings → Environment Variables
//   VITE_SUPABASE_URL       = https://xxxxx.supabase.co
//   VITE_SUPABASE_ANON_KEY  = eyJ... (your anon public key)
// ═══════════════════════════════════════════════════════════════════
const SUPABASE_URL = 'https://qdcmrirwkfesaqxvhxgh.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_0-OVkFmYhQpomL9UDpA-3g_A5E3qwds';

const supabaseConfigured =
  !SUPABASE_URL.includes('YOUR-PROJECT') &&
  SUPABASE_ANON_KEY.length > 20 &&
  !SUPABASE_ANON_KEY.includes('YOUR-ANON');

// ═══════════════════════════════════════════════════════════════════
// VERSION
// ═══════════════════════════════════════════════════════════════════
const APP_VERSION = 'v3.32';

// ═══════════════════════════════════════════════════════════════════
// PWA / RESPONSIVE HELPERS (PR-1)
// ═══════════════════════════════════════════════════════════════════
// Hook: matchea una media query y reacciona a cambios de tamaño/orientación.
function useMediaQuery(query) {
  const [matches, setMatches] = useState(
    () => typeof window !== 'undefined' && window.matchMedia(query).matches
  );
  useEffect(() => {
    if (typeof window === 'undefined') return;
    const mql = window.matchMedia(query);
    const handler = (e) => setMatches(e.matches);
    setMatches(mql.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, [query]);
  return matches;
}

// true = la app corre como PWA instalada (standalone) Y en un celular.
// Se usa SOLO para decidir la landing inicial (arrancar en Dashboard).
//   - standalone: display-mode standalone (Android/desktop) o navigator.standalone (iOS)
//   - móvil: viewport angosto (< 768px, breakpoint md de Tailwind)
// Nota: si la PWA se instala en una PC, mobile=false → no aplica (arranca en carga).
function isLaunchedAsInstalledMobile() {
  if (typeof window === 'undefined') return false;
  const standalone =
    window.matchMedia?.('(display-mode: standalone)').matches ||
    window.navigator.standalone === true;
  const mobile = window.matchMedia?.('(max-width: 767px)').matches;
  return !!(standalone && mobile);
}

// ═══════════════════════════════════════════════════════════════════
// VERSION GATE (Punto 2 — bloqueo de versiones desactualizadas)
// La app lee `min_version` de la tabla app_config al arrancar y compara
// contra APP_VERSION. Si la local es menor, bloquea la UI y redirige a la
// URL de producción. Si el fetch falla (red, tabla inexistente) → fail-open
// (no bloquea, loguea). Protege de v3.4 en adelante; el cliente v2.1 ya suelto
// no tiene este código y se ataca del lado servidor (Punto 3).
// Formato de versión: 'vMAJOR.MINOR' → entero MAJOR*100 + MINOR (v3.4 → 304).
const PROD_URL = 'https://mantenimiento-ferring.vercel.app';
const parseVersion = (v) => {
  if (!v) return 0;
  const m = String(v).replace(/^v/i, '').match(/(\d+)(?:\.(\d+))?/);
  if (!m) return 0;
  return parseInt(m[1], 10) * 100 + parseInt(m[2] || '0', 10);
};
// ═══════════════════════════════════════════════════════════════════
// V2.9 — ID único para entradas del timeline
// Formato: tl_xxxxxx (6 chars alfanuméricos random).
// Sirve para identificar inequívocamente cada entrada al detectar diffs
// y propagar cambios admin retroactivos a reportes posteriores.
// Las entradas históricas (pre-V2.9) ya tienen id asignado vía
// migración SQL one-shot. Las nuevas se generan acá.
// ═══════════════════════════════════════════════════════════════════
const generateTimelineId = () => {
  const chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  let id = 'tl_';
  for (let i = 0; i < 6; i++) {
    id += chars[Math.floor(Math.random() * chars.length)];
  }
  return id;
};


// ═══════════════════════════════════════════════════════════════════
// #19 (v3.10) — Último avance real de una OT
// Devuelve la entrada de timeline más reciente que NO sea ruido, o null.
// "Ruido": texto vacío o compuesto solo de puntos/espacios (ej. ".", "..",
// "...  "). Es el ruido histórico que v3.9 dejó de generar pero no limpió
// de la base (BACKLOG #27/#19). El timeline llega ordenado por timestamp
// ascendente desde hydrate/dedupCorrective, así que recorremos de atrás
// hacia adelante y devolvemos la primera entrada no-ruido.
// ═══════════════════════════════════════════════════════════════════
const isNoiseAdvance = (text) => /^[.\s]*$/.test((text || ''));

const lastRealAdvance = (timeline) => {
  const tl = timeline || [];
  for (let k = tl.length - 1; k >= 0; k--) {
    if (!isNoiseAdvance(tl[k].text)) return tl[k];
  }
  return null;
};


// ═══════════════════════════════════════════════════════════════════
// V2.6 — MODO ADMINISTRADOR
// Password hardcoded para acceso a edición/eliminación avanzada.
// Sirve como barrera contra clicks accidentales, NO es control de
// acceso real (el password queda visible en GitHub).
// ═══════════════════════════════════════════════════════════════════
const ADMIN_PASSWORD = 'FerringBiomas2026';

// BACKLOG #42 (Fase 1) — Password del rol PLANIFICADOR (pool de OTs).
// MISMO nivel que ADMIN_PASSWORD: está en el repo, NO es control de acceso real,
// es una barrera anti-clicks accidentales. Deliberadamente SEPARADA de la de admin:
// el planificador no debe pasar por una pantalla que puede borrar reportes.
const POOL_PASSWORD = 'Planificador2026';

// ═══════════════════════════════════════════════════════════════════
// BACKLOG #46 (v3.25) — EXTRAS: solicitud y aprobación de horas extras
// ═══════════════════════════════════════════════════════════════════
// MISMO nivel de garantía que ADMIN_PASSWORD y POOL_PASSWORD, y conviene ser
// explícito porque acá el flujo es una APROBACIÓN (algo que después se paga):
//
//   Estas credenciales viajan en el bundle JS que sirve Vercel. Cualquiera con
//   la URL de la app las lee desde DevTools en diez segundos. NO son control de
//   acceso. Son (a) una barrera contra clicks accidentales y (b) ATRIBUCIÓN
//   NOMINAL: quién dice ser el que carga o el que aprueba.
//
// El filtrado "cada encargado ve solo lo suyo" es partición de UI, NO de datos:
// la policy RLS de `horas_extras` es `USING (true)`, igual que el resto de las
// tablas del proyecto, así que la tabla entera se puede leer vía REST con la
// publishable key. El registro sirve como control interno del sector; no
// respalda una aprobación formal ante RRHH.
//
// Decidido a sabiendas con Leo el 2026-08-23 (opción A1 sobre A2/A3). La
// migración a Supabase Auth + RLS por auth.uid() queda en BACKLOG #47.
//
// SUPUESTO A VERIFICAR: el mapeo usuario → nombre se infirió de las direcciones
// de mail que pasó Leo. Si alguno no corresponde, corregir acá: el `nombre` es
// lo que queda escrito en cada solicitud como solicitante o resolutor.

// ═══════════════════════════════════════════════════════════════════
// MODO DE LA APP Y SECTOR DE CASA (#62, v3.29)
// ═══════════════════════════════════════════════════════════════════
// Extras dejó de ser exclusivo de Mantenimiento: Facilities lo usa también,
// con su propio jefe, sus propios Encargado/Supervisor y su propio personal.
// El requisito de Leo fue que Facilities NO vea el reporte de turno.
//
// SON DOS COSAS SEPARADAS, y conviene no mezclarlas:
//
//  1. EL SECTOR particiona los DATOS. Cada fila de `horas_extras` nace con su
//     `sector` y toda consulta filtra por él. Un sector nunca ve las horas del
//     otro, ni en el listado, ni en el dashboard, ni en el aviso de pendientes,
//     ni en el Excel.
//
//  2. EL MODO decide qué se RENDERIZA. En 'full' la app es la de siempre
//     (reporte de turno + los tres roles). En 'extras' arranca directo en el
//     login de horas extras y no existe ninguna otra solapa.
//
// Cómo se decide: por HOSTNAME, no por variable de entorno. El proyecto ya
// tuvo env vars que no se inyectaban al build — es la razón por la que las
// keys de Supabase están hardcodeadas (decisión crítica #1) — y una env var
// que no llega no falla el build: sirve la app en modo equivocado, o sea el
// sector de Facilities viendo el reporte de turno. El hostname es verificable
// en el bundle servido y no depende de la config de Vercel.
//
// La regla es "el hostname CONTIENE el slug", no una igualdad: así cubre de
// una sola vez el dominio de producción del segundo proyecto
// (extras-facilities.vercel.app) y todos sus previews de branch
// (extras-facilities-git-dev-….vercel.app), que es donde se prueba.
//
// ⚠️ EL SEGUNDO PROYECTO DE VERCEL TIENE QUE LLAMARSE `extras-facilities`.
//    Vercel deriva el dominio del nombre del proyecto: si se llama distinto,
//    el hostname no matchea, la app cae al default (Mantenimiento en modo
//    full) y Facilities termina viendo el reporte de turno. No falla nada
//    visible — por eso está escrito acá en mayúsculas.
//
// ⚠️ ESTO NO ES CONTROL DE ACCESO, igual que todo lo demás en este módulo. El
//    código del reporte de turno sigue estando en el bundle del deploy de
//    Facilities aunque no se renderice, y quien tenga la URL de producción la
//    abre y ve todo. Evita el ruido, que es lo que se pidió; la separación
//    real sigue siendo BACKLOG #47.
const EXTRAS_ONLY_SLUG = 'extras-facilities';

const APP_HOSTNAME = (typeof window !== 'undefined' && window.location)
  ? (window.location.hostname || '').toLowerCase()
  : '';

// Escape hatch SOLO para desarrollo local: permite probar el modo Extras sin
// tener que crear el proyecto de Vercel. Acotado a localhost a propósito — un
// query param que funcione en producción convertiría la separación en algo
// que se saltea escribiendo en la barra de direcciones.
const APP_IS_LOCALHOST = APP_HOSTNAME === 'localhost' || APP_HOSTNAME === '127.0.0.1';
const APP_LOCAL_OVERRIDE = (() => {
  if (!APP_IS_LOCALHOST || typeof window === 'undefined') return '';
  try {
    return new URLSearchParams(window.location.search).get('modo') || '';
  } catch { return ''; }
})();

const APP_MODE =
  (APP_HOSTNAME.includes(EXTRAS_ONLY_SLUG) || APP_LOCAL_OVERRIDE === 'extras')
    ? 'extras'
    : 'full';

// El sector de casa lo decide el DEPLOY, no el usuario logueado. Tiene que ser
// así porque el aviso de pendientes se muestra ANTES del login: sin sesión la
// app no sabe quién está del otro lado, pero la URL sí dice de qué sector es.
const APP_SECTOR = APP_MODE === 'extras' ? 'Facilities' : 'Mantenimiento';

// `index.html` es el MISMO archivo para los dos deploys (mismo build): el
// <title> estático queda fijo en "Reporte Diario de Mantenimiento" para los
// dos. Se corrige en runtime solo para el modo Extras, mismo texto que ya usa
// el header (`Horas Extras · {APP_SECTOR}`). En modo full no se toca nada:
// coincide con el <title> por defecto.
if (typeof document !== 'undefined' && APP_MODE === 'extras') {
  document.title = `Horas Extras · ${APP_SECTOR}`;
}

// ═══════════════════════════════════════════════════════════════════
// CATÁLOGOS (matching the Excel template)
// ═══════════════════════════════════════════════════════════════════
const RESPONSABLES = [
  { id: 1, name: 'ALASIA, Juan' },
  { id: 2, name: 'FIORETTI, Luciano' },
  { id: 3, name: 'PARE, Gustavo' }
];

const TECNICOS = [
  { id: 1, name: 'OLIVARES, Victor' },
  { id: 22, name: 'BAGGIO, Christian' },
  { id: 2, name: 'BARRIOS, Martin' },
  { id: 3, name: 'TERAN, Cesar' },
  { id: 4, name: 'VILLASANTE, Eduardo' },
  { id: 5, name: 'LAGOS, Nicolas' },
  { id: 6, name: 'LEMA, Sergio' },
  { id: 7, name: 'CAÑETE, Martin' },
  { id: 8, name: 'FIGUEIRA, Gastón' },
  { id: 9, name: 'MORENO, Jorge' },
  { id: 10, name: 'VERGARA, Antonio' },
  { id: 11, name: 'ECHAZARRETA, Ricardo' },
  { id: 12, name: 'CACERES, Daniel' },
  { id: 13, name: 'VALDEZ, Sergio' },
  { id: 14, name: 'SUAREZ, Alan' },
  { id: 15, name: 'MEDINA, Emanuel' },
  { id: 16, name: 'GOLINO, Santiago' },
  { id: 17, name: 'RIVERO, Cristian' },
  { id: 18, name: 'LEDESMA, Emanuel' },
  { id: 19, name: 'ZAVALA, Emmanuel' },
  { id: 20, name: 'RAMILO, Rodrigo' },
  { id: 21, name: 'YEGROS, Lucas' }
];
const TECNICO_NAMES = TECNICOS.map(t => t.name);
const findTecnicoId = (name) => TECNICOS.find(t => t.name === name)?.id || '';

// Subset who can act as Foguista (Planta de Efluentes y Caldera)
const FOGUISTAS = ['FIGUEIRA, Gastón', 'MORENO, Jorge', 'MEDINA, Emanuel', 'YEGROS, Lucas'];

// ═══════════════════════════════════════════════════════════════════
// PERSONAL HABILITADO PARA HORAS EXTRAS (solapa Extras, #46)
// ═══════════════════════════════════════════════════════════════════
// Universo PROPIO de Extras. NO es TECNICOS y no debe serlo: incluye a los tres
// supervisores (RESPONSABLES) y a personal que reporta horas extras pero NO
// ejecuta OTs de mantenimiento. Ese personal no puede aparecer en "Equipo del
// Turno", ni en Preventivos, ni en el export a Excel del reporte de turno —
// por eso es una lista separada y no un alta en TECNICOS.
//
// Se DERIVA de los catálogos existentes en vez de copiar nombres: un alta o una
// corrección en TECNICOS o RESPONSABLES se propaga sola acá. Lo único
// hardcodeado es el personal que no vive en ningún otro catálogo.
//
// Orden: técnicos (orden del catálogo) → supervisores → personal solo-extras.
// Dedup defensivo por si alguien pasa a estar en dos listas a la vez.
//
// Los que NO son técnicos no tienen `id`: `findTecnicoId` devuelve '' y en
// `horas_extras` la fila queda con `tecnico_id = NULL` y el nombre en
// `tecnico_nombre` (NOT NULL, snapshot), que es la fuente real del registro.
// La tabla ya lo admite — `tecnico_id INTEGER` es nullable. Sin cambios en
// Supabase, sin migración.
const EXTRAS_SOLO_PERSONAL = [
  'INGINO, Matias',
  'NIETO, Ignacio',
  'CUENCA, Sergio',
  'LEZCANO, Nahuel'
];
const EXTRAS_PERSONAL_MANTENIMIENTO = [
  ...TECNICO_NAMES,
  ...RESPONSABLES.map(r => r.name),
  ...EXTRAS_SOLO_PERSONAL
].filter((n, i, arr) => arr.indexOf(n) === i);

// ═══════════════════════════════════════════════════════════════════
// PERSONAL DE FACILITIES (#62, v3.29) — ⚠️ PENDIENTE: ETAPA 2
// ═══════════════════════════════════════════════════════════════════
// Catálogo PROPIO, sin ninguna relación con TECNICOS ni RESPONSABLES:
// Facilities no carga reporte de turno, así que su gente no debe aparecer en
// "Equipo del Turno", ni en Preventivos, ni en el export del reporte. Por eso
// es una lista literal y no un derivado — no hay de qué derivarla.
//
// LO QUE FALTA para que el deploy de Facilities sirva de algo (etapa 2):
//  1. El personal a cargo de cada Encargado/Supervisor. Leo lo va a pasar.
//  2. Los mails y las contraseñas de los cinco usuarios.
//
// Estructura confirmada con Leo el 2026-08-28:
//  · Jefe: ALARCON, Fernando — carga a cualquiera del sector y autoaprueba.
//  · Encargado/Supervisor (4): GROVAS, Leandro · GALLEGO, Sergio ·
//    URUEÑA, Gerardo · AVIO, Raúl. Cada uno con su gente a cargo.
//  · Los cinco pueden tener horas a su propio nombre y se las carga ALARCON,
//    igual que los tres supervisores de Mantenimiento (que no tienen
//    encargado a propósito y los carga el jefe).
//
// ⚠️ `URUEÑA` lleva ñ, igual que `CAÑETE`. Cuando se importe su histórico de
//    RRHH hay que verificarlo por `length()` contra este catálogo: si la ñ se
//    rompe en el pegado, el UNIQUE de `horas_extras_importadas` no lo detecta
//    —serían nombres distintos— y el síntoma aparece meses después como una
//    persona en cero sin ninguna explicación.
//
// Mientras las listas estén vacías el deploy de Facilities levanta pero no
// tiene con qué loguear. Es el estado esperado al terminar la etapa 1: la
// arquitectura entra en producción sin Facilities, así Mantenimiento se
// verifica aislado.
const EXTRAS_ENCARGADOS_FACILITIES = [
  { user: 'sega2@ferring.com', pass: 'Galle2026', nombre: 'GALLEGO, Sergio',   rol: 'encargado' },
  { user: 'legr@ferring.com',  pass: 'Lea2026',   nombre: 'GROVAS, Leandro',   rol: 'encargado' },
  { user: 'geur@ferring.com',  pass: 'geur2026',  nombre: 'URUEÑA, Gerardo',   rol: 'encargado' },
  { user: 'raav@ferring.com',  pass: 'raav2026',  nombre: 'AVIO, Raúl',        rol: 'encargado' },
  { user: 'feal2@ferring.com', pass: 'Fer2026',   nombre: 'ALARCON, Fernando', rol: 'jefe' }
];
// Datos recibidos de Leo el 2026-09-02, y grafías corregidas el mismo día
// contra el dato real (no contra el catálogo original) tras cruzar la
// planilla de RRHH: ARGAÑARAS→ARGAÑARAZ, MAZOLA→MAZZOLA, Jonhatan→Jonathan,
// PÉREZ→PEREZ (sin acento). Mismo criterio que URUEÑA en su momento (regla
// #10 de v3.28): se confirma contra el dato real, no contra cómo lo
// escribió la planilla o el catálogo previo.
// ZEBALLOS, Yonatan (a cargo de Gallego) causó baja de la empresa — sin
// filas en horas_extras (verificado por SQL), así que sacarlo del catálogo
// no deja nada huérfano.
// ALVARADO, Agustín y QUINTANA, Walter Fabian: de licencia médica, no
// estaban en el catálogo original — alta confirmada por Leo. Van al grupo
// compartido de Urueña/Avio porque así vinieron en la planilla de RRHH que
// mandó Leo (misma hoja que las otras 19 personas de ese grupo) — supuesto
// declarado, no confirmado letra por letra con Leo.
const EXTRAS_PERSONAL_FACILITIES = [
  // usuarios (5)
  'ALARCON, Fernando', 'GALLEGO, Sergio', 'GROVAS, Leandro',
  'URUEÑA, Gerardo', 'AVIO, Raúl',
  // a cargo de Gallego (5 — ZEBALLOS de baja)
  'RIOS, Carlos', 'SUAREZ, Juan Francisco',
  'MORLAS, Matias', 'AHUMADA, Cristian', 'LOBOS, Roy',
  // a cargo de Grovas, además de Urueña y Avio (3)
  'MORENO, Matias', 'SANTA ANA, Damian', 'MORAS, Leonardo',
  // a cargo COMPARTIDO de Urueña y Avio (21) — ver nota en aCargo más abajo
  'LUQUEZ, Natanael', 'OLEAS, Fabian', 'AMAYA, Lucas', 'ARGAÑARAZ, Federico',
  'MONZON, Lucas', 'MAZZOLA, Leandro', 'RUGNIA, Elisa', 'LAZO, Mirelys',
  'ACOSTA, Dari', 'ZANONI, Ariel', 'CABRERA, Angel', 'PRADO, Brian',
  'FRENKEL, Franco', 'CORDOBA, Matias', 'GODOY, Jonathan', 'ANADON, Tomas',
  'PEREZ, Brian', 'ZARATE, Federico', 'FERNANDEZ, Gustavo',
  'ALVARADO, Agustín', 'QUINTANA, Walter Fabian'
];

// ═══════════════════════════════════════════════════════════════════
// CONFIGURACIÓN POR SECTOR (#62, v3.29)
// ═══════════════════════════════════════════════════════════════════
// Un sector = un catálogo de personal, un catálogo de usuarios, una asignación
// de gente a cargo, un corte de fuente y sus umbrales. Todo lo que abajo se
// resuelve contra APP_SECTOR sale de acá.
//
// El `label` de cada clave es EXACTAMENTE el string que se guarda en la
// columna `sector` de `horas_extras`, y está replicado en el CHECK
// `horas_extras_sector_chk` de Postgres. AGREGAR UN SECTOR ACÁ EXIGE
// ACTUALIZAR LA CONSTRAINT EN EL MISMO PASO, o el insert lo rechaza la base
// con un error crudo de la API — mismo régimen que EXTRAS_MOTIVO_CATEGORIAS.
//
// 'Facilities' se escribe igual que el label de FAC en SECTORES_OT a
// propósito: es la misma área de la planta. Son cosas distintas igual (el
// sector de una OT no tiene relación con el sector de una hora extra), pero
// dos nombres para la misma área es cómo se empiezan a desalinear los datos.
const EXTRAS_SECTORES = {
  Mantenimiento: {
    label: 'Mantenimiento',
    // Cómo se llama el nivel intermedio EN PANTALLA. En Mantenimiento son
    // "encargados"; en Facilities el grupo se llama "Encargado/Supervisor".
    etiquetaEncargado: 'encargado',
    personal: EXTRAS_PERSONAL_MANTENIMIENTO,
    usuarios: [
      { user: 'jual3@ferring.com', pass: 'juan2026',     nombre: 'ALASIA, Juan',        rol: 'encargado' },
      { user: 'lufi2@ferring.com', pass: 'lufi2',        nombre: 'FIORETTI, Luciano',   rol: 'encargado' },
      { user: 'gtp@ferring.com',   pass: 'gtp2026',      nombre: 'PARE, Gustavo',       rol: 'encargado' },
      { user: 'lgar@ferring.com',  pass: 'Extrasbiomas', nombre: 'GARIBOTTO, Leonardo', rol: 'jefe' }
    ],
    // ─── ASIGNACIÓN DE PERSONAL A ENCARGADOS (#58, v3.28) ───────────
    // Define qué gente tiene a cargo cada encargado. Gobierna DOS cosas:
    //  1. Qué solicitudes ve en el listado — su gente, sin importar quién las
    //     cargó. Antes veía lo que él mismo había cargado; el criterio cambió
    //     a "por persona" en v3.28.
    //  2. A quién puede cargarle extras — solo a su gente. El resto lo carga
    //     el jefe, que ve todo.
    //
    // Las 22 personas de TECNICOS están asignadas, sin solapes ni faltantes
    // (verificado contra el catálogo al definirlo). Las 7 restantes del
    // personal del sector —los tres supervisores y los cuatro de
    // EXTRAS_SOLO_PERSONAL— NO tienen encargado a propósito: las carga y las
    // ve únicamente el jefe.
    //
    // La clave es el `user`, no el nombre: el nombre es dato de presentación
    // y podría cambiar de formato.
    aCargo: {
      'jual3@ferring.com': [
        'OLIVARES, Victor', 'BARRIOS, Martin', 'BAGGIO, Christian', 'VILLASANTE, Eduardo',
        'LAGOS, Nicolas', 'TERAN, Cesar', 'LEMA, Sergio', 'FIGUEIRA, Gastón',
        'MORENO, Jorge', 'CAÑETE, Martin'
      ],
      'lufi2@ferring.com': [
        'ECHAZARRETA, Ricardo', 'VERGARA, Antonio', 'MEDINA, Emanuel',
        'VALDEZ, Sergio', 'SUAREZ, Alan', 'CACERES, Daniel'
      ],
      'gtp@ferring.com': [
        'GOLINO, Santiago', 'RIVERO, Cristian', 'RAMILO, Rodrigo',
        'ZAVALA, Emmanuel', 'LEDESMA, Emanuel', 'YEGROS, Lucas'
      ]
    },
    // Corte entre fuentes. Hasta el período RRHH de agosto 2026 inclusive
    // manda lo IMPORTADO de la planilla de RRHH; desde septiembre 2026 (11/08
    // en adelante) manda lo registrado en la app. Nunca se suman las dos
    // fuentes para un mismo período: eso contaría dos veces las mismas horas.
    corteApp: { anio: 2026, mes: 9 },
    alertaMes: 20,
    alertaAnio: 200
  },

  Facilities: {
    label: 'Facilities',
    etiquetaEncargado: 'Encargado/Supervisor',
    personal: EXTRAS_PERSONAL_FACILITIES,
    usuarios: EXTRAS_ENCARGADOS_FACILITIES,
    // ⚠️ DIFERENCIA DE DISEÑO respecto de Mantenimiento (confirmado con Leo,
    // 2026-09-02): acá la asignación NO es sin solapes. Urueña y Avio
    // comparten literalmente la misma lista de 19 personas — los dos pueden
    // ver y cargarle horas a cualquiera de ellas, y los dos pueden aprobarle
    // la misma solicitud a la misma persona. En Mantenimiento el reparto es
    // sin solapes a propósito (#58) para que no haya ambigüedad de quién
    // aprueba; en Facilities esa ambigüedad queda aceptada tal cual la pidió
    // Leo. ALARCON (jefe) no tiene entrada acá: ve y carga a todo el sector,
    // mismo patrón que el jefe de Mantenimiento.
    aCargo: {
      'sega2@ferring.com': [ // Gallego — ZEBALLOS de baja (2026-09-02)
        'RIOS, Carlos', 'SUAREZ, Juan Francisco',
        'MORLAS, Matias', 'AHUMADA, Cristian', 'LOBOS, Roy'
      ],
      'legr@ferring.com': [ // Grovas
        'URUEÑA, Gerardo', 'AVIO, Raúl',
        'MORENO, Matias', 'SANTA ANA, Damian', 'MORAS, Leonardo'
      ],
      'geur@ferring.com': [ // Urueña — lista compartida con Avio
        'LUQUEZ, Natanael', 'OLEAS, Fabian', 'AMAYA, Lucas', 'ARGAÑARAZ, Federico',
        'MONZON, Lucas', 'MAZZOLA, Leandro', 'RUGNIA, Elisa', 'LAZO, Mirelys',
        'ACOSTA, Dari', 'ZANONI, Ariel', 'CABRERA, Angel', 'PRADO, Brian',
        'FRENKEL, Franco', 'CORDOBA, Matias', 'GODOY, Jonathan', 'ANADON, Tomas',
        'PEREZ, Brian', 'ZARATE, Federico', 'FERNANDEZ, Gustavo',
        'ALVARADO, Agustín', 'QUINTANA, Walter Fabian'
      ],
      'raav@ferring.com': [ // Avio — misma lista que Urueña, a propósito
        'LUQUEZ, Natanael', 'OLEAS, Fabian', 'AMAYA, Lucas', 'ARGAÑARAZ, Federico',
        'MONZON, Lucas', 'MAZZOLA, Leandro', 'RUGNIA, Elisa', 'LAZO, Mirelys',
        'ACOSTA, Dari', 'ZANONI, Ariel', 'CABRERA, Angel', 'PRADO, Brian',
        'FRENKEL, Franco', 'CORDOBA, Matias', 'GODOY, Jonathan', 'ANADON, Tomas',
        'PEREZ, Brian', 'ZARATE, Federico', 'FERNANDEZ, Gustavo',
        'ALVARADO, Agustín', 'QUINTANA, Walter Fabian'
      ]
    },
    // ⚠️ PROVISORIO — ETAPA 2. Facilities no tiene histórico importado, así
    // que hoy TODO su acumulado sale de la app. El corte queda en el pasado
    // para que ningún período quede esperando una fuente importada que no
    // existe. Si Leo importa el histórico de RRHH de Facilities, este corte
    // pasa a ser el mes en que arranca a cargar en la app — y NO es el mismo
    // que el de Mantenimiento.
    corteApp: { anio: 2026, mes: 1 },
    // Leo confirmó que Facilities liquida con el mismo corte 11→10 y usa los
    // mismos umbrales. El corte 11→10 es política interna confirmada
    // verbalmente, no algo verificable desde los archivos.
    alertaMes: 20,
    alertaAnio: 200
  }
};

// ─── Resolución contra el sector de casa ─────────────────────────────
// Todo lo que sigue son los mismos nombres que usaba el resto del módulo
// antes de #62, ahora resueltos contra APP_SECTOR. Los ~20 puntos de uso
// (ExtrasView, ExtrasDashboard, el header) no cambian: leen estas constantes
// como siempre y reciben lo del sector que corresponde.
const EXTRAS_SECTOR_CONF = EXTRAS_SECTORES[APP_SECTOR] || EXTRAS_SECTORES.Mantenimiento;
const EXTRAS_PERSONAL_NAMES = EXTRAS_SECTOR_CONF.personal;
const EXTRAS_USUARIOS = EXTRAS_SECTOR_CONF.usuarios;
const EXTRAS_A_CARGO = EXTRAS_SECTOR_CONF.aCargo;
const EXTRAS_ETIQUETA_ENCARGADO = EXTRAS_SECTOR_CONF.etiquetaEncargado;

// A quién puede CARGARLE extras un encargado: solo su gente.
const extrasPersonalDe = (user) => EXTRAS_A_CARGO[user] || [];

// Qué solicitudes VE un encargado: su gente MÁS las suyas propias. Ver las
// horas extras que a uno le cargaron es razonable aunque no pueda tocarlas;
// editarlas y anularlas sigue exigiendo ser el autor.
const extrasVisiblesDe = (user, nombre) => {
  const base = extrasPersonalDe(user);
  return nombre && !base.includes(nombre) ? [...base, nombre] : base;
};

// Autenticación local contra el catálogo del SECTOR DE CASA (#62). Devuelve la
// sesión (sin la password) o null. La comparación de usuario es
// case-insensitive y trimmed porque el teclado del celular capitaliza la
// primera letra de un mail solo.
//
// Que solo acepte usuarios del sector del deploy es deliberado: el jefe de
// Facilities no puede loguearse en la app de Mantenimiento aunque le llegue la
// URL, y al revés tampoco. Refuerza la separación de ruido sin costo — y como
// Leo confirmó que no necesita el consolidado de los dos sectores, no saca
// nada que se use.
const extrasAuth = (user, pass) => {
  const u = (user || '').trim().toLowerCase();
  const hit = EXTRAS_USUARIOS.find(x => x.user.toLowerCase() === u && x.pass === pass);
  return hit ? { user: hit.user, nombre: hit.nombre, rol: hit.rol, sector: APP_SECTOR } : null;
};

// ═══════════════════════════════════════════════════════════════════
// MOTIVOS DE HORAS EXTRAS (#54, v3.27)
// ═══════════════════════════════════════════════════════════════════
// El motivo dejó de ser texto libre: ahora es categoría cerrada + detalle.
// La razón es el dashboard — agrupar por texto libre obligaba a normalizar
// strings y "Cubrir licencia" y "cubrir licencia médica" caían en grupos
// distintos. Con categoría el agrupamiento es exacto.
//
// El dominio está replicado en Postgres como CHECK constraint
// (`horas_extras_motivo_categoria_chk`). SI SE AGREGA O RENOMBRA UNA
// CATEGORÍA ACÁ, HAY QUE ACTUALIZAR LA CONSTRAINT EN EL MISMO PASO o el
// insert lo rechaza la base con un error crudo de la API.
const EXTRAS_MOTIVO_CATEGORIAS = [
  'Cubrir vacaciones',
  'Cubrir licencia médica',
  'Cubrir vacante',
  'Cubrir feriado',
  'Trabajos específicos',
  'Finalización de trabajos en curso'
  // 'Otros' se eliminó (v3.30): quedó reemplazado por "Trabajos específicos".
  // Sin filas históricas afectadas (verificado: 0 filas con esta categoría).
];

// Categorías que exigen detalle. Las cuatro de cobertura se explican solas;
// forzar texto ahí solo genera "-" y repeticiones de la categoría.
// Replicado en `horas_extras_motivo_detalle_chk`.
const EXTRAS_MOTIVO_REQUIERE_DETALLE = [
  'Trabajos específicos',
  'Finalización de trabajos en curso'
];
const extrasRequiereDetalle = (cat) => EXTRAS_MOTIVO_REQUIERE_DETALLE.includes(cat);

// Motivos REACTIVOS: el trabajo se ejecuta ANTES de que el jefe apruebe, así
// que la aprobación es un acto administrativo posterior al hecho. Los
// planificados se piden y se aprueban antes. La distinción no cambia el
// guardado — cambia cómo se leen las métricas de proceso en el dashboard,
// donde promediar ambos tipos juntos no significa nada.
const EXTRAS_MOTIVOS_REACTIVOS = [
  'Trabajos específicos',
  'Finalización de trabajos en curso'
];
const extrasEsReactivo = (cat) => EXTRAS_MOTIVOS_REACTIVOS.includes(cat);

// ── Detección de solapamiento (#54) ────────────────────────────────
// Una misma persona puede tener varios extras el mismo día y eso es legítimo
// (un trabajo que se extiende se carga como fila aparte y las horas se suman).
// Lo que NO se distingue solo es una duplicación por error de una cadena
// legítima, y en un registro que alimenta liquidación eso se paga dos veces.
// De ahí el aviso: detecta, avisa, y deja pasar si el usuario confirma.
const extrasVentana = (r) => ({
  ini: `${r.fecha}T${(r.hora_inicio || '').slice(0, 5)}`,
  fin: `${r.fecha_fin}T${(r.hora_fin || '').slice(0, 5)}`
});
// Comparación lexicográfica de ISO: válida porque el formato es de ancho fijo.
// Bordes que se tocan (una termina 14:00, la otra arranca 14:00) NO solapan.
const extrasSolapan = (a, b) => a.ini < b.fin && b.ini < a.fin;

// ── ¿Ya se ejecutó? (v3.30, edición post-aprobación) ────────────────
// Determina qué régimen de edición aplica sobre una fila aprobada: libre
// (todavía no pasó, #1) o asimétrico solo-a-la-baja (ya pasó, #2). Se
// deriva de la ventana horaria de la propia fila contra el reloj — sin
// campo nuevo ni marca manual que el encargado se pueda olvidar de tildar.
const extrasYaEjecutada = (r) => {
  const fin = new Date(`${r.fecha_fin}T${(r.hora_fin || '00:00').slice(0, 5)}:00`);
  return fin.getTime() <= Date.now();
};

// ── ¿Se cargó en un día posterior a la ejecución? (v3.31) ──────────
// Señal de control interno para la columna "Cargada": compara el DÍA
// calendario LOCAL en que quedó asentada la solicitud (`created_at`)
// contra el día en que terminó el trabajo (`fecha_fin`). Comparación
// por DÍA, no por hora exacta — cargar el mismo día que terminó el
// turno no cuenta como tardío, aunque sea horas después (decisión
// explícita, 2026-09-01). Comparación lexicográfica de ISO YYYY-MM-DD:
// válida porque el formato es de ancho fijo (mismo patrón que
// extrasSolapan). Se excluye "Finalización de trabajos en curso": por
// definición ese motivo se carga después de terminado el trabajo — no
// es una carga tardía, es la naturaleza del motivo (a diferencia de
// "Trabajos específicos", que es igual de reactivo por diseño pero
// Leo pidió dejarlo afuera de la excepción).
const extrasCargaTardia = (r) => {
  if (r.motivo_categoria === 'Finalización de trabajos en curso') return false;
  if (!r.created_at || !r.fecha_fin) return false;
  const c = new Date(r.created_at);
  if (isNaN(c.getTime())) return false;
  const cargaISO = `${c.getFullYear()}-${String(c.getMonth() + 1).padStart(2, '0')}-${String(c.getDate()).padStart(2, '0')}`;
  return cargaISO > r.fecha_fin;
};

// ── Períodos del dashboard de Extras (#49, v3.27 · reemplazado por #66) ──
// Hasta v3.31 acá vivían EXTRAS_CUATRIS y extrasPeriodo(tipo, offset), que
// armaban el período en mes/cuatrimestre/año CALENDARIO para los KPIs y el
// ranking. Desde #66 esos bloques pasan a período RRHH (ver más abajo), así
// que quedaron sin uso y se retiraron. isoLocalYMD se mantiene: la sigue
// usando extrasRangoRRHH.
const isoLocalYMD = (y, m, d) =>
  `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;

// ── Períodos RRHH (#59, v3.28) ─────────────────────────────────────
// RRHH liquida del 11 de un mes al 10 del siguiente: el período "agosto" va
// del 11/07 al 10/08. NO se usa en los KPIs ni en la evolución, que son de
// mes calendario — se usa solo en la tabla de acumulado, que es la que se le
// comparte a RRHH.
// Devuelve { anio, mes } con mes 1..12, imputando por la fecha de INICIO.
function extrasPeriodoRRHH(fechaISO) {
  const [y, m, d] = (fechaISO || '').split('-').map(Number);
  if (!y || !m || !d) return null;
  // Del 11 en adelante cae en el período del mes siguiente.
  const mm = d >= 11 ? m + 1 : m;
  return mm > 12 ? { anio: y + 1, mes: 1 } : { anio: y, mes: mm };
}

// Rango de fechas calendario que abarca un período RRHH completo:
// mes N del año A va del 11 del mes N-1 al 10 del mes N.
function extrasRangoRRHH(anio, mes) {
  const ini = new Date(anio, mes - 2, 11);
  return {
    desde: isoLocalYMD(ini.getFullYear(), ini.getMonth(), 11),
    hasta: isoLocalYMD(anio, mes - 1, 10)
  };
}

// Corte entre fuentes, POR SECTOR desde #62. En Mantenimiento: hasta el
// período RRHH de agosto 2026 inclusive manda lo IMPORTADO de la planilla de
// RRHH; desde septiembre 2026 (11/08 en adelante) manda lo registrado en la
// app. Nunca se suman las dos fuentes para un mismo período: eso contaría dos
// veces las mismas horas. Cada sector tiene el suyo — el de Facilities no
// puede ser el mismo, porque su histórico todavía no está importado.
const EXTRAS_CORTE_APP = EXTRAS_SECTOR_CONF.corteApp;
const extrasFuenteEsApp = (anio, mes) =>
  anio > EXTRAS_CORTE_APP.anio ||
  (anio === EXTRAS_CORTE_APP.anio && mes >= EXTRAS_CORTE_APP.mes);

// Umbrales de alerta del acumulado (#59), por sector desde #62. Señal visual,
// NO un límite: no bloquean nada, no cambian ningún total y no impiden cargar.
// Si alguna vez tienen que ser un tope real, eso es otra feature y necesita su
// decisión (BACKLOG #61).
const EXTRAS_ALERTA_MES = EXTRAS_SECTOR_CONF.alertaMes;
const EXTRAS_ALERTA_ANIO = EXTRAS_SECTOR_CONF.alertaAnio;

// Duración legible para los tiempos de resolución del dashboard.
const formatDuracion = (horas) => {
  const h = Number(horas) || 0;
  if (h < 1) return `${Math.round(h * 60)} min`;
  if (h < 48) return `${h.toFixed(1).replace('.', ',')} h`;
  return `${(h / 24).toFixed(1).replace('.', ',')} días`;
};

// Tope de filas que trae `listExtras` para el LISTADO. El dashboard no lo usa:
// consulta por rango con su propio tope (ver `listExtrasRango`).
//
// DIMENSIONAMIENTO MEDIDO (2026-08-27), sobre el acumulado real de RRHH de
// ene–ago 2026 para las 22 personas del catálogo que figuran en esa planilla:
// 4.645 h en 8 meses, promedio 581 h/mes, con picos de 864 h (enero) y 852 h
// (julio). A 7–9 h por solicitud eso da 65 a 83 cargas por mes y entre 775 y
// 995 al año — más, porque los 3 supervisores y los 4 de EXTRAS_SOLO_PERSONAL
// no estaban en esa planilla.
//
// O sea que 500 se agota en 6 o 7 meses. Para el LISTADO se acepta: está
// ordenado por fecha descendente y lo que se corta es lo más viejo. Pero los
// contadores de arriba del listado se calculan sobre lo cargado, así que un
// filtro a un rango antiguo va a mostrar de menos. Anotado como BACKLOG #57.
const EXTRAS_LIST_LIMIT = 500;

// Tope del dashboard, por período consultado. Un año completo al ritmo actual
// entra cómodo; el banner avisa si alguna vez no entrara.
const EXTRAS_DASHBOARD_LIMIT = 2000;

// V2.4 — Sectores válidos para N° OT (según SOP 10.3.2)
// Formato: XXX-YYYYY (sector-correlativo de 5 dígitos)
const SECTORES_OT = [
  { code: 'FOA1', label: 'Fracciones O, A y 1' },
  { code: 'FB2',  label: 'Fracciones B y 2' },
  { code: 'RO',   label: 'Recepción de Orina' },
  { code: 'BIO',  label: 'Bioterio' },
  { code: 'DEP',  label: 'Depósito' },
  { code: 'PP',   label: 'Planta Piloto' },
  { code: 'PAD',  label: 'Planta de Agua Deionizada' },
  { code: 'LIM',  label: 'Sector Limpieza' },
  { code: 'EHS',  label: 'Sector EHS' },
  { code: 'MAN',  label: 'Mantenimiento' },
  { code: 'FAC',  label: 'Facilities' },
  { code: 'PTEL', label: 'Planta de Tratamiento de Efluentes Líquidos' },
];
const SECTORES_CODES = SECTORES_OT.map(s => s.code);

// V3.0 — Cutoff para KPIs admin de performance por turno.
// Las OTs creadas o cerradas antes de esta fecha NO se cuentan en:
//   - "OTs dejadas pendientes por turno de origen"
//   - "OTs heredadas cerradas por turno"
// Motivo: limpieza retroactiva de OTs admin pre-20/05 que distorsionaba las métricas.
// Si querés mover el cutoff a futuro, cambiá esta constante y redeploy.
const KPI_CUTOFF_DATE = '2026-05-20';

// Valida formato XXX-YYYYY donde XXX es uno de los sectores y YYYYY exactamente 5 dígitos
const isValidOT = (ot) => {
  if (!ot || typeof ot !== 'string') return false;
  const trimmed = ot.trim();
  const match = trimmed.match(/^([A-Z0-9]+)-(\d{5})$/);
  if (!match) return false;
  return SECTORES_CODES.includes(match[1]);
};

// v3.22 (#39) — Normalización canónica del N° de OT.
// Reduce cualquier string a la forma SECTOR-NNNNN, o devuelve '' si no puede
// identificar el sector contra SECTORES_CODES. NUNCA adivina el sector.
// Casos que resuelve (todos presentes en el histórico):
//   'OT-MAN-00677'  → 'MAN-00677'   (prefijo "OT" que trae el papel)
//   '0T-MAN-00354'  → 'MAN-00354'   (cero en vez de O)
//   'FOA1 01382'    → 'FOA1-01382'  (separador espacio)
//   'OT-FO-A1 01413'→ 'FOA1-01413'  (separadores múltiples)
//   'RO-3255'       → 'RO-03255'    (padding faltante)
//   'FOA1-01395'    → 'FOA1-01395'  (idempotente sobre lo ya canónico)
// Casos que NO resuelve (devuelven '' — se conserva el raw, no se inventa):
//   '02324'  (número pelado, sin sector)
//   'MAN'    (sector sin número)
//   'FACILTY'(sector no reconocible)
const canonOT = (ot) => {
  if (!ot || typeof ot !== 'string') return '';
  // Uppercase → saca prefijo "OT"/"0T" inicial → elimina todo separador.
  const t = ot.trim().toUpperCase()
    .replace(/^[O0]T[^A-Z0-9]*/, '')
    .replace(/[^A-Z0-9]/g, '');
  if (!t) return '';
  // Match del código de sector más largo primero, para que un código que sea
  // prefijo de otro no gane por casualidad si el catálogo cambia a futuro.
  const sector = [...SECTORES_CODES]
    .sort((a, b) => b.length - a.length)
    .find(code => t.startsWith(code));
  if (!sector) return '';
  const num = t.slice(sector.length);
  if (!/^\d{1,5}$/.test(num)) return '';
  return `${sector}-${num.padStart(5, '0')}`;
};

// Clave de identidad de una OT dentro de un reporte (dedup y unicidad).
// Usa la forma canónica cuando se puede resolver; si no, cae al string crudo
// trimmeado (comportamiento previo a v3.22). '' = sin número, no se dedupea.
const otKey = (ot) => canonOT(ot) || (ot || '').trim();

// Parsea "FOA1-01395" → { sector: 'FOA1', numero: '01395' }
// IMPORTANTE: en el camino directo NO padea el número. Solo lo devuelve tal cual.
// El padding se aplica únicamente al perder foco (handleNumeroBlur).
// Esto permite escribir digito por digito sin que se autocompleten ceros.
const parseOT = (ot) => {
  if (!ot) return { sector: '', numero: '' };
  // Camino directo: ya viene estructurado. Es el caso de alguien tipeando
  // (buildOT produce "MAN-1", "MAN-12"…), así que acá NO se padea.
  const match = ot.trim().match(/^([A-Z0-9]+)-(\d{1,5})$/);
  if (match && SECTORES_CODES.includes(match[1])) {
    return { sector: match[1], numero: match[2] };
  }
  // v3.22 (#39) — Camino tolerante: sólo se alcanza cuando el valor NO tiene
  // forma estructurada, es decir cuando viene de dato ya guardado y sucio.
  // Acá sí se padea, porque nadie está tipeando. Permite que el input
  // estructurado edite una OT legacy y que se autocure al primer guardado.
  const canon = canonOT(ot);
  if (canon) {
    const m = canon.match(/^([A-Z0-9]+)-(\d{5})$/);
    if (m) return { sector: m[1], numero: m[2] };
  }
  return { sector: '', numero: '' };
};

// Compone "FOA1" + "1395" → "FOA1-1395" (sin padding mientras escribís).
// El padding a 5 dígitos se hace solo al perder foco en handleNumeroBlur.
const buildOT = (sector, numero) => {
  if (!sector || !numero) return '';
  const num = String(numero).replace(/\D/g, '').slice(0, 5);
  if (!num) return '';
  return `${sector}-${num}`;
};

const TURNOS = ['Mañana', 'Tarde', 'Noche'];
// Orden cronológico de turnos dentro de una jornada. El turno Noche arranca la víspera
// (~23h) y se etiqueta con la fecha del día en que termina, por lo que cronológicamente
// es el PRIMER turno de ese día calendario: Noche → Mañana → Tarde.
// Convención consistente en todo el histórico. Afecta carry-over, dedup y "último estado".
const shiftOrder = (s) => ({ 'Noche': '0', 'Mañana': '1', 'Tarde': '2' }[s] || '9');
const ESTADOS_OT = ['Sin Iniciar', 'En Curso', 'Realizada'];
const ESTADOS_SERVICIO = ['Operativo', 'No Operativo'];
const ESTADOS_PLANTA = ['Operativa', 'No Operativa'];
const NIVELES_CISTERNAS = ['Alto', 'Medio', 'Bajo'];
const ESTADOS_CISTERNAS = ['Ingreso Normal', 'Ingreso Limitado', 'Sin Ingreso'];
const PRIORIDADES = ['Normal', 'Urgente'];
const FRECUENCIAS = ['Diaria', 'Semanal', 'Quincenal', 'Mensual', 'Trimestral', 'Semestral', 'Anual'];

const COMPRESORES = ['SACB 002', 'SACB-C-002', 'SACB-C-003', 'Compresor PTEL'];
const GRUPOS_ELECTROGENOS = ['MANB 001', 'MANB 002', 'MANB 003', 'Grupo Depósito 10'];

// ═══════════════════════════════════════════════════════════════════
// DATE HELPERS (formato dd/mmm/aa para visualización)
// ═══════════════════════════════════════════════════════════════════
const MESES_CORTOS = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];

// Formatea YYYY-MM-DD a dd/mmm/aa (ej: "2026-05-08" -> "08/may/26")
const formatDateShort = (isoDate) => {
  if (!isoDate) return '—';
  const [y, m, d] = isoDate.split('-').map(Number);
  if (!y || !m || !d) return isoDate;
  const dd = String(d).padStart(2, '0');
  const mmm = MESES_CORTOS[m - 1] || '???';
  const yy = String(y).slice(-2);
  return `${dd}/${mmm}/${yy}`;
};

// Formatea YYYY-MM-DD a fecha larga (ej: "viernes, 8 de mayo de 2026")
const formatDateLong = (isoDate) => {
  if (!isoDate) return '';
  const [y, m, d] = isoDate.split('-').map(Number);
  if (!y || !m || !d) return isoDate;
  return new Date(y, m - 1, d).toLocaleDateString('es-AR', {
    weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
  });
};

// #11 (v3.15) — Fecha de HOY en calendario LOCAL (no UTC). Reemplaza el viejo
// new Date().toISOString().slice(0,10), que devolvía la fecha en UTC: en Argentina
// (UTC−3) eso adelantaba el día a partir de las 21:00 hora local, proponiendo "mañana"
// como fecha por defecto en cargas vespertinas (causa raíz de parte de #11) y rompiendo
// el arranque del turno Noche (~23h). Usar SIEMPRE este helper para "hoy" como fecha.
const todayLocalISO = () => {
  const t = new Date();
  return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`;
};

// BACKLOG #42 (Fase 1) — Turno sugerido segun la hora local, con los horarios de
// corte acordados: Noche 23:00-06:00 / Manana 06:00-15:00 / Tarde 15:00-23:00.
// Es solo un DEFAULT editable del alta al pool: el turno_origen lo elige el
// planificador a mano, porque la OT puede haberse emitido en otro momento.
// NO se usa en el flujo de reportes; ahi la fecha/turno siguen viniendo del form.
const currentShiftFromClock = () => {
  const h = new Date().getHours();
  if (h >= 23 || h < 6) return 'Noche';
  if (h < 15) return 'Mañana';
  return 'Tarde';
};

// #11 (v3.15) — true si la fecha es POSTERIOR a hoy (calendario local). Complemento de
// isWithinEditWindow: #9 acota quién edita el PASADO; #11 bloquea el FUTURO. Chequeo aparte
// a propósito (no se mezcla en isWithinEditWindow, que deja pasar futuras por diseño —
// reglas y sujetos distintos: el futuro no aplica a admin, que sí puede precargar).
// Sin fecha o inválida → false (no bloquear reportes a medio armar). El turno Noche carga
// con fecha = hoy local (nunca mañana), así que ">hoy" no le genera falso positivo.
const isFutureDate = (isoDate) => {
  if (!isoDate) return false;
  const [y, m, d] = isoDate.split('-').map(Number);
  if (!y || !m || !d) return false;
  const target = new Date(y, m - 1, d);
  target.setHours(0, 0, 0, 0);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return target > today;
};

// #46 (v3.25) — Helpers de horas extras.
// Suma días a una fecha ISO en calendario LOCAL (mismo criterio que todayLocalISO:
// nada de UTC, que corre el día después de las 21h en Argentina).
const addDaysISO = (isoDate, days) => {
  if (!isoDate) return '';
  const [y, m, d] = isoDate.split('-').map(Number);
  if (!y || !m || !d) return isoDate;
  const t = new Date(y, m - 1, d + days);
  return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`;
};

// REGLA DE MEDIANOCHE (acordada el 2026-08-23): si la hora de fin es menor o
// IGUAL a la de inicio, el extra cruza medianoche y termina al día siguiente.
// El caso típico es el turno Noche (23:00 → 02:00). El "igual" también cae acá
// a propósito: 20:00 → 20:00 no es un extra de cero horas, es uno de 24, y el
// CHECK de la tabla lo rebota como ventana inválida solo si fuese cero.
// La UI muestra un chip "+1 día" cuando esto se dispara, para que no sea sorpresa.
const extrasCruzaMedianoche = (horaInicio, horaFin) =>
  !!horaInicio && !!horaFin && horaFin <= horaInicio;

const extrasFechaFin = (fecha, horaInicio, horaFin) =>
  extrasCruzaMedianoche(horaInicio, horaFin) ? addDaysISO(fecha, 1) : fecha;

// Duración en horas decimales. Es SOLO para previsualizar en el form y para
// totalizar en pantalla: la columna `horas` de la tabla es GENERATED ALWAYS en
// Postgres, así que el número que se guarda lo calcula la base, no el browser.
// Se replica la fórmula acá para que el usuario vea el mismo valor antes de
// guardar; si alguna vez divergen, la base manda.
const extrasHorasCalc = (fecha, horaInicio, horaFin) => {
  if (!fecha || !horaInicio || !horaFin) return 0;
  const fin = extrasFechaFin(fecha, horaInicio, horaFin);
  const ini = new Date(`${fecha}T${horaInicio}`);
  const end = new Date(`${fin}T${horaFin}`);
  if (isNaN(ini) || isNaN(end)) return 0;
  return Math.round(((end - ini) / 3600000) * 100) / 100;
};

// Fecha de auditoría (solicitud, resolución, anulación) en calendario LOCAL:
// "24/ago/26". Se muestra solo la FECHA, sin hora — decidido el 2026-08-24.
// La base sigue guardando el timestamptz completo: tirar la hora en la columna
// sería irreversible y no ahorra nada, mientras que ocultarla en pantalla es
// esto. Si algún día hace falta la hora, está.
//
// La conversión a local NO es cosmética: los timestamptz vuelven de Postgres en
// UTC, y en Argentina (UTC−3) todo lo cargado después de las 21:00 cae al día
// SIGUIENTE si se lee crudo. Una solicitud del lunes 23:30 se mostraría como
// martes. new Date() + getters locales lo resuelve; slice()/replace() sobre el
// string ISO no — ese era el bug de la primera versión de #46.
const formatFechaAudit = (iso) => {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '—';
  const dd = String(d.getDate()).padStart(2, '0');
  const mmm = MESES_CORTOS[d.getMonth()] || '???';
  const yy = String(d.getFullYear()).slice(-2);
  return `${dd}/${mmm}/${yy}`;
};

// Equivalente en 12 h de un "HH:MM" de 24 h, solo como referencia visual
// al lado del selector: "17:30" → "5:30 pm".
const to12h = (hhmm) => {
  if (!hhmm || !hhmm.includes(':')) return '';
  const [h, m] = hhmm.split(':').map(Number);
  if (isNaN(h) || isNaN(m)) return '';
  const suf = h < 12 ? 'am' : 'pm';
  const h12 = h % 12 === 0 ? 12 : h % 12;
  return `${h12}:${String(m).padStart(2, '0')} ${suf}`;
};

// Formato de horas para pantalla: 2.5 → "2,5 h" (coma decimal, es-AR).
const formatHoras = (h) => {
  const n = Number(h) || 0;
  return `${n.toFixed(2).replace(/\.?0+$/, '').replace('.', ',')} h`;
};

// #9 (v3.14) — Ventana de edición para no-admin: solo HOY o AYER (fecha calendario local).
// Reportes más viejos que ayer son read-only para no-admin (corrección retroactiva = solo admin).
// "Hoy + ayer" (no "hoy" literal) cubre el cruce de medianoche del turno Noche, que arranca
// la víspera (~23h) y se etiqueta con el día en que termina. Admin no pasa por acá (siempre edita).
const isWithinEditWindow = (isoDate) => {
  if (!isoDate) return true; // sin fecha (reporte nuevo a medio armar): no bloquear
  const [y, m, d] = isoDate.split('-').map(Number);
  if (!y || !m || !d) return true;
  const target = new Date(y, m - 1, d);
  target.setHours(0, 0, 0, 0);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diffDays = Math.round((today - target) / 86400000);
  // diffDays === 0 → hoy; === 1 → ayer; negativo → futuro (lo deja pasar, #11 se ocupa de futuras)
  return diffDays <= 1;
};

// Day-of-week 0=Domingo, 1=Lunes, ..., 6=Sábado
const dayOfWeek = (isoDate) => {
  const [y, m, d] = isoDate.split('-').map(Number);
  return new Date(y, m - 1, d).getDay();
};

// Suma N días a una fecha YYYY-MM-DD y devuelve nueva fecha YYYY-MM-DD
const addDays = (isoDate, days) => {
  const [y, m, d] = isoDate.split('-').map(Number);
  const dt = new Date(y, m - 1, d);
  dt.setDate(dt.getDate() + days);
  return dt.toISOString().slice(0, 10);
};

// Compara reportes para ordenar de más viejo a más nuevo: fecha+turno
const reportSortKey = (r) => `${r.date}-${shiftOrder(r.shift)}`;

// ═══════════════════════════════════════════════════════════════════
// EMPTY REPORT — V2.0
//   - plantaCaldera: schema NUEVO (PTEL + Caldera + Ablandadores)
//   - preventivosResumen: NUEVO (asignados, realizados, porTecnico)
// ═══════════════════════════════════════════════════════════════════
const emptyReport = () => ({
  date: todayLocalISO(), // #11 (v3.15) — fecha local, no UTC (evita salto a "mañana" después de las 21h AR)
  shift: 'Mañana',
  responsable: '',
  team: [],                    // [string] tech names
  corrective: [],              // V2.4: [{ot, equipoCodigo, task, technicians:[], state, createdInShift, lastModifiedInShift, timeline:[]}]
  preventive: [],              // [{codigoTarea, equipoCodigo, equipoDescripcion, task, comments, otCorrectivaAsociada, technicians:[], frequency}]
  servicios: {
    plantaCaldera: {
      tecnicos: [],            // [string] foguistas del turno (multi-select)
      estado: 'Operativa',
      // PTEL
      caudal: '',              // m³/h
      vacio: '',               // Vacío del equipo
      deltaT: '',              // ΔT entre torres (°C)
      tk1: '',                 // % Nivel TK1
      tk2: '',                 // % Nivel TK2
      tk7: '',                 // % Nivel TK7  (NUEVO en V2.0)
      // Caldera
      conductividadCaldera: '', // mS  (NUEVO en V2.0)
      pHCaldera: '',            //     (NUEVO en V2.0)
      // Agua Ablandadores
      conductividadAblandador: '', // mS (NUEVO en V2.0)
      pHAblandador: ''             //     (NUEVO en V2.0)
    },
    compresores: COMPRESORES.map(c => ({ code: c, state: 'Operativo' })),
    gruposElectrogenos: GRUPOS_ELECTROGENOS.map(g => ({ code: g, state: 'Operativo' })),
    cisternas: { nivel: 'Alto', estado: 'Ingreso Normal' },
    aguaPozo: { cloroPozo3: '', cloroPozo6: '' },
    proveedores: []            // [{provider, task}]
  },
  comments: [],                // [{text, priority}]
  preventivosResumen: {        // NUEVO en V2.0
    asignados: '',
    realizados: '',
    porTecnico: []             // V2.4: [{tecnicos:[], cantidad}] — multi-select por grupo
  }
});

// #26 (v3.7) — Deduplicación intra-reporte de OTs correctivas.
// Un reporte NUNCA debería tener dos entradas con el mismo número de OT, pero
// pasó (ver #25): el flujo permitió recargar una OT ya heredada por el carry-over
// como si fuera nueva, dejando dos versiones (ej. una Realizada + una En Curso).
// Eso rompía a TODOS los consumidores que dedup por número con "el último gana"
// (computePending, stats, dashboard). En vez de parchear cada consumidor, saneamos
// acá: como todo reporte pasa por hydrate al cargarse, los consumidores reciben
// datos ya limpios desde un único punto.
//
// Regla de fusión cuando hay duplicados del mismo número:
//   - Estado ganador: el más "cerrado" (Realizada > En Curso > Sin Iniciar).
//     Ante empate de estado, gana la última posicional (suele ser la más nueva).
//   - createdInShift: el MÁS ANTIGUO entre las versiones (primera aparición real).
//   - lastModifiedInShift, equipoCodigo, task, technicians: los de la ganadora.
//   - timeline: UNIÓN de todas las entradas (dedup por id), ordenadas por timestamp.
//   - La entrada fusionada queda en la POSICIÓN de la primera aparición (no reordena).
// IMPORTANTE: las OTs sin número (ot === '' o ausente) NO se deduplican — cada una
// es una entrada legítima distinta. Esa basura (Clase 2 de #25) se limpia por SQL,
// no acá. hydrate las deja pasar todas tal cual.
const STATE_RANK = { 'Realizada': 3, 'En Curso': 2, 'Sin Iniciar': 1 };

const dedupCorrective = (corrective) => {
  const list = corrective || [];
  const indexByOt = new Map();   // ot# -> índice en `result` donde vive la entrada fusionada
  const result = [];
  list.forEach(c => {
    // v3.22 (#39) — Clave canónica: 'RO-3255' y 'RO-03255' son la MISMA OT.
    // Antes se comparaba el string crudo, así que las variantes de formato
    // convivían como entradas distintas (duplicado invisible).
    const key = otKey(c.ot);
    const entry = { ...c, timeline: c.timeline || [] };
    if (!key) { result.push(entry); return; }   // sin número: no se dedup
    if (!indexByOt.has(key)) {
      indexByOt.set(key, result.length);
      result.push(entry);
      return;
    }
    // Ya existe una entrada con este número: fusionar.
    const idx = indexByOt.get(key);
    const prev = result[idx];
    const prevRank = STATE_RANK[prev.state] || 0;
    const curRank = STATE_RANK[entry.state] || 0;
    // Ganadora: mayor rank; ante empate, la actual (más nueva posicionalmente).
    const winner = curRank >= prevRank ? entry : prev;
    const loser  = winner === entry ? prev : entry;
    // Timeline combinado: unión por id, ordenado por timestamp.
    const seen = new Set();
    const mergedTimeline = [...(prev.timeline || []), ...(entry.timeline || [])]
      .filter(t => {
        const tid = t && t.id ? t.id : JSON.stringify(t);
        if (seen.has(tid)) return false;
        seen.add(tid);
        return true;
      })
      .sort((a, b) => (a.timestamp || '').localeCompare(b.timestamp || ''));
    // createdInShift más antiguo entre las dos versiones (primera aparición real).
    // OJO: el orden lexicográfico de "YYYY-MM-DD-Turno" NO es cronológico (el turno
    // va como texto y "Mañana" < "Noche" alfabéticamente, pero Noche es anterior).
    // Comparamos con shiftOrder: clave = fecha + shiftOrder(turno).
    const sortKeyOfShiftId = (sid) => {
      if (!sid) return '\uffff';   // sin valor: lo mandamos al final
      const i = sid.lastIndexOf('-');
      const fecha = i >= 0 ? sid.slice(0, i) : sid;
      const turno = i >= 0 ? sid.slice(i + 1) : '';
      return `${fecha}-${shiftOrder(turno)}`;
    };
    const createdCandidates = [prev.createdInShift, entry.createdInShift].filter(Boolean);
    const createdInShift = createdCandidates.length
      ? createdCandidates.sort((a, b) => sortKeyOfShiftId(a).localeCompare(sortKeyOfShiftId(b)))[0]
      : winner.createdInShift;
    result[idx] = {
      ...winner,
      createdInShift,
      timeline: mergedTimeline
    };
  });
  return result;
};

// V2.4 — Hidrata un reporte asegurando estructura completa.
// Esto cubre reportes guardados con schemas anteriores (V1.0 a V2.3):
//   - Correctivos sin `timeline` → se inicializa como []
//   - Grupos del resumen con `tecnico` (singular) → se migran a `tecnicos: [...]`
//   - Servicios y subobjetos faltantes se completan con defaults
//   - #26 (v3.7): deduplica OTs correctivas con el mismo número (dato corrupto)
const hydrate = (raw) => {
  const base = emptyReport();
  if (!raw) return base;
  return {
    ...base,
    ...raw,
    // V2.4 — asegurar timeline en cada OT correctiva
    // #26 (v3.7) — y deduplicar OTs con el mismo número (dato corrupto intra-reporte)
    corrective: dedupCorrective(raw.corrective || []),
    servicios: {
      ...base.servicios,
      ...(raw.servicios || {}),
      plantaCaldera: {
        ...base.servicios.plantaCaldera,
        ...(raw.servicios?.plantaCaldera || {})
      },
      cisternas: {
        ...base.servicios.cisternas,
        ...(raw.servicios?.cisternas || {})
      },
      aguaPozo: {
        ...base.servicios.aguaPozo,
        ...(raw.servicios?.aguaPozo || {})
      },
      compresores: raw.servicios?.compresores || base.servicios.compresores,
      gruposElectrogenos: raw.servicios?.gruposElectrogenos || base.servicios.gruposElectrogenos,
      proveedores: raw.servicios?.proveedores || []
    },
    preventivosResumen: {
      ...base.preventivosResumen,
      ...(raw.preventivosResumen || {}),
      // V2.4 — migrar grupos viejos {tecnico, cantidad} → {tecnicos:[tecnico], cantidad}
      porTecnico: ((raw.preventivosResumen?.porTecnico) || []).map(t => {
        if (t.tecnicos && Array.isArray(t.tecnicos)) return t;
        if (t.tecnico) return { tecnicos: [t.tecnico], cantidad: t.cantidad };
        return { tecnicos: [], cantidad: t.cantidad || '' };
      })
    }
  };
};

// ═══════════════════════════════════════════════════════════════════
// STORAGE LAYER (Supabase REST API or local fallback)
// ═══════════════════════════════════════════════════════════════════
const sbHeaders = () => ({
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  'Content-Type': 'application/json'
});

// #62 — Filtro de sector para TODAS las consultas de horas extras.
//
// Va en el QUERY, no en el render, y esa diferencia importa por dos motivos
// que no son cosméticos:
//  1. `listExtras` tiene un tope de filas (EXTRAS_LIST_LIMIT). Si los dos
//     sectores compitieran por esas 500 filas, el tope se agotaría en la
//     mitad de tiempo y cada sector vería de menos SIN NINGÚN AVISO — que es
//     exactamente el modo de falla de BACKLOG #57, adelantado.
//  2. Filtrar en el render significa haber traído las filas del otro sector
//     hasta el browser. No cambia nada de seguridad (la tabla se lee entera
//     vía REST igual, ver #47), pero es payload que nadie mira.
const SECTOR_QS = `sector=eq.${encodeURIComponent(APP_SECTOR)}`;

const storage = {
  configured: supabaseConfigured,

  async list() {
    if (supabaseConfigured) {
      const res = await fetch(
        `${SUPABASE_URL}/rest/v1/reportes?select=*&order=fecha.desc,turno.asc`,
        { headers: sbHeaders() }
      );
      if (!res.ok) throw new Error(`Supabase: ${res.status} ${await res.text()}`);
      const rows = await res.json();
      // V3.5 (#22) — exponemos updated_at junto al data para el chequeo de concurrencia (optimistic locking).
      return rows.map(r => ({ ...r.data, _updatedAt: r.updated_at }));
    } else {
      // Fallback: browser localStorage (per-device, not shared)
      const all = [];
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && k.startsWith('rep:')) {
          try { all.push(JSON.parse(localStorage.getItem(k))); } catch {}
        }
      }
      return all.filter(Boolean).sort((a, b) => (b.date + b.shift).localeCompare(a.date + a.shift));
    }
  },

  async save(report) {
    const id = `${report.date}-${report.shift}`;
    if (supabaseConfigured) {
      const res = await fetch(
        `${SUPABASE_URL}/rest/v1/reportes?on_conflict=id`,
        {
          method: 'POST',
          headers: { ...sbHeaders(), Prefer: 'resolution=merge-duplicates,return=representation' },
          body: JSON.stringify({
            id,
            fecha: report.date,
            turno: report.shift,
            responsable: report.responsable || null,
            data: report,
            updated_at: new Date().toISOString()
          })
        }
      );
      if (!res.ok) throw new Error(`Supabase: ${res.status} ${await res.text()}`);
    } else {
      localStorage.setItem(`rep:${id}`, JSON.stringify(report));
    }
  },

  // BACKLOG #21 (v3.12) — Backup de la versión anterior antes de sobreescribir.
  // Inserta en reportes_historial un snapshot del reporte VIEJO (el que se está
  // por pisar), tal como estaba en la base. Se llama desde doSaveReport DESPUÉS
  // del UPSERT, fail-open: si esto lanza, el caller loguea y sigue (el guardado
  // normal ya ocurrió; el backup es red de seguridad, no debe bloquear la operación).
  //   reporteId      — "YYYY-MM-DD-Turno" del reporte pisado
  //   dataAnterior   — objeto reporte viejo completo (lo que había en la base)
  //   updatedAtAnterior — updated_at que tenía esa versión (puede ser null)
  //   motivo         — 'save_normal' | 'overwrite_turno' | 'concurrency'
  // Solo escribe en Supabase. Si no está configurado (fallback localStorage), no hace
  // historial: el modo local es per-device y no es el escenario que #21 protege.
  async saveHistorial(reporteId, dataAnterior, updatedAtAnterior, motivo) {
    if (!supabaseConfigured) return;
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/reportes_historial`,
      {
        method: 'POST',
        headers: { ...sbHeaders(), Prefer: 'return=minimal' },
        body: JSON.stringify({
          reporte_id: reporteId,
          data_anterior: dataAnterior,
          updated_at_anterior: updatedAtAnterior || null,
          motivo
        })
      }
    );
    if (!res.ok) throw new Error(`Supabase historial: ${res.status} ${await res.text()}`);
  },

  // V2.6 — Eliminar reporte completo (modo admin)
  async delete(date, shift) {
    const id = `${date}-${shift}`;
    if (supabaseConfigured) {
      const res = await fetch(
        `${SUPABASE_URL}/rest/v1/reportes?id=eq.${encodeURIComponent(id)}`,
        { method: 'DELETE', headers: sbHeaders() }
      );
      if (!res.ok) throw new Error(`Supabase: ${res.status} ${await res.text()}`);
    } else {
      localStorage.removeItem(`rep:${id}`);
    }
  },

  // Punto 2 — Lee un valor de la tabla app_config. Devuelve null si no existe,
  // no está configurado Supabase, o falla la consulta (el caller hace fail-open).
  async getConfig(key) {
    if (!supabaseConfigured) return null;
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/app_config?select=value&key=eq.${encodeURIComponent(key)}`,
      { headers: sbHeaders() }
    );
    if (!res.ok) throw new Error(`Supabase app_config: ${res.status} ${await res.text()}`);
    const rows = await res.json();
    return rows.length ? rows[0].value : null;
  },

  // ─────────────────────────────────────────────────────────────────
  // BACKLOG #42 (Fase 1) — POOL DE OTs (tabla `ordenes_pool`)
  // ─────────────────────────────────────────────────────────────────
  // Arquitectura (i), decidida el 2026-07-30: `ordenes_pool` guarda SOLO las
  // altas del planificador. NO se importan las OTs que entran por los reportes.
  // El denominador completo se deriva al vuelo (pool ∪ OTs de `reportes`), asi
  // que no hay sincronizacion posible de romper ni duplicados que reconciliar.
  // NO hay fallback a localStorage a proposito: el pool es inherentemente
  // compartido entre personas y dispositivos; un pool per-device no significa nada.
  async listPool(limit = 200) {
    if (!supabaseConfigured) return [];
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/ordenes_pool?select=*&order=created_at.desc&limit=${limit}`,
      { headers: sbHeaders() }
    );
    if (!res.ok) throw new Error(`Supabase pool: ${res.status} ${await res.text()}`);
    return res.json();
  },

  // Alta. El unique parcial `ordenes_pool_ot_activa_uniq` (ot WHERE anulada_at IS NULL)
  // rebota un numero ya activo con 409/23505; se traduce a un error con code
  // 'DUPLICADA' para que la UI muestre un mensaje util en vez del texto de Postgres.
  async insertPool(row) {
    if (!supabaseConfigured) throw new Error('El pool requiere Supabase configurado');
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/ordenes_pool`,
      {
        method: 'POST',
        headers: { ...sbHeaders(), Prefer: 'return=representation' },
        body: JSON.stringify(row)
      }
    );
    if (!res.ok) {
      const txt = await res.text();
      if (res.status === 409 || txt.includes('ordenes_pool_ot_activa_uniq') || txt.includes('23505')) {
        const e = new Error('DUPLICADA');
        e.code = 'DUPLICADA';
        throw e;
      }
      throw new Error(`Supabase pool: ${res.status} ${txt}`);
    }
    const rows = await res.json();
    return rows[0];
  },

  // Anulacion = soft-delete con motivo obligatorio. NUNCA DELETE: el CHECK
  // `ordenes_pool_anulacion` de la tabla ya impide anular sin motivo, pero se
  // valida tambien en la UI para dar el mensaje antes del round-trip.
  async anularPool(id, motivo) {
    if (!supabaseConfigured) throw new Error('El pool requiere Supabase configurado');
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/ordenes_pool?id=eq.${encodeURIComponent(id)}`,
      {
        method: 'PATCH',
        headers: { ...sbHeaders(), Prefer: 'return=representation' },
        body: JSON.stringify({ anulada_at: new Date().toISOString(), anulada_motivo: motivo })
      }
    );
    if (!res.ok) throw new Error(`Supabase pool: ${res.status} ${await res.text()}`);
    const rows = await res.json();
    return rows[0];
  },

  // ─────────────────────────────────────────────────────────────────
  // BACKLOG #46 (v3.25) — HORAS EXTRAS (tabla `horas_extras`)
  // ─────────────────────────────────────────────────────────────────
  // Igual que el pool: NO hay fallback a localStorage. Un registro de horas
  // extras per-device no significa nada — el encargado solicita en una máquina
  // y el jefe aprueba en otra. Sin Supabase, la solapa se muestra deshabilitada.
  //
  // La tabla NO tiene GRANT de DELETE (ver el DDL): la corrección es siempre
  // soft-delete con motivo. Si alguna vez hace falta purgar de verdad, se hace
  // desde el SQL Editor como `postgres`, con backup previo.
  async listExtras(limit = EXTRAS_LIST_LIMIT) {
    if (!supabaseConfigured) return [];
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/horas_extras?select=*&${SECTOR_QS}&order=fecha.desc,hora_inicio.desc&limit=${limit}`,
      { headers: sbHeaders() }
    );
    if (!res.ok) throw new Error(`Supabase extras: ${res.status} ${await res.text()}`);
    return res.json();
  },

  // #49 — Consulta acotada al período que está mirando el dashboard.
  // El dashboard NO calcula sobre `extras` (el listado): con el volumen real
  // del sector — ~65 a 83 solicitudes por mes, ~1000 al año — el tope del
  // listado se agota en 6 o 7 meses y la vista anual quedaría truncada.
  // Trayendo solo el rango, el payload es proporcional a lo que se mira y la
  // vista anual pesa lo mismo dentro de cinco años que hoy.
  //
  // Se imputa por `fecha` (inicio del extra), coherente con el resto del
  // módulo: una solicitud que cruza medianoche cuenta entera en el período
  // donde arrancó.
  async listExtrasRango(desde, hasta, limit = EXTRAS_DASHBOARD_LIMIT) {
    if (!supabaseConfigured) return [];
    const qs = [
      'select=*',
      SECTOR_QS,
      `fecha=gte.${encodeURIComponent(desde)}`,
      `fecha=lte.${encodeURIComponent(hasta)}`,
      'order=fecha.desc,hora_inicio.desc',
      `limit=${limit}`
    ].join('&');
    const res = await fetch(`${SUPABASE_URL}/rest/v1/horas_extras?${qs}`, { headers: sbHeaders() });
    if (!res.ok) throw new Error(`Supabase extras rango: ${res.status} ${await res.text()}`);
    return res.json();
  },

  // #49 — Serie del gráfico de evolución: 12 meses, proyección MÍNIMA.
  // Solo 4 columnas en vez de las 22 de `select=*`. Un año entero con esta
  // proyección pesa menos que un mes con la fila completa, así que traer 12
  // meses para el gráfico sale más barato que traer el período para los KPIs.
  async listExtrasSerie(desde, hasta, limit = EXTRAS_DASHBOARD_LIMIT) {
    if (!supabaseConfigured) return [];
    const qs = [
      'select=fecha,horas,estado,anulada_at,tecnico_nombre',
      SECTOR_QS,
      `fecha=gte.${encodeURIComponent(desde)}`,
      `fecha=lte.${encodeURIComponent(hasta)}`,
      'order=fecha.asc',
      `limit=${limit}`
    ].join('&');
    const res = await fetch(`${SUPABASE_URL}/rest/v1/horas_extras?${qs}`, { headers: sbHeaders() });
    if (!res.ok) throw new Error(`Supabase extras serie: ${res.status} ${await res.text()}`);
    return res.json();
  },

  // #59 — Horas extras HISTÓRICAS importadas de la planilla de RRHH.
  // Tabla `horas_extras_importadas`, de solo lectura para la app: la carga se
  // hace por SQL. No tiene GRANT de INSERT/UPDATE/DELETE para `anon` a
  // propósito — un dato ya liquidado no se toca desde la UI.
  // Los totales son MENSUALES por persona en período RRHH (11→10), sin
  // detalle diario: por eso solo alimentan la tabla de acumulado y no los
  // bloques de mes calendario.
  async listImportadas(anio) {
    if (!supabaseConfigured) return [];
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/horas_extras_importadas?select=persona,anio,mes,horas,origen&${SECTOR_QS}&anio=eq.${anio}&limit=2000`,
      { headers: sbHeaders() }
    );
    if (!res.ok) throw new Error(`Supabase importadas: ${res.status} ${await res.text()}`);
    return res.json();
  },

  // #50 (v3.26) — Sonda para el aviso del botón de rol del header.
  // Corre para TODOS los usuarios al arrancar la app y en cada refresh, entren
  // o no a Extras: es el único modo de que el jefe vea que hay algo pendiente
  // sin loguearse. Por eso tiene que ser el query más barato posible.
  //
  // Devuelve BOOLEAN, no un conteo, por dos motivos:
  //  1. La UI decidida no muestra número (solo el color), así que contar sería
  //     traer un dato que nadie mira.
  //  2. `limit=1` + `select=id` trae como máximo una fila de un entero. La
  //     alternativa (`HEAD` con `Prefer: count=exact`) cuesta lo mismo en
  //     round-trip pero obliga a parsear el header `Content-Range` y depende de
  //     que CORS lo exponga. Si algún día hace falta el número, es cambiar
  //     `limit=1` por el count: la firma del método es lo único que se toca.
  //
  // Población: `estado = 'pendiente'` Y `anulada_at IS NULL`. Las anuladas no
  // suman NUNCA — mismo criterio que los contadores del listado de Extras. Las
  // de fecha futura sí cuentan: son el caso normal (se solicita antes de
  // trabajar), no una anomalía a filtrar.
  //
  // #62 — Filtrada por SECTOR DE CASA. Sin esto el aviso sería cruzado: una
  // pendiente de Facilities le teñiría el botón al jefe de Mantenimiento y al
  // revés. El sector sale del deploy y no de la sesión, que es justamente lo
  // que hace posible que esto funcione ANTES del login.
  //
  // FAIL-SILENT a propósito: si esto falla no debe ensuciar `connError` ni
  // bloquear nada. Es un aviso, no un dato operativo. Ante la duda, false:
  // preferimos no avisar de más que teñir el header por un error de red.
  async hasExtrasPendientes() {
    if (!supabaseConfigured) return false;
    // v3.30 (#64): 'modificada' también necesita que el jefe la mire — es un
    // ajuste a la baja esperando confirmación, mismo criterio que pendiente.
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/horas_extras?select=id&${SECTOR_QS}&estado=in.(pendiente,modificada)&anulada_at=is.null&limit=1`,
      { headers: sbHeaders() }
    );
    if (!res.ok) throw new Error(`Supabase extras (sonda): ${res.status} ${await res.text()}`);
    const rows = await res.json();
    return rows.length > 0;
  },

  async insertExtra(row) {
    if (!supabaseConfigured) throw new Error('Extras requiere Supabase configurado');
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/horas_extras`,
      {
        method: 'POST',
        headers: { ...sbHeaders(), Prefer: 'return=representation' },
        body: JSON.stringify(row)
      }
    );
    if (!res.ok) throw new Error(`Supabase extras: ${res.status} ${await res.text()}`);
    const rows = await res.json();
    return rows[0];
  },

  // PATCH genérico. Lo usan los tres caminos de escritura posteriores al alta
  // (editar mientras está pendiente, resolver, anular). `updated_at` se pisa
  // siempre acá y no en cada call site, para que no se pueda olvidar.
  async updateExtra(id, patch) {
    if (!supabaseConfigured) throw new Error('Extras requiere Supabase configurado');
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/horas_extras?id=eq.${encodeURIComponent(id)}`,
      {
        method: 'PATCH',
        headers: { ...sbHeaders(), Prefer: 'return=representation' },
        body: JSON.stringify({ ...patch, updated_at: new Date().toISOString() })
      }
    );
    if (!res.ok) throw new Error(`Supabase extras: ${res.status} ${await res.text()}`);
    const rows = await res.json();
    if (!rows.length) {
      // 0 filas devueltas con 200 = la fila existe pero el PATCH no matcheó el
      // filtro, o la borró otro. No es un caso esperado; se avisa fuerte.
      throw new Error('La solicitud ya no existe o fue modificada por otra sesión');
    }
    return rows[0];
  }
};

// ═══════════════════════════════════════════════════════════════════
// DRAFT STORE — autoguardado local de borrador (BACKLOG #7, v3.6)
// ═══════════════════════════════════════════════════════════════════
// Red de seguridad LOCAL y POR DISPOSITIVO contra la pérdida de un borrador
// en curso (refresh / cierre de pestaña / caída antes de apretar "Guardar").
// NO reemplaza al guardado en Supabase ni al optimistic locking de #22:
// - #22 previene que dos sesiones se pisen entre sí (server-side, por timestamp).
// - #7 evita que UNA sesión pierda lo tipeado si se cae antes de guardar (local).
// Es localStorage puro: no toca Supabase, no es registro auditable, es UX.
// Una key por turno: `draft:YYYY-MM-DD-Turno`. Expira a 48h.
// Todo fail-silent (try/catch): si localStorage no está disponible — modo
// privado de iOS, cuota llena — no rompe el flujo de carga.
const DRAFT_PREFIX = 'draft:';
const DRAFT_TTL_MS = 48 * 60 * 60 * 1000; // 48 horas

const draftStore = {
  key(id) { return `${DRAFT_PREFIX}${id}`; },

  // Guarda el borrador del reporte bajo su id (date-shift).
  save(id, report) {
    try {
      localStorage.setItem(this.key(id), JSON.stringify({
        report,
        savedAt: new Date().toISOString(),
        version: APP_VERSION
      }));
    } catch { /* fail-silent: sin localStorage, no hay red local pero la app sigue */ }
  },

  // Devuelve { report, savedAt, version } o null si no existe / está corrupto.
  load(id) {
    try {
      const raw = localStorage.getItem(this.key(id));
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (!parsed || !parsed.report) return null;
      return parsed;
    } catch { return null; }
  },

  // Borra el borrador de un turno. Se llama tras un save exitoso a Supabase.
  clear(id) {
    try { localStorage.removeItem(this.key(id)); } catch { /* fail-silent */ }
  },

  // Borra borradores de más de 48h. Se llama una vez al arrancar la app.
  // Evita que se acumulen borradores viejos de turnos ya cerrados.
  purgeOld() {
    try {
      const now = Date.now();
      const toRemove = [];
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (!k || !k.startsWith(DRAFT_PREFIX)) continue;
        try {
          const parsed = JSON.parse(localStorage.getItem(k));
          const t = parsed?.savedAt ? Date.parse(parsed.savedAt) : NaN;
          if (!Number.isFinite(t) || (now - t) > DRAFT_TTL_MS) toRemove.push(k);
        } catch { toRemove.push(k); } // borrador corrupto → descartar
      }
      toRemove.forEach(k => localStorage.removeItem(k));
    } catch { /* fail-silent */ }
  }
};

// ═══════════════════════════════════════════════════════════════════
// hasUserWork — ¿el form tiene trabajo REAL del usuario? (BACKLOG #7, v3.6)
// ═══════════════════════════════════════════════════════════════════
// El form casi nunca está 100% vacío: el carry-over trae correctivos pendientes
// automáticamente (decisión crítica #4). Necesitamos distinguir "form con
// trabajo del usuario" de "form recién abierto solo con carry-over", para no
// ofrecer recuperar un borrador que en realidad nadie tocó.
//
// Señal de trabajo real (cualquiera alcanza):
//   - responsable / equipo del turno cargados
//   - preventivos (array detallado) o comentarios cargados
//   - resumen de preventivos con datos (asignados/realizados/porTecnico)
//   - servicios tocados respecto del default (caldera, agua de pozo, proveedores)
//   - algún correctivo TOCADO en este turno: lastModifiedInShift === date-shift
//     (las pendientes intactas del carry-over NO marcan esto — misma señal que
//     usa el Dashboard para mostrar solo lo del turno, V2.3), o con una entrada
//     de timeline cuyo shiftKey sea el turno actual.
const hasUserWork = (report) => {
  if (!report) return false;
  const id = `${report.date}-${report.shift}`;
  const base = emptyReport();

  if ((report.responsable || '').trim()) return true;
  if ((report.team || []).length > 0) return true;
  if ((report.preventive || []).length > 0) return true;
  if ((report.comments || []).length > 0) return true;

  const pr = report.preventivosResumen || {};
  if ((pr.asignados ?? '') !== '' || (pr.realizados ?? '') !== '') return true;
  if ((pr.porTecnico || []).length > 0) return true;

  const s = report.servicios || {};
  const pc = s.plantaCaldera || {};
  const bpc = base.servicios.plantaCaldera;
  // PTEL + caldera + ablandadores: cualquier medición cargada, foguistas, o estado != default
  const calderaFields = ['caudal','vacio','deltaT','tk1','tk2','tk7',
    'conductividadCaldera','pHCaldera','conductividadAblandador','pHAblandador'];
  if (calderaFields.some(f => (pc[f] ?? '') !== '')) return true;
  if ((pc.tecnicos || []).length > 0) return true;
  if ((pc.estado ?? bpc.estado) !== bpc.estado) return true;
  const ap = s.aguaPozo || {};
  if ((ap.cloroPozo3 ?? '') !== '' || (ap.cloroPozo6 ?? '') !== '') return true;
  if ((s.proveedores || []).length > 0) return true;
  // Compresores / grupos / cisternas con estado distinto del default
  if ((s.compresores || []).some(c => c.state && c.state !== 'Operativo')) return true;
  if ((s.gruposElectrogenos || []).some(g => g.state && g.state !== 'Operativo')) return true;
  const cis = s.cisternas || {};
  if ((cis.nivel ?? base.servicios.cisternas.nivel) !== base.servicios.cisternas.nivel) return true;
  if ((cis.estado ?? base.servicios.cisternas.estado) !== base.servicios.cisternas.estado) return true;

  // Correctivos tocados en ESTE turno (no los del carry-over intactos)
  const corrTouched = (report.corrective || []).some(c =>
    c.lastModifiedInShift === id ||
    (c.timeline || []).some(t => t.shiftKey === id)
  );
  if (corrTouched) return true;

  return false;
};

// reportsEqual — compara dos reportes por contenido, ignorando _updatedAt
// (metadato de concurrencia que no es parte del reporte en sí). Se usa para
// decidir si el form tiene cambios sin guardar respecto del último snapshot
// guardado. (BACKLOG #7, v3.6)
const reportsEqual = (a, b) => {
  if (!a || !b) return false;
  const strip = (r) => { const { _updatedAt, ...rest } = r; return rest; };
  try {
    return JSON.stringify(strip(a)) === JSON.stringify(strip(b));
  } catch { return false; }
};

// ═══════════════════════════════════════════════════════════════════
// UI PRIMITIVES
// ═══════════════════════════════════════════════════════════════════
const Card = ({ children, className = '', ...rest }) => (
  <div className={`bg-white border border-slate-200 rounded-xl shadow-sm ${className}`} {...rest}>{children}</div>
);

const SectionTitle = ({ icon: Icon, children, accent = 'sky' }) => {
  const accentMap = {
    sky: 'bg-sky-500', emerald: 'bg-emerald-500', orange: 'bg-orange-500',
    violet: 'bg-violet-500', indigo: 'bg-indigo-500', cyan: 'bg-cyan-500',
    slate: 'bg-slate-500', amber: 'bg-amber-500'
  };
  return (
    <div className="flex items-center gap-2.5 mb-4">
      <div className={`w-1 h-6 rounded-full ${accentMap[accent]}`} />
      <h2 className="text-[15px] font-semibold text-slate-800 tracking-tight uppercase inline-flex items-center gap-2">
        <Icon className="w-4 h-4 text-slate-500" />{children}
      </h2>
    </div>
  );
};

const Field = ({ label, children, className = '' }) => (
  <label className={`flex flex-col gap-1.5 ${className}`}>
    <span className="text-xs font-medium text-slate-600 uppercase tracking-wide">{label}</span>
    {children}
  </label>
);

const inputCls = "w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500/40 focus:border-sky-500 transition";
const buttonCls = "inline-flex items-center gap-1.5 px-3 py-2 text-sm font-medium rounded-lg transition";
const EmptyHint = ({ children }) => (
  <div className="text-sm text-slate-400 italic py-3 text-center border border-dashed border-slate-200 rounded-lg">{children}</div>
);

// ═══════════════════════════════════════════════════════════════════
// MULTI-SELECT
// ═══════════════════════════════════════════════════════════════════
function MultiSelect({ options, value = [], onChange, placeholder = 'Seleccionar…' }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState('');
  const wrapRef = useRef(null);

  useEffect(() => {
    const onClick = (e) => { if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false); };
    document.addEventListener('mousedown', onClick);
    return () => document.removeEventListener('mousedown', onClick);
  }, []);

  const toggle = (opt) => onChange(value.includes(opt) ? value.filter(v => v !== opt) : [...value, opt]);
  const filtered = options.filter(o => o.toLowerCase().includes(search.toLowerCase()));

  return (
    <div ref={wrapRef} className="relative">
      <button type="button" onClick={() => setOpen(o => !o)}
        className={`${inputCls} text-left flex items-center justify-between min-h-[38px] cursor-pointer`}>
        <div className="flex flex-wrap gap-1 flex-1">
          {value.length === 0 && <span className="text-slate-400">{placeholder}</span>}
          {value.map(v => (
            <span key={v} onClick={(e) => { e.stopPropagation(); toggle(v); }}
              className="inline-flex items-center gap-1 px-2 py-0.5 bg-sky-100 text-sky-800 rounded text-xs font-medium hover:bg-sky-200">
              {v}<X className="w-3 h-3" />
            </span>
          ))}
        </div>
        <ChevronDown className={`w-4 h-4 text-slate-400 transition ${open ? 'rotate-180' : ''}`} />
      </button>
      {open && (
        <div className="absolute z-30 mt-1 w-full bg-white border border-slate-200 rounded-lg shadow-xl max-h-72 overflow-hidden flex flex-col">
          <div className="p-2 border-b border-slate-100 sticky top-0 bg-white">
            <div className="relative">
              <Search className="w-4 h-4 absolute left-2 top-1/2 -translate-y-1/2 text-slate-400" />
              <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Buscar…"
                className="w-full pl-8 pr-2 py-1.5 text-sm border border-slate-200 rounded focus:outline-none focus:border-sky-400" />
            </div>
          </div>
          <div className="overflow-y-auto">
            {filtered.length === 0 && <div className="px-3 py-4 text-sm text-slate-400 text-center">Sin resultados</div>}
            {filtered.map(o => (
              <label key={o} className="flex items-center gap-2 px-3 py-2 text-sm hover:bg-sky-50 cursor-pointer">
                <input type="checkbox" checked={value.includes(o)} onChange={() => toggle(o)} className="rounded text-sky-600" />
                <span>{o}</span>
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// STATE PILL
// ═══════════════════════════════════════════════════════════════════
const stateColor = (state) => {
  switch (state) {
    case 'Realizada': case 'Operativo': case 'Operativa': case 'Ingreso Normal': case 'Alto':
      return 'text-emerald-700 bg-emerald-50 ring-emerald-200';
    case 'En Curso': case 'Ingreso Limitado': case 'Medio':
      return 'text-amber-700 bg-amber-50 ring-amber-200';
    case 'Sin Iniciar': case 'No Operativo': case 'No Operativa': case 'Sin Ingreso': case 'Bajo':
      return 'text-red-700 bg-red-50 ring-red-200';
    default:
      return 'text-slate-700 bg-slate-50 ring-slate-200';
  }
};
const StatePill = ({ state }) => (
  <span className={`text-xs px-2 py-0.5 rounded-full ring-1 font-medium ${stateColor(state)}`}>{state}</span>
);

// ═══════════════════════════════════════════════════════════════════
// OT NUMBER INPUT — V2.4, reformado en v3.22 (#39)
// Input compuesto: dropdown sector + input numérico de 5 dígitos.
// Auto-completa con ceros a la izquierda al perder foco.
// v3.22: ya NO existe un modo "legacy" de texto libre. Todas las OTs se editan
// con el input estructurado, que por construcción no puede producir un número
// no canónico. Los formatos sucios del histórico se recuperan vía parseOT
// tolerante; los irrecuperables se muestran al lado como referencia read-only.
// ═══════════════════════════════════════════════════════════════════
function OTNumberInput({ value, onChange, hasError, disabled }) {
  const parsed = parseOT(value);
  const [sector, setSector] = useState(parsed.sector);
  const [numero, setNumero] = useState(parsed.numero);

  // Re-parsear si cambia value desde afuera
  useEffect(() => {
    const p = parseOT(value);
    setSector(p.sector);
    setNumero(p.numero);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value]);

  // v3.22 (#39) — Se ELIMINÓ la rama de texto libre para OTs "legacy".
  // Era un vector de corrupción: una OT heredada con número no canónico recibía
  // un <input> sin ningún filtro (onChange={e => onChange(e.target.value)}), lo
  // que la mantenía no canónica para siempre — y por lo tanto en texto libre el
  // turno siguiente, y el siguiente. Cada turno que la tocaba podía empeorarla.
  // Ese loop es la causa raíz de los 191 números mal formados del histórico.
  // Ahora TODAS las OTs se editan con el input estructurado (select de sector +
  // input de solo dígitos), que no puede producir un valor no canónico.
  //
  // Cuando el valor guardado no se puede resolver a un sector conocido
  // (ej. '02324' pelado, 'FACILTY'), el input estructurado queda vacío y el
  // string original se muestra al lado como referencia READ-ONLY, para que la
  // información no se pierda de pantalla y nadie borre un número que no puede leer.
  const unresolved = !!value && !parsed.sector;

  const handleSectorChange = (e) => {
    const newSector = e.target.value;
    setSector(newSector);
    onChange(buildOT(newSector, numero));
  };

  const handleNumeroChange = (e) => {
    // Solo dígitos, máx 5
    const cleaned = e.target.value.replace(/\D/g, '').slice(0, 5);
    setNumero(cleaned);
    onChange(buildOT(sector, cleaned));
  };

  const handleNumeroBlur = () => {
    if (numero && numero.length > 0 && numero.length < 5) {
      const padded = numero.padStart(5, '0');
      setNumero(padded);
      onChange(buildOT(sector, padded));
    }
  };

  return (
    <div className="flex flex-col gap-1">
      {/* v3.22 (#39) — Valor original irrecuperable: se muestra como referencia
          read-only para no perder el dato. No es editable a propósito. */}
      {unresolved && (
        <div className="flex items-center gap-1" title={`Formato no reconocido — valor original: ${value}`}>
          <span className="text-[9px] px-1 py-0.5 bg-amber-100 text-amber-700 rounded font-bold shrink-0">L</span>
          <span className="text-[10px] text-amber-700 num truncate">{value}</span>
        </div>
      )}
      <div className="flex items-center gap-1">
        <select
          className={`px-1.5 py-2 text-xs bg-white border rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500/40 transition num font-semibold ${(hasError || unresolved) && !sector ? 'border-red-400' : 'border-slate-300'}`}
          value={sector}
          onChange={handleSectorChange}
          disabled={disabled}
          style={{ width: '70px' }}
        >
          <option value="">—</option>
          {SECTORES_OT.map(s => (
            <option key={s.code} value={s.code} title={s.label}>{s.code}</option>
          ))}
        </select>
        <span className="text-slate-400 text-sm font-bold">-</span>
        <input
          type="text"
          inputMode="numeric"
          pattern="[0-9]*"
          maxLength={5}
          className={`${inputCls} num flex-1 ${hasError && (numero.length < 5) ? 'border-red-400' : ''}`}
          value={numero}
          onChange={handleNumeroChange}
          onBlur={handleNumeroBlur}
          placeholder="00000"
          disabled={disabled}
        />
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════
export default function App() {
  const [tab, setTab] = useState('form');
  const [report, setReport] = useState(emptyReport());
  const [history, setHistory] = useState([]);
  const [loading, setLoading] = useState(true);
  // Punto 2 — Version gate. true = versión local desactualizada → bloquear UI.
  const [versionBlocked, setVersionBlocked] = useState(false);
  const [saveMsg, setSaveMsg] = useState('');
  const [saving, setSaving] = useState(false);
  const [connError, setConnError] = useState('');

  // V2.5 — Estado para el modal de confirmación al guardar con OTs/preventivos vacíos.
  // Cuando hay vacíos detectados, se abre el modal y se posterga el guardado real
  // hasta que el usuario confirme.
  const [emptyConfirm, setEmptyConfirm] = useState(null); // null | { emptyCorr, emptyPrev, cleanedReport }

  // V2.6 — Modo administrador (no persistente, solo en sesión actual del navegador)
  const [adminMode, setAdminMode] = useState(false);

  // V2.9 — Snapshot del reporte original al momento de cargarlo.
  // Se captura siempre que se abre un reporte histórico (vía setDateShift o
  // al montar el form con un reporte existente). Solo se USA cuando adminMode
  // está activo al momento de guardar, para detectar diffs y propagar cambios
  // a reportes posteriores. null = no hay snapshot (reporte nuevo o ya limpio).
  const [originalReport, setOriginalReport] = useState(null);
  const [adminLoginOpen, setAdminLoginOpen] = useState(false);

  // BACKLOG #42 (Fase 1) — Rol planificador y pool de OTs.
  // poolMode es independiente de adminMode: son dos roles distintos y no se implican.
  const [poolMode, setPoolMode] = useState(false);
  const [poolLoginOpen, setPoolLoginOpen] = useState(false);
  const [pool, setPool] = useState([]);
  const [poolLoading, setPoolLoading] = useState(false);
  const [poolError, setPoolError] = useState('');

  // BACKLOG #46 (v3.25) — Sesión de Extras. Independiente de adminMode y de
  // poolMode: son tres roles distintos y ninguno implica a los otros.
  // extrasUser = null | { user, nombre, rol }. No persiste entre recargas a
  // propósito (mismo criterio que los otros dos roles).
  const [extrasUser, setExtrasUser] = useState(null);
  const [extrasLoginOpen, setExtrasLoginOpen] = useState(false);
  const [extras, setExtras] = useState([]);
  const [extrasLoading, setExtrasLoading] = useState(false);
  const [extrasError, setExtrasError] = useState('');
  // BACKLOG #50 (v3.26) — ¿Hay solicitudes de horas extras sin resolver?
  // Independiente de `extras`: esto se sabe SIN estar logueado al rol, que es
  // justamente el punto. `extras` solo tiene datos después del login.
  const [extrasPendientes, setExtrasPendientes] = useState(false);
  // V2.6 — Confirmación de eliminación de reporte completo
  const [deleteReportConfirm, setDeleteReportConfirm] = useState(null); // null | { date, shift, source }

  // V2.8 — Conflictos de OT cerrada detectados al guardar.
  // Cuando se intenta guardar un reporte que contiene una OT en estado
  // Sin Iniciar / En Curso pero esa misma OT ya está Realizada en otro reporte
  // posterior en Supabase, se abre este modal para que el usuario decida qué hacer.
  // Cada conflicto: { otNumber, otTask, closedIn: {date, shift, responsable}, currentStateInForm }
  // Estructura: null | { conflicts: [...], onResolve: (decisions) => void }
  const [closedConflicts, setClosedConflicts] = useState(null);

  // v3.16 — Índices de OTs con error de validación. Se setean cuando validateReport
  // falla, se usan en FormView para resaltar las tarjetas problemáticas en rojo.
  // Se limpian al guardar exitosamente o cuando el usuario edita cualquier OT.
  const [otErrorIndices, setOtErrorIndices] = useState(new Set());
  // Tipo de error: 'avance' | 'otro'. Cuando es 'avance', se resalta además la
  // sección de Estado de avance dentro de cada tarjeta afectada.
  const [otErrorType, setOtErrorType] = useState('');

  // V2.9 — Modal de propagación admin.
  // Estructura: null | { diffs, affectedReports, fixedReport }
  //   - diffs: salida de detectChangesForPropagation
  //   - affectedReports: salida de findAffectedLaterReports
  //   - fixedReport: el reporte editado que se va a guardar después de confirmar
  const [propagationModal, setPropagationModal] = useState(null);

  // V3.3 — Guard de sobreescritura (BACKLOG #20).
  // Se abre cuando el destino (date+shift) ya existe en Supabase con datos,
  // y es distinto del reporte que se tenía abierto originalmente (originalReport).
  // Estructura: null | { reportToSave, existingN: número de correctivos del reporte existente }
  const [overwriteConfirm, setOverwriteConfirm] = useState(null);
  // V3.5 (#22) — Modal de concurrencia (optimistic locking). null | { reportToSave, freshUpdatedAt, mineUpdatedAt }
  const [concurrencyConflict, setConcurrencyConflict] = useState(null);

  // V2.4 — Override del Dashboard: cuando está seteado, el Dashboard muestra
  // ese reporte en lugar del activo. Se resetea automáticamente si el usuario
  // modifica el reporte activo (Pregunta 1 opción B).
  const [dashboardOverride, setDashboardOverride] = useState(null);

  // Wrapper de setReport que también resetea el override (porque el usuario
  // está editando "Cargar Reporte" y queremos que vuelva al turno actual).
  const setReportAndResetOverride = useCallback((nextReportOrFn) => {
    setReport(nextReportOrFn);
    setDashboardOverride(null);
  }, []);

  // BACKLOG #50 (v3.26) — Sonda de pendientes de Extras.
  // try/catch PROPIO y separado del de `refresh`: un fallo acá no debe setear
  // `connError` ni abortar la carga del histórico, que es el dato que la app
  // realmente necesita para funcionar.
  const loadExtrasPendientes = useCallback(async () => {
    try {
      setExtrasPendientes(await storage.hasExtrasPendientes());
    } catch (e) {
      console.warn('Extras: no se pudo verificar si hay pendientes (fail-silent):', e);
      setExtrasPendientes(false);
    }
  }, []);

  const refresh = useCallback(async () => {
    // La sonda va acá y no en un useEffect propio para que quede colgada del
    // mismo gesto que ya refresca todo: el botón ↻ del header y el arranque de
    // la app. Un camino, no dos.
    loadExtrasPendientes();
    // #62 — En el deploy solo-Extras no se carga el histórico de reportes.
    // No es una optimización cosmética: `storage.list()` trae TODOS los
    // reportes de turno, y en el browser de Facilities eso sería descargar en
    // cada arranque un dato que esa pantalla no usa y que ese sector no tiene
    // por qué recibir. Que la tabla sea legible vía REST igual (#47) no es
    // razón para mandársela.
    if (APP_MODE === 'extras') return [];
    try {
      setConnError('');
      const data = await storage.list();
      // Hidratar todos los reportes para garantizar estructura V2.0
      const hydrated = data.map(hydrate);
      setHistory(hydrated);
      return hydrated;
    } catch (e) {
      setConnError(e.message || 'Error de conexión');
      console.error(e);
      return [];
    }
  }, [loadExtrasPendientes]);

  // Punto 2 — Chequeo de versión al arrancar. Fail-open: si no se puede
  // verificar (red caída, tabla inexistente), NO bloquea y deja trabajar.
  const checkVersion = useCallback(async () => {
    try {
      const min = await storage.getConfig('min_version');
      if (min != null && parseVersion(APP_VERSION) < parseInt(min, 10)) {
        setVersionBlocked(true);
      }
    } catch (e) {
      console.warn('Version gate: no se pudo verificar la versión (fail-open):', e);
    }
  }, []);

  useEffect(() => { (async () => {
    draftStore.purgeOld();
    await checkVersion();
    const hydrated = await refresh();
    // Landing PR-1: en PWA instalada + móvil, arrancar en Dashboard mostrando
    // el último reporte guardado (modo viewer). En PC/desktop NO aplica.
    // Mismo criterio de "último" que el Dashboard: fecha + shiftOrder.
    if (isLaunchedAsInstalledMobile() && hydrated && hydrated.length > 0) {
      const sorted = [...hydrated].sort((a, b) =>
        `${a.date}-${shiftOrder(a.shift)}`.localeCompare(`${b.date}-${shiftOrder(b.shift)}`)
      );
      const last = sorted[sorted.length - 1];
      if (last) { setDashboardOverride(last); setTab('dashboard'); }
    }
    setLoading(false);
  })(); }, [checkVersion, refresh]);

  // Validaciones antes de guardar (V2.0)
  // v3.16 — Devuelve { message: string, errorIndices: Set<number> }
  // message: '' si todo OK, string descriptivo si hay error
  // errorIndices: índices (en r.corrective) de las OTs con problemas
  const validateReport = (r) => {
    const ok = { message: '', errorIndices: new Set() };
    const currentShiftKey = `${r.date}-${r.shift}`;

    // V2.4 — 1. OTs nuevas (creadas en este turno) deben tener formato XXX-YYYYY válido
    const indicesInvalidas = [];
    (r.corrective || []).forEach((c, i) => {
      if (c.createdInShift === currentShiftKey && !isValidOT(c.ot)) indicesInvalidas.push(i);
    });
    if (indicesInvalidas.length > 0) {
      return {
        message: `${indicesInvalidas.length} OT nueva con formato inválido. Formato requerido: XXX-YYYYY (sector + 5 dígitos). Ej: FOA1-01395`,
        errorIndices: new Set(indicesInvalidas)
      };
    }

    // V2.5 — 2. Avance del turno obligatorio SOLO cuando hay cambio de estado en este turno.
    // Reemplaza la regla V2.4 que exigía avance para CUALQUIER OT en "En Curso"
    // (generaba ruido con heredadas en curso que el turno actual no trabajó).
    //
    // Reglas:
    //  - OT nueva (no existe en history) creada como Sin Iniciar  → no requiere avance
    //  - OT nueva creada como En Curso o Realizada                → requiere avance
    //  - OT existente en history:
    //      Sin Iniciar → Sin Iniciar : no requiere
    //      Sin Iniciar → En Curso    : requiere
    //      Sin Iniciar → Realizada   : requiere
    //      En Curso    → En Curso    : no requiere (sigue igual)
    //      En Curso    → Realizada   : requiere
    //      Realizada   → cualquiera  : (no debería pasar; lo dejamos pasar)
    //
    // Para detectar el estado previo, buscamos la OT (por número) en el reporte
    // más reciente del history que sea ANTERIOR al turno actual.
    const previousStateOf = (otNumber) => {
      if (!otNumber) return null;
      // Reportes anteriores estrictamente al turno actual (no incluye el actual)
      const previousReports = history
        .filter(rep => {
          if (rep.date < r.date) return true;
          if (rep.date === r.date && shiftOrder(rep.shift) < shiftOrder(r.shift)) return true;
          return false;
        })
        .sort((a, b) =>
          `${b.date}-${shiftOrder(b.shift)}`.localeCompare(`${a.date}-${shiftOrder(a.shift)}`)
        ); // descendente (más reciente primero)
      for (const rep of previousReports) {
        const found = (rep.corrective || []).find(c => c.ot === otNumber);
        if (found) return found.state || null;
      }
      return null; // no aparece antes → es nueva
    };

    const requiresAdvance = (c) => {
      const prev = previousStateOf(c.ot);
      const curr = c.state;
      // Si no existe en history previa: es nueva. Requiere avance si NO está en "Sin Iniciar".
      if (prev === null) {
        return curr === 'En Curso' || curr === 'Realizada';
      }
      // Cambios de estado que requieren avance:
      if (prev === 'Sin Iniciar' && (curr === 'En Curso' || curr === 'Realizada')) return true;
      if (prev === 'En Curso' && curr === 'Realizada') return true;
      // Cualquier otro caso (mismo estado, o no aplica): no requiere
      return false;
    };

    const indicesSinAvance = [];
    (r.corrective || []).forEach((c, i) => {
      if (!requiresAdvance(c)) return;
      const tl = c.timeline || [];
      const hasFromCurrent = tl.some(e => e.shiftKey === currentShiftKey);
      if (!hasFromCurrent) indicesSinAvance.push(i);
    });
    if (indicesSinAvance.length > 0) {
      return {
        message: `${indicesSinAvance.length} OT con cambio de estado en este turno sin entrada de Estado de avance. Cargá el avance antes de guardar.`,
        errorIndices: new Set(indicesSinAvance)
      };
    }

   // #10 (v3.17) — Bypass admin: se aplica DESPUÉS de las reglas críticas (2 y 3).
    // Técnico en En Curso/Realizada y avance al cambiar estado son no-bypasseables incluso en admin.
    // Las reglas contextuales (formato OT, preventivos, duplicadas) sí se bypasean.
    if (adminMode && originalReport) return ok;

    // V2.5 — 3. Técnico obligatorio en OTs "En Curso" y "Realizada".
    // #35 (v3.17) — "Sin Iniciar" ya no bloquea: la OT todavía no fue tomada por nadie,
    // el técnico se asigna cuando arranca el trabajo.
    const indicesSinTecnico = [];
    (r.corrective || []).forEach((c, i) => {
      if (c.state === 'Sin Iniciar') return;
      if (!c.technicians || c.technicians.length === 0) indicesSinTecnico.push(i);
    });
    if (indicesSinTecnico.length > 0) {
      return {
        message: `${indicesSinTecnico.length} OT correctiva sin técnico asignado. Las OTs "En Curso" y "Realizada" deben tener técnico antes de guardar.`,
        errorIndices: new Set(indicesSinTecnico)
      };
    }

    // V2.5 — 4. TODOS los preventivos cargados deben tener al menos un técnico.
    const preventivosSinTecnico = (r.preventive || []).filter(
      p => !p.technicians || p.technicians.length === 0
    );
    if (preventivosSinTecnico.length > 0) {
      return {
        message: `${preventivosSinTecnico.length} preventivo sin técnico asignado. Asigná técnicos antes de guardar.`,
        errorIndices: new Set()
      };
    }

    // 5. Resumen de preventivos: si hay realizados > 0, la suma del detalle por técnico debe coincidir
    const realizados = Number(r.preventivosResumen?.realizados);
    if (realizados > 0) {
      // V2.4 — opción C: cada fila puede tener N técnicos en grupo. La cantidad cuenta UNA VEZ
      // (no se multiplica por la cantidad de técnicos del grupo).
      const sumaPorGrupo = (r.preventivosResumen?.porTecnico || [])
        .reduce((s, t) => s + (Number(t.cantidad) || 0), 0);
      if (sumaPorGrupo !== realizados) {
        return {
          message: `Resumen de preventivos: la suma del detalle (${sumaPorGrupo}) no coincide con "Preventivos realizados" (${realizados}).`,
          errorIndices: new Set()
        };
      }
      // Validar que no haya filas sin técnicos o con cantidad <= 0
      const filasMalas = (r.preventivosResumen?.porTecnico || []).filter(t => {
        const tecnicos = t.tecnicos || (t.tecnico ? [t.tecnico] : []);  // compat hacia atrás
        return tecnicos.length === 0 || !t.cantidad || Number(t.cantidad) <= 0;
      });
      if (filasMalas.length > 0) {
        return {
          message: `Resumen de preventivos: hay filas sin técnicos o con cantidad inválida.`,
          errorIndices: new Set()
        };
      }
    }

    // #26 (v3.7) — 6. Unicidad de número de OT dentro del reporte.
    // Un reporte NO debe tener dos OTs correctivas con el mismo número. El caso
    // típico: el carry-over trajo una OT En Curso y el usuario cargó OTRA con el
    // mismo número en vez de editar la heredada (causa raíz de #25). La dedup de
    // hydrate sanea esto en LECTURA, pero acá lo frenamos en ESCRITURA para que no
    // se genere el dato corrupto y el usuario corrija en el momento.
    // Solo cuenta OTs con número (las vacías son caso aparte, se ignoran).
    // v3.22 (#39) — La comparación pasa a ser por clave canónica, así que
    // 'RO-3255' y 'RO-03255' cuentan como el mismo número (antes escapaban).
    const otCounts = {};
    (r.corrective || []).forEach((c, i) => {
      const key = otKey(c.ot);
      if (!key) return;
      if (!otCounts[key]) otCounts[key] = [];
      otCounts[key].push(i);
    });
    const duplicadas = Object.keys(otCounts).filter(k => otCounts[k].length > 1);
    if (duplicadas.length > 0) {
      const lista = duplicadas.slice(0, 3).join(', ') + (duplicadas.length > 3 ? '…' : '');
      const indicesDup = new Set(duplicadas.flatMap(k => otCounts[k]));
      return {
        message: `Hay ${duplicadas.length} número${duplicadas.length === 1 ? '' : 's'} de OT repetido${duplicadas.length === 1 ? '' : 's'} en el reporte (${lista}). Cada OT debe figurar una sola vez: si una vino del turno anterior, editá esa en vez de cargarla de nuevo.`,
        errorIndices: indicesDup
      };
    }

    return ok;
  };
  //   Correctivo vacío: sin OT, sin task, sin técnicos.
  //   Preventivo vacío: sin equipoCodigo, sin task, sin técnicos.
  // Devuelve { emptyCorrIdx, emptyPrevIdx, cleanedReport } o null si no hay vacíos.
  const detectEmptyEntries = (r) => {
    const emptyCorrIdx = [];
    (r.corrective || []).forEach((c, i) => {
      const noOT = !c.ot || !String(c.ot).trim();
      const noTask = !c.task || !String(c.task).trim();
      const noTech = !c.technicians || c.technicians.length === 0;
      if (noOT && noTask && noTech) emptyCorrIdx.push(i);
    });
    const emptyPrevIdx = [];
    (r.preventive || []).forEach((p, i) => {
      const noEq = !p.equipoCodigo || !String(p.equipoCodigo).trim();
      const noTask = !p.task || !String(p.task).trim();
      const noTech = !p.technicians || p.technicians.length === 0;
      if (noEq && noTask && noTech) emptyPrevIdx.push(i);
    });
    if (emptyCorrIdx.length === 0 && emptyPrevIdx.length === 0) return null;
    const cleanedReport = {
      ...r,
      corrective: (r.corrective || []).filter((_, i) => !emptyCorrIdx.includes(i)),
      preventive: (r.preventive || []).filter((_, i) => !emptyPrevIdx.includes(i))
    };
    return { emptyCorr: emptyCorrIdx.length, emptyPrev: emptyPrevIdx.length, cleanedReport };
  };

  // Ejecuta el guardado real (sin modal). Se llama directo si no hay vacíos,
  // o desde el modal después de confirmar.
  // V2.8 — Detección de OTs cerradas en otros turnos (carry-over stale guard).
  //
  // Problema que resuelve: si un turno A marca una OT como Realizada y guarda, pero
  // el turno B tenía el form ya abierto (con la OT En Curso heredada del carry-over
  // viejo), al guardar B con la OT En Curso PISA la lógica de cierre de A.
  //
  // Esta función detecta esos casos comparando el reporte que se va a guardar
  // contra el history fresco recién traído de Supabase.
  //
  // Devuelve array de conflictos con shape:
  //   { otNumber, otTask, formState, closedIn: {date, shift, responsable, when} }
  // Solo considera OTs con número de OT no vacío (las legacy se incluyen porque
  // también pueden estar en estado contradictorio).
  //
  // Política: una OT se considera "ya cerrada" si EN OTRO REPORTE del history fresco
  // (que NO sea el actual que estamos por guardar) tiene state === 'Realizada'.
  //
  // Importante: el chequeo es por número de OT exacto (string match). OTs con
  // exactamente el mismo número (incluso legacy con espacios) caen al mismo bucket.
  const detectClosedConflicts = (reportToCheck, freshHistory) => {
    const conflicts = [];
    const corrective = reportToCheck.corrective || [];

    // Helper: buscar en freshHistory si la OT está Realizada en algún reporte
    // distinto al actual.
    const findClosingReport = (otNumber) => {
      // Iteramos en orden descendente (más reciente primero) para reportar el cierre más reciente
      const sorted = [...freshHistory].sort((a, b) => {
        const ka = `${b.date}-${shiftOrder(b.shift)}`;
        const kb = `${a.date}-${shiftOrder(a.shift)}`;
        return ka.localeCompare(kb);
      });
      for (const rep of sorted) {
        // Excluir el reporte actual (mismo date+shift)
        if (rep.date === reportToCheck.date && rep.shift === reportToCheck.shift) continue;
        const match = (rep.corrective || []).find(c => c.ot === otNumber && c.state === 'Realizada');
        if (match) return { date: rep.date, shift: rep.shift, responsable: rep.responsable, ot: match };
      }
      return null;
    };

    corrective.forEach(c => {
      if (!c.ot) return; // Sin número de OT no hay forma de detectar conflicto
      if (c.state === 'Realizada') return; // Si el form ya tiene Realizada, no hay conflicto
      const closing = findClosingReport(c.ot);
      if (closing) {
        conflicts.push({
          otNumber: c.ot,
          otTask: c.task || '(sin descripción)',
          formState: c.state,
          closedIn: {
            date: closing.date,
            shift: closing.shift,
            responsable: closing.responsable || '(sin responsable)'
          }
        });
      }
    });

    return conflicts;
  };


  // V2.9 — Detecta cambios entre el reporte original (snapshot al cargar) y el
  // reporte actual (form en pantalla). Devuelve un array de diffs por OT.
  //
  // Solo considera cambios "propagables":
  //   - state: la OT cambió de Sin Iniciar/En Curso/Realizada a otro estado
  //   - timeline.added: hay entradas nuevas en el timeline (matching por id)
  //   - timeline.deleted: faltan entradas que estaban en el original (matching por id)
  //
  // Cada diff: { ot, otTask, stateChange: {from, to} | null, addedEntries: [...], deletedEntries: [...] }
  // Solo se incluye una OT en el resultado si tiene al menos un cambio.
  // Las OTs nuevas (que no existían en el original) NO se incluyen — no hay nada
  // que propagar de algo que recién apareció.
  const detectChangesForPropagation = (original, current) => {
    if (!original || !current) return [];
    const origByOt = new Map();
    (original.corrective || []).forEach(c => {
      if (c.ot) origByOt.set(c.ot, c);
    });
    const diffs = [];
    (current.corrective || []).forEach(currOt => {
      if (!currOt.ot) return;
      const origOt = origByOt.get(currOt.ot);
      if (!origOt) return; // OT nueva: no hay original, no hay propagación

      // Cambio de state
      let stateChange = null;
      if (origOt.state !== currOt.state) {
        stateChange = { from: origOt.state, to: currOt.state };
      }

      // Diff de timeline por id
      const origIds = new Set((origOt.timeline || []).map(e => e.id).filter(Boolean));
      const currIds = new Set((currOt.timeline || []).map(e => e.id).filter(Boolean));
      const addedEntries = (currOt.timeline || []).filter(e => e.id && !origIds.has(e.id));
      const deletedEntries = (origOt.timeline || []).filter(e => e.id && !currIds.has(e.id));

      if (stateChange || addedEntries.length > 0 || deletedEntries.length > 0) {
        diffs.push({
          ot: currOt.ot,
          otTask: currOt.task || '(sin descripción)',
          stateChange,
          addedEntries,
          deletedEntries
        });
      }
    });
    return diffs;
  };

  // V2.9 — Identifica reportes posteriores al reporte editado que contienen
  // las OTs afectadas por los diffs.
  //
  // "Posterior" = (date > editedReport.date) OR
  //               (date == editedReport.date AND shiftOrder(shift) > shiftOrder(edited.shift))
  //
  // Por cada OT afectada, devuelve los reportes posteriores donde aparece esa OT
  // (matching por número exacto). Si la OT no aparece en ningún reporte posterior,
  // no se incluye en el resultado.
  //
  // Devuelve: array de { report, affectedOts: [{ ot, currentState, currentTimelineIds, diff }] }
  //   - report: el reporte posterior completo (con date, shift, id de Supabase, etc.)
  //   - affectedOts: las OTs de ese reporte que matchean con algún diff, junto con
  //     su estado actual en ese reporte y el diff correspondiente.
  const findAffectedLaterReports = (editedReport, diffs, allReports) => {
    if (!diffs || diffs.length === 0) return [];
    const editedKey = `${editedReport.date}-${shiftOrder(editedReport.shift)}`;
    const diffByOt = new Map(diffs.map(d => [d.ot, d]));

    const laterReports = allReports
      .filter(r => {
        const k = `${r.date}-${shiftOrder(r.shift)}`;
        return k > editedKey;
      })
      .sort((a, b) => {
        const ka = `${a.date}-${shiftOrder(a.shift)}`;
        const kb = `${b.date}-${shiftOrder(b.shift)}`;
        return ka.localeCompare(kb); // cronológico ascendente
      });

    const result = [];
    laterReports.forEach(r => {
      const affectedOts = [];
      (r.corrective || []).forEach(c => {
        if (!c.ot) return;
        const diff = diffByOt.get(c.ot);
        if (!diff) return;
        affectedOts.push({
          ot: c.ot,
          currentState: c.state,
          currentTimelineIds: new Set((c.timeline || []).map(e => e.id).filter(Boolean)),
          diff
        });
      });
      if (affectedOts.length > 0) {
        result.push({ report: r, affectedOts });
      }
    });
    return result;
  };
  const doSaveReport = async (reportToSave, overwriteConfirmed = false, concurrencyConfirmed = false) => {
    if (!reportToSave.date || !reportToSave.shift) { setSaveMsg('Falta fecha o turno'); return; }
    // #9 (v3.14) — Bloqueo duro: un no-admin NO puede guardar reportes anteriores a ayer.
    // Editar hoy/ayer es operación normal del turno en curso; corregir el pasado es
    // edición retroactiva → solo admin (donde la propagación V2.9 funciona). Red de
    // seguridad real: aunque el form esté en read-only y el botón oculto, este guard
    // impide el guardado por cualquier vía. Admin pasa siempre.
    if (!adminMode && !isWithinEditWindow(reportToSave.date)) {
      setSaveMsg('Error: solo se pueden editar reportes de hoy o ayer. Para corregir reportes anteriores, ingresá como admin.');
      return;
    }
    // #11 (v3.15) — Bloqueo duro: un no-admin NO puede guardar un reporte con fecha FUTURA.
    // Una fecha posterior a hoy es siempre un error de carga (el reporte de un turno se
    // carga durante o al cierre del turno, nunca antes de que ocurra). Evita data sucia
    // que después se renombra por SQL. Admin pasa siempre (puede precargar, ej. una tarea
    // preventiva para un día puntual). Chequeo independiente del de #9 (futuro ≠ pasado).
    if (!adminMode && isFutureDate(reportToSave.date)) {
      setSaveMsg('Error: no se puede guardar un reporte con fecha futura. Verificá la fecha del reporte.');
      return;
    }
    const validationResult = validateReport(reportToSave);
    if (validationResult.message) {
      setSaveMsg(`Error: ${validationResult.message}`);
      // v3.16 — Marcar OTs con error y determinar tipo para resaltar sección de avance
      setOtErrorIndices(validationResult.errorIndices);
      const isAvanceError = validationResult.message.includes('Estado de avance');
      const isTecnicoError = validationResult.message.includes('sin técnico');
      setOtErrorType(isAvanceError ? 'avance' : isTecnicoError ? 'tecnico' : 'otro');
      // v3.16 — Scroll a la primera OT con error (si hay alguna)
      if (validationResult.errorIndices.size > 0) {
        const firstIdx = Math.min(...validationResult.errorIndices);
        const corrective = reportToSave.corrective || [];
        const firstOt = corrective[firstIdx];
        const elementId = firstOt
          ? `form-ot-${firstOt.ot || `idx-${firstIdx}`}`
          : `form-ot-idx-${firstIdx}`;
        setTimeout(() => {
          const el = document.getElementById(elementId);
          if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }, 80);
      } else if (validationResult.message.startsWith('Resumen de preventivos')) {
        // v3.16 (A) — Error del Resumen de Preventivos (regla 5). No tiene índice de
        // OT; el bloque ya tiene feedback visual propio (borde rojo + cartel), solo
        // le faltaba que la pantalla llevara al usuario hasta ahí.
        setTimeout(() => {
          const el = document.getElementById('form-preventivos');
          if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }, 80);
      }
      return;
    }
    // Validación OK: limpiar errores previos
    setOtErrorIndices(new Set());
    setOtErrorType('');
    setSaving(true);
    setSaveMsg('Verificando…');

    // V2.8 — Antes de guardar, recargamos history fresh para detectar conflictos
    // de carry-over stale (turno B intenta pisar OT que A ya cerró).
    let freshHistory;
    try {
      freshHistory = (await storage.list()).map(hydrate);
    } catch (e) {
      // Si falla la recarga, igual permitimos guardar (no bloquear al usuario por
      // un problema de red en la verificación; el guardado en sí puede funcionar).
      console.warn('No se pudo recargar history para verificar conflictos:', e);
      freshHistory = history;
    }

    // V3.5 (#22) — Chequeo de concurrencia (optimistic locking).
    // Solo aplica al re-guardar un reporte que YA teníamos abierto (mismo date+shift que
    // el snapshot original). Si el updated_at en la base cambió respecto al que cargamos,
    // significa que otra sesión guardó en el medio → abrimos modal antes de pisar.
    // No aplica si: es un reporte nuevo, cambiaste de turno (de eso se ocupa el guard #20),
    // ya confirmaste sobreescribir la concurrencia, o no hay timestamp de referencia.
    // Fail-open coherente con el resto: si freshHistory cayó al history viejo (catch de red),
    // mineUpdatedAt/freshUpdatedAt pueden no ser confiables, pero como comparamos igualdad
    // estricta y el caso de red ya no recargó, no genera falso positivo relevante.
    if (!concurrencyConfirmed && originalReport) {
      const destId = `${reportToSave.date}-${reportToSave.shift}`;
      const originalId = `${originalReport.date}-${originalReport.shift}`;
      if (destId === originalId) {
        const mineUpdatedAt = originalReport._updatedAt || null;
        const fresh = freshHistory.find(r => `${r.date}-${r.shift}` === destId);
        const freshUpdatedAt = fresh ? (fresh._updatedAt || null) : null;
        if (mineUpdatedAt && freshUpdatedAt && mineUpdatedAt !== freshUpdatedAt) {
          setSaving(false);
          setSaveMsg('');
          setConcurrencyConflict({ reportToSave, freshUpdatedAt, mineUpdatedAt });
          return;
        }
      }
    }
    // V3.3 — Guard de sobreescritura (BACKLOG #20).
    // Si el destino (date+shift) ya existe en freshHistory con datos, y NO es el mismo
    // reporte que teníamos abierto originalmente (originalReport), mostrar modal de
    // confirmación antes de pisar. Esto previene el incidente del 2026-05-21 donde
    // se guardó encima de Mañana sin querer al no cambiar el selector de turno.
    // No dispara si: re-guardás el mismo reporte abierto, guardás uno nuevo, admin
    // edita con snapshot, o ya confirmaste la sobreescritura (overwriteConfirmed).
    if (!overwriteConfirmed && !(adminMode && originalReport)) {
      const destId = `${reportToSave.date}-${reportToSave.shift}`;
      const originalId = originalReport ? `${originalReport.date}-${originalReport.shift}` : null;
      const existing = freshHistory.find(r => `${r.date}-${r.shift}` === destId);
      const hasData = existing && (
        (existing.corrective || []).length > 0 ||
        (existing.team || []).length > 0 ||
        (existing.comments || []).length > 0
      );
      if (hasData && destId !== originalId) {
        setSaving(false);
        setSaveMsg('');
        setOverwriteConfirm({
          reportToSave,
          existingN: (existing.corrective || []).length,
          date: reportToSave.date,
          shift: reportToSave.shift,
        });
        return;
      }
    }

    // V2.9 — Si admin está editando un reporte histórico (tiene snapshot original),
    // saltar la detección V2.8 de conflictos. V2.8 está pensada para responsables
    // que tenían el form abierto cuando otro turno cerró la OT; admin con snapshot
    // ya sabe que edita un reporte pasado y la propagación V2.9 maneja el caso.
    const conflicts = (adminMode && originalReport) ? [] : detectClosedConflicts(reportToSave, freshHistory);
    if (conflicts.length > 0) {
      // Hay conflicto: abrir modal y postergar guardado real
      setSaving(false);
      setSaveMsg('');
      setClosedConflicts({
        conflicts,
        reportToSave, // guardamos referencia para usar después de resolver
      });
      return;
    }

    // Sin conflictos: proceder al guardado real
    // V2.9 — Propagación admin retroactiva.
    // Solo entra acá si adminMode && hay snapshot original (= reporte histórico editado).
    // Detectamos diffs vs original; si hay cambios:
    //   - Cambios SOLO de timeline → propagar silenciosamente (decisión B).
    //   - Cambios de state (con o sin timeline) → abrir modal para decisión caso por caso.
    // Si no hay reportes posteriores afectados, guardar normal sin propagar.
    if (adminMode && originalReport) {
      const diffs = detectChangesForPropagation(originalReport, reportToSave);
      if (diffs.length > 0) {
        const affectedReports = findAffectedLaterReports(reportToSave, diffs, freshHistory);
        if (affectedReports.length > 0) {
          const hasStateChanges = diffs.some(d => d.stateChange);
          if (hasStateChanges) {
            // Abrir modal para decisión caso por caso
            setSaving(false);
            setSaveMsg('');
            setPropagationModal({ diffs, affectedReports, fixedReport: reportToSave });
            return;
          }
          // Solo timeline changes → propagar silenciosamente
          setSaveMsg('Propagando cambios…');
          try {
            // Generar entrada de auditoría en el reporte original
            const nowIso = new Date().toISOString();
            const shiftKey = `${reportToSave.date}-${reportToSave.shift}`;
            const fixedWithAudit = {
              ...reportToSave,
              corrective: (reportToSave.corrective || []).map(c => {
                const diff = diffs.find(d => d.ot === c.ot);
                if (!diff) return c;
                const auditParts = [];
                if (diff.addedEntries.length > 0) {
                  auditParts.push(`${diff.addedEntries.length} entrada${diff.addedEntries.length === 1 ? '' : 's'} agregada${diff.addedEntries.length === 1 ? '' : 's'}`);
                }
                if (diff.deletedEntries.length > 0) {
                  auditParts.push(`${diff.deletedEntries.length} entrada${diff.deletedEntries.length === 1 ? '' : 's'} borrada${diff.deletedEntries.length === 1 ? '' : 's'}`);
                }
                const propagatedCount = affectedReports.filter(({ affectedOts }) =>
                  affectedOts.some(a => a.ot === c.ot)
                ).length;
                const auditText = `[Edición admin] ${auditParts.join(' · ')}. Propagado a ${propagatedCount} reporte${propagatedCount === 1 ? '' : 's'} posterior${propagatedCount === 1 ? '' : 'es'}.`;
                const auditEntry = {
                  id: generateTimelineId(),
                  shiftKey,
                  date: reportToSave.date,
                  shift: reportToSave.shift,
                  author: reportToSave.responsable || '(admin)',
                  text: auditText,
                  timestamp: nowIso
                };
                return { ...c, timeline: [...(c.timeline || []), auditEntry] };
              })
            };
            await propagateChanges(fixedWithAudit, diffs, affectedReports, {});
            await refresh();
            setReport(fixedWithAudit);
            setOriginalReport(JSON.parse(JSON.stringify(fixedWithAudit)));
            setSaveMsg('✓ Reporte guardado y propagado');
            setTimeout(() => setSaveMsg(''), 3000);
            setSaving(false);
            return;
          } catch (e) {
            setSaveMsg(`Error: ${e.message}`);
            setSaving(false);
            return;
          }
        }
        // Hay diffs pero no hay reportes posteriores afectados → guardar normal.
        // (Ej: admin editó el último reporte cargado, o la OT no aparece en ningún posterior).
      }
    }
    setSaveMsg('Guardando…');
    // BACKLOG #21 (v3.12) — Capturamos la versión vieja del destino ANTES de pisar,
    // tomándola de freshHistory (lo que hay en la base ahora). Solo si existe CON datos:
    // un destino vacío o inexistente no tiene nada que respaldar (primer guardado del turno).
    // El INSERT al historial se hace DESPUÉS del UPSERT exitoso (más abajo), fail-open.
    // El motivo se clasifica por el flag con el que entró doSaveReport (NO por si el guard
    // saltó: en la 2da pasada confirmada el guard ya no dispara, ver #20/#22):
    //   concurrencyConfirmed → 'concurrency' (confirmó pisar pese al conflicto de #22)
    //   overwriteConfirmed   → 'overwrite_turno' (confirmó pisar turno con datos, #20)
    //   ninguno              → 'save_normal' (re-guardado del propio turno que ya tenía datos)
    let _histPrev = null;
    {
      const destId = `${reportToSave.date}-${reportToSave.shift}`;
      const prev = freshHistory.find(r => `${r.date}-${r.shift}` === destId);
      const prevHasData = prev && (
        (prev.corrective || []).length > 0 ||
        (prev.team || []).length > 0 ||
        (prev.comments || []).length > 0
      );
      if (prevHasData) {
        const motivo = concurrencyConfirmed ? 'concurrency'
          : overwriteConfirmed ? 'overwrite_turno'
          : 'save_normal';
        // Snapshot del data crudo viejo, sin los campos internos de runtime (_updatedAt).
        const { _updatedAt, ...dataAnterior } = prev;
        _histPrev = { destId, dataAnterior, updatedAtAnterior: _updatedAt || null, motivo };
      }
    }
    try {
      await storage.save(reportToSave);
      // BACKLOG #21 — Backup de la versión pisada. Fail-open: si el INSERT al historial
      // falla, NO revierte ni bloquea el guardado (que ya ocurrió). Solo se loguea.
      if (_histPrev) {
        try {
          await storage.saveHistorial(_histPrev.destId, _histPrev.dataAnterior, _histPrev.updatedAtAnterior, _histPrev.motivo);
        } catch (eh) {
          console.warn('No se pudo respaldar la versión anterior (historial #21):', eh);
        }
      }
      await refresh();
      // V2.5 — Si se limpiaron entradas vacías, persistir el cambio en el state local
      // así el usuario ve el form sin las filas vacías
      if (reportToSave !== report) setReport(reportToSave);
      setOriginalReport(JSON.parse(JSON.stringify(reportToSave)));  // V2.9 — actualizar snapshot al guardado nuevo
      draftStore.clear(`${reportToSave.date}-${reportToSave.shift}`);  // #7 v3.6 — guardado OK: limpiar borrador local
      // v3.16 — Limpiar marcas de error al guardar exitosamente
      setOtErrorIndices(new Set());
      setOtErrorType('');
      setSaveMsg('✓ Reporte guardado');
      setTimeout(() => setSaveMsg(''), 2500);
    } catch (e) {
      setSaveMsg(`Error: ${e.message}`);
    }
    setSaving(false);
  };

  // V2.8 — Resolución del modal de conflictos.
  // decisions: Array<{ otNumber, action: 'remove' | 'reopen', reopenReason?: string }>
  // - 'remove': la OT se elimina del array corrective del reporte antes de guardar
  // - 'reopen': la OT se mantiene como está + se agrega entrada al timeline
  //             documentando la reapertura. Solo permitido en adminMode.
  // V2.9 — Ejecuta la propagación de cambios admin a reportes posteriores.
  // Llama a la RPC propagate_admin_changes en Supabase en una transacción atómica.
  //
  // Toma las decisiones del PropagationModal (qué state changes propagar y cuáles no),
  // arma el payload y lo manda. Los cambios de timeline (add/delete) siempre se propagan
  // de forma uniforme a todos los reportes posteriores afectados (decidido en B).
  //
  // Si la RPC falla, no se guarda nada (rollback automático en Supabase).
  // Si tiene éxito, refrescamos el history para reflejar los cambios.
  const propagateChanges = async (fixedReport, diffs, affectedReports, stateDecisions) => {
    if (!supabaseConfigured) {
      throw new Error('Propagación solo disponible con Supabase configurado');
    }
    const originalReportId = `${fixedReport.date}-${fixedReport.shift}`;

    // Armar lista de operaciones por reporte posterior
    const propagations = affectedReports.map(({ report, affectedOts }) => {
      const operations = [];
      affectedOts.forEach(({ ot, currentTimelineIds, diff }) => {
        // Add timeline (uniforme): por cada entrada nueva del original, agregar a este reporte
        // si todavía no la tiene (matching por id).
        diff.addedEntries.forEach(entry => {
          if (!currentTimelineIds.has(entry.id)) {
            operations.push({ type: 'add_timeline', ot, entry });
          }
        });
        // Delete timeline (uniforme): por cada entrada borrada del original, borrar de este reporte
        // si la tiene (matching por id).
        diff.deletedEntries.forEach(entry => {
          if (currentTimelineIds.has(entry.id)) {
            operations.push({ type: 'delete_timeline', ot, entry_id: entry.id });
          }
        });
        // Change state (caso por caso): solo si admin lo decidió en el modal.
        if (diff.stateChange) {
          const key = `${report.date}|${report.shift}|${ot}`;
          if (stateDecisions[key] === 'propagate') {
            operations.push({ type: 'change_state', ot, new_state: diff.stateChange.to });
          }
        }
      });
      return {
        report_id: `${report.date}-${report.shift}`,
        operations
      };
    }).filter(p => p.operations.length > 0); // descartar reportes sin operaciones efectivas

    const payload = {
      original_report_id: originalReportId,
      original_report_data: fixedReport,
      propagations
    };

    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/rpc/propagate_admin_changes`,
      {
        method: 'POST',
        headers: { ...sbHeaders(), 'Content-Type': 'application/json' },
        body: JSON.stringify({ payload })
      }
    );
    if (!res.ok) {
      throw new Error(`Propagación falló: ${res.status} ${await res.text()}`);
    }
    return await res.json();
  };

  // V2.9 — Callback del PropagationModal cuando admin confirma.
  // Aplica los cambios del modal + ejecuta la propagación + refresca history.
  const handlePropagationConfirm = async ({ stateDecisions }) => {
    if (!propagationModal) return;
    const { diffs, affectedReports, fixedReport } = propagationModal;
    setPropagationModal(null);
    setSaving(true);
    setSaveMsg('Propagando cambios…');
    try {
      // V2.9 — Agregar entrada de auditoría [Edición admin] en el reporte original
      // ANTES de propagar. Se agrega en cada OT que tuvo cambios.
      const nowIso = new Date().toISOString();
      const shiftKey = `${fixedReport.date}-${fixedReport.shift}`;
      const fixedWithAudit = {
        ...fixedReport,
        corrective: (fixedReport.corrective || []).map(c => {
          const diff = diffs.find(d => d.ot === c.ot);
          if (!diff) return c;
          const auditParts = [];
          if (diff.stateChange) {
            auditParts.push(`estado ${diff.stateChange.from} → ${diff.stateChange.to}`);
          }
          if (diff.addedEntries.length > 0) {
            auditParts.push(`${diff.addedEntries.length} entrada${diff.addedEntries.length === 1 ? '' : 's'} agregada${diff.addedEntries.length === 1 ? '' : 's'}`);
          }
          if (diff.deletedEntries.length > 0) {
            auditParts.push(`${diff.deletedEntries.length} entrada${diff.deletedEntries.length === 1 ? '' : 's'} borrada${diff.deletedEntries.length === 1 ? '' : 's'}`);
          }
          const propagatedCount = affectedReports.filter(({ affectedOts }) =>
            affectedOts.some(a => a.ot === c.ot)
          ).length;
          const auditText = `[Edición admin] ${auditParts.join(' · ')}. Propagado a ${propagatedCount} reporte${propagatedCount === 1 ? '' : 's'} posterior${propagatedCount === 1 ? '' : 'es'}.`;
          const auditEntry = {
            id: generateTimelineId(),
            shiftKey,
            date: fixedReport.date,
            shift: fixedReport.shift,
            author: fixedReport.responsable || '(admin)',
            text: auditText,
            timestamp: nowIso
          };
          return { ...c, timeline: [...(c.timeline || []), auditEntry] };
        })
      };

      await propagateChanges(fixedWithAudit, diffs, affectedReports, stateDecisions);
      await refresh();
      setReport(fixedWithAudit);
      setOriginalReport(JSON.parse(JSON.stringify(fixedWithAudit)));
      setSaveMsg('✓ Reporte guardado y propagado');
      setTimeout(() => setSaveMsg(''), 3000);
    } catch (e) {
      setSaveMsg(`Error: ${e.message}`);
    }
    setSaving(false);
  };

  const handlePropagationCancel = () => {
    setPropagationModal(null);
    setSaveMsg('');
  };
  const handleConflictResolve = async (decisions) => {
    if (!closedConflicts) return;
    const { reportToSave } = closedConflicts;
    setClosedConflicts(null);

    // Aplicar decisiones al reporte
    const decisionByOt = new Map(decisions.map(d => [d.otNumber, d]));
    const nowIso = new Date().toISOString();
    const shiftKey = `${reportToSave.date}-${reportToSave.shift}`;

    const newCorrective = (reportToSave.corrective || []).reduce((acc, c) => {
      const decision = decisionByOt.get(c.ot);
      if (!decision) {
        acc.push(c);
        return acc;
      }
      if (decision.action === 'remove') {
        // No se incluye en el reporte (se elimina)
        return acc;
      }
      if (decision.action === 'reopen') {
        // Mantener la OT + agregar entrada al timeline documentando la reapertura
        const reopenEntry = {
          id: generateTimelineId(),                // V2.9 — id único
          date: reportToSave.date,
          text: `[Reapertura admin] Motivo: ${decision.reopenReason || '(sin motivo)'}`,
          shift: reportToSave.shift,
          author: reportToSave.responsable || '(admin)',
          shiftKey,
          timestamp: nowIso
        };
        acc.push({
          ...c,
          timeline: [...(c.timeline || []), reopenEntry]
        });
        return acc;
      }
      acc.push(c);
      return acc;
    }, []);

    const fixedReport = { ...reportToSave, corrective: newCorrective };
    // Actualizar también el state local así el usuario ve los cambios aplicados
    setReport(fixedReport);

    // Llamar al guardado real saltando la verificación (ya la pasamos).
    // Para evitar recursión infinita usamos un guardado directo sin re-detección.
    setSaving(true);
    setSaveMsg('Guardando…');
    try {
      await storage.save(fixedReport);
      await refresh();
      setOriginalReport(JSON.parse(JSON.stringify(fixedReport)));  // V2.9 — actualizar snapshot al guardado nuevo
      draftStore.clear(`${fixedReport.date}-${fixedReport.shift}`);  // #7 v3.6 — guardado OK: limpiar borrador local
      setSaveMsg('✓ Reporte guardado');
      setTimeout(() => setSaveMsg(''), 2500);
    } catch (e) {
      setSaveMsg(`Error: ${e.message}`);
    }
    setSaving(false);
  };

  const handleConflictCancel = () => {
    setClosedConflicts(null);
    setSaveMsg('');
  };

  const saveReport = async () => {
    if (!report.date || !report.shift) { setSaveMsg('Falta fecha o turno'); return; }
    // v3.16 — Limpiar errores anteriores al iniciar un nuevo intento de guardado
    setOtErrorIndices(new Set());
    setOtErrorType('');
    // V2.5 — Antes de validar, detectar entradas vacías y abrir modal si las hay
    const detection = detectEmptyEntries(report);
    if (detection) {
      setEmptyConfirm(detection);
      return;
    }
    await doSaveReport(report);
  };

  // V2.5 — Handlers del modal
  const handleConfirmEmpty = async () => {
    if (!emptyConfirm) return;
    const cleaned = emptyConfirm.cleanedReport;
    setEmptyConfirm(null);
    await doSaveReport(cleaned);
  };
  const handleCancelEmpty = () => setEmptyConfirm(null);

  // V2.6 — Handlers de modo administrador
  // BACKLOG #42 (Fase 1) — carga del pool. Se dispara al entrar como planificador.
  const loadPool = useCallback(async () => {
    if (!supabaseConfigured) { setPool([]); return; }
    setPoolLoading(true);
    setPoolError('');
    try {
      setPool(await storage.listPool());
    } catch (e) {
      setPoolError(e.message || 'No se pudo cargar el pool');
    } finally {
      setPoolLoading(false);
    }
  }, []);

  useEffect(() => { if (poolMode) loadPool(); }, [poolMode, loadPool]);

  const handlePoolLogin = (passwordTry) => {
    if (passwordTry === POOL_PASSWORD) {
      setPoolMode(true);
      setPoolLoginOpen(false);
      return true;
    }
    return false;
  };

  const handlePoolLogout = () => {
    setPoolMode(false);
    setPool([]);
    // Si estaba parado en la pestaña del pool, vuelve a Carga: la tab deja de existir.
    setTab(t => (t === 'pool' ? 'form' : t));
  };

  const handlePoolAdd = async (row) => {
    const created = await storage.insertPool(row);
    setPool(prev => [created, ...prev]);
    return created;
  };

  const handlePoolAnular = async (id, motivo) => {
    const updated = await storage.anularPool(id, motivo);
    setPool(prev => prev.map(r => (r.id === id ? updated : r)));
    return updated;
  };

  // BACKLOG #46 (v3.25) — Carga y handlers de horas extras.
  // Se trae SIEMPRE la tabla completa y se filtra en el cliente por rol. Esto
  // es una decisión consciente y su límite hay que tenerlo presente: filtrar
  // server-side (`?solicitado_por=eq.X`) daría exactamente la misma garantía,
  // que es ninguna, porque la RLS está abierta y cualquiera puede pedir la
  // tabla entera igual. Traer todo simplifica el refresh y evita dos caminos
  // de carga distintos según el rol.
  const loadExtras = useCallback(async () => {
    if (!supabaseConfigured) { setExtras([]); return; }
    setExtrasLoading(true);
    setExtrasError('');
    try {
      setExtras(await storage.listExtras());
    } catch (e) {
      setExtrasError(e.message || 'No se pudieron cargar las horas extras');
    } finally {
      setExtrasLoading(false);
    }
  }, []);

  useEffect(() => { if (extrasUser) loadExtras(); }, [extrasUser, loadExtras]);

  const handleExtrasLogin = (user, pass) => {
    const sess = extrasAuth(user, pass);
    if (sess) {
      setExtrasUser(sess);
      setExtrasLoginOpen(false);
      setTab('extras');
      return true;
    }
    return false;
  };

  const handleExtrasLogout = () => {
    setExtrasUser(null);
    setExtras([]);
    // Si estaba parado en la solapa de Extras, vuelve a Carga: la tab deja de existir.
    setTab(t => (t === 'extras' ? 'form' : t));
    // #50 — Revalidar la sonda justo cuando el botón vuelve a ser visible.
    // Durante la sesión el aviso NO se renderiza (en su lugar está el badge del
    // rol), así que no hace falta recalcularlo en cada alta o resolución: alcanza
    // con hacerlo en el único momento en que el resultado vuelve a mirarse.
    // Si el jefe acaba de aprobar la última pendiente, el naranja se apaga acá.
    loadExtrasPendientes();
  };

  const handleExtrasAdd = async (row) => {
    const created = await storage.insertExtra(row);
    setExtras(prev => [created, ...prev]);
    return created;
  };

  const handleExtrasUpdate = async (id, patch) => {
    const updated = await storage.updateExtra(id, patch);
    setExtras(prev => prev.map(r => (r.id === id ? updated : r)));
    return updated;
  };

  const handleAdminLogin = (passwordTry) => {
    if (passwordTry === ADMIN_PASSWORD) {
      setAdminMode(true);
      setAdminLoginOpen(false);
      return true;
    }
    return false;
  };
  const handleAdminLogout = () => {
    setAdminMode(false);
  };

  // V2.6 — Eliminar reporte completo (modo admin)
  // `source` puede ser 'form' (estás editando ese reporte) o 'history' (eliminás desde la lista)
  const requestDeleteReport = (date, shift, source = 'form') => {
    setDeleteReportConfirm({ date, shift, source });
  };
  const confirmDeleteReport = async () => {
    if (!deleteReportConfirm) return;
    const { date, shift, source } = deleteReportConfirm;
    setDeleteReportConfirm(null);
    try {
      await storage.delete(date, shift);
      await refresh();
      // Si estabas editando el reporte que se eliminó, limpiar el form
      if (source === 'form' && report.date === date && report.shift === shift) {
        setReport(emptyReport());
        setDashboardOverride(null);
      }
      setSaveMsg('✓ Reporte eliminado');
      setTimeout(() => setSaveMsg(''), 2500);
    } catch (e) {
      setSaveMsg(`Error: ${e.message}`);
    }
  };
  const cancelDeleteReport = () => setDeleteReportConfirm(null);

  // V2.6 — Redirección desde Dashboard a Cargar Reporte para editar una OT específica.
  // Cuando el admin clickea una OT en el Dashboard, se carga el reporte
  // correspondiente en el form y se cambia de pestaña.
  // scrollTarget opcional: 'preventivos' o 'ot:<numero>' para scrollear a esa sección.
  const editFromDashboard = (reportData, scrollTarget = null) => {
    setReport(hydrate(reportData));
    setDashboardOverride(null);
    setTab('form');
    if (scrollTarget) {
      // Pequeño delay para que el form se monte primero.
      // Reintenta una vez si el elemento aún no está en el DOM (form recién montado en dispositivos lentos).
      const id = scrollTarget === 'preventivos'
        ? 'form-preventivos'
        : `form-ot-${scrollTarget.replace(/^ot:/, '')}`;
      const tryScroll = (attempt = 0) => {
        const el = document.getElementById(id);
        if (el) {
          el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        } else if (attempt < 3) {
          setTimeout(() => tryScroll(attempt + 1), 100);
        }
      };
      setTimeout(() => tryScroll(0), 80);
    }
  };

  // ── Excel exports (matching template format + V2.0 additions) ────────────────────
  const exportFull = () => {
    if (!history.length) { alert('No hay reportes guardados.'); return; }
    const wb = XLSX.utils.book_new();

    // 1. Tecnicos presentes
    const tecPres = [];
    history.forEach(r => (r.team || []).forEach(name => {
      tecPres.push({
        TecnicoID: findTecnicoId(name),
        TecnicoNombre: name,
        TecnicoPresente: 'Presente',
        TecnicoFechaControl: r.date
      });
    }));
    addSheet(wb, tecPres.length ? tecPres : [{ TecnicoID: '', TecnicoNombre: '', TecnicoPresente: '', TecnicoFechaControl: '' }], 'Tecnicos presentes');

    // 2. Correctivos
    const corr = [];
    let cId = 1;
    history.forEach(r => (r.corrective || []).forEach(c => {
      // V2.4 — Serializar timeline en una columna (formato legible para auditoría)
      const timelineStr = (c.timeline || [])
        .map(e => `[${e.date} ${e.shift}${e.author ? ' / ' + e.author : ''}] ${e.text}`)
        .join('\n');
      corr.push({
        OrdenID: cId++,
        OrdenNumero: c.ot || '',
        SectorID: '',
        EquipoID: '',
        EquipoCodigo: c.equipoCodigo || '',
        OrdenTecnicoAsignado: (c.technicians || []).join(', '),
        Turno: r.shift,
        FechaRealizacion: r.date,
        Descripcion: c.task || '',
        Estado: c.state || '',
        EstadoAvance: timelineStr
      });
    }));
    addSheet(wb, corr.length ? corr : [{ OrdenID: '', OrdenNumero: '', SectorID: '', EquipoID: '', EquipoCodigo: '', OrdenTecnicoAsignado: '', Turno: '', FechaRealizacion: '', Descripcion: '', Estado: '', EstadoAvance: '' }], 'Correctivos');

    // 3. Preventivos (detalle individual — formato original mantenido)
    const prev = [];
    let pId = 1;
    history.forEach(r => (r.preventive || []).forEach(p => {
      prev.push({
        OrdenID: pId++,
        OrdenEquipoID: p.codigoTarea || '',
        OrdenEquipoCodigo: p.equipoCodigo || '',
        OrdenEquipoDescripcion: p.equipoDescripcion || '',
        OrdenTecnicoAsignado: (p.technicians || []).join(', '),
        OrdenTurno: r.shift,
        OrdenFechaRealizacion: r.date,
        OrdenFrecuencia: p.frequency || '',
        OrdenDescripcion: p.task || '',
        OrdenComentarios: p.comments || '',
        OrdenCorrectivaAsociada: p.otCorrectivaAsociada || ''
      });
    }));
    addSheet(wb, prev.length ? prev : [{ OrdenID: '', OrdenEquipoID: '', OrdenEquipoCodigo: '', OrdenEquipoDescripcion: '', OrdenTecnicoAsignado: '', OrdenTurno: '', OrdenFechaRealizacion: '', OrdenFrecuencia: '', OrdenDescripcion: '', OrdenComentarios: '', OrdenCorrectivaAsociada: '' }], 'Preventivos');

    // 3b. Resumen Preventivos por Turno (NUEVO V2.0)
    const resumen = [];
    let resId = 1;
    history.forEach(r => {
      const pr = r.preventivosResumen;
      if (!pr) return;
      const asig = pr.asignados, real = pr.realizados;
      if ((asig === '' || asig == null) && (real === '' || real == null)) return;
      resumen.push({
        ResumenID: resId++,
        Fecha: r.date,
        Turno: r.shift,
        Asignados: asig !== '' && asig != null ? Number(asig) : '',
        Realizados: real !== '' && real != null ? Number(real) : ''
      });
    });
    addSheet(wb, resumen.length ? resumen : [{ ResumenID: '', Fecha: '', Turno: '', Asignados: '', Realizados: '' }], 'Resumen Preventivos Turno');

    // 3c. Preventivos por Técnico (NUEVO V2.0)
    const porTec = [];
    let ptId = 1;
    history.forEach(r => {
      (r.preventivosResumen?.porTecnico || []).forEach(t => {
        // V2.4 — Schema multi-técnico: cada grupo tiene `tecnicos: []` y cantidad.
        // Para el Excel, generamos UNA FILA POR TÉCNICO del grupo, con la cantidad
        // del grupo. Así "Juan + Pedro hicieron 4" → 2 filas (una para Juan con 4,
        // otra para Pedro con 4). Esto coincide con cómo se acreditan las stats
        // individuales.
        const tecnicos = t.tecnicos || (t.tecnico ? [t.tecnico] : []);
        const cantidad = Number(t.cantidad) || 0;
        if (tecnicos.length === 0 || cantidad === 0) return;
        tecnicos.forEach(name => {
          porTec.push({
            RegistroID: ptId++,
            Fecha: r.date,
            Turno: r.shift,
            Tecnico: name,
            TecnicoID: findTecnicoId(name),
            Cantidad: cantidad,
            EnGrupoCon: tecnicos.filter(n => n !== name).join(', ') || '—'
          });
        });
      });
    });
    addSheet(wb, porTec.length ? porTec : [{ RegistroID: '', Fecha: '', Turno: '', Tecnico: '', TecnicoID: '', Cantidad: '', EnGrupoCon: '' }], 'Preventivos por Tecnico');

    // 4. Estado de cisternas
    const cist = [];
    let cisId = 1;
    history.forEach(r => {
      const c = r.servicios?.cisternas;
      if (c && (c.nivel || c.estado)) {
        cist.push({
          CisternaEstadoID: cisId++,
          CisternaEstadoNivel: c.nivel || '',
          CisternaEstadoIngreso: c.estado || '',
          CisternaFechaControl: r.date,
          CisternaTurno: r.shift
        });
      }
    });
    addSheet(wb, cist.length ? cist : [{ CisternaEstadoID: '', CisternaEstadoNivel: '', CisternaEstadoIngreso: '', CisternaFechaControl: '', CisternaTurno: '' }], 'Estado de cisternas');

    // 4b. Agua de Pozo
    const pozo = [];
    let pozoId = 1;
    history.forEach(r => {
      const a = r.servicios?.aguaPozo;
      if (!a) return;
      const v3 = a.cloroPozo3, v6 = a.cloroPozo6;
      if ((v3 === '' || v3 == null) && (v6 === '' || v6 == null)) return;
      pozo.push({
        AguaPozoRegistroID: pozoId++,
        AguaPozoFecha: r.date,
        AguaPozoTurno: r.shift,
        NivelCloroPozo3: v3 !== '' && v3 != null ? Number(v3) : '',
        NivelCloroPozo6: v6 !== '' && v6 != null ? Number(v6) : ''
      });
    });
    addSheet(wb, pozo.length ? pozo : [{ AguaPozoRegistroID: '', AguaPozoFecha: '', AguaPozoTurno: '', NivelCloroPozo3: '', NivelCloroPozo6: '' }], 'Agua de Pozo');

    // 5. Compresores
    const comp = [];
    let compId = 1;
    history.forEach(r => (r.servicios?.compresores || []).forEach(c => {
      comp.push({
        CompresorRegistroId: compId++,
        CompresorCodigoEquipo: c.code,
        CompresorEstado: c.state,
        CompresorRegistroTurno: r.shift,
        CompresorRegistroFecha: r.date
      });
    }));
    addSheet(wb, comp.length ? comp : [{ CompresorRegistroId: '', CompresorCodigoEquipo: '', CompresorEstado: '', CompresorRegistroTurno: '', CompresorRegistroFecha: '' }], 'Compresores');

    // 6. Grupos Electrogenos
    const gru = [];
    let gId = 1;
    history.forEach(r => (r.servicios?.gruposElectrogenos || []).forEach(g => {
      gru.push({
        GrupoRegistroId: gId++,
        GrupoCodigoEquipo: g.code,
        GrupoEstado: g.state,
        GrupoRegistroTurno: r.shift,
        GrupoRegistroFecha: r.date
      });
    }));
    addSheet(wb, gru.length ? gru : [{ GrupoRegistroId: '', GrupoCodigoEquipo: '', GrupoEstado: '', GrupoRegistroTurno: '', GrupoRegistroFecha: '' }], 'Grupos Electrogenos');

    // 7. Planta de efluentes — V2.0: NUEVAS columnas (PTEL + Caldera + Ablandadores)
    // Las columnas viejas (Ablandador, TKEmergencia, TK4) se eliminan en este export.
    const planta = [];
    let plId = 1;
    history.forEach(r => {
      const p = r.servicios?.plantaCaldera;
      if (!p) return;
      const foguistas = (p.tecnicos && p.tecnicos.length > 0)
        ? p.tecnicos
        : (p.tecnico ? [p.tecnico] : ['']);
      if (foguistas[0] === '' && !p.estado) return;
      foguistas.forEach(name => {
        planta.push({
          CalderaRegistroID: plId++,
          CalderaRegistroTurno: r.shift,
          CalderaTecnicoNombre: name || '',
          TecnicoID: findTecnicoId(name),
          CalderaFechaControl: r.date,
          CalderaEstado: p.estado || '',
          // PTEL
          PTEL_Caudal_m3h: p.caudal !== '' && p.caudal != null ? Number(p.caudal) : '',
          PTEL_Vacio: p.vacio !== '' && p.vacio != null ? Number(p.vacio) : '',
          PTEL_DeltaT_C: p.deltaT !== '' && p.deltaT != null ? Number(p.deltaT) : '',
          PTEL_TK1_pct: p.tk1 !== '' && p.tk1 != null ? Number(p.tk1) : '',
          PTEL_TK2_pct: p.tk2 !== '' && p.tk2 != null ? Number(p.tk2) : '',
          PTEL_TK7_pct: p.tk7 !== '' && p.tk7 != null ? Number(p.tk7) : '',
          // Caldera
          Caldera_Conductividad_mS: p.conductividadCaldera !== '' && p.conductividadCaldera != null ? Number(p.conductividadCaldera) : '',
          Caldera_pH: p.pHCaldera !== '' && p.pHCaldera != null ? Number(p.pHCaldera) : '',
          // Ablandadores
          Ablandador_Conductividad_mS: p.conductividadAblandador !== '' && p.conductividadAblandador != null ? Number(p.conductividadAblandador) : '',
          Ablandador_pH: p.pHAblandador !== '' && p.pHAblandador != null ? Number(p.pHAblandador) : ''
        });
      });
    });
    addSheet(wb, planta.length ? planta : [{
      CalderaRegistroID: '', CalderaRegistroTurno: '', CalderaTecnicoNombre: '', TecnicoID: '', CalderaFechaControl: '', CalderaEstado: '',
      PTEL_Caudal_m3h: '', PTEL_Vacio: '', PTEL_DeltaT_C: '', PTEL_TK1_pct: '', PTEL_TK2_pct: '', PTEL_TK7_pct: '',
      Caldera_Conductividad_mS: '', Caldera_pH: '',
      Ablandador_Conductividad_mS: '', Ablandador_pH: ''
    }], 'Planta de efluentes');

    // 8. Servicios externo
    const ext = [];
    let eId = 1;
    history.forEach(r => (r.servicios?.proveedores || []).forEach(p => {
      ext.push({
        ServicioExternoID: eId++,
        ServicioExternoNombre: p.provider,
        ServicioExternoTarea: p.task,
        ServicioExternoFecha: r.date,
        ServicioExternoTurno: r.shift
      });
    }));
    addSheet(wb, ext.length ? ext : [{ ServicioExternoID: '', ServicioExternoNombre: '', ServicioExternoTarea: '', ServicioExternoFecha: '', ServicioExternoTurno: '' }], 'Servicios externo');

    // 9. Comentarios adicionales
    const comm = [];
    let comId = 1;
    history.forEach(r => (r.comments || []).forEach(c => {
      comm.push({
        ComentarioAdicionalID: comId++,
        ComentarioAdicionalTexto: c.text,
        ComentarioAdicionalPrioridad: c.priority,
        ComentarioAdicionalFecha: r.date,
        ComentarioAdicionalTurno: r.shift
      });
    }));
    addSheet(wb, comm.length ? comm : [{ ComentarioAdicionalID: '', ComentarioAdicionalTexto: '', ComentarioAdicionalPrioridad: '', ComentarioAdicionalFecha: '', ComentarioAdicionalTurno: '' }], 'Comentarios adicionales');

    // 10. Responsables
    const resp = [];
    let rId = 1;
    history.forEach(r => {
      if (r.responsable) {
        resp.push({
          ResponsableID: rId++,
          ResponsableNombre: r.responsable,
          ResponsableTurno: r.shift,
          ResponsableFecha: r.date
        });
      }
    });
    addSheet(wb, resp.length ? resp : [{ ResponsableID: '', ResponsableNombre: '', ResponsableTurno: '' }], 'Responsables');

    XLSX.writeFile(wb, `ReporteDiario_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const exportSingleSheet = (kind) => {
    if (!history.length) { alert('No hay reportes guardados.'); return; }
    if (kind === 'correctivos') {
      const rows = [];
      let i = 1;
      history.forEach(r => (r.corrective || []).forEach(c => {
        const timelineStr = (c.timeline || [])
          .map(e => `[${e.date} ${e.shift}${e.author ? ' / ' + e.author : ''}] ${e.text}`)
          .join('\n');
        rows.push({
          OrdenID: i++,
          OrdenNumero: c.ot || '',
          SectorID: '',
          EquipoID: '',
          EquipoCodigo: c.equipoCodigo || '',
          OrdenTecnicoAsignado: (c.technicians || []).join(', '),
          Turno: r.shift,
          FechaRealizacion: r.date,
          Descripcion: c.task || '',
          Estado: c.state || '',
          EstadoAvance: timelineStr
        });
      }));
      if (!rows.length) { alert('Sin correctivos.'); return; }
      downloadSingle(rows, 'Correctivos', `Correctivos_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } else if (kind === 'preventivos') {
      const rows = [];
      let i = 1;
      history.forEach(r => (r.preventive || []).forEach(p => rows.push({
        OrdenID: i++,
        OrdenEquipoID: p.codigoTarea || '',
        OrdenEquipoCodigo: p.equipoCodigo || '',
        OrdenEquipoDescripcion: p.equipoDescripcion || '',
        OrdenTecnicoAsignado: (p.technicians || []).join(', '),
        OrdenTurno: r.shift,
        OrdenFechaRealizacion: r.date,
        OrdenFrecuencia: p.frequency || '',
        OrdenDescripcion: p.task || '',
        OrdenComentarios: p.comments || '',
        OrdenCorrectivaAsociada: p.otCorrectivaAsociada || ''
      })));
      if (!rows.length) { alert('Sin preventivos.'); return; }
      downloadSingle(rows, 'Preventivos', `Preventivos_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } else if (kind === 'comentarios') {
      // NUEVO V2.0
      const rows = [];
      let i = 1;
      history.forEach(r => (r.comments || []).forEach(c => rows.push({
        ComentarioAdicionalID: i++,
        Fecha: r.date,
        Turno: r.shift,
        ComentarioAdicionalTexto: c.text || '',
        ComentarioAdicionalPrioridad: c.priority || ''
      })));
      if (!rows.length) { alert('Sin comentarios.'); return; }
      downloadSingle(rows, 'Comentarios', `Comentarios_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } else if (kind === 'proveedores') {
      // NUEVO V2.0
      const rows = [];
      let i = 1;
      history.forEach(r => (r.servicios?.proveedores || []).forEach(p => rows.push({
        ServicioExternoID: i++,
        Fecha: r.date,
        Turno: r.shift,
        Proveedor: p.provider || '',
        Tarea: p.task || ''
      })));
      if (!rows.length) { alert('Sin proveedores.'); return; }
      downloadSingle(rows, 'Proveedores', `Proveedores_${new Date().toISOString().slice(0, 10)}.xlsx`);
    }
  };

  // Punto 2 — Pantalla de bloqueo por versión desactualizada (hard gate).
  // Early-return: si la versión local es vieja, no se renderiza nada del árbol
  // normal — no hay forma de cargar ni guardar. El link redirige a producción.
  if (versionBlocked) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4"
           style={{ fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif" }}>
        <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full p-7 text-center">
          <div className="w-14 h-14 rounded-full bg-amber-100 flex items-center justify-center mx-auto mb-4">
            <AlertTriangle className="w-7 h-7 text-amber-600" />
          </div>
          <h1 className="text-lg font-bold text-slate-900 mb-2">
            Versión desactualizada
          </h1>
          <p className="text-sm text-slate-600 leading-relaxed mb-1">
            Estás usando una versión vieja de la app
            (<span className="num font-semibold">{APP_VERSION}</span>) que ya no está habilitada para cargar reportes.
          </p>
          <p className="text-sm text-slate-600 leading-relaxed mb-5">
            Abrí la versión actual desde el siguiente enlace y, si tenés un acceso directo viejo en el escritorio, reemplazalo por este:
          </p>
          <a href={PROD_URL}
             className="inline-flex items-center justify-center w-full px-4 py-3 text-sm font-semibold text-white bg-sky-600 hover:bg-sky-700 rounded-xl transition mb-3">
            Abrir la versión actual
          </a>
          <div className="text-xs text-slate-400 break-all select-all">
            {PROD_URL}
          </div>
        </div>
      </div>
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // #62 — MODO SOLO-EXTRAS (deploy de Facilities)
  // ═══════════════════════════════════════════════════════════════════
  // Early-return con árbol propio, mismo patrón que `versionBlocked`. Es a
  // propósito que sea un return y no condicionales salpicados por el header y
  // las tabs: así el reporte de turno NO SE RENDERIZA POR CONSTRUCCIÓN, y no
  // por acordarse de envolver ocho lugares. Un lugar nuevo que alguien agregue
  // más adelante al árbol normal tampoco aparece acá.
  //
  // Va DESPUÉS de todos los hooks de App y después del gate de versión: el
  // gate aplica igual en este deploy, porque los dos comparten `app_config`.
  //
  // Recordatorio, porque es la parte que se malinterpreta: esto oculta, no
  // protege. El código del reporte sigue en el bundle. Ver #47.
  if (APP_MODE === 'extras') {
    const sinUsuarios = EXTRAS_USUARIOS.length === 0;
    return (
      <div className="min-h-screen bg-slate-50" style={{ fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif" }}>
        <style>{`
          @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap');
          body { font-family: 'Plus Jakarta Sans', system-ui, sans-serif; }
          .num { font-family: 'JetBrains Mono', monospace; font-feature-settings: 'tnum', 'zero'; }
        `}</style>

        <header className="bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900 text-white border-b border-slate-800 sticky top-0 z-40 shadow-lg">
          <div className="max-w-[1600px] mx-auto px-6 py-2 md:py-4 flex items-center justify-between flex-wrap gap-2 md:gap-3">
            <div className="flex items-center gap-2 md:gap-4">
              <div className="w-9 h-9 md:w-14 md:h-14 rounded-xl bg-white ring-1 ring-slate-200 flex items-center justify-center p-1 md:p-1.5 shadow-sm flex-shrink-0">
                <img src="/logo-biomas.jpg" alt="Biomas" className="w-full h-full object-contain"
                  onError={(e) => { e.currentTarget.style.display = 'none'; }} />
              </div>
              <div className="min-w-0">
                <h1 className="text-sm md:text-lg font-bold tracking-tight leading-tight flex items-center gap-2">
                  Horas Extras · {APP_SECTOR}
                  <span className="px-1.5 py-0.5 bg-slate-700 text-slate-200 rounded text-[10px] font-semibold num md:hidden">{APP_VERSION}</span>
                </h1>
                <p className="hidden md:block text-[11px] text-slate-300 mt-0.5">
                  Solicitud y aprobación de horas extras
                  <span className="ml-2 px-1.5 py-0.5 bg-slate-700 text-slate-200 rounded text-[10px] font-semibold num">{APP_VERSION}</span>
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2 text-xs">
              {supabaseConfigured ? (
                connError
                  ? <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-red-500/20 text-red-200 rounded ring-1 ring-red-400/30">
                    <CloudOff className="w-3.5 h-3.5" /><span className="hidden md:inline">Sin conexión</span>
                  </span>
                  : <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-emerald-500/20 text-emerald-200 rounded ring-1 ring-emerald-400/30">
                    <Cloud className="w-3.5 h-3.5" /><span className="hidden md:inline">Supabase conectado</span>
                  </span>
              ) : (
                <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-amber-500/20 text-amber-200 rounded ring-1 ring-amber-400/30">
                  <Settings className="w-3.5 h-3.5" /><span className="hidden md:inline">Modo local</span>
                </span>
              )}
              <button onClick={refresh} className="p-1.5 hover:bg-white/10 rounded transition" title="Refrescar">
                <RefreshCw className="w-4 h-4" />
              </button>
              {extrasUser && (
                <>
                  <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded ring-1 font-semibold ${extrasUser.rol === 'jefe' ? 'bg-teal-500/30 text-teal-100 ring-teal-400/50' : 'bg-cyan-500/30 text-cyan-100 ring-cyan-400/50'}`}
                        title={`${extrasUser.nombre} · ${extrasUser.rol === 'jefe' ? 'jefe' : EXTRAS_ETIQUETA_ENCARGADO}`}>
                    <Timer className="w-3.5 h-3.5" />
                    <span className="hidden md:inline">{extrasUser.nombre}</span>
                    <span className="md:hidden">{extrasUser.rol === 'jefe' ? 'JEFE' : 'EXTRAS'}</span>
                  </span>
                  <button onClick={handleExtrasLogout}
                    className="inline-flex items-center gap-1 px-2 py-1 bg-white/10 hover:bg-white/20 text-slate-200 rounded transition text-[10px]"
                    title="Cerrar sesión">
                    <LogOut className="w-3 h-3" />Salir
                  </button>
                </>
              )}
            </div>
          </div>
        </header>

        <main className="max-w-[1600px] mx-auto px-6 py-6">
          {extrasUser ? (
            <ExtrasView
              sesion={extrasUser}
              extras={extras}
              extrasLoading={extrasLoading}
              extrasError={extrasError}
              onAdd={handleExtrasAdd}
              onUpdate={handleExtrasUpdate}
              onRefresh={loadExtras}
            />
          ) : (
            /* Pantalla de ingreso. Acá vive el CHIP PRE-LOGIN: es el
               equivalente del botón naranja del header en la app completa.
               No puede decir de quién son las pendientes —todavía no hay
               sesión— así que habla del sector. El dato preciso por persona
               lo da el contador post-login, ya adentro. */
            <div className="max-w-md mx-auto mt-8">
              <Card className="p-7 text-center">
                <div className="w-14 h-14 rounded-full bg-cyan-100 flex items-center justify-center mx-auto mb-4">
                  <Timer className="w-7 h-7 text-cyan-600" />
                </div>
                <h2 className="text-lg font-bold text-slate-900 mb-1">Horas extras · {APP_SECTOR}</h2>
                <p className="text-sm text-slate-600 leading-relaxed mb-5">
                  Ingresá con tu usuario para cargar o resolver solicitudes.
                </p>

                {extrasPendientes && (
                  <div className="mb-5 inline-flex items-center gap-2 px-3 py-2 bg-orange-50 text-orange-800 ring-1 ring-orange-300 rounded-lg text-xs font-semibold">
                    <AlertTriangle className="w-4 h-4 flex-shrink-0" />
                    Hay solicitudes sin resolver en el sector
                  </div>
                )}

                {sinUsuarios ? (
                  /* Estado esperado al terminar la etapa 1 de #62: la
                     arquitectura está deployada pero el catálogo del sector
                     todavía no se cargó. Se dice explícito en vez de mostrar
                     un login que rebota cualquier credencial. */
                  <div className="flex items-start gap-2 px-3 py-2 bg-amber-50 border border-amber-200 rounded-lg text-left">
                    <AlertTriangle className="w-4 h-4 text-amber-600 flex-shrink-0 mt-0.5" />
                    <span className="text-xs text-amber-800 leading-relaxed">
                      Todavía no hay usuarios cargados para {APP_SECTOR}. El módulo queda operativo
                      cuando se den de alta el jefe y los Encargado/Supervisor con su personal a cargo.
                    </span>
                  </div>
                ) : (
                  <button onClick={() => setExtrasLoginOpen(true)}
                    className="inline-flex items-center justify-center gap-2 w-full px-4 py-3 text-sm font-semibold text-white bg-cyan-600 hover:bg-cyan-500 rounded-xl transition">
                    <Lock className="w-4 h-4" />Ingresar
                  </button>
                )}
              </Card>
            </div>
          )}
        </main>

        {extrasLoginOpen && (
          <ExtrasLoginDialog
            onConfirm={handleExtrasLogin}
            onCancel={() => setExtrasLoginOpen(false)}
          />
        )}
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50" style={{ fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap');
        body { font-family: 'Plus Jakarta Sans', system-ui, sans-serif; }
        .num { font-family: 'JetBrains Mono', monospace; font-feature-settings: 'tnum', 'zero'; }
      `}</style>

      {/* HEADER — V2.6: sticky para acceso permanente a tabs y modo admin durante scroll */}
      <header className="bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900 text-white border-b border-slate-800 sticky top-0 z-40 shadow-lg">
        <div className="max-w-[1600px] mx-auto px-6 py-2 md:py-4 flex items-center justify-between flex-wrap gap-2 md:gap-3">
          <div className="flex items-center gap-2 md:gap-4">
            {/* LOGO BIOMAS — V2.0: reemplaza el icono del casco. Fondo blanco según pedido. */}
            <div className="w-9 h-9 md:w-14 md:h-14 rounded-xl bg-white ring-1 ring-slate-200 flex items-center justify-center p-1 md:p-1.5 shadow-sm flex-shrink-0">
              <img
                src="/logo-biomas.jpg"
                alt="Biomas"
                className="w-full h-full object-contain"
                onError={(e) => {
                  // Si el logo no carga (por ej. en preview antes del primer deploy con el archivo),
                  // fallback al icono original sin romper el layout.
                  e.currentTarget.style.display = 'none';
                  e.currentTarget.parentElement.classList.remove('bg-white', 'ring-slate-200');
                  e.currentTarget.parentElement.classList.add('bg-sky-500/20', 'ring-sky-400/40');
                  const icon = document.createElement('div');
                  icon.className = 'w-5 h-5 text-sky-300';
                  e.currentTarget.parentElement.appendChild(icon);
                }}
              />
            </div>
            <div className="min-w-0">
              <h1 className="text-sm md:text-lg font-bold tracking-tight leading-tight flex items-center gap-2">
                Reporte Diario de Mantenimiento
                <span className="px-1.5 py-0.5 bg-slate-700 text-slate-200 rounded text-[10px] font-semibold num md:hidden">{APP_VERSION}</span>
              </h1>
              <p className="hidden md:block text-[11px] text-slate-300 mt-0.5">
                Sistema integral · carga, dashboard, estadísticas y exportación
                <span className="ml-2 px-1.5 py-0.5 bg-slate-700 text-slate-200 rounded text-[10px] font-semibold num">{APP_VERSION}</span>
              </p>
            </div>
          </div>
          <div className="flex items-center gap-4 text-xs">
            <div className="flex flex-row md:flex-col items-center md:items-end gap-2 md:gap-1">
              <div className="flex items-center gap-2">
                {supabaseConfigured ? (
                  connError
                    ? <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-red-500/20 text-red-200 rounded ring-1 ring-red-400/30">
                      <CloudOff className="w-3.5 h-3.5" /><span className="hidden md:inline">Sin conexión</span>
                    </span>
                    : <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-emerald-500/20 text-emerald-200 rounded ring-1 ring-emerald-400/30">
                      <Cloud className="w-3.5 h-3.5" /><span className="hidden md:inline">Supabase conectado</span>
                    </span>
                ) : (
                  <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-amber-500/20 text-amber-200 rounded ring-1 ring-amber-400/30">
                    <Settings className="w-3.5 h-3.5" /><span className="hidden md:inline">Modo local</span>
                  </span>
                )}
                <button onClick={refresh} className="p-1.5 hover:bg-white/10 rounded transition" title="Refrescar">
                  <RefreshCw className="w-4 h-4" />
                </button>
              </div>
              {/* v3.25 — Los tres roles (Admin · Planificador · Extras) en FILA
                  horizontal debajo del badge de Supabase. Antes iban apilados y
                  con el tercero quedaba un header de tres pisos.
                  Van acá y no al lado de "Guardar reporte" porque ese bloque está
                  condicionado a `tab === 'form'`: colgados de ahí desaparecerían
                  en Dashboard, Estadísticas e Histórico, que es justo desde donde
                  se necesita entrar a Extras. */}
              <div className="flex flex-wrap items-center justify-end gap-1.5">
              {adminMode ? (
                <div className="flex items-center gap-1.5">
                  <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-red-500/30 text-red-100 rounded ring-1 ring-red-400/50 font-semibold">
                    <Shield className="w-3.5 h-3.5" /><span className="md:hidden">ADMIN</span><span className="hidden md:inline">MODO ADMIN</span>
                  </span>
                  <button onClick={handleAdminLogout}
                    className="inline-flex items-center gap-1 px-2 py-1 bg-white/10 hover:bg-white/20 text-slate-200 rounded transition text-[10px]"
                    title="Salir de modo admin">
                    <LogOut className="w-3 h-3" /><span className="hidden md:inline">Salir</span>
                  </button>
                </div>
              ) : (
                <button onClick={() => setAdminLoginOpen(true)}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-white/5 hover:bg-white/10 text-slate-300 rounded ring-1 ring-white/10 transition">
                  <Lock className="w-3.5 h-3.5" />Admin
                </button>
              )}
              {/* #42 (Fase 1) — Rol planificador, separado de admin */}
              {poolMode ? (
                <div className="flex items-center gap-1.5">
                  <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-violet-500/30 text-violet-100 rounded ring-1 ring-violet-400/50 font-semibold">
                    <Inbox className="w-3.5 h-3.5" /><span className="md:hidden">PLANIF</span><span className="hidden md:inline">PLANIFICADOR</span>
                  </span>
                  <button onClick={handlePoolLogout}
                    className="inline-flex items-center gap-1 px-2 py-1 bg-white/10 hover:bg-white/20 text-slate-200 rounded transition text-[10px]"
                    title="Salir del rol planificador">
                    <LogOut className="w-3 h-3" /><span className="hidden md:inline">Salir</span>
                  </button>
                </div>
              ) : (
                <button onClick={() => setPoolLoginOpen(true)}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-white/5 hover:bg-white/10 text-slate-300 rounded ring-1 ring-white/10 transition">
                  <Inbox className="w-3.5 h-3.5" />Planificador
                </button>
              )}
              {/* #46 (v3.25) — Rol Extras. Mismo patrón que el pool: la solapa no
                  existe hasta loguearse. El badge muestra el nombre, no el usuario:
                  lo que importa es quién es, no con qué mail entró.

                  #50 (v3.26) — El botón se pone NARANJA cuando hay solicitudes
                  sin resolver (`estado = 'pendiente'` y no anuladas).
                  Aplica SOLO al botón con el rol inactivo: una vez logueado, el
                  badge conserva su teal/cyan por rol. El aviso existe justamente
                  para que se vea ANTES de entrar — es la única forma de que el
                  jefe sepa que tiene algo que aprobar sin loguearse.
                  Lo ven todos, no solo el jefe: sin sesión la app no sabe quién
                  está del otro lado, así que discriminar por rol es imposible
                  hasta el login, momento en el que el aviso ya no se muestra.
                  Naranja y no ámbar a propósito: el ámbar ya está tomado en este
                  mismo header por el badge "Modo local". */}
              {extrasUser ? (
                <div className="flex items-center gap-1.5">
                  <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded ring-1 font-semibold ${extrasUser.rol === 'jefe' ? 'bg-teal-500/30 text-teal-100 ring-teal-400/50' : 'bg-cyan-500/30 text-cyan-100 ring-cyan-400/50'}`}
                        title={`${extrasUser.nombre} · ${extrasUser.rol}`}>
                    <Timer className="w-3.5 h-3.5" />
                    <span className="md:hidden">{extrasUser.rol === 'jefe' ? 'JEFE' : 'EXTRAS'}</span>
                    <span className="hidden md:inline">{extrasUser.rol === 'jefe' ? 'EXTRAS · JEFE' : 'EXTRAS'}</span>
                  </span>
                  <button onClick={handleExtrasLogout}
                    className="inline-flex items-center gap-1 px-2 py-1 bg-white/10 hover:bg-white/20 text-slate-200 rounded transition text-[10px]"
                    title="Cerrar sesión de Extras">
                    <LogOut className="w-3 h-3" /><span className="hidden md:inline">Salir</span>
                  </button>
                </div>
              ) : (
                <button onClick={() => setExtrasLoginOpen(true)}
                  className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded ring-1 transition ${
                    extrasPendientes
                      ? 'bg-orange-500/30 hover:bg-orange-500/40 text-orange-100 ring-orange-400/50 font-semibold'
                      : 'bg-white/5 hover:bg-white/10 text-slate-300 ring-white/10'
                  }`}
                  title={extrasPendientes
                    ? 'Hay solicitudes de horas extras sin resolver'
                    : 'Ingresar al módulo de horas extras'}>
                  <Timer className="w-3.5 h-3.5" />Extras
                </button>
              )}
              </div>
            </div>
            <div className="text-right">
              <div className="num text-sm font-semibold text-white capitalize">
                <span className="hidden md:inline">{new Date().toLocaleDateString('es-AR', { weekday: 'long', day: 'numeric', month: 'long' })}</span>
                <span className="md:hidden">{(() => {
                  const d = new Date();
                  const dias = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
                  const dd = String(d.getDate()).padStart(2, '0');
                  const mm = String(d.getMonth() + 1).padStart(2, '0');
                  return `${dias[d.getDay()]} ${dd}/${mm}`;
                })()}</span>
              </div>
              <div className="hidden md:block text-slate-300"><span className="num">{history.length}</span> {history.length === 1 ? 'reporte' : 'reportes'}</div>
            </div>
            {/* v3.13 (#12) — Guardar reporte en el header global, visible siempre (sin scroll)
                solo cuando la tab activa es "Cargar Reporte". Botón grande, en línea a la
                derecha (no fuerza 2da línea para no ensanchar el header). El saveMsg lo
                acompaña a la izquierda. Limpiar y Eliminar NO suben acá: van al FAB admin
                dentro de FormView (Limpiar fue el detonante del incidente 21/05). */}
            {tab === 'form' && (
              <div className="flex items-center gap-2 md:gap-3 pl-2 md:pl-4 md:border-l md:border-white/10">
                {saveMsg && (
                  <span className={`hidden md:inline text-xs font-medium max-w-[180px] truncate ${saveMsg.startsWith('Error') ? 'text-red-300' : saveMsg.startsWith('✓') ? 'text-emerald-300' : 'text-slate-300'}`}>
                    {saveMsg}
                  </span>
                )}
                {/* #9 (v3.14) — el botón Guardar se oculta si el reporte es read-only para
                    no-admin (anterior a ayer). El banner del form explica por qué. */}
                {!(!adminMode && !isWithinEditWindow(report.date)) && (
                  <button onClick={saveReport} disabled={saving}
                    className="inline-flex items-center gap-2 px-4 md:px-6 py-2 md:py-2.5 bg-sky-500 hover:bg-sky-400 text-white rounded-xl font-bold text-sm md:text-base transition disabled:opacity-50 shadow-md flex-shrink-0"
                    title="Guardar reporte">
                    <Save className="w-5 h-5" /><span className="hidden sm:inline">Guardar reporte</span><span className="sm:hidden">Guardar</span>
                  </button>
                )}
              </div>
            )}
          </div>
        </div>
        <div className="max-w-[1600px] mx-auto px-6">
          <nav className="flex gap-1 overflow-x-auto">
            {[
              { id: 'form', label: 'Cargar Reporte', icon: ClipboardList },
              { id: 'dashboard', label: 'Dashboard', icon: BarChart3 },
              { id: 'stats', label: 'Estadísticas', icon: TrendingUp },
              { id: 'history', label: 'Histórico & Excel', icon: FileSpreadsheet },
              // #42 (Fase 1) — la pestaña del pool solo existe con el rol activo
              ...(poolMode ? [{ id: 'pool', label: 'Pool de OTs', icon: Inbox }] : []),
              // #46 (v3.25) — ídem Extras: patrón del pool, la solapa aparece al loguearse
              ...(extrasUser ? [{ id: 'extras', label: 'Extras', icon: Timer }] : [])
            ].map(t => (
              <button key={t.id} onClick={() => setTab(t.id)}
                className={`flex items-center gap-2 px-5 py-2.5 text-sm font-medium border-b-2 transition whitespace-nowrap ${tab === t.id ? 'border-sky-400 text-white' : 'border-transparent text-slate-300 hover:text-white'}`}>
                <t.icon className="w-4 h-4" />{t.label}
              </button>
            ))}
          </nav>
        </div>
      </header>

      {/* SETUP NOTICE */}
      {!supabaseConfigured && (
        <div className="bg-amber-50 border-b border-amber-200">
          <div className="max-w-[1600px] mx-auto px-6 py-2 text-xs text-amber-800 flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 flex-shrink-0" />
            <span>
              <strong>Modo local:</strong> los datos se guardan solo en este navegador. Para compartir entre turnos, configurá Supabase
              (editá <code className="bg-amber-100 px-1 rounded">SUPABASE_URL</code> y <code className="bg-amber-100 px-1 rounded">SUPABASE_ANON_KEY</code> al inicio del archivo).
            </span>
          </div>
        </div>
      )}

     <main className="max-w-[1600px] mx-auto px-6 py-5">
       {loading && <div className="text-center text-slate-500 py-20">Cargando…</div>}
        {!loading && tab === 'form' && <FormView
          report={report}
          setReport={setReportAndResetOverride}
          onSave={saveReport}
          saveMsg={saveMsg}
          setSaveMsg={setSaveMsg}
          saving={saving}
          history={history}
          adminMode={adminMode}
          originalReport={originalReport}
          setOriginalReport={setOriginalReport}                                
          onDeleteReport={() => requestDeleteReport(report.date, report.shift, 'form')}
          otErrorIndices={otErrorIndices}
          otErrorType={otErrorType}
          onClearOtErrors={() => { setOtErrorIndices(new Set()); setOtErrorType(''); }}
        />}
        {!loading && tab === 'dashboard' && <DashboardView
          report={dashboardOverride || report}
          history={history}
          activeReport={report}
          dashboardOverride={dashboardOverride}
          setDashboardOverride={setDashboardOverride}
          adminMode={adminMode}
          onEditFromDashboard={editFromDashboard}
        />}
        {!loading && tab === 'stats' && <StatsView history={history} adminMode={adminMode} />}
        {!loading && tab === 'history' && <HistoryView history={history}
          onExportCorrectives={() => exportSingleSheet('correctivos')}
          onExportPreventives={() => exportSingleSheet('preventivos')}
          onExportComments={() => exportSingleSheet('comentarios')}
          onExportProviders={() => exportSingleSheet('proveedores')}
          onExportFull={exportFull}
          adminMode={adminMode}
          onDeleteReport={(date, shift) => requestDeleteReport(date, shift, 'history')}
        />}
        {!loading && tab === 'extras' && extrasUser && <ExtrasView
          sesion={extrasUser}
          extras={extras}
          extrasLoading={extrasLoading}
          extrasError={extrasError}
          onAdd={handleExtrasAdd}
          onUpdate={handleExtrasUpdate}
          onRefresh={loadExtras}
        />}

        {!loading && tab === 'pool' && poolMode && <PoolView
          pool={pool}
          poolLoading={poolLoading}
          poolError={poolError}
          onAdd={handlePoolAdd}
          onAnular={handlePoolAnular}
          onRefresh={loadPool}
        />}
      </main>

      {/* V2.5 — Modal de confirmación cuando se quiere guardar con entradas vacías */}
      {emptyConfirm && (
        <EmptyEntriesConfirmDialog
          emptyCorr={emptyConfirm.emptyCorr}
          emptyPrev={emptyConfirm.emptyPrev}
          onConfirm={handleConfirmEmpty}
          onCancel={handleCancelEmpty}
        />
      )}

      {/* V2.6 — Modal de login admin */}
      {adminLoginOpen && (
        <AdminLoginDialog
          onConfirm={handleAdminLogin}
          onCancel={() => setAdminLoginOpen(false)}
        />
      )}

      {/* #42 (Fase 1) — Modal de login del rol planificador */}
      {poolLoginOpen && (
        <PoolLoginDialog
          onConfirm={handlePoolLogin}
          onCancel={() => setPoolLoginOpen(false)}
        />
      )}

      {/* #46 (v3.25) — Modal de login de Extras (usuario + password) */}
      {extrasLoginOpen && (
        <ExtrasLoginDialog
          onConfirm={handleExtrasLogin}
          onCancel={() => setExtrasLoginOpen(false)}
        />
      )}

      {/* V2.6 — Modal de confirmación de eliminación de reporte */}
      {deleteReportConfirm && (
        <DeleteReportConfirmDialog
          date={deleteReportConfirm.date}
          shift={deleteReportConfirm.shift}
          onConfirm={confirmDeleteReport}
          onCancel={cancelDeleteReport}
        />
      )}

      {/* V2.8 — Modal de conflictos de OT cerrada (carry-over stale guard).
          Se abre cuando al guardar se detecta que el reporte contiene OTs en estado
          Sin Iniciar/En Curso que YA están como Realizada en otro reporte del history. */}
      {closedConflicts && (
        <ClosedConflictDialog
          conflicts={closedConflicts.conflicts}
          adminMode={adminMode}
          onResolve={handleConflictResolve}
          onCancel={handleConflictCancel}
        />
      )}

      {/* V2.9 — Modal de propagación admin retroactiva.
          Se abre al guardar un reporte histórico (admin) si hay cambios de state
          con reportes posteriores afectados. Cambios solo de timeline propagan
          silenciosamente (decisión B). */}
      {propagationModal && (
        <PropagationModal
          diffs={propagationModal.diffs}
          affectedReports={propagationModal.affectedReports}
          onConfirm={handlePropagationConfirm}
          onCancel={handlePropagationCancel}
        />
      )}

      {/* V3.3 — Modal guard de sobreescritura (BACKLOG #20).
          Se abre cuando el destino ya tiene un reporte guardado con datos
          y es distinto del que se tenía abierto originalmente. */}
      {overwriteConfirm && (
        <OverwriteConfirmDialog
          date={overwriteConfirm.date}
          shift={overwriteConfirm.shift}
          existingN={overwriteConfirm.existingN}
          onConfirm={() => {
            const rts = overwriteConfirm.reportToSave;
            setOverwriteConfirm(null);
            // V3.3 — Reintentar el guardado con la sobreescritura ya confirmada.
            // NO tocamos originalReport (tiene semántica de snapshot admin); usamos
            // el flag overwriteConfirmed para que el guard no vuelva a disparar.
            doSaveReport(rts, true);
          }}
          onCancel={() => { setOverwriteConfirm(null); setSaveMsg(''); }}
        />
      )}

      {concurrencyConflict && (
        <ConcurrencyConflictDialog
          date={concurrencyConflict.reportToSave.date}
          shift={concurrencyConflict.reportToSave.shift}
          onDiscard={async () => {
            // Acción segura (default): descartar lo que tengo en pantalla y recargar el fresco de la base.
            const { date, shift } = concurrencyConflict.reportToSave;
            setConcurrencyConflict(null);
            await refresh();
            const fresh = (await storage.list()).map(hydrate).find(r => r.date === date && r.shift === shift);
            if (fresh) {
              setReport(fresh);
              setOriginalReport(JSON.parse(JSON.stringify(fresh)));
              setSaveMsg('Recargado desde la base. Revisá los datos antes de volver a guardar.');
              setTimeout(() => setSaveMsg(''), 4000);
            }
          }}
          onOverwrite={() => {
            // Acción destructiva: pisar igual. Reintenta con el flag de concurrencia confirmada.
            const rts = concurrencyConflict.reportToSave;
            setConcurrencyConflict(null);
            doSaveReport(rts, false, true);
          }}
        />
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// V2.6 — MODAL DE LOGIN ADMIN
// Pide el password de administrador antes de activar el modo admin.
// ═══════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════
// #42 (Fase 1) — POOL DE OTs
// ═══════════════════════════════════════════════════════════════════
// Registro paralelo de las solicitudes de trabajo correctivas. Existe porque
// las OTs nacen FUERA de la app (papel preimpreso emitido por los sectores) y
// la app solo conoce las que alguien cargó en el reporte de un turno: las que
// nadie carga no existen en ningún lado del sistema. El pool es el DENOMINADOR.
//
// Arquitectura (i): el pool guarda SOLO altas del planificador. Las OTs que
// entran por los reportes NO se importan acá; el universo se deriva al vuelo
// como pool ∪ OTs de `reportes`. Por eso no hay duplicados posibles: si el
// planificador carga una OT que ya está en un reporte, la unión la dedupea por
// clave canónica y no pasa nada. La Fase 1 no necesita avisar nada al respecto.
//
// Clase A = la ejecuta Mantenimiento. Clase B = la ejecuta otra área; ocupa el
// número y nada más. El ejecutor es INDEPENDIENTE del prefijo de sector.
function PoolLoginDialog({ onConfirm, onCancel }) {
  const [pwd, setPwd] = useState('');
  const [error, setError] = useState('');
  const inputRef = useRef(null);

  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onCancel(); };
    document.addEventListener('keydown', onKey);
    inputRef.current?.focus();
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel]);

  const submit = () => {
    const ok = onConfirm(pwd);
    if (!ok) setError('Password incorrecto');
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-2xl max-w-sm w-full p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-violet-600 flex items-center justify-center flex-shrink-0">
            <Inbox className="w-5 h-5 text-white" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">Rol planificador</h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              Habilita el registro de solicitudes de trabajo en el pool de OTs.
            </p>
          </div>
        </div>
        <input ref={inputRef} type="password" value={pwd}
          onChange={e => { setPwd(e.target.value); setError(''); }}
          onKeyDown={e => { if (e.key === 'Enter') submit(); }}
          placeholder="Password"
          className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-violet-500/40 focus:border-violet-500 transition" />
        {error && <div className="mt-2 text-xs text-red-600">{error}</div>}
        <div className="flex items-center justify-end gap-2 mt-5">
          <button onClick={onCancel}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition">
            Cancelar
          </button>
          <button onClick={submit}
            className="px-4 py-2 text-sm font-semibold text-white bg-violet-600 hover:bg-violet-500 rounded-lg transition">
            Activar
          </button>
        </div>
      </div>
    </div>
  );
}

// Anulación con motivo obligatorio. Es el ÚNICO camino de corrección: el
// planificador no edita ni cierra OTs. Soft-delete, nunca DELETE — el CHECK
// `ordenes_pool_anulacion` de la tabla también lo exige del lado de la base.
function PoolAnularDialog({ row, onConfirm, onCancel }) {
  const [motivo, setMotivo] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const inputRef = useRef(null);

  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onCancel(); };
    document.addEventListener('keydown', onKey);
    inputRef.current?.focus();
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel]);

  const submit = async () => {
    const m = motivo.trim();
    if (!m) { setError('El motivo es obligatorio'); return; }
    setBusy(true);
    try { await onConfirm(m); }
    catch (e) { setError(e.message || 'No se pudo anular'); setBusy(false); }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-2xl max-w-sm w-full p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-amber-500 flex items-center justify-center flex-shrink-0">
            <Ban className="w-5 h-5 text-white" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">
              Anular <span className="num">{row.ot}</span>
            </h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              La OT queda registrada como anulada, con el motivo. No se borra.
            </p>
          </div>
        </div>
        <input ref={inputRef} type="text" value={motivo}
          onChange={e => { setMotivo(e.target.value); setError(''); }}
          onKeyDown={e => { if (e.key === 'Enter') submit(); }}
          placeholder="Motivo de la anulación"
          className={inputCls} />
        {error && <div className="mt-2 text-xs text-red-600">{error}</div>}
        <div className="flex items-center justify-end gap-2 mt-5">
          <button onClick={onCancel} disabled={busy}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition disabled:opacity-50">
            Cancelar
          </button>
          <button onClick={submit} disabled={busy}
            className="px-4 py-2 text-sm font-semibold text-white bg-amber-600 hover:bg-amber-500 rounded-lg transition disabled:opacity-50">
            {busy ? 'Anulando…' : 'Anular'}
          </button>
        </div>
      </div>
    </div>
  );
}

function PoolView({ pool, poolLoading, poolError, onAdd, onAnular, onRefresh }) {
  const [ot, setOt] = useState('');
  const [clase, setClase] = useState('A');
  const [ejecutor, setEjecutor] = useState('');
  const [fecha, setFecha] = useState(todayLocalISO());
  const [turno, setTurno] = useState(currentShiftFromClock());
  const [descripcion, setDescripcion] = useState('');
  const [cargadaPor, setCargadaPor] = useState('');
  const [msg, setMsg] = useState('');
  const [saving, setSaving] = useState(false);
  const [filtroSector, setFiltroSector] = useState('');
  const [verAnuladas, setVerAnuladas] = useState(false);
  const [anularTarget, setAnularTarget] = useState(null);

  const activas = pool.filter(r => !r.anulada_at);
  const listado = pool
    .filter(r => (verAnuladas ? true : !r.anulada_at))
    .filter(r => (filtroSector ? r.sector === filtroSector : true));

  const submit = async () => {
    // canonOT nunca adivina el sector: si no matchea el catálogo devuelve ''.
    const canon = canonOT(ot);
    if (!canon || !isValidOT(canon)) { setMsg('Error: elegí sector y número de OT.'); return; }
    if (!fecha) { setMsg('Error: falta la fecha.'); return; }
    if (isFutureDate(fecha)) { setMsg('Error: no se puede registrar una OT con fecha futura.'); return; }
    if (!turno) { setMsg('Error: falta el turno de origen.'); return; }

    const [sec, num] = canon.split('-');
    setSaving(true);
    setMsg('');
    try {
      await onAdd({
        ot: canon,
        sector: sec,
        correlativo: parseInt(num, 10),
        clase,
        // El ejecutor solo aplica a Clase B; en Clase A ejecuta Mantenimiento por definición.
        ejecutor: clase === 'B' ? (ejecutor.trim() || null) : null,
        fecha,
        turno_origen: turno,
        descripcion: descripcion.trim() || null,
        cargada_por: cargadaPor || null
      });
      // Reset del N° de OT y de los campos propios de esa OT. Fecha, turno y
      // "cargada por" se conservan (son de la sesión de trabajo). El SECTOR se
      // resetea a propósito: un sector pegajoso puede guardar una OT bien
      // formada pero mal atribuida, sin ninguna señal. Ver decisión 35.
      setOt('');
      setDescripcion('');
      setEjecutor('');
      setMsg(`✓ ${canon} registrada en el pool`);
    } catch (e) {
      setMsg(e.code === 'DUPLICADA'
        ? `Error: ${canon} ya está activa en el pool`
        : `Error: ${e.message || 'no se pudo registrar'}`);
    } finally {
      setSaving(false);
    }
  };

  const doAnular = async (motivo) => {
    await onAnular(anularTarget.id, motivo);
    setAnularTarget(null);
    setMsg('✓ OT anulada');
  };

  if (!storage.configured) {
    return (
      <Card className="p-5">
        <div className="flex items-center gap-2 text-sm text-amber-800">
          <AlertTriangle className="w-4 h-4 flex-shrink-0" />
          El pool de OTs requiere Supabase configurado. En modo local no está disponible.
        </div>
      </Card>
    );
  }

  return (
    <div className="flex flex-col gap-5">
      <Card className="p-5">
        <SectionTitle icon={Inbox} accent="violet">Registrar OT en el pool</SectionTitle>

        <div className="grid grid-cols-1 lg:grid-cols-4 gap-4">
          <Field label="N° de OT">
            <OTNumberInput value={ot} onChange={setOt} />
          </Field>

          <Field label="Clase">
            <div className="flex gap-2">
              <button type="button" onClick={() => setClase('A')}
                className={`${buttonCls} flex-1 justify-center ${clase === 'A' ? 'bg-sky-500 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                title="La ejecuta Mantenimiento — entra al carry-over y a estadísticas">
                A · Mantenimiento
              </button>
              <button type="button" onClick={() => setClase('B')}
                className={`${buttonCls} flex-1 justify-center ${clase === 'B' ? 'bg-slate-600 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                title="La ejecuta otra área — solo ocupa el número">
                B · Otra área
              </button>
            </div>
          </Field>

          <Field label="Fecha de emisión">
            <input type="date" className={`${inputCls} num`} value={fecha}
              max={todayLocalISO()} onChange={e => setFecha(e.target.value)} />
          </Field>

          <Field label="Turno de origen">
            <select className={inputCls} value={turno} onChange={e => setTurno(e.target.value)}>
              {TURNOS.map(t => <option key={t} value={t}>{t}</option>)}
            </select>
          </Field>

          {clase === 'B' && (
            <Field label="Ejecuta (área)">
              <input type="text" className={inputCls} value={ejecutor}
                onChange={e => setEjecutor(e.target.value)}
                placeholder="Ej: Facilities" />
            </Field>
          )}

          <Field label="Descripción (opcional)" className={clase === 'B' ? 'lg:col-span-2' : 'lg:col-span-3'}>
            <input type="text" className={inputCls} value={descripcion}
              onChange={e => setDescripcion(e.target.value)}
              placeholder="Lo que dice el papel" />
          </Field>

          <Field label="Cargada por">
            <select className={inputCls} value={cargadaPor} onChange={e => setCargadaPor(e.target.value)}>
              <option value="">—</option>
              {RESPONSABLES.map(r => <option key={r.id} value={r.name}>{r.name}</option>)}
            </select>
          </Field>
        </div>

        <div className="flex items-center gap-3 mt-4">
          <button onClick={submit} disabled={saving}
            className="inline-flex items-center gap-2 px-5 py-2.5 bg-violet-600 hover:bg-violet-500 text-white rounded-xl font-bold text-sm transition disabled:opacity-50 shadow-md">
            <Plus className="w-4 h-4" />{saving ? 'Registrando…' : 'Registrar OT'}
          </button>
          {msg && (
            <span className={`text-xs font-medium ${msg.startsWith('Error') ? 'text-red-600' : 'text-emerald-600'}`}>
              {msg}
            </span>
          )}
        </div>

        <p className="text-[11px] text-slate-400 mt-3 leading-relaxed">
          El pool registra la solicitud, no el trabajo. La OT se sigue cargando y cerrando
          en el reporte del turno como siempre; si ya está en un reporte, registrarla acá no la duplica.
        </p>
      </Card>

      <Card className="p-5">
        <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
          <SectionTitle icon={ListChecks} accent="slate">
            OTs en el pool
          </SectionTitle>
          <div className="flex items-center gap-2">
            <select className="px-2 py-1.5 text-xs bg-white border border-slate-300 rounded-lg num"
              value={filtroSector} onChange={e => setFiltroSector(e.target.value)}>
              <option value="">Todos los sectores</option>
              {SECTORES_OT.map(sx => <option key={sx.code} value={sx.code}>{sx.code}</option>)}
            </select>
            <label className="flex items-center gap-1.5 text-xs text-slate-600">
              <input type="checkbox" checked={verAnuladas} onChange={e => setVerAnuladas(e.target.checked)} />
              Ver anuladas
            </label>
            <button onClick={onRefresh} className="p-1.5 hover:bg-slate-100 rounded transition" title="Refrescar">
              <RefreshCw className="w-4 h-4 text-slate-500" />
            </button>
          </div>
        </div>

        <div className="flex flex-wrap gap-4 mb-4 text-xs text-slate-600">
          <span><strong className="num text-slate-800">{activas.length}</strong> activas</span>
          <span><strong className="num text-slate-800">{activas.filter(r => r.clase === 'A').length}</strong> Clase A</span>
          <span><strong className="num text-slate-800">{activas.filter(r => r.clase === 'B').length}</strong> Clase B</span>
          <span><strong className="num text-slate-800">{pool.length - activas.length}</strong> anuladas</span>
        </div>

        {poolError && (
          <div className="mb-3 text-xs text-red-600 flex items-center gap-1.5">
            <AlertTriangle className="w-3.5 h-3.5" />{poolError}
          </div>
        )}

        {poolLoading ? (
          <div className="text-center text-slate-500 py-10 text-sm">Cargando pool…</div>
        ) : listado.length === 0 ? (
          <EmptyHint>No hay OTs registradas en el pool con este filtro.</EmptyHint>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="py-2 pr-3">N° OT</th>
                  <th className="py-2 pr-3">Clase</th>
                  <th className="py-2 pr-3">Ejecuta</th>
                  <th className="py-2 pr-3">Emisión</th>
                  <th className="py-2 pr-3">Descripción</th>
                  <th className="py-2 pr-3">Cargada por</th>
                  <th className="py-2 pr-3"></th>
                </tr>
              </thead>
              <tbody>
                {listado.map(r => (
                  <tr key={r.id} className={`border-b border-slate-100 ${r.anulada_at ? 'opacity-50' : ''}`}>
                    <td className="py-2 pr-3 num font-semibold text-slate-800">
                      {r.anulada_at ? <s>{r.ot}</s> : r.ot}
                    </td>
                    <td className="py-2 pr-3">
                      <span className={`text-[10px] px-1.5 py-0.5 rounded font-bold ${r.clase === 'A' ? 'bg-sky-100 text-sky-700' : 'bg-slate-200 text-slate-600'}`}>
                        {r.clase}
                      </span>
                    </td>
                    <td className="py-2 pr-3 text-slate-600">{r.clase === 'A' ? 'Mantenimiento' : (r.ejecutor || '—')}</td>
                    <td className="py-2 pr-3 num text-slate-600 whitespace-nowrap">{r.fecha} · {r.turno_origen}</td>
                    <td className="py-2 pr-3 text-slate-600 max-w-[280px] truncate" title={r.descripcion || ''}>{r.descripcion || '—'}</td>
                    <td className="py-2 pr-3 text-slate-500 text-xs">{r.cargada_por || '—'}</td>
                    <td className="py-2 pr-3 text-right">
                      {r.anulada_at ? (
                        <span className="text-[10px] text-amber-700" title={r.anulada_motivo || ''}>
                          anulada
                        </span>
                      ) : (
                        <button onClick={() => setAnularTarget(r)}
                          className="p-1 hover:bg-amber-50 rounded transition" title="Anular con motivo">
                          <Ban className="w-4 h-4 text-amber-600" />
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      {anularTarget && (
        <PoolAnularDialog
          row={anularTarget}
          onConfirm={doAnular}
          onCancel={() => setAnularTarget(null)}
        />
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// BACKLOG #46 (v3.25) — EXTRAS: solicitud y aprobación de horas extras
// ═══════════════════════════════════════════════════════════════════
// Modelo del flujo, para que se entienda leyendo solo esto:
//
//   - Tres ENCARGADOS cargan solicitudes de horas extras. Cada uno ve
//     únicamente las que cargó él (filtro por AUTOR, no por turno: los turnos
//     rotan y un turno fijo por usuario se rompería en la primera rotación).
//   - Un JEFE ve todas, aprueba o rechaza. También puede cargar, y lo que
//     carga nace ya aprobado (se autoaprueba: no tiene a quién elevarlo).
//   - Las fechas FUTURAS son válidas y son el caso normal: se solicita antes
//     de hacer el trabajo, no después. Al revés que los reportes, que bloquean
//     el futuro (#11).
//   - Una solicitud `pendiente` es editable y anulable por su autor. Una vez
//     resuelta queda congelada; corregir = anular con motivo y cargar de nuevo.
//   - Anular es soft-delete con motivo obligatorio, NUNCA DELETE (la tabla ni
//     siquiera tiene GRANT de DELETE). El encargado anula solo las propias y
//     pendientes; el jefe anula cualquiera, en cualquier estado.
//
// Lo que este módulo NO garantiza está documentado arriba de EXTRAS_USUARIOS.

function ExtrasLoginDialog({ onConfirm, onCancel }) {
  const [user, setUser] = useState('');
  const [pwd, setPwd] = useState('');
  const [error, setError] = useState('');
  const inputRef = useRef(null);

  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onCancel(); };
    document.addEventListener('keydown', onKey);
    inputRef.current?.focus();
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel]);

  const submit = () => {
    if (!user.trim() || !pwd) { setError('Completá usuario y contraseña'); return; }
    const ok = onConfirm(user, pwd);
    if (!ok) setError('Usuario o contraseña incorrectos');
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-2xl max-w-sm w-full p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-cyan-600 flex items-center justify-center flex-shrink-0">
            <Timer className="w-5 h-5 text-white" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">Horas extras</h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              Solicitud y aprobación de horas extras. Ingresá con tu usuario.
            </p>
          </div>
        </div>
        <div className="flex flex-col gap-2">
          <input ref={inputRef} type="text" value={user} autoComplete="username"
            autoCapitalize="none" autoCorrect="off" spellCheck={false}
            onChange={e => { setUser(e.target.value); setError(''); }}
            onKeyDown={e => { if (e.key === 'Enter') submit(); }}
            placeholder="Usuario"
            className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-cyan-500/40 focus:border-cyan-500 transition" />
          <input type="password" value={pwd} autoComplete="current-password"
            onChange={e => { setPwd(e.target.value); setError(''); }}
            onKeyDown={e => { if (e.key === 'Enter') submit(); }}
            placeholder="Contraseña"
            className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-cyan-500/40 focus:border-cyan-500 transition" />
        </div>
        {error && <div className="mt-2 text-xs text-red-600">{error}</div>}
        <div className="flex items-center justify-end gap-2 mt-5">
          <button onClick={onCancel}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition">
            Cancelar
          </button>
          <button onClick={submit}
            className="px-4 py-2 text-sm font-semibold text-white bg-cyan-600 hover:bg-cyan-500 rounded-lg transition">
            Ingresar
          </button>
        </div>
      </div>
    </div>
  );
}

// Diálogo de motivo, compartido por anulación (motivo OBLIGATORIO, lo exige
// también el CHECK `horas_extras_anulacion` de la tabla) y rechazo (motivo
// opcional, decidido así el 2026-08-23).
function ExtrasMotivoDialog({ titulo, descripcion, placeholder, requerido, cta, ctaColor, iconColor, icon: Icon, onConfirm, onCancel }) {
  const [motivo, setMotivo] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const inputRef = useRef(null);

  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onCancel(); };
    document.addEventListener('keydown', onKey);
    inputRef.current?.focus();
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel]);

  const submit = async () => {
    const m = motivo.trim();
    if (requerido && !m) { setError('El motivo es obligatorio'); return; }
    setBusy(true);
    try { await onConfirm(m); }
    catch (e) { setError(e.message || 'No se pudo completar la operación'); setBusy(false); }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-2xl max-w-sm w-full p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className={`w-10 h-10 rounded-full ${iconColor} flex items-center justify-center flex-shrink-0`}>
            <Icon className="w-5 h-5 text-white" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">{titulo}</h3>
            <p className="text-sm text-slate-600 leading-relaxed">{descripcion}</p>
          </div>
        </div>
        <input ref={inputRef} type="text" value={motivo}
          onChange={e => { setMotivo(e.target.value); setError(''); }}
          onKeyDown={e => { if (e.key === 'Enter') submit(); }}
          placeholder={placeholder}
          className={inputCls} />
        {error && <div className="mt-2 text-xs text-red-600">{error}</div>}
        <div className="flex items-center justify-end gap-2 mt-5">
          <button onClick={onCancel} disabled={busy}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition disabled:opacity-50">
            Cancelar
          </button>
          <button onClick={submit} disabled={busy}
            className={`px-4 py-2 text-sm font-semibold text-white rounded-lg transition disabled:opacity-50 ${ctaColor}`}>
            {busy ? 'Guardando…' : cta}
          </button>
        </div>
      </div>
    </div>
  );
}

// Selector de hora en 24 h. Reemplaza a <input type="time">, cuyo formato lo
// decide el LOCALE del navegador/SO y no se puede forzar desde el HTML: con el
// navegador en inglés aparece el picker de AM/PM, y no hay atributo que lo
// cambie. Con dos selects el formato es el mismo para todos los usuarios,
// independiente de cómo tenga configurada la máquina cada encargado.
// El valor sigue siendo "HH:MM" de 24 h, igual que antes — la columna `time`
// de Postgres no se entera de este cambio.
// Los minutos son SOLO :00 y :30 — las horas extras no se fraccionan por debajo
// de la media hora. Al elegir la hora, los minutos se completan en '00' solos.
const HORAS_24 = Array.from({ length: 24 }, (_, i) => String(i).padStart(2, '0'));
const MINUTOS_VALIDOS = ['00', '30'];

function TimeInput24({ value, onChange, disabled }) {
  const [h, m] = (value || '').split(':');
  // Elegir solo una de las dos mitades completa la otra en '00' en vez de dejar
  // un valor a medio formar que después rebota en la validación.
  const setH = (nh) => onChange(nh ? `${nh}:${m || '00'}` : '');
  const setM = (nm) => onChange(`${h || '00'}:${nm}`);
  // Si una fila guardada trae un minuto fuera de la lista (por ejemplo :45 de
  // una carga anterior a esta regla), se agrega como opción en vez de dejar el
  // select vacío: editar el motivo de esa fila no le puede cambiar la hora en
  // silencio. La opción extra desaparece en cuanto se elige :00 o :30.
  const opcionesMin = m && !MINUTOS_VALIDOS.includes(m)
    ? [...MINUTOS_VALIDOS, m].sort()
    : MINUTOS_VALIDOS;
  const selCls = "px-2 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-cyan-500/40 focus:border-cyan-500 transition num font-semibold";

  return (
    <div className="flex items-center gap-1">
      <select className={selCls} value={h || ''} onChange={e => setH(e.target.value)} disabled={disabled}>
        <option value="">--</option>
        {HORAS_24.map(x => <option key={x} value={x}>{x}</option>)}
      </select>
      <span className="text-slate-400 font-bold">:</span>
      <select className={selCls} value={m || ''} onChange={e => setM(e.target.value)} disabled={disabled}>
        <option value="">--</option>
        {opcionesMin.map(x => <option key={x} value={x}>{x}</option>)}
      </select>
      {value && value.includes(':') && (
        <span className="text-[10px] text-slate-400 whitespace-nowrap ml-0.5">{to12h(value)}</span>
      )}
    </div>
  );
}

const EXTRAS_ESTADO_STYLE = {
  pendiente:  'bg-amber-100 text-amber-700',
  aprobada:   'bg-emerald-100 text-emerald-700',
  rechazada:  'bg-red-100 text-red-700',
  // Ajuste a la baja sobre algo ya ejecutado, esperando que el jefe lo
  // confirme (v3.30). Cian para no confundirse con "pendiente" (ámbar):
  // es una corrección sobre algo que ya se había aprobado, no una carga
  // nueva sin resolver.
  modificada: 'bg-cyan-100 text-cyan-700'
};

// ═══════════════════════════════════════════════════════════════════
// DASHBOARD DE HORAS EXTRAS (#49, v3.27) — SOLO ROL JEFE
// ═══════════════════════════════════════════════════════════════════
// Todo se calcula en el CLIENTE sobre `extras` ya cargado: cero llamadas
// nuevas a Supabase, cero SQL. La tabla es chica y la alternativa (vistas
// agregadas en Postgres) obligaría a mantener la lógica en dos lugares.
//
// Reglas de población, válidas para TODOS los bloques:
//  · Las ANULADAS no suman nunca, en ningún bloque, aunque estén vigentes en
//    el listado con el checkbox. Una hora anulada no se trabaja ni se paga.
//  · Aprobadas y pendientes se muestran SEPARADAS y jamás sumadas: lo que se
//    paga son las aprobadas; las pendientes son proyección.
//  · Se agrupa por `tecnico_nombre`, NUNCA por `tecnico_id` — desde #52 hay
//    7 personas sin id y un GROUP BY por id las perdería en silencio.
//  · Una solicitud entra al período por su `fecha` (inicio). Las que cruzan
//    medianoche el último día del período cuentan enteras en ese período.
// `soloPersonas` (#58): null para el jefe, que ve todo. Para el encargado, la
// lista de su gente a cargo. El recorte es de UI y no de datos: el fetch trae
// el período completo y el filtro se aplica acá.
function ExtrasDashboard({ soloPersonas }) {
  // #66 — Selector único de período RRHH (Año + Mes), reemplaza al viejo
  // Mensual/Cuatrimestral/Anual en mes calendario. Unifica KPIs, ranking y
  // métricas de Proceso bajo el mismo recorte que ya usaban Evolución y
  // Acumulado RRHH — antes convivían dos calendarios distintos en la misma
  // pantalla. Se pierden las vistas Cuatrimestral/Anual de antes; volver a
  // tenerlas como agregado de períodos RRHH, si hace falta, es aparte.
  const hoyPeriodo = extrasPeriodoRRHH(todayLocalISO())
    || { anio: new Date().getFullYear(), mes: new Date().getMonth() + 1 };
  const [anioSel, setAnioSel] = useState(hoyPeriodo.anio);
  const [mesSel, setMesSel] = useState(hoyPeriodo.mes);
  const [verTodos, setVerTodos] = useState(false);

  // Recorte por gente a cargo. Se aplica a TODOS los bloques.
  const enAlcance = useCallback(
    (nombre) => !soloPersonas || soloPersonas.includes(nombre),
    [soloPersonas]
  );

  // Datos propios, acotados al período. No se usa `extras` (el listado) porque
  // ese viene topeado en EXTRAS_LIST_LIMIT y con el volumen real del sector se
  // agota en 6 o 7 meses — la vista anual quedaría truncada.
  const [datos, setDatos] = useState([]);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState('');

  // Serie del gráfico de evolución: ventana propia de 12 períodos, con su
  // modo y un filtro opcional por persona ('' = todas).
  const [modo12, setModo12] = useState('anio');
  const [personaSel, setPersonaSel] = useState('');

  // Avanza/retrocede un período RRHH completo (con acarreo de año).
  const irAPeriodo = (delta) => {
    let m = mesSel + delta, a = anioSel;
    if (m < 1) { m = 12; a -= 1; } else if (m > 12) { m = 1; a += 1; }
    setMesSel(m); setAnioSel(a);
  };

  const periodo = useMemo(() => {
    const r = extrasRangoRRHH(anioSel, mesSel);
    return { ...r, anio: anioSel, mes: mesSel, label: `${MESES_CORTOS[mesSel - 1]} ${anioSel}` };
  }, [anioSel, mesSel]);

  useEffect(() => {
    let vigente = true;
    setCargando(true);
    setError('');
    storage.listExtrasRango(periodo.desde, periodo.hasta)
      .then(r => { if (vigente) { setDatos(r); setCargando(false); } })
      .catch(e => {
        if (vigente) { setError(e.message || 'no se pudieron cargar los datos'); setCargando(false); }
      });
    // Descarta respuestas de un período que ya no se está mirando: si se
    // navega rápido, la consulta vieja puede llegar después de la nueva.
    return () => { vigente = false; };
  }, [periodo.desde, periodo.hasta]);

  useEffect(() => { setVerTodos(false); }, [periodo.desde]);

  const truncado = datos.length >= EXTRAS_DASHBOARD_LIMIT;

  // ═════════════════════════════════════════════════════════════════
  // DATOS EN PERÍODOS RRHH (#59) — alimentan la evolución Y el acumulado
  // ═════════════════════════════════════════════════════════════════
  // Los tres bloques (KPIs/ranking, evolución, acumulado) trabajan en
  // períodos RRHH (11→10) desde #66. La razón de fondo es la misma de
  // siempre: las horas IMPORTADAS solo existen como total mensual del 11 al
  // 10, sin detalle diario, así que repartirlas en meses calendario exigiría
  // inventar una distribución. Antes los KPIs eran de mes calendario y el
  // gráfico aclaraba el desfasaje; ahora todo comparte el mismo recorte.
  const anioRRHH = periodo.anio;

  // Los 12 períodos RRHH de la ventana del gráfico.
  const periodos12 = useMemo(() => {
    if (modo12 === 'anio') {
      return Array.from({ length: 12 }, (_, i) => ({ anio: anioRRHH, mes: i + 1 }));
    }
    // Móvil: los 12 que terminan en el período RRHH seleccionado arriba.
    // Anclar al final evita cortar la serie antes de lo que se está mirando.
    const fin = extrasPeriodoRRHH(periodo.hasta) || { anio: anioRRHH, mes: 12 };
    return Array.from({ length: 12 }, (_, k) => {
      const t = fin.anio * 12 + (fin.mes - 1) - 11 + k;
      return { anio: Math.floor(t / 12), mes: (t % 12) + 1 };
    });
  }, [modo12, anioRRHH, periodo.hasta]);

  // Los 12 del acumulado son siempre el año RRHH completo, como la planilla.
  const periodosAcum = useMemo(
    () => Array.from({ length: 12 }, (_, i) => ({ anio: anioRRHH, mes: i + 1 })),
    [anioRRHH]
  );

  // Una sola consulta cubre las dos ventanas: se toma la unión de rangos.
  const [importadas, setImportadas] = useState([]);
  const [appRRHH, setAppRRHH] = useState([]);
  const [rrhhCargando, setRrhhCargando] = useState(true);

  const anios = useMemo(() => {
    const s = new Set([...periodos12, ...periodosAcum].map(p => p.anio));
    return [...s].sort();
  }, [periodos12, periodosAcum]);

  const rangoRRHH = useMemo(() => {
    const todos = [...periodos12, ...periodosAcum];
    const rangos = todos.map(p => extrasRangoRRHH(p.anio, p.mes));
    return {
      desde: rangos.reduce((a, r) => (r.desde < a ? r.desde : a), rangos[0].desde),
      hasta: rangos.reduce((a, r) => (r.hasta > a ? r.hasta : a), rangos[0].hasta)
    };
  }, [periodos12, periodosAcum]);

  useEffect(() => {
    let vigente = true;
    setRrhhCargando(true);
    Promise.all([
      Promise.all(anios.map(a => storage.listImportadas(a))).then(r => r.flat()),
      storage.listExtrasSerie(rangoRRHH.desde, rangoRRHH.hasta)
    ])
      .then(([imp, app]) => {
        if (!vigente) return;
        setImportadas(imp); setAppRRHH(app); setRrhhCargando(false);
      })
      .catch(() => {
        if (!vigente) return;
        setImportadas([]); setAppRRHH([]); setRrhhCargando(false);
      });
    return () => { vigente = false; };
  }, [anios, rangoRRHH.desde, rangoRRHH.hasta]);

  // ── Índice unificado: período RRHH + persona → { aprobadas, pendientes }
  // Cada celda toma su valor de UNA sola fuente, nunca de las dos: hasta
  // ago-2026 lo importado, desde sep-2026 la app. Sin doble conteo posible.
  // Las importadas cuentan como aprobadas: ya están liquidadas.
  const indiceRRHH = useMemo(() => {
    const map = new Map();
    const k = (a, m, p) => `${a}|${m}|${p}`;
    const add = (a, m, p, campo, h) => {
      const key = k(a, m, p);
      const o = map.get(key) || { anio: a, mes: m, persona: p, aprobadas: 0, pendientes: 0, importado: false };
      o[campo] += h;
      if (campo === 'aprobadas' && !extrasFuenteEsApp(a, m)) o.importado = true;
      map.set(key, o);
    };
    importadas.forEach(r => {
      if (!enAlcance(r.persona) || extrasFuenteEsApp(r.anio, r.mes)) return;
      add(r.anio, r.mes, r.persona, 'aprobadas', Number(r.horas) || 0);
    });
    appRRHH.forEach(r => {
      if (r.anulada_at || !enAlcance(r.tecnico_nombre)) return;
      // 'modificada' cuenta como aprobada: son horas que ya estaban liquidadas
      // y el encargado las bajó tras la ejecución (v3.30). Usa el valor YA
      // bajado — nunca sobreestima lo que corresponde pagar mientras el jefe
      // confirma.
      if (r.estado !== 'aprobada' && r.estado !== 'modificada' && r.estado !== 'pendiente') return;
      const p = extrasPeriodoRRHH(r.fecha);
      if (!p || !extrasFuenteEsApp(p.anio, p.mes)) return;
      add(p.anio, p.mes, r.tecnico_nombre,
        (r.estado === 'aprobada' || r.estado === 'modificada') ? 'aprobadas' : 'pendientes', Number(r.horas) || 0);
    });
    return [...map.values()];
  }, [importadas, appRRHH, enAlcance]);

  // Personas disponibles para el filtro del gráfico.
  const personasRRHH = useMemo(
    () => [...new Set(indiceRRHH.map(r => r.persona))].sort(),
    [indiceRRHH]
  );

  useEffect(() => {
    if (personaSel && !personasRRHH.includes(personaSel)) setPersonaSel('');
  }, [personasRRHH, personaSel]);

  // Grilla del acumulado: persona x 12 períodos del año RRHH.
  const acumRRHH = useMemo(() => {
    const map = new Map();
    indiceRRHH.forEach(r => {
      if (r.anio !== anioRRHH) return;
      const c = map.get(r.persona) || { persona: r.persona, meses: Array(12).fill(null) };
      const v = r.aprobadas;
      c.meses[r.mes - 1] = (c.meses[r.mes - 1] || 0) + v;
      map.set(r.persona, c);
    });
    return [...map.values()].map(c => {
      const conDato = c.meses.filter(v => v !== null);
      const total = conDato.reduce((a, v) => a + v, 0);
      return {
        ...c,
        total,
        // Promedio sobre los meses CON DATO, no sobre 12: es lo que hace la
        // planilla de RRHH. BAGGIO ingresó en julio, así que su promedio se
        // calcula sobre 2 meses (19,8 h) y no sobre 12 (3,3 h).
        promedio: conDato.length ? total / conDato.length : 0,
        nMeses: conDato.length
      };
    }).sort((a, b) => b.total - a.total);
  }, [indiceRRHH, anioRRHH]);

  const totalRRHH = useMemo(() => acumRRHH.reduce((a, c) => a + c.total, 0), [acumRRHH]);

  const vivas = useMemo(
    () => datos.filter(r => !r.anulada_at && enAlcance(r.tecnico_nombre)),
    [datos, enAlcance]
  );

  const sumH = (arr) => arr.reduce((a, r) => a + (Number(r.horas) || 0), 0);

  const kpis = useMemo(() => {
    // 'modificada' se cuenta junto a las aprobadas (v3.30): son horas que ya
    // estaban aprobadas y se corrigieron a la baja tras ejecutarse.
    const ap = vivas.filter(r => r.estado === 'aprobada' || r.estado === 'modificada');
    const pe = vivas.filter(r => r.estado === 'pendiente');
    const re = vivas.filter(r => r.estado === 'rechazada');
    return {
      hAprobadas: sumH(ap), nAprobadas: ap.length,
      hPendientes: sumH(pe), nPendientes: pe.length,
      hRechazadas: sumH(re), nRechazadas: re.length,
      personas: new Set(vivas.map(r => r.tecnico_nombre)).size
    };
  }, [vivas]);

  // Ranking por persona. Ordenado por aprobadas, que es lo que se paga.
  const porPersona = useMemo(() => {
    const map = new Map();
    vivas.forEach(r => {
      const k = r.tecnico_nombre || '(sin nombre)';
      const o = map.get(k) || { nombre: k, aprobadas: 0, pendientes: 0 };
      if (r.estado === 'aprobada' || r.estado === 'modificada') o.aprobadas += Number(r.horas) || 0;
      else if (r.estado === 'pendiente') o.pendientes += Number(r.horas) || 0;
      map.set(k, o);
    });
    return [...map.values()]
      .filter(o => o.aprobadas > 0 || o.pendientes > 0)
      .sort((a, b) => (b.aprobadas + b.pendientes) - (a.aprobadas + a.pendientes));
  }, [vivas]);

  const porPersonaTop = useMemo(
    () => (verTodos ? porPersona : porPersona.slice(0, 15)),
    [porPersona, verTodos]
  );

  // Evolución: SIEMPRE 12 períodos RRHH, sobre el índice unificado. Una sola
  // barra no es una evolución. Se recorren los 12 de la ventana y no los que
  // tienen datos: un período en cero tiene que verse como cero.
  // `enPeriodo` marca la barra que coincide con el período RRHH elegido
  // arriba. Desde #66 el período elegido ES un período RRHH exacto (antes
  // era un recorte calendario que solo se solapaba con alguno de los 12),
  // así que esto da match en como mucho una barra de las 12.
  const porMes = useMemo(() => periodos12.map(({ anio, mes }) => {
    const filas = indiceRRHH.filter(r =>
      r.anio === anio && r.mes === mes && (!personaSel || r.persona === personaSel)
    );
    const r = extrasRangoRRHH(anio, mes);
    return {
      mes: `${MESES_CORTOS[mes - 1]} ${String(anio).slice(2)}`,
      enPeriodo: r.desde <= periodo.hasta && r.hasta >= periodo.desde,
      importado: !extrasFuenteEsApp(anio, mes),
      aprobadas: Number(filas.reduce((a, x) => a + x.aprobadas, 0).toFixed(2)),
      pendientes: Number(filas.reduce((a, x) => a + x.pendientes, 0).toFixed(2))
    };
  }), [indiceRRHH, periodos12, periodo, personaSel]);

  // Horas por categoría de motivo. Exacto desde #54: antes era texto libre y
  // había que normalizar strings, con lo que "Cubrir licencia" y "cubrir
  // licencia médica" caían en grupos distintos.
  const porMotivo = useMemo(() => {
    const map = new Map();
    vivas.forEach(r => {
      const k = r.motivo_categoria || '(sin categoría)';
      const o = map.get(k) || { motivo: k, aprobadas: 0, pendientes: 0, n: 0 };
      if (r.estado === 'aprobada' || r.estado === 'modificada') o.aprobadas += Number(r.horas) || 0;
      else if (r.estado === 'pendiente') o.pendientes += Number(r.horas) || 0;
      o.n += 1;
      map.set(k, o);
    });
    return [...map.values()].sort((a, b) => (b.aprobadas + b.pendientes) - (a.aprobadas + a.pendientes));
  }, [vivas]);

  // ── Métricas de proceso ──────────────────────────────────────────
  // Población: solo las RESUELTAS (aprobadas o rechazadas). Las pendientes
  // no tienen desenlace todavía y meterlas bajaría artificialmente la tasa.
  //
  // Las cargas del jefe SE INCLUYEN, por decisión explícita. Como se
  // autoaprueban, `resuelto_at` se sella en el cliente ANTES de que el insert
  // ponga su `created_at`, y la diferencia da NEGATIVA. Se cortan en cero: un
  // tiempo negativo no es un dato, es un artefacto del orden de sellado.
  const proceso = useMemo(() => {
    const resueltas = vivas.filter(r => r.estado === 'aprobada' || r.estado === 'rechazada');
    const conTiempo = resueltas.filter(r => r.created_at && r.resuelto_at);
    const horasDe = (r) => Math.max(0,
      (new Date(r.resuelto_at) - new Date(r.created_at)) / 3600000);
    const prom = (arr) => (arr.length ? arr.reduce((a, r) => a + horasDe(r), 0) / arr.length : null);
    const planificadas = conTiempo.filter(r => !extrasEsReactivo(r.motivo_categoria));
    const reactivas = conTiempo.filter(r => extrasEsReactivo(r.motivo_categoria));
    return {
      nResueltas: resueltas.length,
      tasa: resueltas.length
        ? (resueltas.filter(r => r.estado === 'aprobada').length / resueltas.length) * 100
        : null,
      tPlanificadas: prom(planificadas), nPlanificadas: planificadas.length,
      tReactivas: prom(reactivas), nReactivas: reactivas.length,
      autoaprobadas: resueltas.filter(r => r.solicitado_por === r.resuelto_por).length
    };
  }, [vivas]);

  // Años ofrecidos en el selector: ventana fija alrededor del actual, más el
  // año elegido si por algún motivo cae fuera (deep-link, etc.).
  const anioOptions = useMemo(() => {
    const actual = new Date().getFullYear();
    const arr = [];
    for (let y = actual - 2; y <= actual + 1; y++) arr.push(y);
    if (!arr.includes(anioSel)) arr.push(anioSel);
    return arr.sort((a, b) => a - b);
  }, [anioSel]);

  return (
    <div className="flex flex-col gap-5">
      {/* ── Selector de período RRHH ───────────────────────────── */}
      <Card className="p-4">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <button onClick={() => irAPeriodo(-1)}
              className="p-1.5 hover:bg-slate-100 rounded-lg transition border border-slate-300" title="Período RRHH anterior">
              <ChevronLeft className="w-4 h-4 text-slate-600" />
            </button>
            <select value={mesSel} onChange={e => setMesSel(Number(e.target.value))}
              className="px-2 py-1.5 text-sm font-semibold bg-white border border-slate-300 rounded-lg capitalize">
              {MESES_CORTOS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
            </select>
            <select value={anioSel} onChange={e => setAnioSel(Number(e.target.value))}
              className="px-2 py-1.5 text-sm font-semibold bg-white border border-slate-300 rounded-lg">
              {anioOptions.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
            <button onClick={() => irAPeriodo(1)}
              className="p-1.5 hover:bg-slate-100 rounded-lg transition border border-slate-300" title="Período RRHH siguiente">
              <ChevronRight className="w-4 h-4 text-slate-600" />
            </button>
            {(anioSel !== hoyPeriodo.anio || mesSel !== hoyPeriodo.mes) && (
              <button onClick={() => { setAnioSel(hoyPeriodo.anio); setMesSel(hoyPeriodo.mes); }}
                className="px-2 py-1 text-[11px] font-semibold text-emerald-700 hover:bg-emerald-50 rounded-lg transition">
                Hoy
              </button>
            )}
          </div>
        </div>
        <p className="text-[11px] text-slate-400 mt-2 num">
          Período RRHH (11 al 10): {formatDateShort(periodo.desde)} — {formatDateShort(periodo.hasta)}
        </p>
      </Card>

      {/* Truncamiento: sin esto, un período largo mostraría totales
          incompletos sin ninguna señal. */}
      {truncado && (
        <div className="flex items-start gap-2 px-3 py-2 bg-red-50 border border-red-200 rounded-lg">
          <AlertTriangle className="w-4 h-4 text-red-600 flex-shrink-0 mt-0.5" />
          <div className="text-xs text-red-800 leading-relaxed">
            <strong>Datos incompletos.</strong> Se alcanzó el tope de {EXTRAS_DASHBOARD_LIMIT} solicitudes para
            este período, así que los números de abajo pueden estar por debajo del real.
            Hay que subir <code>EXTRAS_DASHBOARD_LIMIT</code> o paginar antes de usar esto para liquidación.
          </div>
        </div>
      )}

      {error && (
        <div className="flex items-start gap-2 px-3 py-2 bg-red-50 border border-red-200 rounded-lg">
          <AlertTriangle className="w-4 h-4 text-red-600 flex-shrink-0 mt-0.5" />
          <div className="text-xs text-red-800">No se pudieron cargar los datos del período: {error}</div>
        </div>
      )}

      {/* ── Evolución: SIEMPRE 12 meses, ventana propia ─────────────
          Va FUERA del condicional del período: aunque el período elegido
          esté vacío, los otros 11 meses siguen teniendo algo que decir. */}
      <Card className="p-5">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <SectionTitle icon={BarChart3} accent="emerald">Evolución · 12 períodos RRHH</SectionTitle>
          <div className="flex flex-wrap items-center gap-1 mb-4">
            <select value={personaSel} onChange={e => setPersonaSel(e.target.value)}
              className="px-2 py-1 text-[11px] font-semibold border border-slate-300 rounded-lg bg-white text-slate-700 max-w-[190px]"
              title="Filtrar la evolución por persona">
              <option value="">Todas ({personasRRHH.length})</option>
              {personasRRHH.map(n => <option key={n} value={n}>{n}</option>)}
            </select>
            <button onClick={() => setModo12('anio')}
              className={`px-2.5 py-1 text-[11px] font-semibold rounded-lg transition ${
                modo12 === 'anio' ? 'bg-emerald-600 text-white' : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-300'
              }`} title="Los 12 períodos RRHH del año">
              Año RRHH
            </button>
            <button onClick={() => setModo12('movil')}
              className={`px-2.5 py-1 text-[11px] font-semibold rounded-lg transition ${
                modo12 === 'movil' ? 'bg-emerald-600 text-white' : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-300'
              }`} title="Los 12 períodos que terminan en el seleccionado">
              Últimos 12
            </button>
          </div>
        </div>
        {rrhhCargando ? (
          <div className="text-center text-slate-500 py-16 text-sm">Cargando evolución…</div>
        ) : (
          <>
            <ResponsiveContainer width="100%" height={250}>
              <BarChart data={porMes}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis dataKey="mes" tick={{ fontSize: 10 }} interval={0} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip formatter={(v) => formatHoras(v)} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                {/* Los meses fuera del período se atenúan: la ventana es de 12
                    pero los KPIs de abajo son del período. Sin esta señal, los
                    dos bloques parecen hablar del mismo recorte y no lo hacen. */}
                <Bar dataKey="aprobadas" name="Aprobadas" radius={[3, 3, 0, 0]}>
                  {porMes.map((d, i) => (
                    <Cell key={i} fill="#059669" fillOpacity={d.enPeriodo ? 1 : 0.35} />
                  ))}
                </Bar>
                <Bar dataKey="pendientes" name="Pendientes" radius={[3, 3, 0, 0]}>
                  {porMes.map((d, i) => (
                    <Cell key={i} fill="#d97706" fillOpacity={d.enPeriodo ? 1 : 0.35} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
            <p className="text-[11px] text-slate-400 mt-2 leading-relaxed">
              Períodos RRHH: cada barra va del 11 del mes anterior al 10 del propio mes. En color pleno, los que se
              solapan con el período seleccionado ({periodo.label}); el resto va atenuado.
              {personaSel && <> Filtrado por <strong>{personaSel}</strong>.</>}
              {' '}Hasta ago 2026 las horas son importadas de RRHH; desde sep 2026 salen de la app.
              Ningún período mezcla las dos fuentes.
            </p>
          </>
        )}
      </Card>

      {cargando ? (
        <Card className="p-5">
          <div className="text-center text-slate-500 py-10 text-sm">Cargando {periodo.label}…</div>
        </Card>
      ) : vivas.length === 0 ? (
        <Card className="p-5">
          <EmptyHint>No hay solicitudes en {periodo.label}.</EmptyHint>
        </Card>
      ) : (
        <>
          {/* ── KPIs ──────────────────────────────────────────────── */}
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            <Card className="p-4">
              <div className="text-[11px] uppercase tracking-wide text-slate-500 font-semibold">Aprobadas</div>
              <div className="text-2xl font-bold text-emerald-700 num mt-1">{formatHoras(kpis.hAprobadas)}</div>
              <div className="text-[11px] text-slate-400 num">{kpis.nAprobadas} solicitudes · se pagan</div>
            </Card>
            <Card className="p-4">
              <div className="text-[11px] uppercase tracking-wide text-slate-500 font-semibold">Pendientes</div>
              <div className="text-2xl font-bold text-amber-700 num mt-1">{formatHoras(kpis.hPendientes)}</div>
              <div className="text-[11px] text-slate-400 num">{kpis.nPendientes} solicitudes · proyección</div>
            </Card>
            <Card className="p-4">
              <div className="text-[11px] uppercase tracking-wide text-slate-500 font-semibold">Rechazadas</div>
              <div className="text-2xl font-bold text-slate-400 num mt-1">{formatHoras(kpis.hRechazadas)}</div>
              <div className="text-[11px] text-slate-400 num">{kpis.nRechazadas} solicitudes</div>
            </Card>
            <Card className="p-4">
              <div className="text-[11px] uppercase tracking-wide text-slate-500 font-semibold">Personas</div>
              <div className="text-2xl font-bold text-slate-800 num mt-1">{kpis.personas}</div>
              <div className="text-[11px] text-slate-400">con extras en el período</div>
            </Card>
          </div>

          {/* ── Ranking por persona ───────────────────────────────── */}
          <Card className="p-5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <SectionTitle icon={Users} accent="emerald">Horas por persona</SectionTitle>
              {porPersona.length > 15 && (
                <button onClick={() => setVerTodos(v => !v)}
                  className="text-xs font-semibold text-emerald-700 hover:text-emerald-800 mb-4">
                  {verTodos ? 'Ver solo los primeros 15' : `Ver todos (${porPersona.length})`}
                </button>
              )}
            </div>
            <ResponsiveContainer width="100%" height={Math.max(200, porPersonaTop.length * 30)}>
              <BarChart data={porPersonaTop} layout="vertical" margin={{ top: 5, right: 20, bottom: 5, left: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis type="number" tick={{ fontSize: 11 }} />
                <YAxis type="category" dataKey="nombre" width={150} tick={{ fontSize: 11 }} />
                <Tooltip formatter={(v) => formatHoras(v)} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                {/* barSize fijo: sin esto Recharts engorda la franja cuando
                    hay poca gente. Mismo grosor que las columnas de la
                    evolución, sea cual sea la cantidad de personas. */}
                <Bar dataKey="aprobadas" name="Aprobadas" stackId="h" fill="#059669" barSize={14} />
                <Bar dataKey="pendientes" name="Pendientes" stackId="h" fill="#d97706" barSize={14} />
              </BarChart>
            </ResponsiveContainer>
            {!verTodos && porPersona.length > 15 && (
              <p className="text-[11px] text-slate-400 mt-2">
                Mostrando 15 de {porPersona.length}, ordenadas por total. Los KPIs de arriba
                consideran a las {porPersona.length}.
              </p>
            )}
          </Card>

          {/* ── Acumulado RRHH (#59) ─────────────────────────────────────
          Períodos 11→10, no calendario. Va fuera del condicional del
          período: la tabla es del año entero, no del período elegido. */}
      <Card className="p-5">
        <SectionTitle icon={FileSpreadsheet} accent="emerald">
          Acumulado RRHH {anioRRHH} · períodos 11→10
        </SectionTitle>
        <p className="text-[11px] text-slate-500 -mt-2 mb-3 leading-relaxed">
          Cada columna va del <strong>11 del mes anterior al 10 del propio mes</strong>: "Ago" es del 11/07 al 10/08.
          Es el <strong>año RRHH completo</strong> — a diferencia de los bloques de arriba, que muestran solo el
          período (mes) elegido.
        </p>
        {rrhhCargando ? (
          <div className="text-center text-slate-500 py-10 text-sm">Cargando acumulado…</div>
        ) : acumRRHH.length === 0 ? (
          <EmptyHint>No hay horas acumuladas en {anioRRHH}.</EmptyHint>
        ) : (
          <>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="py-2 pr-3 sticky left-0 bg-white">Personal</th>
                    {MESES_CORTOS.map(m => (
                      <th key={m} className="py-2 px-1 text-right capitalize">{m}</th>
                    ))}
                    <th className="py-2 pl-2 text-right">Total</th>
                    <th className="py-2 pl-2 text-right">Prom.</th>
                  </tr>
                </thead>
                <tbody>
                  {acumRRHH.map(c => (
                    <tr key={c.persona} className="border-b border-slate-100">
                      <td className="py-1.5 pr-3 font-semibold text-slate-800 whitespace-nowrap sticky left-0 bg-white">
                        {c.persona}
                      </td>
                      {c.meses.map((v, i) => {
                        // Umbral mensual: más de EXTRAS_ALERTA_MES horas en un
                        // período se sombrea. Es señal visual, no un límite:
                        // no bloquea nada ni cambia ningún total.
                        const alto = v !== null && v > EXTRAS_ALERTA_MES;
                        return (
                          <td key={i}
                              className={`py-1.5 px-1 num text-right ${alto ? 'bg-red-500/15 font-semibold text-red-800' : v === null ? 'text-slate-300' : 'text-slate-600'}`}
                              title={v === null ? 'Sin registro'
                                : `${extrasFuenteEsApp(anioRRHH, i + 1) ? 'Registrado en la app' : 'Importado de RRHH'}${alto ? ` · supera las ${EXTRAS_ALERTA_MES} h del período` : ''}`}>
                            {v === null ? '·' : v.toFixed(2).replace(/\.?0+$/, '').replace('.', ',')}
                          </td>
                        );
                      })}
                      <td className={`py-1.5 pl-2 num text-right font-bold ${c.total > EXTRAS_ALERTA_ANIO ? 'bg-red-500/15 text-red-800' : 'text-slate-900'}`}
                          title={c.total > EXTRAS_ALERTA_ANIO ? `Supera las ${EXTRAS_ALERTA_ANIO} h en el año` : undefined}>
                        {c.total.toFixed(2).replace(/\.?0+$/, '').replace('.', ',')}
                      </td>
                      <td className="py-1.5 pl-2 num text-right font-semibold text-emerald-700"
                          title={`${c.nMeses} ${c.nMeses === 1 ? 'mes' : 'meses'} con registro`}>
                        {c.promedio.toFixed(1).replace('.', ',')}
                      </td>
                    </tr>
                  ))}
                  <tr className="border-t-2 border-slate-300">
                    <td className="py-2 pr-3 font-bold text-slate-900 sticky left-0 bg-white">TOTAL</td>
                    {MESES_CORTOS.map((_, i) => {
                      const s = acumRRHH.reduce((a, c) => a + (c.meses[i] || 0), 0);
                      return (
                        <td key={i} className="py-2 px-1 num text-right font-semibold text-slate-700">
                          {s ? s.toFixed(2).replace(/\.?0+$/, '').replace('.', ',') : '·'}
                        </td>
                      );
                    })}
                    <td className="py-2 pl-2 num text-right font-bold text-slate-900">
                      {totalRRHH.toFixed(2).replace(/\.?0+$/, '').replace('.', ',')}
                    </td>
                    <td className="py-2 pl-2"></td>
                  </tr>
                </tbody>
              </table>
            </div>

            <div className="mt-3 text-[11px] text-slate-500 leading-relaxed space-y-1">
              <p>
                · Hasta el período de <strong>agosto 2026</strong> las horas son <strong>importadas</strong> de la
                planilla acumulada de RRHH: son totales mensuales sin detalle diario. Desde
                <strong> septiembre 2026</strong> salen de lo registrado en la app. Ninguna celda mezcla
                las dos fuentes, así que nada se cuenta dos veces. Pasá el cursor sobre una celda para ver su origen.
              </p>
              <p>
                · De la app entran solo las <strong>aprobadas</strong>: la tabla replica lo que se liquida, y una
                solicitud pendiente todavía no lo es.
              </p>
              <p>
                · El <strong>promedio</strong> se calcula sobre los meses con registro, no sobre 12. Quien ingresó
                a mitad de año no queda con un promedio artificialmente bajo. Un punto (·) es ausencia de registro,
                distinto de un cero.
              </p>
              <p>
                · Sombreado rojo: más de <strong>{EXTRAS_ALERTA_MES} h en un período</strong> o más de
                <strong> {EXTRAS_ALERTA_ANIO} h en el año</strong>. Es señal visual, no un límite — no bloquea
                nada ni altera ningún total.
              </p>
              <p>
                · Las horas importadas <strong>no entran</strong> en los KPIs ni en el ranking de arriba — esos
                bloques solo leen lo cargado en la app, aunque ahora compartan el mismo recorte por período RRHH.
                Sí entran en la evolución, que combina las dos fuentes.
              </p>
            </div>
          </>
        )}
      </Card>

          {/* ── Motivos ───────────────────────────────────────────── */}
          <Card className="p-5">
            <SectionTitle icon={ListChecks} accent="emerald">Horas por motivo</SectionTitle>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="py-2 pr-3">Categoría</th>
                    <th className="py-2 pr-3 text-right">Aprobadas</th>
                    <th className="py-2 pr-3 text-right">Pendientes</th>
                    <th className="py-2 pr-3 text-right">Solicitudes</th>
                  </tr>
                </thead>
                <tbody>
                  {porMotivo.map(m => (
                    <tr key={m.motivo} className="border-b border-slate-100">
                      <td className="py-2 pr-3 text-slate-700">
                        {m.motivo}
                        {extrasEsReactivo(m.motivo) && (
                          <span className="ml-1.5 text-[9px] px-1 py-0.5 bg-indigo-100 text-indigo-700 rounded font-bold"
                                title="Se ejecuta antes de la aprobación">reactivo</span>
                        )}
                      </td>
                      <td className="py-2 pr-3 num text-right font-semibold text-emerald-700">{formatHoras(m.aprobadas)}</td>
                      <td className="py-2 pr-3 num text-right text-amber-700">{formatHoras(m.pendientes)}</td>
                      <td className="py-2 pr-3 num text-right text-slate-500">{m.n}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>

          {/* ── Proceso ───────────────────────────────────────────── */}
          <Card className="p-5">
            <SectionTitle icon={Timer} accent="emerald">Proceso de aprobación</SectionTitle>
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-3">
              <div className="px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg">
                <div className="text-[11px] uppercase tracking-wide text-slate-500 font-semibold">Tasa de aprobación</div>
                <div className="text-xl font-bold text-slate-800 num mt-1">
                  {proceso.tasa === null ? '—' : `${proceso.tasa.toFixed(0)}%`}
                </div>
                <div className="text-[11px] text-slate-400 num">sobre {proceso.nResueltas} resueltas</div>
              </div>
              <div className="px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg">
                <div className="text-[11px] uppercase tracking-wide text-slate-500 font-semibold">Resolución · planificadas</div>
                <div className="text-xl font-bold text-slate-800 num mt-1">
                  {proceso.tPlanificadas === null ? '—' : formatDuracion(proceso.tPlanificadas)}
                </div>
                <div className="text-[11px] text-slate-400 num">{proceso.nPlanificadas} solicitudes</div>
              </div>
              <div className="px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg">
                <div className="text-[11px] uppercase tracking-wide text-slate-500 font-semibold">Resolución · reactivas</div>
                <div className="text-xl font-bold text-slate-800 num mt-1">
                  {proceso.tReactivas === null ? '—' : formatDuracion(proceso.tReactivas)}
                </div>
                <div className="text-[11px] text-slate-400 num">{proceso.nReactivas} solicitudes</div>
              </div>
            </div>

            {/* Notas al pie. Son parte del dato, no decoración: sin esto los
                números de arriba se leen como algo que no son. */}
            <div className="mt-4 text-[11px] text-slate-500 leading-relaxed space-y-1">
              <p>
                · Planificadas y reactivas se miden por separado a propósito. En las reactivas el trabajo ya se
                hizo cuando llega la aprobación, así que el tiempo mide demora administrativa, no velocidad de respuesta.
              </p>
              <p>
                · Las {proceso.autoaprobadas} solicitudes cargadas y aprobadas por la misma persona
                <strong> están incluidas</strong> en ambas métricas. Se autoaprueban en el momento, así que
                bajan el promedio de resolución y suben la tasa sin haber pasado por ninguna decisión.
              </p>
              <p>
                · Los tiempos negativos se cortan en cero: en las autoaprobadas el sello de resolución se pone
                milisegundos antes que el de creación, y eso es un artefacto del orden de guardado, no un dato.
              </p>
              <p>· Las anuladas no suman en ningún bloque. Aprobadas y pendientes nunca se suman entre sí.</p>
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

function ExtrasView({ sesion, extras, extrasLoading, extrasError, onAdd, onUpdate, onRefresh }) {
  const esJefe = sesion.rol === 'jefe';

  // Sub-vista (#49). Solo el jefe tiene dashboard; el encargado ve el listado
  // siempre y no necesita este estado, pero se declara igual para no meter un
  // hook condicional.
  const [vista, setVista] = useState('listado');

  // Form
  const [tecnico, setTecnico] = useState('');
  const [fecha, setFecha] = useState(todayLocalISO());
  const [horaInicio, setHoraInicio] = useState('');
  const [horaFin, setHoraFin] = useState('');
  const [motivoCat, setMotivoCat] = useState('');
  const [motivo, setMotivo] = useState('');
  const [editId, setEditId] = useState(null);
  // Fila completa en edición (v3.30): el submit necesita el estado, las horas
  // y la ventana ORIGINALES para decidir qué régimen de edición aplica
  // (pendiente: libre · aprobada no ejecutada: libre · aprobada/modificada ya
  // ejecutada: solo a la baja). editId solo no alcanza para esa decisión.
  const [editRow, setEditRow] = useState(null);
  const [msg, setMsg] = useState('');
  const [saving, setSaving] = useState(false);
  // Segundo click para aceptar un solapamiento detectado (#54).
  const [solapeOk, setSolapeOk] = useState(false);

  // Filtros del listado (#66 — también acotan lo que sale al Excel). El
  // rango de fechas libre de antes (filtroDesde/filtroHasta) se reemplaza por
  // Año+Mes de período RRHH, más un filtro por persona. Default: el período
  // RRHH actual.
  const hoyPeriodoListado = extrasPeriodoRRHH(todayLocalISO())
    || { anio: new Date().getFullYear(), mes: new Date().getMonth() + 1 };
  const [filtroAnio, setFiltroAnio] = useState(hoyPeriodoListado.anio);
  const [filtroMes, setFiltroMes] = useState(hoyPeriodoListado.mes);
  const [filtroPersona, setFiltroPersona] = useState('');
  const [filtroEstado, setFiltroEstado] = useState('');
  const [verAnuladas, setVerAnuladas] = useState(false);

  // Diálogos
  const [anularTarget, setAnularTarget] = useState(null);
  const [rechazarTarget, setRechazarTarget] = useState(null);

  // ── Visibilidad por rol (partición de UI, no de datos: ver nota en
  //    EXTRAS_USUARIOS). v3.28 (#58): el criterio pasó de POR AUTOR a POR
  //    PERSONA. El encargado ve las solicitudes de su gente a cargo aunque
  //    las haya cargado el jefe, y ya no ve las que él mismo cargó para
  //    alguien de otro equipo — porque tampoco puede cargarlas.
  //    Incluye las suyas propias: ver lo que a uno le cargaron es razonable
  //    aunque no pueda tocarlo.
  const aCargo = useMemo(
    () => extrasVisiblesDe(sesion.user, sesion.nombre),
    [sesion.user, sesion.nombre]
  );

  // A quién se le puede cargar: el jefe a cualquiera, el encargado solo a su
  // gente. No incluye su propio nombre aunque lo vea en el listado. Mismo
  // catálogo que alimenta el selector de "Persona" del filtro (#66).
  const personalCargable = useMemo(
    () => (esJefe ? EXTRAS_PERSONAL_NAMES : extrasPersonalDe(sesion.user)),
    [esJefe, sesion.user]
  );

  // ── Datos del listado por período RRHH (#66) ──────────────────────────
  // Antes este bloque leía del `extras` cacheado a nivel app (tope
  // EXTRAS_LIST_LIMIT=500, sin acotar por fecha): con el volumen real de
  // Facilities+Mantenimiento un período viejo podía quedar afuera del cupo
  // sin ningún aviso (BACKLOG #57). Ahora se pide por rango, mismo mecanismo
  // que ya usa el dashboard (`storage.listExtrasRango`), así que cada
  // período tiene su propia consulta. `extras` (el prop) se sigue usando tal
  // cual para el aviso de solapamiento del formulario, que necesita ver más
  // allá del período filtrado acá.
  const rangoListado = useMemo(() => extrasRangoRRHH(filtroAnio, filtroMes), [filtroAnio, filtroMes]);
  const [datosListado, setDatosListado] = useState([]);
  const [listadoLoading, setListadoLoading] = useState(true);
  const [listadoError, setListadoError] = useState('');

  const cargarListado = useCallback(() => {
    setListadoLoading(true);
    setListadoError('');
    return storage.listExtrasRango(rangoListado.desde, rangoListado.hasta)
      .then(r => setDatosListado(r))
      .catch(e => setListadoError(e.message || 'no se pudieron cargar las solicitudes'))
      .finally(() => setListadoLoading(false));
  }, [rangoListado.desde, rangoListado.hasta]);

  useEffect(() => { cargarListado(); }, [cargarListado]);

  // Años ofrecidos en el selector del listado: mismo criterio que el del
  // dashboard (ventana fija alrededor del actual + el elegido si cae afuera).
  const filtroAnioOptions = useMemo(() => {
    const actual = new Date().getFullYear();
    const arr = [];
    for (let y = actual - 2; y <= actual + 1; y++) arr.push(y);
    if (!arr.includes(filtroAnio)) arr.push(filtroAnio);
    return arr.sort((a, b) => a - b);
  }, [filtroAnio]);

  const visibles = useMemo(
    () => (esJefe ? datosListado : datosListado.filter(r => aCargo.includes(r.tecnico_nombre))),
    [datosListado, esJefe, aCargo]
  );

  const listado = useMemo(() => visibles
    .filter(r => (verAnuladas ? true : !r.anulada_at))
    .filter(r => (filtroEstado ? r.estado === filtroEstado : true))
    .filter(r => (filtroPersona ? r.tecnico_nombre === filtroPersona : true)),
    [visibles, verAnuladas, filtroEstado, filtroPersona]);

  // Totales sobre lo que se está viendo. Las anuladas NO suman nunca, aunque
  // estén visibles con el checkbox: una hora anulada no se trabaja ni se paga.
  const totales = useMemo(() => {
    const vivas = listado.filter(r => !r.anulada_at);
    const sum = (est) => vivas
      .filter(r => r.estado === est)
      .reduce((a, r) => a + (Number(r.horas) || 0), 0);
    return {
      pendientes: sum('pendiente'),
      // 'modificada' suma junto a aprobadas: ya son horas aprobadas, con un
      // ajuste a la baja esperando confirmación (v3.30).
      aprobadas: sum('aprobada') + sum('modificada'),
      nPendientes: vivas.filter(r => r.estado === 'pendiente' || r.estado === 'modificada').length
    };
  }, [listado]);

  // ── Contador post-login (#62). El aviso naranja del header y el chip de la
  //    pantalla de login son PRE-login: sin sesión la app no sabe quién está
  //    del otro lado, así que solo pueden hablar del sector entero. Esto es lo
  //    contrario — ya hay sesión, así que dice exactamente lo que le toca a
  //    quien está mirando: al jefe lo que tiene para resolver, al encargado
  //    cuántas de su gente están esperando aprobación.
  //
  //    Se calcula sobre `visibles`, NO sobre `listado`: los filtros de estado
  //    y fecha del listado son para explorar, y un aviso que desaparece porque
  //    alguien filtró por "aprobadas" no avisa nada. Las anuladas nunca suman.
  const pendientesPropios = useMemo(
    () => visibles.filter(r => !r.anulada_at && (r.estado === 'pendiente' || r.estado === 'modificada')),
    [visibles]
  );
  const horasPendientes = useMemo(
    () => pendientesPropios.reduce((a, r) => a + (Number(r.horas) || 0), 0),
    [pendientesPropios]
  );

  const cruza = extrasCruzaMedianoche(horaInicio, horaFin);
  const horasPreview = extrasHorasCalc(fecha, horaInicio, horaFin);

  // ── Solapamiento con extras ya cargados de la misma persona (#54).
  //    Se busca sobre `extras` COMPLETO, no sobre `visibles`: el encargado no
  //    ve lo que cargó otro encargado, pero igual tiene que ser advertido si
  //    se pisa con ello — si no, la duplicación queda invisible justo para
  //    quien la está creando. Es el único lugar donde la partición de UI se
  //    saltea a propósito.
  //    No cuentan las anuladas ni las rechazadas: esas horas no se pagan.
  const solapes = useMemo(() => {
    if (!tecnico || !fecha || !horaInicio || !horaFin || horasPreview <= 0) return [];
    const nueva = {
      ini: `${fecha}T${horaInicio}`,
      fin: `${extrasFechaFin(fecha, horaInicio, horaFin)}T${horaFin}`
    };
    return extras.filter(r =>
      r.tecnico_nombre === tecnico &&
      !r.anulada_at &&
      r.estado !== 'rechazada' &&
      r.id !== editId &&
      extrasSolapan(nueva, extrasVentana(r))
    );
  }, [extras, tecnico, fecha, horaInicio, horaFin, horasPreview, editId]);

  // Cualquier cambio en la ventana o la persona invalida la confirmación
  // anterior: si no, se acepta un solapamiento y después se cambia la hora.
  useEffect(() => { setSolapeOk(false); }, [tecnico, fecha, horaInicio, horaFin]);

  const resetForm = () => {
    setEditId(null);
    setEditRow(null);
    setTecnico('');
    setHoraInicio('');
    setHoraFin('');
    setMotivoCat('');
    setMotivo('');
    setSolapeOk(false);
    // La fecha NO se resetea: al cargar varios extras del mismo día seguidos,
    // volver a tipearla cada vez es la forma más rápida de equivocarse.
  };

  const cargarParaEditar = (r) => {
    setEditId(r.id);
    setEditRow(r);
    setTecnico(r.tecnico_nombre || '');
    setFecha(r.fecha);
    setHoraInicio((r.hora_inicio || '').slice(0, 5));
    setHoraFin((r.hora_fin || '').slice(0, 5));
    setMotivoCat(r.motivo_categoria || '');
    setMotivo(r.motivo || '');
    setSolapeOk(false);
    setMsg('');
    if (typeof window !== 'undefined') window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const submit = async () => {
    if (!tecnico) { setMsg('Error: elegí a la persona.'); return; }
    // Defensa además del select: si el nombre quedó de una edición previa y
    // ya no está a cargo, no se guarda. El select solo oculta la opción.
    if (!personalCargable.includes(tecnico)) {
      setMsg(`Error: ${tecnico} no está a tu cargo. Esa carga la hace el jefe.`);
      return;
    }
    if (!fecha) { setMsg('Error: falta la fecha.'); return; }
    if (!horaInicio || !horaFin) { setMsg('Error: faltan las horas de inicio y fin.'); return; }
    if (horasPreview <= 0) { setMsg('Error: la ventana horaria es inválida.'); return; }
    if (horasPreview > 24) { setMsg('Error: más de 24 h en un solo extra — revisá las horas.'); return; }
    if (!motivoCat) { setMsg('Error: elegí la categoría del motivo.'); return; }
    if (extrasRequiereDetalle(motivoCat) && !motivo.trim()) {
      setMsg(`Error: "${motivoCat}" necesita que detalles el trabajo.`);
      return;
    }

    // Aviso blando de solapamiento (#54): no bloquea, pide un segundo click.
    // Deliberadamente NO es un hard-block: encadenar extras sobre un trabajo
    // que se extendió es un caso legítimo y frecuente.
    if (solapes.length > 0 && !solapeOk) {
      const cual = solapes[0];
      setSolapeOk(true);
      setMsg(`Error: ${tecnico} ya tiene un extra que se pisa con este (${formatDateShort(cual.fecha)} ${(cual.hora_inicio || '').slice(0, 5)}–${(cual.hora_fin || '').slice(0, 5)}). Si es correcto, volvé a tocar el botón para confirmar.`);
      return;
    }

    setSaving(true);
    setMsg('');
    try {
      const datos = {
        tecnico_id: findTecnicoId(tecnico) || null,
        tecnico_nombre: tecnico,
        fecha,
        hora_inicio: horaInicio,
        fecha_fin: extrasFechaFin(fecha, horaInicio, horaFin),
        hora_fin: horaFin,
        motivo_categoria: motivoCat,
        // Vacío es válido para las categorías de cobertura. La constraint
        // `horas_extras_motivo_no_vacio` se dropeó en v3.27 justamente para
        // permitirlo; el Excel muestra la categoría cuando el detalle falta.
        motivo: motivo.trim()
        // `ots` NO se manda (#55): la columna sigue existiendo en la base con
        // DEFAULT '{}' NOT NULL, así que las filas nuevas nacen con array
        // vacío sin que el cliente la toque.
      };

      if (editId) {
        // v3.30 (#64) — Tres regímenes de edición según el estado ORIGINAL
        // de la fila (editRow), no del valor que se está por guardar:
        //
        // 1. pendiente: como siempre, solo datos. El estado no se toca.
        //
        // 2. aprobada, todavía no se ejecutó (fecha_fin+hora_fin en el
        //    futuro): edición libre, igual que pendiente, pero vuelve a
        //    'pendiente' — el jefe tiene que volver a mirarla. Resetea la
        //    traza de resolución porque la constraint de la base exige
        //    resuelto_por/resuelto_at = NULL para estado='pendiente'.
        //
        // 3. aprobada o modificada, YA se ejecutó: el trabajo ya pasó, así
        //    que solo se puede CORREGIR A LA BAJA (el encargado detectó que
        //    se hicieron menos horas de las aprobadas). Si el valor nuevo
        //    es igual o mayor, se bloquea acá mismo — el excedente se carga
        //    como fila nueva con "Finalización de trabajos en curso", no
        //    editando esta. Pasa a 'modificada' y conserva resuelto_por/
        //    resuelto_at (quién había aprobado la versión anterior): el
        //    jefe la confirma con el mismo botón de aprobar, sin rechazo
        //    (no tiene sentido rechazar un hecho ya ejecutado).
        //    horas_previas guarda el valor ANTERIOR solo la primera vez
        //    (viniendo de 'aprobada'): si ya estaba en 'modificada' y se
        //    vuelve a bajar, se conserva el original, no el intermedio.
        if (editRow && editRow.estado === 'aprobada' && !extrasYaEjecutada(editRow)) {
          await onUpdate(editId, {
            ...datos,
            estado: 'pendiente',
            resuelto_por: null,
            resuelto_por_nombre: null,
            resuelto_at: null
          });
          setMsg('✓ Solicitud actualizada · vuelve a estar pendiente de aprobación');
        } else if (editRow && (editRow.estado === 'aprobada' || editRow.estado === 'modificada') && extrasYaEjecutada(editRow)) {
          const horasAntes = Number(editRow.horas) || 0;
          if (horasPreview >= horasAntes) {
            setMsg(`Error: ya se ejecutó — solo se puede corregir a la baja (tenía ${formatHoras(horasAntes)}). Si se hicieron MÁS horas, cargá una solicitud nueva con "Finalización de trabajos en curso".`);
            setSaving(false);
            return;
          }
          await onUpdate(editId, {
            ...datos,
            estado: 'modificada',
            ...(editRow.estado === 'aprobada' ? { horas_previas: horasAntes } : {})
          });
          setMsg('✓ Ajuste registrado · esperando que el jefe lo confirme');
        } else {
          // Edición de una pendiente propia: solo datos. El estado, el
          // solicitante y la traza de resolución no se tocan desde acá.
          await onUpdate(editId, datos);
          setMsg('✓ Solicitud actualizada');
        }
      } else {
        const ahora = new Date().toISOString();
        await onAdd({
          ...datos,
          // #62 — El sector lo pone el DEPLOY, no el usuario ni un selector.
          // No se manda en el PATCH de edición: el sector de una fila no
          // cambia nunca, y dejarlo fuera del update lo vuelve imposible de
          // pisar por accidente desde la pantalla de edición.
          sector: APP_SECTOR,
          solicitado_por: sesion.user,
          solicitado_por_nombre: sesion.nombre,
          // El jefe autoaprueba lo que carga: no tiene a quién elevárselo.
          // Queda registrado como resuelto por él mismo, que es lo que pasó.
          estado: esJefe ? 'aprobada' : 'pendiente',
          resuelto_por: esJefe ? sesion.user : null,
          resuelto_por_nombre: esJefe ? sesion.nombre : null,
          resuelto_at: esJefe ? ahora : null
        });
        setMsg(esJefe ? '✓ Horas extras cargadas y aprobadas' : '✓ Solicitud enviada · queda pendiente de aprobación');
      }
      resetForm();
      cargarListado();
    } catch (e) {
      setMsg(`Error: ${e.message || 'no se pudo guardar'}`);
    } finally {
      setSaving(false);
    }
  };

  // Sirve para dos casos (v3.30): aprobar una pendiente nueva, y confirmar
  // el ajuste a la baja de una 'modificada' — mismo destino (estado
  // 'aprobada'), mismo botón, mensaje distinto solo para orientar al jefe.
  const aprobar = async (r) => {
    try {
      await onUpdate(r.id, {
        estado: 'aprobada',
        resuelto_por: sesion.user,
        resuelto_por_nombre: sesion.nombre,
        resuelto_at: new Date().toISOString(),
        rechazo_motivo: null,
        horas_previas: null
      });
      setMsg(r.estado === 'modificada' ? `✓ Confirmada · ${r.tecnico_nombre}` : `✓ Aprobada · ${r.tecnico_nombre}`);
      cargarListado();
    } catch (e) {
      setMsg(`Error: ${e.message || 'no se pudo aprobar'}`);
    }
  };

  const doRechazar = async (motivoRechazo) => {
    await onUpdate(rechazarTarget.id, {
      estado: 'rechazada',
      resuelto_por: sesion.user,
      resuelto_por_nombre: sesion.nombre,
      resuelto_at: new Date().toISOString(),
      rechazo_motivo: motivoRechazo || null
    });
    setRechazarTarget(null);
    setMsg('✓ Solicitud rechazada');
    cargarListado();
  };

  const doAnular = async (motivoAnulacion) => {
    await onUpdate(anularTarget.id, {
      anulada_at: new Date().toISOString(),
      anulada_por: sesion.nombre,
      anulada_motivo: motivoAnulacion
    });
    setAnularTarget(null);
    setMsg('✓ Solicitud anulada');
    cargarListado();
  };

  // Permisos de fila. El encargado toca solo lo propio; el jefe resuelve lo
  // pendiente/modificado de otros y anula cualquier cosa, en cualquier estado.
  //
  // v3.30 (#64) — puedeEditar se extiende a 'aprobada' y 'modificada': el
  // botón se muestra en los tres estados, y es el submit el que decide el
  // régimen exacto (libre → pendiente, o solo-a-la-baja → modificada) según
  // si la fila ya se ejecutó. 'rechazada' sigue sin edición.
  const puedeEditar = (r) =>
    !r.anulada_at &&
    r.solicitado_por === sesion.user &&
    (r.estado === 'pendiente' || r.estado === 'aprobada' || r.estado === 'modificada');
  const puedeAnular = (r) => !r.anulada_at && (esJefe || (r.estado === 'pendiente' && r.solicitado_por === sesion.user));
  // 'modificada' se resuelve solo con Confirmar (ver JSX): no tiene sentido
  // "rechazar" horas que ya se ejecutaron y se corrigieron a la baja.
  const puedeResolver = (r) => esJefe && !r.anulada_at && (r.estado === 'pendiente' || r.estado === 'modificada');

  // ── Export a Excel. Sale EXACTAMENTE lo que se está viendo (mismo rol, mismos
  //    filtros), para que el archivo no pueda contener algo que en pantalla no
  //    estaba. Una fila por solicitud.
  const exportar = () => {
    if (listado.length === 0) { setMsg('Error: no hay filas para exportar con estos filtros.'); return; }
    const rows = listado.map(r => ({
      // El encabezado sigue diciendo "Técnico" a propósito (#53): renombrarlo
      // rompería cualquier planilla o proceso que consuma la columna por
      // nombre. En pantalla sí dice "Personal".
      'Técnico': r.tecnico_nombre,
      'Fecha': r.fecha,
      'Hora inicio': (r.hora_inicio || '').slice(0, 5),
      'Fecha fin': r.fecha_fin,
      'Hora fin': (r.hora_fin || '').slice(0, 5),
      'Horas': Number(r.horas) || 0,
      // Solo tiene valor en 'modificada' (v3.30): las horas aprobadas antes
      // de la corrección a la baja. Vacío en el resto — no repite 'Horas'.
      'Horas previas': r.horas_previas != null ? Number(r.horas_previas) : '',
      'Categoría': r.motivo_categoria || '',
      // Fallback deliberado: las categorías de cobertura no exigen detalle, y
      // una celda vacía en una planilla que puede ir a RRHH se lee como dato
      // faltante. En la BASE `motivo` queda vacío — no se duplica la categoría
      // — para que siga siendo consultable cuáles tienen detalle real.
      'Motivo': (r.motivo || '').trim() || r.motivo_categoria || '',
      'Estado': r.anulada_at ? 'ANULADA' : r.estado,
      'Solicitada por': r.solicitado_por_nombre || '',
      // Trazabilidad: en qué FECHA se pidió y en qué fecha se resolvió, en
      // calendario local. Antes se exportaba el timestamptz crudo, que sale en
      // UTC y corría el día en todo lo cargado después de las 21:00.
      'Fecha de solicitud': formatFechaAudit(r.created_at),
      'Resuelta por': r.resuelto_por_nombre || '',
      'Fecha de resolución': r.resuelto_at ? formatFechaAudit(r.resuelto_at) : '',
      'Motivo de rechazo': r.rechazo_motivo || '',
      'Anulada por': r.anulada_por || '',
      'Fecha de anulación': r.anulada_at ? formatFechaAudit(r.anulada_at) : '',
      'Motivo de anulación': r.anulada_motivo || ''
    }));

    // #66 — Segunda pestaña: totales por persona sobre EXACTAMENTE las
    // mismas filas de la pestaña 1 (mismos filtros de estado/persona/
    // anuladas ya aplicados en `listado`). Si se filtra por un estado
    // puntual, el total refleja ese estado — no hay un criterio fijo de
    // "solo aprobadas" que ignore el filtro elegido en pantalla. Orden
    // alfabético, igual que las planillas de RRHH que arma Leo a mano.
    const porPersonaMap = new Map();
    listado.forEach(r => {
      const k = r.tecnico_nombre || '(sin nombre)';
      const o = porPersonaMap.get(k) || { persona: k, solicitudes: 0, horas: 0 };
      o.solicitudes += 1;
      o.horas += Number(r.horas) || 0;
      porPersonaMap.set(k, o);
    });
    const totalesRows = [...porPersonaMap.values()]
      .sort((a, b) => a.persona.localeCompare(b.persona, 'es'))
      .map(o => ({
        'Persona': o.persona,
        'Cantidad de solicitudes': o.solicitudes,
        'Horas totales': Number(o.horas.toFixed(2))
      }));

    // #62 — El sector va en el NOMBRE DEL ARCHIVO, no en una columna nueva
    // (Leo confirmó que no hay consolidado entre sectores). #66 — el rango
    // de fechas libre del nombre pasa a Año-Mes del período RRHH filtrado.
    const periodoArchivo = `${filtroAnio}-${String(filtroMes).padStart(2, '0')}`;
    downloadStyledExtras(
      [
        { rows, name: 'Horas extras' },
        { rows: totalesRows, name: 'Totales por persona' }
      ],
      `HorasExtras_${APP_SECTOR}_${periodoArchivo}.xlsx`
    );
    setMsg(`✓ Exportadas ${rows.length} solicitudes`);
  };

  if (!storage.configured) {
    return (
      <Card className="p-5">
        <div className="flex items-center gap-2 text-sm text-amber-800">
          <AlertTriangle className="w-4 h-4 flex-shrink-0" />
          Las horas extras requieren Supabase configurado. En modo local no están disponibles.
        </div>
      </Card>
    );
  }

  return (
    <div className="flex flex-col gap-5">
      {/* ── Toggle de sub-vista (#49). Desde v3.28 también el encargado,
             con el dashboard recortado a su gente a cargo (#58). ─────── */}
      <div className="flex items-center gap-2">
        <button onClick={() => setVista('listado')}
          className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg transition ${
            vista === 'listado' ? 'bg-slate-800 text-white shadow-sm' : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-300'
          }`}>
          <ListChecks className="w-3.5 h-3.5" />Listado
        </button>
        <button onClick={() => setVista('dashboard')}
          className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg transition ${
            vista === 'dashboard' ? 'bg-slate-800 text-white shadow-sm' : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-300'
          }`}>
          <BarChart3 className="w-3.5 h-3.5" />Dashboard
        </button>
        {!esJefe && vista === 'dashboard' && (
          <span className="text-[11px] text-slate-500">
            Tu gente a cargo · {extrasPersonalDe(sesion.user).length} personas
          </span>
        )}

        {/* Contador post-login (#62). Visible en las dos sub-vistas, y no se
            renderiza cuando no hay nada pendiente: un badge en cero es ruido
            permanente y deja de leerse justo cuando importa. */}
        {pendientesPropios.length > 0 && (
          <span
            className={`ml-auto inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] font-semibold rounded-lg ring-1 ${
              esJefe
                ? 'bg-orange-50 text-orange-800 ring-orange-300'
                : 'bg-amber-50 text-amber-800 ring-amber-300'
            }`}
            title={esJefe
              ? 'Solicitudes de tu sector esperando que las apruebes o rechaces'
              : 'Solicitudes de tu gente a cargo esperando aprobación del jefe'}>
            <AlertTriangle className="w-3.5 h-3.5" />
            <span className="num">{pendientesPropios.length}</span>
            {esJefe
              ? (pendientesPropios.length === 1 ? ' solicitud para resolver' : ' solicitudes para resolver')
              : (pendientesPropios.length === 1 ? ' de tu gente sin aprobar' : ' de tu gente sin aprobar')}
            <span className="text-[10px] font-normal opacity-75">· {formatHoras(horasPendientes)}</span>
          </span>
        )}
      </div>

      {vista === 'dashboard' ? (
        <ExtrasDashboard soloPersonas={esJefe ? null : aCargo} />
      ) : (
      <>
      {/* ── ALTA / EDICIÓN ─────────────────────────────────────────── */}
      <Card className="p-5">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <SectionTitle icon={Timer} accent={esJefe ? 'emerald' : 'cyan'}>
            {editId ? 'Editar solicitud' : (esJefe ? 'Cargar horas extras' : 'Solicitar horas extras')}
          </SectionTitle>
          <span className="text-xs text-slate-500 mb-4">
            {sesion.nombre} · <span className="font-semibold">{esJefe ? 'jefe' : EXTRAS_ETIQUETA_ENCARGADO}</span>
            <span className="text-slate-400"> · {APP_SECTOR}</span>
          </span>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-4 gap-4">
          <Field label="Personal">
            {/* Jefe: EXTRAS_PERSONAL_NAMES completo. Encargado: solo su gente
                a cargo (#58). El resto lo carga el jefe. */}
            <select className={inputCls} value={tecnico} onChange={e => setTecnico(e.target.value)}>
              <option value="">—</option>
              {personalCargable.map(n => <option key={n} value={n}>{n}</option>)}
            </select>
          </Field>

          {/* Sin `max`: las fechas futuras son válidas y son el caso normal
              (se solicita antes de hacer el trabajo). Ver #11 para el contraste
              con los reportes, donde el futuro sí está bloqueado. */}
          <Field label="Fecha del extra">
            <input type="date" className={`${inputCls} num`} value={fecha}
              onChange={e => setFecha(e.target.value)} />
          </Field>

          <Field label="Hora de inicio">
            <TimeInput24 value={horaInicio} onChange={setHoraInicio} />
          </Field>

          <Field label="Hora de fin">
            <div className="flex items-center gap-2 flex-wrap">
              <TimeInput24 value={horaFin} onChange={setHoraFin} />
              {cruza && (
                <span className="text-[10px] px-1.5 py-1 bg-indigo-100 text-indigo-700 rounded font-bold whitespace-nowrap"
                      title={`Cruza medianoche: termina el ${formatDateShort(extrasFechaFin(fecha, horaInicio, horaFin))}`}>
                  +1 día
                </span>
              )}
            </div>
          </Field>

          <Field label="Motivo">
            <select className={inputCls} value={motivoCat}
              onChange={e => { setMotivoCat(e.target.value); }}>
              <option value="">—</option>
              {EXTRAS_MOTIVO_CATEGORIAS.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </Field>

          <Field
            label={extrasRequiereDetalle(motivoCat) ? 'Detalle del trabajo *' : 'Detalle (opcional)'}
            className="lg:col-span-2">
            <input type="text" className={inputCls} value={motivo}
              onChange={e => setMotivo(e.target.value)}
              placeholder={
                !motivoCat ? 'Elegí primero una categoría…'
                  : extrasRequiereDetalle(motivoCat)
                    ? 'Qué trabajo se hizo — obligatorio para esta categoría'
                    : 'Opcional: a quién cubre, o cualquier aclaración'
              } />
          </Field>

          <Field label="Duración">
            <div className="px-3 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg num font-semibold text-slate-700">
              {horasPreview > 0 ? formatHoras(horasPreview) : '—'}
            </div>
          </Field>
        </div>

        {/* Aviso de solapamiento (#54). Informativo: no impide guardar. */}
        {solapes.length > 0 && (
          <div className="mt-4 flex items-start gap-2 px-3 py-2 bg-amber-50 border border-amber-200 rounded-lg">
            <AlertTriangle className="w-4 h-4 text-amber-600 flex-shrink-0 mt-0.5" />
            <div className="text-xs text-amber-800 leading-relaxed">
              <strong>{tecnico}</strong> ya tiene {solapes.length === 1 ? 'un extra' : `${solapes.length} extras`} que se
              {solapes.length === 1 ? ' pisa' : ' pisan'} con esta ventana:
              <ul className="mt-1 space-y-0.5">
                {solapes.slice(0, 4).map(s => (
                  <li key={s.id} className="num">
                    · {formatDateShort(s.fecha)} {(s.hora_inicio || '').slice(0, 5)}–{(s.hora_fin || '').slice(0, 5)}
                    <span className="font-sans"> · {s.motivo_categoria || s.motivo || 'sin motivo'} · {s.estado}</span>
                  </li>
                ))}
                {solapes.length > 4 && <li className="text-amber-700">· y {solapes.length - 4} más</li>}
              </ul>
              <div className="mt-1">
                Si el trabajo se extendió y esto es una continuación, está bien — confirmá y seguí.
                Si es una carga repetida, corregila: se pagaría dos veces.
              </div>
            </div>
          </div>
        )}

        {/* Los motivos reactivos se ejecutan antes de la aprobación. */}
        {extrasEsReactivo(motivoCat) && (
          <p className="text-[11px] text-slate-500 mt-3 leading-relaxed">
            Este motivo se registra normalmente <strong>después</strong> de hecho el trabajo. La aprobación queda como
            acto administrativo posterior, y el dashboard lo cuenta aparte de los motivos planificados.
          </p>
        )}

        <div className="flex flex-wrap items-center gap-3 mt-4">
          <button onClick={submit} disabled={saving}
            className={`inline-flex items-center gap-2 px-5 py-2.5 text-white rounded-xl font-bold text-sm transition disabled:opacity-50 shadow-md ${esJefe ? 'bg-emerald-600 hover:bg-emerald-500' : 'bg-cyan-600 hover:bg-cyan-500'}`}>
            {editId ? <Save className="w-4 h-4" /> : <Plus className="w-4 h-4" />}
            {saving ? 'Guardando…' : (editId ? 'Guardar cambios' : (esJefe ? 'Cargar y aprobar' : 'Solicitar'))}
          </button>
          {editId && (
            <button onClick={() => { resetForm(); setMsg(''); }}
              className="px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition">
              Cancelar edición
            </button>
          )}
          {msg && (
            <span className={`text-xs font-medium ${msg.startsWith('Error') ? 'text-red-600' : 'text-emerald-600'}`}>
              {msg}
            </span>
          )}
        </div>

        <p className="text-[11px] text-slate-400 mt-3 leading-relaxed">
          {esJefe
            ? 'Lo que cargues acá queda aprobado en el momento, registrado a tu nombre.'
            : `La solicitud queda pendiente hasta que el jefe la apruebe. Mientras esté pendiente la podés editar o anular. Solo podés cargar a tu gente a cargo (${extrasPersonalDe(sesion.user).length} personas); el resto lo carga el jefe.`}
        </p>
        {/* #62 — Qué sector está registrando. En el deploy solo-Extras es
            redundante con el header, pero en la app completa es la única
            señal de que este registro es de Mantenimiento y no del otro
            sector. Barato, y evita la pregunta. */}
        <p className="text-[11px] text-slate-400 mt-1 leading-relaxed">
          Sector: <strong className="text-slate-500">{APP_SECTOR}</strong>. Cada sector ve y exporta
          únicamente sus propias horas extras.
        </p>
      </Card>

      {/* ── LISTADO ────────────────────────────────────────────────── */}
      <Card className="p-5">
        <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
          <SectionTitle icon={ListChecks} accent="slate">
            {esJefe ? 'Todas las solicitudes' : 'Mi gente a cargo'}
          </SectionTitle>
          <div className="flex flex-wrap items-center gap-2">
            <select className="px-2 py-1.5 text-xs bg-white border border-slate-300 rounded-lg"
              value={filtroEstado} onChange={e => setFiltroEstado(e.target.value)}>
              <option value="">Todos los estados</option>
              <option value="pendiente">Pendientes</option>
              <option value="aprobada">Aprobadas</option>
              <option value="rechazada">Rechazadas</option>
            </select>
            {/* #66 — Persona (mismo criterio que el selector de carga: jefe ve
                todo el personal, encargado solo su gente a cargo) y Año/Mes
                período RRHH, en vez del rango de fechas libre de antes. */}
            <select className="px-2 py-1.5 text-xs bg-white border border-slate-300 rounded-lg max-w-[160px]"
              value={filtroPersona} onChange={e => setFiltroPersona(e.target.value)}>
              <option value="">Toda la gente</option>
              {personalCargable.map(n => <option key={n} value={n}>{n}</option>)}
            </select>
            <select className="px-2 py-1.5 text-xs bg-white border border-slate-300 rounded-lg capitalize"
              value={filtroMes} onChange={e => setFiltroMes(Number(e.target.value))}>
              {MESES_CORTOS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
            </select>
            <select className="px-2 py-1.5 text-xs bg-white border border-slate-300 rounded-lg"
              value={filtroAnio} onChange={e => setFiltroAnio(Number(e.target.value))}>
              {filtroAnioOptions.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
            <label className="flex items-center gap-1.5 text-xs text-slate-600">
              <input type="checkbox" checked={verAnuladas} onChange={e => setVerAnuladas(e.target.checked)} />
              Ver anuladas
            </label>
            <button onClick={exportar}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-semibold bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg transition"
              title="Exportar a Excel lo que se ve con estos filtros">
              <Download className="w-3.5 h-3.5" />Excel
            </button>
            <button onClick={() => { onRefresh(); cargarListado(); }}
              className="p-1.5 hover:bg-slate-100 rounded transition" title="Refrescar">
              <RefreshCw className="w-4 h-4 text-slate-500" />
            </button>
          </div>
        </div>

        <p className="text-[11px] text-slate-400 -mt-2 mb-3 num">
          Período RRHH (11 al 10): {formatDateShort(rangoListado.desde)} — {formatDateShort(rangoListado.hasta)}
        </p>

        <div className="flex flex-wrap gap-4 mb-4 text-xs text-slate-600">
          <span><strong className="num text-amber-700">{formatHoras(totales.pendientes)}</strong> pendientes de aprobación
            {totales.nPendientes > 0 && <span className="num"> ({totales.nPendientes})</span>}</span>
          <span><strong className="num text-emerald-700">{formatHoras(totales.aprobadas)}</strong> aprobadas</span>
          <span><strong className="num text-slate-800">{listado.length}</strong> solicitudes en pantalla</span>
        </div>

        {listadoError && (
          <div className="mb-3 text-xs text-red-600 flex items-center gap-1.5">
            <AlertTriangle className="w-3.5 h-3.5" />{listadoError}
          </div>
        )}

        {listadoLoading ? (
          <div className="text-center text-slate-500 py-10 text-sm">Cargando horas extras…</div>
        ) : listado.length === 0 ? (
          <EmptyHint>No hay solicitudes con estos filtros.</EmptyHint>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="py-2 pr-3">Personal</th>
                  <th className="py-2 pr-3">Fecha</th>
                  <th className="py-2 pr-3">Horario</th>
                  <th className="py-2 pr-3 text-right">Horas</th>
                  <th className="py-2 pr-3">Motivo</th>
                  <th className="py-2 pr-3">Estado</th>
                  {/* v3.28 (#58): visible también para el encargado. Ahora ve
                      filas que cargó el jefe, y solo puede editar las propias:
                      sin esta columna no puede saber cuáles son cuáles. */}
                  <th className="py-2 pr-3">Solicitó</th>
                  <th className="py-2 pr-3">Cargada</th>
                  <th className="py-2 pr-3"></th>
                </tr>
              </thead>
              <tbody>
                {listado.map(r => {
                  const cruzaRow = r.fecha_fin !== r.fecha;
                  return (
                    <tr key={r.id} className={`border-b border-slate-100 ${r.anulada_at ? 'opacity-50' : ''}`}>
                      <td className="py-2 pr-3 font-semibold text-slate-800 whitespace-nowrap">
                        {r.anulada_at ? <s>{r.tecnico_nombre}</s> : r.tecnico_nombre}
                      </td>
                      <td className="py-2 pr-3 num text-slate-600 whitespace-nowrap">{formatDateShort(r.fecha)}</td>
                      <td className="py-2 pr-3 num text-slate-600 whitespace-nowrap">
                        {(r.hora_inicio || '').slice(0, 5)}–{(r.hora_fin || '').slice(0, 5)}
                        {cruzaRow && (
                          <span className="ml-1 text-[9px] px-1 py-0.5 bg-indigo-100 text-indigo-700 rounded font-bold"
                                title={`Termina el ${formatDateShort(r.fecha_fin)}`}>+1</span>
                        )}
                      </td>
                      <td className="py-2 pr-3 num text-right font-semibold text-slate-800 whitespace-nowrap">
                        {r.estado === 'modificada' && r.horas_previas != null ? (
                          <span title="Ajustado tras la ejecución — esperando confirmación">
                            <span className="text-slate-400 line-through font-normal">{formatHoras(r.horas_previas)}</span>
                            {' → '}{formatHoras(r.horas)}
                          </span>
                        ) : formatHoras(r.horas)}
                      </td>
                      <td className="py-2 pr-3 text-slate-600 max-w-[260px]"
                          title={[r.motivo_categoria, r.motivo].filter(Boolean).join(' — ')}>
                        <div className="font-medium text-slate-700 truncate">
                          {r.motivo_categoria || <span className="text-slate-400 italic">sin categoría</span>}
                        </div>
                        {(r.motivo || '').trim() && (
                          <div className="text-[11px] text-slate-500 truncate">{r.motivo}</div>
                        )}
                      </td>
                      <td className="py-2 pr-3 whitespace-nowrap">
                        <span className={`text-[10px] px-1.5 py-0.5 rounded font-bold ${EXTRAS_ESTADO_STYLE[r.estado] || 'bg-slate-200 text-slate-600'}`}
                              title={[
                                r.resuelto_por_nombre ? `Resuelta por ${r.resuelto_por_nombre}` : '',
                                r.rechazo_motivo ? `Motivo: ${r.rechazo_motivo}` : ''
                              ].filter(Boolean).join(' · ')}>
                          {r.estado}
                        </span>
                        {r.anulada_at && (
                          <span className="ml-1 text-[10px] px-1.5 py-0.5 rounded font-bold bg-slate-200 text-slate-600"
                                title={`${formatFechaAudit(r.anulada_at)}${r.anulada_por ? ` — ${r.anulada_por}` : ''}${r.anulada_motivo ? `: ${r.anulada_motivo}` : ''}`}>
                            anulada
                          </span>
                        )}
                        {/* Cuándo se resolvió y quién, debajo del estado. */}
                        {r.resuelto_at && (
                          <div className="text-[10px] text-slate-400 num whitespace-nowrap mt-0.5"
                               title={r.resuelto_por_nombre ? `Resuelta por ${r.resuelto_por_nombre}` : ''}>
                            {formatFechaAudit(r.resuelto_at)}
                          </div>
                        )}
                      </td>
                      <td className="py-2 pr-3 text-slate-500 text-xs whitespace-nowrap">
                        {r.solicitado_por === sesion.user
                          ? <span className="font-semibold text-slate-700">vos</span>
                          : (r.solicitado_por_nombre || '—')}
                      </td>
                      {/* Momento en que el encargado dejó asentada la solicitud.
                          Sombreado rojo translúcido (v3.31) si se cargó en un día
                          posterior a la ejecución del trabajo — salvo "Finalización
                          de trabajos en curso", donde eso es lo esperado. */}
                      <td className={`py-2 px-1 text-[11px] num whitespace-nowrap rounded ${
                            extrasCargaTardia(r) ? 'bg-red-500/15 text-red-800 font-semibold' : 'text-slate-400'
                          }`}
                          title={extrasCargaTardia(r)
                            ? `Cargada después de que el trabajo terminó (fin: ${formatDateShort(r.fecha_fin)})`
                            : undefined}>
                        {formatFechaAudit(r.created_at)}
                      </td>
                      <td className="py-2 pr-3">
                        <div className="flex items-center justify-end gap-1">
                          {puedeResolver(r) && (
                            <>
                              <button onClick={() => aprobar(r)}
                                className="p-1 hover:bg-emerald-50 rounded transition"
                                title={r.estado === 'modificada' ? 'Confirmar el ajuste' : 'Aprobar'}>
                                <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                              </button>
                              {/* 'modificada' no tiene Rechazar: el trabajo ya se ejecutó, no
                                  hay nada que rechazar, solo tomar nota de la baja. */}
                              {r.estado === 'pendiente' && (
                                <button onClick={() => setRechazarTarget(r)}
                                  className="p-1 hover:bg-red-50 rounded transition" title="Rechazar">
                                  <X className="w-4 h-4 text-red-500" />
                                </button>
                              )}
                            </>
                          )}
                          {puedeEditar(r) && (
                            <button onClick={() => cargarParaEditar(r)}
                              className="p-1 hover:bg-sky-50 rounded transition"
                              title={r.estado === 'pendiente' ? 'Editar' : 'Corregir (solo a la baja si ya se ejecutó)'}>
                              <Edit3 className="w-4 h-4 text-sky-600" />
                            </button>
                          )}
                          {puedeAnular(r) && (
                            <button onClick={() => setAnularTarget(r)}
                              className="p-1 hover:bg-amber-50 rounded transition" title="Anular con motivo">
                              <Ban className="w-4 h-4 text-amber-600" />
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        <p className="text-[11px] text-slate-400 mt-4 leading-relaxed">
          Registro de control interno del sector. Las credenciales de esta solapa son una
          barrera contra cargas accidentales y sirven para dejar asentado quién solicitó y
          quién aprobó — no son un control de acceso ni una firma electrónica.
        </p>
      </Card>
      </>
      )}

      {rechazarTarget && (
        <ExtrasMotivoDialog
          titulo="Rechazar solicitud"
          descripcion={`${rechazarTarget.tecnico_nombre} · ${formatDateShort(rechazarTarget.fecha)} · ${formatHoras(rechazarTarget.horas)}. El motivo es opcional.`}
          placeholder="Motivo del rechazo (opcional)"
          requerido={false}
          cta="Rechazar"
          ctaColor="bg-red-600 hover:bg-red-500"
          iconColor="bg-red-500"
          icon={X}
          onConfirm={doRechazar}
          onCancel={() => setRechazarTarget(null)}
        />
      )}

      {anularTarget && (
        <ExtrasMotivoDialog
          titulo="Anular solicitud"
          descripcion={`${anularTarget.tecnico_nombre} · ${formatDateShort(anularTarget.fecha)} · ${formatHoras(anularTarget.horas)}. Queda registrada como anulada, con el motivo. No se borra.`}
          placeholder="Motivo de la anulación"
          requerido={true}
          cta="Anular"
          ctaColor="bg-amber-600 hover:bg-amber-500"
          iconColor="bg-amber-500"
          icon={Ban}
          onConfirm={doAnular}
          onCancel={() => setAnularTarget(null)}
        />
      )}
    </div>
  );
}

function AdminLoginDialog({ onConfirm, onCancel }) {
  const [pwd, setPwd] = useState('');
  const [error, setError] = useState('');
  const inputRef = useRef(null);

  // Cerrar con Escape, autofocus al abrir
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onCancel(); };
    document.addEventListener('keydown', onKey);
    inputRef.current?.focus();
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel]);

  const submit = () => {
    const ok = onConfirm(pwd);
    if (!ok) setError('Password incorrecto');
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-2xl max-w-sm w-full p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-slate-800 flex items-center justify-center flex-shrink-0">
            <Shield className="w-5 h-5 text-white" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">
              Modo administrador
            </h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              Ingresá el password para activar la edición y eliminación avanzada.
            </p>
          </div>
        </div>
        <input ref={inputRef} type="password" value={pwd}
          onChange={e => { setPwd(e.target.value); setError(''); }}
          onKeyDown={e => { if (e.key === 'Enter') submit(); }}
          placeholder="Password"
          className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500/40 focus:border-sky-500 transition" />
        {error && <div className="mt-2 text-xs text-red-600">{error}</div>}
        <div className="flex items-center justify-end gap-2 mt-5">
          <button onClick={onCancel}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition">
            Cancelar
          </button>
          <button onClick={submit}
            className="px-4 py-2 text-sm font-medium text-white bg-slate-800 hover:bg-slate-700 rounded-lg transition inline-flex items-center gap-1.5">
            <Lock className="w-4 h-4" />
            Ingresar
          </button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// V2.6 — MODAL DE CONFIRMACIÓN DE ELIMINACIÓN DE REPORTE
// Pide confirmación antes de eliminar un reporte completo de Supabase.
// La acción es irreversible.
// ═══════════════════════════════════════════════════════════════════
function DeleteReportConfirmDialog({ date, shift, onConfirm, onCancel }) {
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onCancel(); };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel]);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-2xl max-w-md w-full p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-red-100 flex items-center justify-center flex-shrink-0">
            <Trash2 className="w-5 h-5 text-red-600" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">
              Eliminar reporte completo
            </h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              Vas a eliminar el reporte del <strong>{date}</strong> turno <strong>{shift}</strong>.
            </p>
            <p className="text-sm text-red-700 leading-relaxed mt-2 font-medium">
              Esta acción es irreversible. Se borra toda la información del reporte:
              correctivos, preventivos, servicios, comentarios y avance.
            </p>
          </div>
        </div>
        <div className="flex items-center justify-end gap-2 mt-5">
          <button onClick={onCancel}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition">
            Cancelar
          </button>
          <button onClick={onConfirm}
            className="px-4 py-2 text-sm font-medium text-white bg-red-600 hover:bg-red-700 rounded-lg transition inline-flex items-center gap-1.5">
            <Trash2 className="w-4 h-4" />
            Eliminar reporte
          </button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// V3.3 — MODAL GUARD DE SOBREESCRITURA (BACKLOG #20)
// Se muestra antes de guardar si el destino (date+shift) ya tiene un
// reporte con datos en Supabase y es distinto del que se tenía abierto.
// Previene pisadas accidentales como el incidente del 2026-05-21.
// ═══════════════════════════════════════════════════════════════════
function OverwriteConfirmDialog({ date, shift, existingN, onConfirm, onCancel }) {
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onCancel(); };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel]);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-2xl max-w-md w-full p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-amber-100 flex items-center justify-center flex-shrink-0">
            <AlertTriangle className="w-5 h-5 text-amber-600" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">
              Reporte ya existente
            </h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              Ya hay un reporte guardado para <strong>{date}</strong> turno <strong>{shift}</strong>
              {existingN > 0 && <> con <strong>{existingN} OT{existingN === 1 ? '' : 's'} correctiva{existingN === 1 ? '' : 's'}</strong></>}.
            </p>
            <p className="text-sm text-amber-700 leading-relaxed mt-2 font-medium">
              Si continuás, el contenido actual del formulario va a reemplazar ese reporte. Esta acción no se puede deshacer.
            </p>
          </div>
        </div>
        <div className="flex items-center justify-end gap-2 mt-5">
          <button onClick={onCancel}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition">
            Cancelar
          </button>
          <button onClick={onConfirm}
            className="px-4 py-2 text-sm font-medium text-white bg-amber-600 hover:bg-amber-700 rounded-lg transition inline-flex items-center gap-1.5">
            <Save className="w-4 h-4" />
            Sobreescribir
          </button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// #7 (v3.6) — MODAL DE RECUPERACIÓN DE BORRADOR LOCAL
// Se muestra al posicionarse en un turno para el que hay un borrador sin
// guardar en localStorage (refresh/caída antes de guardar). Default visual =
// recuperar (acción esperada). Si el turno ya tiene un reporte en el servidor,
// se informa para que el usuario decida con contexto.
// ═══════════════════════════════════════════════════════════════════
function DraftRecoveryDialog({ date, shift, savedAt, serverUpdatedAt, onRecover, onDiscard }) {
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onDiscard(); };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onDiscard]);

  const fmt = (iso) => {
    if (!iso) return null;
    try {
      const d = new Date(iso);
      return d.toLocaleString('es-AR', {
        day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit'
      });
    } catch { return null; }
  };
  const savedStr = fmt(savedAt);
  const serverStr = fmt(serverUpdatedAt);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onDiscard}>
      <div className="bg-white rounded-xl shadow-2xl max-w-md w-full p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-sky-100 flex items-center justify-center flex-shrink-0">
            <RotateCcw className="w-5 h-5 text-sky-600" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">
              Borrador sin guardar
            </h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              Hay un borrador sin guardar del turno <strong>{shift}</strong> del{' '}
              <strong>{formatDateShort(date)}</strong>
              {savedStr ? <> (guardado automáticamente el <strong>{savedStr}</strong>)</> : null}.
              {' '}Probablemente quedó de una carga anterior que no llegaste a guardar.
            </p>
            {serverStr && (
              <p className="text-sm text-amber-700 leading-relaxed mt-2">
                Ojo: ya hay un reporte guardado en el servidor para este turno
                (última actualización <strong>{serverStr}</strong>). Si recuperás el
                borrador y guardás, vas a pisar esa versión. Revisá antes de guardar.
              </p>
            )}
          </div>
        </div>
        <div className="flex items-center justify-end gap-2 mt-5">
          <button onClick={onDiscard}
            className="px-4 py-2 text-sm font-medium text-slate-600 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition">
            Descartar borrador
          </button>
          <button onClick={onRecover}
            className="px-4 py-2 text-sm font-medium text-white bg-sky-600 hover:bg-sky-700 rounded-lg transition inline-flex items-center gap-1.5">
            <RotateCcw className="w-4 h-4" />
            Recuperar borrador
          </button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// V3.5 (#22) — MODAL DE CONCURRENCIA (optimistic locking)
// Se muestra cuando otra sesión guardó el mismo reporte mientras lo editabas.
// Default visual = acción segura (Descartar y recargar). Sobreescribir queda
// como acción secundaria y consciente.
// ═══════════════════════════════════════════════════════════════════
function ConcurrencyConflictDialog({ date, shift, onDiscard, onOverwrite }) {
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onDiscard(); };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onDiscard]);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onDiscard}>
      <div className="bg-white rounded-xl shadow-2xl max-w-md w-full p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-amber-100 flex items-center justify-center flex-shrink-0">
            <AlertTriangle className="w-5 h-5 text-amber-600" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">
              Modificado por otra sesión
            </h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              El reporte de <strong>{date}</strong> turno <strong>{shift}</strong> fue
              guardado por otra sesión mientras lo editabas. Si sobreescribís, vas a
              pisar esos cambios.
            </p>
            <p className="text-sm text-amber-700 leading-relaxed mt-2 font-medium">
              Lo recomendado es descartar lo tuyo y recargar la versión más reciente para no perder datos.
            </p>
          </div>
        </div>
        <div className="flex items-center justify-end gap-2 mt-5">
          <button onClick={onOverwrite}
            className="px-4 py-2 text-sm font-medium text-amber-700 bg-white border border-amber-300 hover:bg-amber-50 rounded-lg transition inline-flex items-center gap-1.5">
            <Save className="w-4 h-4" />
            Sobreescribir igual
          </button>
          <button onClick={onDiscard}
            className="px-4 py-2 text-sm font-medium text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg transition">
            Descartar y recargar
          </button>
        </div>
      </div>
    </div>
  );
}
// Se muestra antes de guardar si hay correctivos o preventivos completamente
// vacíos (sin OT/equipo, sin task, sin técnicos). El usuario debe confirmar
// que está OK eliminarlos antes de continuar con el guardado.
// ═══════════════════════════════════════════════════════════════════
function EmptyEntriesConfirmDialog({ emptyCorr, emptyPrev, onConfirm, onCancel }) {
  // Cerrar con Escape
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onCancel(); };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel]);

  const parts = [];
  if (emptyCorr > 0) parts.push(`${emptyCorr} OT correctiva${emptyCorr === 1 ? '' : 's'}`);
  if (emptyPrev > 0) parts.push(`${emptyPrev} preventivo${emptyPrev === 1 ? '' : 's'}`);
  const lista = parts.join(' y ');

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-2xl max-w-md w-full p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-amber-100 flex items-center justify-center flex-shrink-0">
            <AlertTriangle className="w-5 h-5 text-amber-600" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">
              Entradas vacías detectadas
            </h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              Se detectaron <strong>{lista}</strong> sin información cargada
              (sin N° OT/equipo, sin descripción y sin técnicos).
            </p>
            <p className="text-sm text-slate-600 leading-relaxed mt-2">
              Al confirmar, esas entradas se <strong>eliminarán</strong> y se guardará el reporte
              con el resto de la información.
            </p>
          </div>
        </div>
        <div className="flex items-center justify-end gap-2 mt-5">
          <button onClick={onCancel}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition">
            Cancelar
          </button>
          <button onClick={onConfirm}
            className="px-4 py-2 text-sm font-medium text-white bg-slate-800 hover:bg-slate-700 rounded-lg transition inline-flex items-center gap-1.5">
            <Save className="w-4 h-4" />
            Confirmar y guardar
          </button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// V2.8 — MODAL DE CONFLICTOS DE OT CERRADA (carry-over stale guard)
//
// Se abre antes de guardar un reporte si se detecta que contiene OTs cuyo
// estado en el form (Sin Iniciar / En Curso) contradice el estado Realizada
// que ya tienen en otro reporte del history (escenario clásico: el turno
// dejó el form abierto, otro turno cerró la OT, y al guardar el primero
// estaría pisando esa decisión).
//
// Por cada conflicto el usuario elige UNA acción:
//   - "Quitarla del reporte" (recomendado): la OT se elimina del corrective[]
//     de este reporte antes de guardar. El cierre del otro turno queda intacto.
//   - "Reabrir explícitamente" (SOLO ADMIN): la OT se mantiene tal como está
//     en el form (estado En Curso / Sin Iniciar), pero se agrega una entrada
//     al timeline documentando el motivo. La OT queda efectivamente reabierta.
//
// Hasta que NO se decida sobre TODOS los conflictos, el botón Guardar está
// deshabilitado. Cancelar cierra el modal sin guardar (el usuario puede seguir
// editando o decidir más tarde).
// ═══════════════════════════════════════════════════════════════════
function ClosedConflictDialog({ conflicts, adminMode, onResolve, onCancel }) {
  // decisions[i] = { action: 'remove' | 'reopen', reopenReason: string } | null
  const [decisions, setDecisions] = useState(() => conflicts.map(() => null));

  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onCancel(); };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel]);

  const setDecision = (i, partial) => {
    setDecisions(prev => {
      const next = [...prev];
      next[i] = { ...(next[i] || {}), ...partial };
      return next;
    });
  };

  // Validación: todos los conflictos deben tener decisión, y los "reopen" deben tener motivo
  const allResolved = decisions.every((d, i) => {
    if (!d || !d.action) return false;
    if (d.action === 'reopen' && (!d.reopenReason || d.reopenReason.trim().length === 0)) return false;
    return true;
  });

  const submit = () => {
    if (!allResolved) return;
    const payload = conflicts.map((c, i) => ({
      otNumber: c.otNumber,
      action: decisions[i].action,
      reopenReason: decisions[i].reopenReason || ''
    }));
    onResolve(payload);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-2xl max-w-2xl w-full max-h-[90vh] overflow-y-auto p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-amber-100 flex items-center justify-center flex-shrink-0">
            <AlertTriangle className="w-5 h-5 text-amber-600" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">
              {conflicts.length === 1 ? 'Conflicto detectado' : `${conflicts.length} conflictos detectados`}
            </h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              {conflicts.length === 1
                ? 'Una OT de este reporte ya fue cerrada por otro turno. Decidí qué hacer antes de guardar.'
                : 'Estas OTs ya fueron cerradas por otros turnos. Decidí qué hacer con cada una antes de guardar.'}
            </p>
          </div>
        </div>

        <div className="space-y-3 mb-4">
          {conflicts.map((c, i) => {
            const d = decisions[i] || {};
            return (
              <div key={c.otNumber + i} className="border border-slate-200 rounded-lg p-3 bg-slate-50">
                <div className="text-sm font-semibold text-slate-800 mb-1">
                  <span className="num">{c.otNumber}</span>
                  <span className="text-slate-500 font-normal ml-2 text-xs">({c.otTask})</span>
                </div>
                <div className="text-xs text-slate-600 mb-2 leading-relaxed">
                  En tu reporte: <strong className="text-orange-700">{c.formState}</strong>
                  {' · '}
                  Ya cerrada por: <strong>{c.closedIn.responsable}</strong>
                  {' en '}
                  <strong className="num">{c.closedIn.date}</strong> {c.closedIn.shift}
                </div>
                <div className="flex flex-col gap-1.5">
                  <label className="inline-flex items-start gap-2 text-xs cursor-pointer">
                    <input
                      type="radio"
                      name={`conflict-${i}`}
                      checked={d.action === 'remove'}
                      onChange={() => setDecision(i, { action: 'remove' })}
                      className="mt-0.5"
                    />
                    <span className="text-slate-700">
                      <strong>Quitarla de este reporte</strong> (recomendado).
                      La OT se elimina de tu reporte. El cierre del otro turno queda intacto.
                    </span>
                  </label>
                  <label className={`inline-flex items-start gap-2 text-xs ${adminMode ? 'cursor-pointer' : 'cursor-not-allowed opacity-50'}`}>
                    <input
                      type="radio"
                      name={`conflict-${i}`}
                      checked={d.action === 'reopen'}
                      onChange={() => adminMode && setDecision(i, { action: 'reopen' })}
                      disabled={!adminMode}
                      className="mt-0.5"
                    />
                    <span className="text-slate-700">
                      <strong>Reabrir la OT</strong> (solo admin).
                      La OT se mantiene en estado {c.formState} y queda registrado en el timeline.
                      {!adminMode && <span className="text-slate-500 italic"> Logueate como admin para usar esta opción.</span>}
                    </span>
                  </label>
                  {d.action === 'reopen' && (
                    <textarea
                      placeholder="Motivo de reapertura (obligatorio)…"
                      value={d.reopenReason || ''}
                      onChange={(e) => setDecision(i, { reopenReason: e.target.value })}
                      rows={2}
                      className="mt-1 w-full text-xs px-2 py-1.5 border border-slate-300 rounded resize-none focus:outline-none focus:ring-2 focus:ring-amber-400"
                    />
                  )}
                </div>
              </div>
            );
          })}
        </div>

        <div className="flex items-center justify-end gap-2">
          <button onClick={onCancel}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition">
            Cancelar
          </button>
          <button
            onClick={submit}
            disabled={!allResolved}
            className={`px-4 py-2 text-sm font-medium text-white rounded-lg transition inline-flex items-center gap-1.5
              ${allResolved ? 'bg-amber-600 hover:bg-amber-700' : 'bg-slate-300 cursor-not-allowed'}`}>
            Aplicar y guardar
          </button>
        </div>
      </div>
    </div>
  );
}
// ═══════════════════════════════════════════════════════════════════
// V2.9 — MODAL DE PROPAGACIÓN ADMIN
//
// Se abre cuando admin edita un reporte histórico, guarda, y se detectan
// cambios propagables (diffs por OT) Y existen reportes posteriores que
// contienen las mismas OTs.
//
// UI HÍBRIDA:
//   - Sección "Cambios uniformes": resumen compacto de entradas de timeline
//     agregadas/borradas. Aplica a TODOS los reportes posteriores afectados
//     que tengan la OT correspondiente. No requiere confirmación por reporte.
//   - Sección "Cambios de estado": una entrada por (OT × reporte posterior).
//     Cada uno es un radio button: "propagar" o "no propagar". Admin decide
//     caso por caso porque el reporte posterior puede tener un estado distinto
//     ya cargado manualmente que el admin no quiere pisar.
//
// El resultado del modal es un payload que se pasa al callback onResolve
// y se ejecuta vía la RPC propagate_admin_changes en Supabase.
//
// Cancelar: cierra el modal SIN guardar el reporte editado. Admin puede
// seguir editando o decidir más tarde.
// ═══════════════════════════════════════════════════════════════════
function PropagationModal({ diffs, affectedReports, onConfirm, onCancel }) {
  // Decisiones de propagación de state, indexadas por "reportId:ot".
  // Cada valor: 'propagate' | 'skip'. Default 'propagate' para mantener consistencia.
  const [stateDecisions, setStateDecisions] = useState(() => {
    const initial = {};
    affectedReports.forEach(({ report, affectedOts }) => {
      affectedOts.forEach(({ ot, diff }) => {
        if (diff.stateChange) {
          initial[`${report.date}|${report.shift}|${ot}`] = 'propagate';
        }
      });
    });
    return initial;
  });

  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onCancel(); };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel]);

  // Stats agregadas de cambios de timeline (uniformes)
  const totalAdded = diffs.reduce((s, d) => s + d.addedEntries.length, 0);
  const totalDeleted = diffs.reduce((s, d) => s + d.deletedEntries.length, 0);

  // Cambios de state que tienen al menos un reporte posterior con la OT
  const stateChanges = [];
  diffs.forEach(d => {
    if (!d.stateChange) return;
    affectedReports.forEach(({ report, affectedOts }) => {
      const ao = affectedOts.find(a => a.ot === d.ot);
      if (ao) stateChanges.push({ diff: d, report, currentState: ao.currentState });
    });
  });

  const setStateDecision = (key, value) => {
    setStateDecisions(prev => ({ ...prev, [key]: value }));
  };

  const submit = () => {
    onConfirm({ stateDecisions });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
         onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-2xl max-w-3xl w-full max-h-[90vh] overflow-y-auto p-5"
           onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-sky-100 flex items-center justify-center flex-shrink-0">
            <Shield className="w-5 h-5 text-sky-600" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-base font-semibold text-slate-900 mb-1">
              Propagar cambios a reportes posteriores
            </h3>
            <p className="text-sm text-slate-600 leading-relaxed">
              Detectamos cambios en {diffs.length} OT{diffs.length === 1 ? '' : 's'} que afectan a {affectedReports.length} reporte{affectedReports.length === 1 ? '' : 's'} posterior{affectedReports.length === 1 ? '' : 'es'}.
              Revisá qué se va a propagar antes de confirmar.
            </p>
          </div>
        </div>

        {/* SECCIÓN 1 — Cambios uniformes de timeline */}
        {(totalAdded > 0 || totalDeleted > 0) && (
          <div className="mb-5 border border-slate-200 rounded-lg p-3 bg-slate-50">
            <div className="text-xs font-semibold text-slate-700 uppercase tracking-wide mb-2 inline-flex items-center gap-1.5">
              <Activity className="w-3.5 h-3.5 text-amber-500" />
              Cambios uniformes (timeline)
            </div>
            <p className="text-[12px] text-slate-600 leading-snug mb-3">
              Estos cambios se aplican automáticamente a <strong>todos los reportes posteriores</strong> que contengan la OT correspondiente.
            </p>
            <div className="space-y-2">
              {diffs.filter(d => d.addedEntries.length > 0 || d.deletedEntries.length > 0).map(d => (
                <div key={d.ot} className="bg-white border border-slate-200 rounded p-2">
                  <div className="text-[12px] font-semibold text-slate-800 mb-1">
                    <span className="num">{d.ot}</span>
                    <span className="text-slate-500 font-normal ml-2 text-[11px]">({d.otTask})</span>
                  </div>
                  {d.addedEntries.map(e => (
                    <div key={e.id} className="text-[11px] text-emerald-700 bg-emerald-50/60 border-l-2 border-emerald-300 pl-2 py-0.5 mb-1">
                      <span className="font-semibold mr-1">+ Agregar:</span>
                      <span className="text-slate-700">"{e.text}"</span>
                      <span className="text-slate-400 ml-1 text-[10px]">— {e.date} {e.shift}, {e.author}</span>
                    </div>
                  ))}
                  {d.deletedEntries.map(e => (
                    <div key={e.id} className="text-[11px] text-red-700 bg-red-50/60 border-l-2 border-red-300 pl-2 py-0.5 mb-1">
                      <span className="font-semibold mr-1">− Borrar:</span>
                      <span className="text-slate-700 line-through">"{e.text}"</span>
                      <span className="text-slate-400 ml-1 text-[10px]">— {e.date} {e.shift}, {e.author}</span>
                    </div>
                  ))}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* SECCIÓN 2 — Cambios de estado (caso por caso) */}
        {stateChanges.length > 0 && (
          <div className="mb-5 border border-slate-200 rounded-lg p-3 bg-slate-50">
            <div className="text-xs font-semibold text-slate-700 uppercase tracking-wide mb-2 inline-flex items-center gap-1.5">
              <AlertTriangle className="w-3.5 h-3.5 text-amber-500" />
              Cambios de estado (decisión caso por caso)
            </div>
            <p className="text-[12px] text-slate-600 leading-snug mb-3">
              El reporte posterior puede tener un estado distinto cargado manualmente. Decidí si propagar el cambio en cada caso.
            </p>
            <div className="space-y-2">
              {stateChanges.map(({ diff, report, currentState }, idx) => {
                const key = `${report.date}|${report.shift}|${diff.ot}`;
                const decision = stateDecisions[key] || 'propagate';
                return (
                  <div key={idx} className="bg-white border border-slate-200 rounded p-2">
                    <div className="text-[12px] font-semibold text-slate-800 mb-1">
                      <span className="num">{diff.ot}</span>
                      <span className="text-slate-500 font-normal ml-2 text-[11px]">({diff.otTask})</span>
                    </div>
                    <div className="text-[11px] text-slate-600 mb-2">
                      En reporte <strong className="num">{report.date} {report.shift}</strong>
                      {' · '}
                      Estado actual ahí: <StatePill state={currentState} />
                    </div>
                    <div className="text-[11px] text-slate-700 mb-2 leading-snug">
                      Cambio detectado en el reporte editado: <StatePill state={diff.stateChange.from} /> → <StatePill state={diff.stateChange.to} />
                    </div>
                    <div className="flex flex-col gap-1">
                      <label className="inline-flex items-start gap-2 text-[11px] cursor-pointer">
                        <input
                          type="radio"
                          name={`state-${idx}`}
                          checked={decision === 'propagate'}
                          onChange={() => setStateDecision(key, 'propagate')}
                          className="mt-0.5"
                        />
                        <span className="text-slate-700">
                          <strong>Propagar</strong>: cambiar a <StatePill state={diff.stateChange.to} /> también en este reporte posterior.
                        </span>
                      </label>
                      <label className="inline-flex items-start gap-2 text-[11px] cursor-pointer">
                        <input
                          type="radio"
                          name={`state-${idx}`}
                          checked={decision === 'skip'}
                          onChange={() => setStateDecision(key, 'skip')}
                          className="mt-0.5"
                        />
                        <span className="text-slate-700">
                          <strong>No propagar</strong>: mantener <StatePill state={currentState} /> tal como está acá.
                        </span>
                      </label>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* Resumen previo */}
        <div className="mb-4 px-3 py-2 bg-sky-50 border border-sky-200 rounded text-[12px] text-sky-800">
          <strong>Resumen:</strong> {totalAdded > 0 && <>se agregan {totalAdded} entrada{totalAdded === 1 ? '' : 's'} de timeline · </>}
          {totalDeleted > 0 && <>se borran {totalDeleted} entrada{totalDeleted === 1 ? '' : 's'} de timeline · </>}
          {stateChanges.length > 0 && <>{Object.values(stateDecisions).filter(v => v === 'propagate').length} cambio{Object.values(stateDecisions).filter(v => v === 'propagate').length === 1 ? '' : 's'} de estado a propagar · </>}
          afecta {affectedReports.length} reporte{affectedReports.length === 1 ? '' : 's'} posterior{affectedReports.length === 1 ? '' : 'es'}.
        </div>

        <div className="flex items-center justify-end gap-2">
          <button onClick={onCancel}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition">
            Cancelar
          </button>
          <button onClick={submit}
            className="px-4 py-2 text-sm font-medium text-white bg-sky-600 hover:bg-sky-700 rounded-lg transition inline-flex items-center gap-1.5">
            <Save className="w-4 h-4" />
            Guardar y propagar
          </button>
        </div>
      </div>
    </div>
  );
}


// ═══════════════════════════════════════════════════════════════════
// FORM VIEW — V2.0
//   - sin botón eliminar en correctivos
//   - limpiar conserva pendientes (Sin Iniciar + En Curso) y borra Realizadas
//   - Planta de Efluentes y Caldera con schema nuevo (PTEL + Caldera + Ablandadores)
//   - Resumen Preventivos del Turno al final del formulario
// ═══════════════════════════════════════════════════════════════════
function FormView({ report, setReport, onSave, saveMsg, setSaveMsg, saving, history, adminMode, originalReport, setOriginalReport, onDeleteReport, otErrorIndices = new Set(), otErrorType = '', onClearOtErrors }) {
  // #9 (v3.14) — read-only para no-admin si el reporte es anterior a ayer (corrección
  // retroactiva del pasado = solo admin). Hoy/ayer editable (turno en curso, cruce de
  // medianoche del Noche cubierto). Se aplica como fieldset disabled (deshabilita todos
  // los controles hijos de una sola vez), banner de aviso, y botón Guardar oculto.
  // El guard duro de verdad está en doSaveReport; esto es la capa de UI.
  const isReadOnly = !adminMode && !isWithinEditWindow(report.date);
  const update = (patch) => setReport(r => ({ ...r, ...patch }));
  const updateList = (key, fn) => setReport(r => ({ ...r, [key]: fn(r[key]) }));
  const updateServicios = (patch) => setReport(r => ({ ...r, servicios: { ...r.servicios, ...patch } }));
  const updateResumen = (patch) => setReport(r => ({ ...r, preventivosResumen: { ...r.preventivosResumen, ...patch } }));

  // V2.3 — helper que actualiza una OT correctiva por índice y, si efectivamente
  // hubo cambios, setea lastModifiedInShift al turno actual. Esto permite que el
  // Dashboard distinga OTs "tocadas" en el turno actual (visibles) de las del
  // carry-over que nadie modificó (no visibles).
  // V2.6 — En modo admin, las modificaciones NO marcan lastModifiedInShift
  // (porque pueden ser correcciones retroactivas, no trabajo del turno).
  const updateCorrectiveItem = (i, patch) => {
    // v3.16 — Al editar cualquier OT, limpiar las marcas de error de validación
    if (otErrorIndices.size > 0 && onClearOtErrors) onClearOtErrors();
    setReport(r => ({
      ...r,
      corrective: r.corrective.map((x, j) => {
        if (j !== i) return x;
        // si la patch no cambia nada efectivo, no actualizamos lastModifiedInShift
        const changed = Object.keys(patch).some(k => x[k] !== patch[k]);
        if (!changed) return x;
        if (adminMode) {
          // V2.6 — admin: aplicar patch sin tocar lastModifiedInShift
          return { ...x, ...patch };
        }
        return { ...x, ...patch, lastModifiedInShift: `${r.date}-${r.shift}` };
      })
    }));
  };

  // V2.4 — Timeline de Estado de Avance
  // Cada OT correctiva tiene un array timeline:
  //   [{ shiftKey, date, shift, author, text, timestamp }]
  // - Cargado solo cuando el estado es "En Curso" (y al guardar la app valida
  //   que haya entrada del turno actual)
  // - Las entradas anteriores son read-only (no se pueden editar ni borrar)
  // - V2.6 — En modo admin, las entradas son editables y eliminables
  // - timelineDraft mantiene el texto en redacción por índice de OT
  const [timelineDraft, setTimelineDraft] = useState({});
  // #avance (v3.9) — Toggle "¿hubo avance este turno?" por OT (índice).
  // Solo aplica a OTs heredadas que NO requieren avance obligatorio (En Curso → En Curso).
  // true = sí hubo avance (habilita el campo de carga); ausente/false = no (default).
  // Si la OT requiere avance obligatorio (cambio de estado / cierre), el toggle no
  // aplica: el campo se muestra y se exige igual que antes.
  const [avanceToggle, setAvanceToggle] = useState({});
  // PR-2 (v3.20) — breakpoint desktop (lg=1024px) para el layout emparejado.
  const isDesktop = useMediaQuery('(min-width: 1024px)');
  // V2.6 — Edición inline de entradas existentes (solo admin)
  // editingKey es del tipo "i-ei" (índice de OT - índice de entrada)
  const [timelineEditingKey, setTimelineEditingKey] = useState(null);
  const [timelineEditDraft, setTimelineEditDraft] = useState('');

  const addTimelineEntry = (i) => {
    const text = (timelineDraft[i] || '').trim();
    if (!text) return;
    setReport(r => ({
      ...r,
      corrective: r.corrective.map((x, j) => {
        if (j !== i) return x;
        const newEntry = {
          id: generateTimelineId(),                // V2.9 — id único
          shiftKey: `${r.date}-${r.shift}`,
          date: r.date,
          shift: r.shift,
          author: r.responsable || '',
          text,
          timestamp: new Date().toISOString()
        };
        return {
          ...x,
          timeline: [...(x.timeline || []), newEntry],
          lastModifiedInShift: `${r.date}-${r.shift}`
        };
      })
    }));
    setTimelineDraft(d => { const nd = { ...d }; delete nd[i]; return nd; });
  };

  // #avance (v3.9) — Al cambiar de reporte (fecha/turno), resetear el toggle de
  // avance y los borradores en redacción, para que no queden "pegados" por índice
  // de OT entre un reporte y otro (los índices se reusan).
  useEffect(() => {
    setAvanceToggle({});
    setTimelineDraft({});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [report.date, report.shift]);

  const [loadInfo, setLoadInfo] = useState('');
  const initialPendingApplied = useRef(false);

  // ── BACKLOG #7 (v3.6) — Autoguardado local de borrador ──────────────
  // State del modal de recuperación: null | { id, draft, serverUpdatedAt }
  const [draftRecovery, setDraftRecovery] = useState(null);

  // Autoguardado debounced (~1.5s) del borrador en curso a localStorage.
  // Solo guarda si:
  //   - estamos en carga de turno corriente (NO admin editando histórico),
  //   - el form tiene trabajo real del usuario (hasUserWork), no solo carry-over, y
  //   - el form difiere del último estado GUARDADO (originalReport). Si coinciden,
  //     no hay nada sin guardar que respaldar (ya está en Supabase) — esto evita
  //     que el autosave recree el borrador inmediatamente después de un save OK,
  //     que ya lo limpió (la carrera clear-vs-debounce post-guardado).
  // El borrador se limpia tras un save OK (en doSaveReport), no acá.
  useEffect(() => {
    // Excluir edición admin de histórico: el borrador es para carga de turno,
    // no para ediciones retroactivas (que tienen su propio snapshot V2.9).
    if (adminMode && originalReport) return;
    // No autoguardar mientras el modal de recuperación está abierto: el usuario
    // todavía no decidió si recupera o descarta.
    if (draftRecovery) return;
    if (!report.date || !report.shift) return;
    if (!hasUserWork(report)) return;
    // Si el form coincide con el último reporte guardado, no hay borrador pendiente.
    if (originalReport && reportsEqual(report, originalReport)) return;
    const id = `${report.date}-${report.shift}`;
    const t = setTimeout(() => draftStore.save(id, report), 1500);
    return () => clearTimeout(t);
  }, [report, adminMode, originalReport, draftRecovery]);

  // Intenta ofrecer recuperación de borrador para un (date, shift).
  // Devuelve true si abrió el modal de recuperación (el caller corta su flujo),
  // false si no había borrador recuperable (el caller sigue su flujo normal).
  const maybeOfferDraft = (date, shift) => {
    if (adminMode && originalReport) return false;   // no en edición admin
    const id = `${date}-${shift}`;
    const draft = draftStore.load(id);
    if (!draft || !hasUserWork(draft.report)) return false;
    const serverMatch = history.find(r => r.date === date && r.shift === shift);
    setDraftRecovery({
      id,
      draft,
      serverUpdatedAt: serverMatch ? (serverMatch._updatedAt || null) : null
    });
    return true;
  };

  // Aplica el borrador recuperado al form.
  const applyDraftRecovery = () => {
    if (!draftRecovery) return;
    const r = hydrate(draftRecovery.draft.report);
    setReport(r);
    // Si el turno ya existía en el servidor, el snapshot original es el del
    // servidor (para que el guard de sobreescritura / concurrencia siga válido);
    // si no existía, es reporte nuevo (sin snapshot).
    const serverMatch = history.find(x => x.date === r.date && x.shift === r.shift);
    setOriginalReport(serverMatch ? JSON.parse(JSON.stringify(hydrate(serverMatch))) : null);
    setDraftRecovery(null);
    setLoadInfo('↻ Borrador recuperado. Revisá los datos y guardá para confirmar.');
    setTimeout(() => setLoadInfo(''), 5000);
  };

  // Descarta el borrador y sigue el flujo normal (server o nuevo con pendientes).
  const discardDraftRecovery = () => {
    if (!draftRecovery) return;
    const { id } = draftRecovery;
    const [date, shift] = [draftRecovery.draft.report.date, draftRecovery.draft.report.shift];
    draftStore.clear(id);
    setDraftRecovery(null);
    // Reproducir el flujo normal de setDateShift sin volver a ofrecer el draft
    const existing = history.find(r => r.date === date && r.shift === shift);
    if (existing) {
      setReport(hydrate(existing));
      setOriginalReport(JSON.parse(JSON.stringify(hydrate(existing))));
    } else {
      const pending = computePending(date, shift);
      setReport({ ...emptyReport(), date, shift, corrective: pending });
      setOriginalReport(null);
    }
  };
  // ────────────────────────────────────────────────────────────────────

  // V2.4 — Compute correctivos pendientes:
  //   - Recorre todos los reportes hasta (date, shift)
  //   - Mantiene la versión MÁS RECIENTE de cada OT (por número)
  //   - Filtra solo las que quedan en "Sin Iniciar" o "En Curso"
  //   - ORDEN: más recientemente vistas primero (la última que apareció arriba).
  //     Esto deja arriba lo "fresco" del turno anterior y abajo OTs que llevan
  //     varios días sin movimiento.
  //   - Las del carry-over NO tienen createdInShift = turno actual (se preservan
  //     sus flags originales para que el Dashboard las distinga).
  const computePending = (date, shift) => {
    const upTo = history
      .filter(r => (r.date < date) || (r.date === date && shiftOrder(r.shift) <= shiftOrder(shift)))
      .sort((a, b) => (a.date + shiftOrder(a.shift)).localeCompare(b.date + shiftOrder(b.shift)));
    const latestByOT = new Map();   // ot# -> { ot, lastSeenSortKey }
    upTo.forEach(r => (r.corrective || []).forEach(c => {
      if (c.ot) {
        latestByOT.set(c.ot, {
          ot: c,
          lastSeenSortKey: r.date + shiftOrder(r.shift)
        });
      }
    }));
    return [...latestByOT.values()]
      .filter(({ ot }) => ot.state === 'Sin Iniciar' || ot.state === 'En Curso')
      .sort((a, b) => b.lastSeenSortKey.localeCompare(a.lastSeenSortKey))   // descendente
      .map(({ ot }) => ({ ...ot, timeline: ot.timeline || [] }));
  };

  // When date or shift changes, look up history:
  //  - if a saved report exists for that date+shift -> load it as-is
  //  - otherwise, build a new empty report with pending correctivos pre-loaded
  const setDateShift = (newDate, newShift) => {
      // #7 v3.6 — si hay borrador local recuperable para este turno, ofrecer
      // recuperarlo antes de cargar del servidor o armar uno nuevo.
      if (maybeOfferDraft(newDate, newShift)) return;
      const existing = history.find(r => r.date === newDate && r.shift === newShift);
      if (existing) {
      setReport(hydrate(existing));
      setOriginalReport(JSON.parse(JSON.stringify(hydrate(existing))));  // V2.9 — snapshot deep copy  
      setLoadInfo(`✓ Reporte cargado del histórico (${formatDateShort(newDate)} - ${newShift})`);
      setTimeout(() => setLoadInfo(''), 4000);
      return;
    }
    const pending = computePending(newDate, newShift);
    setReport({ ...emptyReport(), date: newDate, shift: newShift, corrective: pending });
    setOriginalReport(null);  // V2.9 — reporte nuevo: sin snapshot
    if (pending.length > 0) {
      setLoadInfo(`↻ ${pending.length} correctivo${pending.length === 1 ? '' : 's'} pendiente${pending.length === 1 ? '' : 's'} traído${pending.length === 1 ? '' : 's'} del turno anterior`);
      setTimeout(() => setLoadInfo(''), 5000);
    }
  };

  // V2.0 — "Limpiar": NO borra correctivos en "Sin Iniciar" / "En Curso".
  // Solo borra los "Realizada" y resetea el resto del formulario.
  // V2.1 — además limpia cualquier mensaje de error/guardado anterior.
  const cleanForm = () => {
    setReport(r => ({
      ...emptyReport(),
      date: r.date,
      shift: r.shift,
      // Conservar correctivos que NO estén en "Realizada"
      corrective: (r.corrective || []).filter(c => c.state !== 'Realizada')
    }));
    // V3.3 — Limpiar desvincula el form del reporte cargado. Así, si después se
    // guarda sobre un destino con datos, el guard de sobreescritura (#20) vuelve
    // a disparar. Antes quedaba apuntando al reporte viejo y el guard no protegía
    // el caso del incidente del 2026-05-21.
    setOriginalReport(null);
    const pendingCount = (report.corrective || []).filter(c => c.state !== 'Realizada').length;
    const removedCount = (report.corrective || []).filter(c => c.state === 'Realizada').length;
    if (pendingCount > 0 && removedCount > 0) {
      setLoadInfo(`↻ Formulario limpio. Se mantienen ${pendingCount} correctivo${pendingCount === 1 ? '' : 's'} pendiente${pendingCount === 1 ? '' : 's'}. Se quitaron ${removedCount} realizado${removedCount === 1 ? '' : 's'}.`);
    } else if (pendingCount > 0) {
      setLoadInfo(`↻ Formulario limpio. Se mantienen ${pendingCount} correctivo${pendingCount === 1 ? '' : 's'} pendiente${pendingCount === 1 ? '' : 's'}.`);
    } else if (removedCount > 0) {
      setLoadInfo(`✓ Formulario limpio. Se quitaron ${removedCount} correctivo${removedCount === 1 ? '' : 's'} realizado${removedCount === 1 ? '' : 's'}.`);
    } else {
      setLoadInfo('✓ Formulario limpio');
    }
    // V2.1 FIX: limpiar mensaje de error/guardado anterior
    if (typeof setSaveMsg === 'function') setSaveMsg('');
    setTimeout(() => setLoadInfo(''), 4500);
  };

  // On first history-loaded mount, auto-apply pending correctivos if the form
  // is still untouched.
  useEffect(() => {
    if (initialPendingApplied.current) return;
    if (history.length === 0) return;
    const isEmpty = !report.responsable && report.team.length === 0 &&
      report.corrective.length === 0 && report.preventive.length === 0 &&
      report.comments.length === 0;
    if (!isEmpty) { initialPendingApplied.current = true; return; }
    // #7 v3.6 — al cargar la app sobre el turno por defecto, ofrecer recuperar
    // borrador si existe para ese turno (antes de cargar server o pendientes).
    if (maybeOfferDraft(report.date, report.shift)) {
      initialPendingApplied.current = true;
      return;
    }
    const existing = history.find(r => r.date === report.date && r.shift === report.shift);
    if (existing) {
      setReport(hydrate(existing));
      setOriginalReport(JSON.parse(JSON.stringify(hydrate(existing))));  // V2.9 — snapshot deep copy
    } else {
      const pending = computePending(report.date, report.shift);
      if (pending.length > 0) {
        setReport(r => ({ ...r, corrective: pending }));
        setOriginalReport(null);  // V2.9 — reporte nuevo: sin snapshot
        setLoadInfo(`↻ ${pending.length} correctivo${pending.length === 1 ? '' : 's'} pendiente${pending.length === 1 ? '' : 's'} del turno anterior`);
        setTimeout(() => setLoadInfo(''), 5000);
      }
    }
    initialPendingApplied.current = true;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [history]);

  // V2.5 — Opciones para asignar técnicos a OTs y preventivos.
  // Reglas:
  //  - Encargados (RESPONSABLES): siempre al principio de la lista.
  //  - Técnicos: SOLO los que están en "Equipo del Turno" actual.
  //    Si no hay equipo cargado, no se muestran técnicos (solo encargados).
  //  - Excepción por OT: si una OT heredada ya tiene asignados técnicos
  //    de turnos previos que no están en el equipo actual, se preservan
  //    en la lista de esa OT puntual (sino se mostrarían los chips pero
  //    no se podrían deseleccionar). Por eso `teamOptionsForOT(c)` recibe
  //    la OT y arma la lista incluyendo los técnicos heredados ya asignados.
  //  - El detalle "por técnico" del Resumen de Preventivos NO usa esta
  //    lista (sigue usando `report.team` directo, sin encargados).
  const encargadosNames = useMemo(() => RESPONSABLES.map(r => r.name), []);
  const teamOptionsForOT = useCallback((ot) => {
    // Base: encargados primero, después técnicos del equipo del turno
    const list = [...encargadosNames];
    (report.team || []).forEach(name => {
      if (!list.includes(name)) list.push(name);
    });
    // Preservar técnicos ya asignados a ESTA OT que no estén en el equipo
    // (ej: técnicos heredados de turnos previos)
    if (ot && Array.isArray(ot.technicians)) {
      ot.technicians.forEach(name => {
        if (!list.includes(name)) list.push(name);
      });
    }
    return list;
  }, [encargadosNames, report.team]);

  // V2.5 — Mapa de "estado previo" de cada OT (por número), buscando en history
  // el reporte más reciente ANTERIOR al turno actual donde aparece la misma OT.
  // Sirve para detectar cambios de estado y decidir si el avance es obligatorio.
  // Si una OT no aparece en history previa, su entry es null (es nueva).
  const previousStateMap = useMemo(() => {
    const map = new Map();
    const previousReports = history
      .filter(rep => {
        if (rep.date < report.date) return true;
        if (rep.date === report.date && shiftOrder(rep.shift) < shiftOrder(report.shift)) return true;
        return false;
      })
      .sort((a, b) =>
        `${b.date}-${shiftOrder(b.shift)}`.localeCompare(`${a.date}-${shiftOrder(a.shift)}`)
      );
    (report.corrective || []).forEach(c => {
      if (!c.ot) return;
      if (map.has(c.ot)) return;
      for (const rep of previousReports) {
        const found = (rep.corrective || []).find(x => x.ot === c.ot);
        if (found) {
          map.set(c.ot, found.state || null);
          break;
        }
      }
      if (!map.has(c.ot)) map.set(c.ot, null); // marca como "nueva"
    });
    return map;
  }, [history, report.date, report.shift, report.corrective]);

  // V2.5 — Helper: decide si una OT requiere entrada de avance en este turno.
  // Regla: solo si HAY cambio de estado en el turno actual respecto a su estado previo.
  //   - Nueva (no existe antes) + Sin Iniciar → no requiere
  //   - Nueva + En Curso o Realizada → requiere
  //   - Sin Iniciar → En Curso o Realizada → requiere
  //   - En Curso → Realizada → requiere
  //   - Cualquier otro (sin cambio) → no requiere
  const requiresAdvanceEntry = (c) => {
    const prev = c.ot ? previousStateMap.get(c.ot) : null;
    const curr = c.state;
    if (prev === undefined || prev === null) {
      return curr === 'En Curso' || curr === 'Realizada';
    }
    if (prev === 'Sin Iniciar' && (curr === 'En Curso' || curr === 'Realizada')) return true;
    if (prev === 'En Curso' && curr === 'Realizada') return true;
    return false;
  };

  // V2.4 — Cuando el equipo cambia, depurar técnicos del resumen que ya no están en el turno.
  // Schema multi-técnico: filtrar dentro de cada grupo los técnicos que ya no están.
  // No borra grupos recién agregados (sin cantidad). Solo limpia técnicos inválidos.
  useEffect(() => {
    const validTecs = new Set(report.team);
    const original = report.preventivosResumen?.porTecnico || [];
    const cleaned = original.map(t => {
      const tecnicos = t.tecnicos || (t.tecnico ? [t.tecnico] : []);
      const filteredTecs = tecnicos.filter(name => validTecs.has(name));
      return { tecnicos: filteredTecs, cantidad: t.cantidad };
    });
    // detectar si hubo cambios reales en los técnicos
    const changed = cleaned.some((t, i) => {
      const o = original[i];
      const ot = o?.tecnicos || (o?.tecnico ? [o.tecnico] : []);
      return JSON.stringify(t.tecnicos) !== JSON.stringify(ot);
    });
    if (changed) {
      updateResumen({ porTecnico: cleaned });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [report.team]);

  // V2.4 — Suma del detalle (cada grupo cuenta su cantidad UNA SOLA VEZ, no se multiplica
  // por la cantidad de técnicos del grupo). Es lo que se compara con "Preventivos realizados".
  const sumaPorTecnico = (report.preventivosResumen?.porTecnico || [])
    .reduce((s, t) => s + (Number(t.cantidad) || 0), 0);
  const realizadosNum = Number(report.preventivosResumen?.realizados) || 0;
  const validacionCruzadaOK = realizadosNum === 0 || sumaPorTecnico === realizadosNum;
  // PR-2 (v3.20) — Tarjeta de OT correctiva extraída a función para renderarla
  // en 2 columnas (Del turno | Heredadas) preservando el índice original i de
  // report.corrective (updateCorrectiveItem, otErrorIndices, timelineDraft,
  // avanceToggle, addTimelineEntry, id form-ot-* siguen usando ese índice).
  const renderOTCard = (c, i) => {
            // V2.5 — Técnico obligatorio en cualquier estado (no solo "Realizada")
            const missingTech = (!c.technicians || c.technicians.length === 0);
            // V2.4 — Determinar si la OT es nueva (creada en este turno).
            // Solo se usa para el flag de error de formato: la validación de
            // formato (regla 4) aplica únicamente a OTs creadas en este turno,
            // las heredadas quedan exentas para no bloquear el turno por dato viejo.
            const currentShiftKey = `${report.date}-${report.shift}`;
            const isNewOT = c.createdInShift === currentShiftKey;
            // v3.22 (#39) — isLegacyFormat se eliminó junto con la rama de texto
            // libre de OTNumberInput. Todas las OTs usan el input estructurado.
            const otHasError = isNewOT && !isValidOT(c.ot);
            // v3.16 — OT marcada como error por la validación al intentar guardar
            const hasValidationError = otErrorIndices.has(i);
            // Borde rojo: preexistente (missingTech / otHasError) o nuevo (validación)
            const cardHasError = missingTech || otHasError || hasValidationError;
            // Para errores de avance faltante: resaltar extra la sección de avance
            const highlightAvance = hasValidationError && otErrorType === 'avance';
            return (
              <div key={i} id={`form-ot-${c.ot || `idx-${i}`}`} className={`border rounded-lg p-3 relative scroll-mt-32 transition-colors ${
                hasValidationError
                  ? 'border-red-500 bg-red-50/60 ring-2 ring-red-300/60'
                  : cardHasError
                  ? 'border-red-300 bg-red-50/40'
                  : 'border-slate-200 bg-slate-50/40'
              }`}>
                {/* v3.16 — Badge de error de validación: visible en el ángulo sup. izquierdo */}
                {hasValidationError && (
                  <div className="absolute top-2 left-2 z-10 inline-flex items-center gap-1 px-1.5 py-0.5 bg-red-600 text-white rounded text-[10px] font-bold">
                    <AlertTriangle className="w-3 h-3" />
                    {otErrorType === 'avance' ? 'Falta avance' : otErrorType === 'tecnico' ? 'Sin técnico' : 'Error'}
                  </div>
                )}
                {/* V2.6 — Botón eliminar OT (solo modo admin) */}
                {adminMode && (
                  <button
                    onClick={() => updateList('corrective', l => l.filter((_, j) => j !== i))}
                    className="absolute top-2 right-2 z-10 text-red-500 hover:text-red-700 hover:bg-red-100 rounded p-1.5 transition ring-1 ring-red-200 bg-white"
                    title="Eliminar OT (modo admin)">
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                )}
                {/* v3.16 — Si hay badge de error, agregar padding arriba para que no tape el N° OT */}
                <div className={`grid grid-cols-12 gap-2 mb-2 ${hasValidationError ? 'mt-6' : ''}`}>
                  <Field label="N° OT *" className="col-span-7 lg:col-span-3">
                    <OTNumberInput
                      value={c.ot}
                      onChange={(newOt) => updateCorrectiveItem(i, { ot: newOt })}
                      hasError={otHasError}
                    />
                  </Field>
                  <Field label="Estado" className="col-span-5 lg:col-span-2 lg:order-3">
                    <select className={`${inputCls} font-semibold ${c.state === 'Sin Iniciar' ? 'text-red-600' : c.state === 'En Curso' ? 'text-amber-600' : 'text-emerald-600'}`}
                      value={c.state}
                      onChange={e => updateCorrectiveItem(i, { state: e.target.value })}>
                      {ESTADOS_OT.map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </Field>
                  <Field label="Equipo / Sector" className="col-span-12 lg:col-span-3 lg:order-2">
                    <input className={inputCls} value={c.equipoCodigo}
                      onChange={e => updateCorrectiveItem(i, { equipoCodigo: e.target.value })} />
                  </Field>
                  <Field label="Técnico/s asignado/s *" className="col-span-12 lg:col-span-4 lg:order-4">
                    <MultiSelect options={teamOptionsForOT(c)} value={c.technicians}
                      onChange={vals => updateCorrectiveItem(i, { technicians: vals })}
                      placeholder="Seleccionar técnico/s o encargado/s…" />
                  </Field>
                  {/* V2.0: BOTÓN ELIMINAR REMOVIDO; V2.6: re-agregado solo en modo admin (arriba derecha) */}
                </div>
                {missingTech && (
                  <div className="text-[11px] text-red-700 mb-1 inline-flex items-center gap-1">
                    <AlertTriangle className="w-3 h-3" />
                    Todas las OTs deben tener al menos un técnico asignado
                  </div>
                )}
                {otHasError && (
                  <div className="text-[11px] text-red-700 mb-1 inline-flex items-center gap-1">
                    <AlertTriangle className="w-3 h-3" />
                    El N° OT debe tener formato XXX-YYYYY (sector + 5 dígitos). Ej: FOA1-01395
                  </div>
                )}
                <Field label="Tarea / descripción">
                  <textarea rows={2} className={inputCls} style={{ fieldSizing: 'content', minHeight: '4rem' }} value={c.task}
                    onChange={e => updateCorrectiveItem(i, { task: e.target.value })} />
                </Field>
                {/* V2.4 — Timeline de Estado de Avance.
                    - Aparece SIEMPRE si la OT tiene entradas previas (las muestra read-only)
                    - Aparece con campo de carga SI el estado actual es "En Curso"
                    - V2.5 — También aparece cuando se pasa a "Realizada" en este turno
                      (para que el responsable cargue el avance final).
                    - Es OBLIGATORIO agregar entrada al guardar SOLO cuando hay
                      cambio de estado en este turno (ver requiresAdvanceEntry). */}
                {(c.state === 'En Curso' || requiresAdvanceEntry(c) || (c.timeline && c.timeline.length > 0)) && (() => {
                  const timeline = c.timeline || [];
                  const currentShiftKey = `${report.date}-${report.shift}`;
                  // Buscar si HAY alguna entrada del turno actual (no solo la última)
                  const hasEntryFromCurrentShift = timeline.some(e => e.shiftKey === currentShiftKey);
                  // V2.5 — Solo requiere cuando hubo cambio de estado real en este turno
                  const requiresEntry = requiresAdvanceEntry(c) && !hasEntryFromCurrentShift;
                  // #avance (v3.9) — ¿es una OT heredada En Curso que NO requiere avance?
                  // (En Curso → En Curso, no nueva). Para esas mostramos el toggle
                  // "¿hubo avance?" en vez de un campo abierto que invita a poner ".".
                  const prevState = c.ot ? previousStateMap.get(c.ot) : null;
                  const esHeredadaSinCambio = c.state === 'En Curso'
                    && !requiresAdvanceEntry(c)
                    && prevState != null;  // existía antes (heredada), no es nueva
                  // Default del toggle: No, salvo que YA exista una entrada del turno
                  // actual (ej. editar un reporte donde sí se cargó avance) → Sí.
                  const toggleVal = avanceToggle[i] !== undefined
                    ? avanceToggle[i]
                    : hasEntryFromCurrentShift;
                  // El campo de carga se muestra si: requiere avance obligatorio (cierre/
                  // cambio de estado, como antes), o el usuario tildó "sí hubo avance".
                  const mostrarCampoCarga = requiresEntry
                    || (c.state === 'En Curso' && (!esHeredadaSinCambio || toggleVal));
                  return (
                    <div className={`mt-2 border rounded-lg p-3 ${
                      highlightAvance
                        ? 'border-red-500 bg-red-50/60 ring-2 ring-red-300/60'
                        : requiresEntry && timelineDraft[i] === undefined
                        ? 'border-amber-300 bg-amber-50/40'
                        : 'border-slate-200 bg-slate-50/30'
                    }`}>
                      <div className="flex items-center justify-between mb-2">
                        <div className="text-[11px] font-semibold text-slate-700 uppercase tracking-wide inline-flex items-center gap-1.5">
                          <Activity className={`w-3.5 h-3.5 ${highlightAvance ? 'text-red-500' : 'text-amber-500'}`} />
                          Estado de avance {(requiresEntry || highlightAvance) && <span className="text-red-600">*</span>}
                        </div>
                        {highlightAvance && !requiresEntry && (
                          <span className="text-[10px] text-red-700 font-medium inline-flex items-center gap-1">
                            <AlertTriangle className="w-3 h-3" />
                            Cargá el avance antes de guardar
                          </span>
                        )}
                        {requiresEntry && !highlightAvance && (
                          <span className="text-[10px] text-amber-700 font-medium inline-flex items-center gap-1">
                            <AlertTriangle className="w-3 h-3" />
                            Requiere entrada del turno actual
                          </span>
                        )}
                      </div>

                      {/* #avance (v3.9) — Toggle "¿hubo avance este turno?" para OTs
                          heredadas En Curso sin cambio de estado. Default No. Evita el
                          ruido de poner "." cuando el turno no trabajó esa OT. */}
                      {esHeredadaSinCambio && !hasEntryFromCurrentShift && (
                        <label className="flex items-center gap-2 mb-2 text-[12px] text-slate-600 cursor-pointer select-none">
                          <input
                            type="checkbox"
                            className="w-4 h-4 rounded border-slate-300 text-amber-500 focus:ring-amber-400"
                            checked={toggleVal}
                            onChange={e => {
                              const val = e.target.checked;
                              setAvanceToggle(t => ({ ...t, [i]: val }));
                              // Al pasar a "No", limpiar el texto en redacción (no se guarda).
                              if (!val) setTimelineDraft(d => { const n = { ...d }; delete n[i]; return n; });
                            }}
                          />
                          ¿Hubo avance en este turno? <span className="text-slate-400">(si no, queda sin novedad)</span>
                        </label>
                      )}

                      {/* Entradas previas (read-only para usuarios normales, editables en modo admin) */}
                      {timeline.length > 0 && (
                        <div className="space-y-1.5 mb-2">
                          {timeline.map((entry, ei) => {
                            // V2.6 — Estado local de edición inline (solo admin)
                            const editKey = `${i}-${ei}`;
                            const isEditing = adminMode && timelineEditingKey === editKey;
                            return (
                              <div key={ei} className={`flex items-start gap-2 text-[12px] border rounded px-2 py-1.5 ${isEditing ? 'border-sky-400 ring-1 ring-sky-200 bg-white' : (entry.text && (entry.text.startsWith('[Edición admin]') || entry.text.startsWith('[Reapertura admin]'))) ? 'bg-slate-100 border-slate-300' : 'bg-white border-slate-200'}`}>
                                <div className="flex-shrink-0 w-32">
                                  <div className="text-[10px] text-slate-500 num font-medium">
                                    {entry.date} · {entry.shift}
                                  </div>
                                  <div className="text-[9px] text-slate-400 truncate" title={entry.author}>
                                    {entry.author || '—'}
                                  </div>
                                </div>
                                <div className="flex-1 min-w-0">
                                  {isEditing ? (
                                    <textarea
                                      rows={2}
                                      autoFocus
                                      className={`${inputCls} text-[12px] w-full`}
                                      value={timelineEditDraft}
                                      onChange={e => setTimelineEditDraft(e.target.value)}
                                    />
                                  ) : (
                                    <div className="text-slate-700 whitespace-pre-wrap break-words">
                                      {entry.text}
                                    </div>
                                  )}
                                </div>
                                {/* V2.6 — Botones admin para editar/eliminar entrada */}
                                {adminMode && !isEditing && (
                                  <div className="flex flex-col gap-1 flex-shrink-0">
                                    <button
                                      onClick={() => {
                                        setTimelineEditingKey(editKey);
                                        setTimelineEditDraft(entry.text || '');
                                      }}
                                      className="text-sky-600 hover:text-sky-800 hover:bg-sky-50 rounded p-1 transition"
                                      title="Editar (admin)">
                                      <Edit3 className="w-3 h-3" />
                                    </button>
                                    <button
                                      onClick={() => {
                                        updateCorrectiveItem(i, {
                                          timeline: timeline.filter((_, j) => j !== ei)
                                        });
                                      }}
                                      className="text-red-500 hover:text-red-700 hover:bg-red-50 rounded p-1 transition"
                                      title="Eliminar entrada (admin)">
                                      <Trash2 className="w-3 h-3" />
                                    </button>
                                  </div>
                                )}
                                {isEditing && (
                                  <div className="flex flex-col gap-1 flex-shrink-0">
                                    <button
                                      onClick={() => {
                                        const newTimeline = timeline.map((e, j) =>
                                          j === ei ? { ...e, text: timelineEditDraft } : e
                                        );
                                        updateCorrectiveItem(i, { timeline: newTimeline });
                                        setTimelineEditingKey(null);
                                        setTimelineEditDraft('');
                                      }}
                                      className="text-emerald-600 hover:text-emerald-800 hover:bg-emerald-50 rounded p-1 transition"
                                      title="Guardar">
                                      <CheckCircle2 className="w-3.5 h-3.5" />
                                    </button>
                                    <button
                                      onClick={() => {
                                        setTimelineEditingKey(null);
                                        setTimelineEditDraft('');
                                      }}
                                      className="text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded p-1 transition"
                                      title="Cancelar">
                                      <X className="w-3.5 h-3.5" />
                                    </button>
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      )}

                      {/* V2.5 — Campo de carga: visible si está En Curso (carga opcional)
                          o si está Realizada y requiere avance (cierre obligatorio del trabajo).
                          #avance (v3.9) — en OTs heredadas En Curso sin cambio, solo si el
                          usuario tildó "¿hubo avance?" (mostrarCampoCarga). */}
                      {mostrarCampoCarga && (
                        <div className="flex items-start gap-2">
                          <textarea
                            rows={2}
                            className={`${inputCls} flex-1 text-[12px] ${requiresEntry && timelineDraft[i] === undefined ? 'border-amber-400' : ''}`}
                            placeholder={hasEntryFromCurrentShift ? "Ya cargaste una entrada este turno. Podés agregar otra opcionalmente…" : "Describí el avance realizado este turno…"}
                            value={timelineDraft[i] || ''}
                            onChange={e => setTimelineDraft(d => ({ ...d, [i]: e.target.value }))}
                          />
                          <button
                            onClick={() => addTimelineEntry(i)}
                            disabled={!timelineDraft[i] || timelineDraft[i].trim().length === 0}
                            className={`${buttonCls} bg-amber-100 text-amber-700 hover:bg-amber-200 disabled:opacity-40 disabled:cursor-not-allowed self-start`}
                          >
                            <Plus className="w-4 h-4" />Guardar avance
                          </button>
                        </div>
                      )}
                    </div>
                  );
                })()}
              </div>
            );
  };

  return (
    <div className="space-y-5">
      {/* #7 v3.6 — Modal de recuperación de borrador local sin guardar */}
      {draftRecovery && (
        <DraftRecoveryDialog
          date={draftRecovery.draft.report.date}
          shift={draftRecovery.draft.report.shift}
          savedAt={draftRecovery.draft.savedAt}
          serverUpdatedAt={draftRecovery.serverUpdatedAt}
          onRecover={applyDraftRecovery}
          onDiscard={discardDraftRecovery}
        />
      )}
      {/* V2.9 — Banner: admin editando reporte histórico (que ya está guardado). */}
      {adminMode && originalReport && (
        <div className="flex items-start gap-2 px-3 py-2.5 bg-sky-50 border border-sky-200 rounded-lg">
          <Shield className="w-4 h-4 text-sky-600 flex-shrink-0 mt-0.5" />
          <div className="text-[13px] text-sky-800 leading-snug">
            <strong>Estás editando un reporte histórico</strong> ({formatDateShort(report.date)} · {report.shift}).
            {' '}Los cambios que hagas sobre OTs (estado o entradas del timeline) pueden propagarse a reportes posteriores que tengan las mismas OTs. Vas a poder revisar antes de confirmar.
          </div>
        </div>
      )}
      {/* v3.13 (#12) — La action bar sticky fue eliminada. Guardar + saveMsg viven
          ahora en el header global (visibles siempre sin scroll). Limpiar + Eliminar
          reporte pasaron al FAB flotante admin (abajo a la izquierda, ver más abajo).
          Acá solo queda la línea informativa de loadInfo, que aparece cuando hay mensaje. */}
      {loadInfo && (
        <div className="flex items-center gap-2 text-sm text-slate-600 bg-white border border-slate-200 rounded-xl px-3 py-2.5 shadow-sm">
          <ClipboardList className="w-5 h-5 text-sky-600 flex-shrink-0" />
          <span className="font-medium">{loadInfo}</span>
        </div>
      )}

      {/* #9 (v3.14) — Banner read-only: no-admin viendo un reporte anterior a ayer.
          No puede editar (corrección retroactiva = solo admin), pero sí mirarlo. El botón
          lo devuelve al turno de hoy para que no quede atrapado (Fecha/Turno también se
          deshabilitan dentro del fieldset). */}
      {isReadOnly && (
        <div className="flex items-start gap-2 px-3 py-2.5 bg-amber-50 border border-amber-200 rounded-lg">
          <Lock className="w-4 h-4 text-amber-600 flex-shrink-0 mt-0.5" />
          <div className="text-[13px] text-amber-800 leading-snug flex-1">
            <strong>Solo lectura</strong> — este reporte ({formatDateShort(report.date)} · {report.shift}) es de un día anterior. Para corregir reportes pasados se necesita acceso admin. Podés editar los reportes de hoy y de ayer.
          </div>
          <button
            onClick={() => { setDateShift(todayLocalISO(), report.shift); }}
            className="flex-shrink-0 inline-flex items-center gap-1.5 px-3 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-xs font-medium transition">
            Ir al reporte de hoy
          </button>
        </div>
      )}

      {/* #9 (v3.14) — fieldset disabled deshabilita TODOS los controles hijos cuando es
          read-only (incluidos Fecha/Turno; por eso el botón "Ir al reporte de hoy" del banner
          de arriba, que queda FUERA del fieldset, para no dejar al usuario atrapado). */}
      <fieldset disabled={isReadOnly} className="space-y-5 min-w-0 border-0 p-0 m-0 disabled:opacity-60">
      {/* v3.13 (#12) — Información del Turno + Equipo del Turno en línea, 50/50 en desktop,
          apiladas en mobile. Antes eran dos Cards a ancho completo, una sobre otra. */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 items-stretch">
        {/* COMPACT TURNO HEADER */}
        <Card className="p-5">
          <SectionTitle icon={Calendar} accent="sky">Información del Turno</SectionTitle>
          <div className="grid grid-cols-3 gap-2 md:gap-3">
            <Field label="Fecha">
              <input type="date" className={inputCls} value={report.date}
                max={adminMode ? undefined : todayLocalISO()}
                onChange={e => setDateShift(e.target.value, report.shift)} />
            </Field>
            <Field label="Turno">
              <select className={inputCls} value={report.shift} onChange={e => setDateShift(report.date, e.target.value)}>
                {TURNOS.map(s => <option key={s}>{s}</option>)}
              </select>
            </Field>
            <Field label="Responsable">
              <select className={inputCls} value={report.responsable} onChange={e => update({ responsable: e.target.value })}>
                <option value="">— Seleccionar —</option>
                {RESPONSABLES.map(r => <option key={r.id} value={r.name}>{r.name}</option>)}
              </select>
            </Field>
          </div>
        </Card>

        {/* EQUIPO (only names) */}
        <Card className="p-5">
          <SectionTitle icon={Users} accent="sky">Equipo del Turno</SectionTitle>
          <Field label="Técnicos del turno (selección múltiple)">
            <MultiSelect options={TECNICO_NAMES} value={report.team} onChange={vals => update({ team: vals })}
              placeholder="Seleccionar técnicos…" />
          </Field>
        </Card>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 items-stretch">
      <Card className="p-5">
        <div className="flex items-center justify-between mb-4">
          <SectionTitle icon={FileText} accent="amber">Comentarios</SectionTitle>
          <button onClick={() => updateList('comments', l => [...l, { text: '', priority: 'Normal' }])}
            className={`${buttonCls} bg-amber-50 text-amber-700 hover:bg-amber-100`}>
            <Plus className="w-4 h-4" />Agregar comentario
          </button>
        </div>
        {report.comments.length === 0 && <EmptyHint>Sin comentarios.</EmptyHint>}
        <div className="space-y-2">
          {report.comments.map((c, i) => (
            <div key={i} className={`grid grid-cols-12 gap-2 items-start p-2 rounded-lg ${c.priority === 'Urgente' ? 'bg-red-50 border border-red-200' : ''}`}>
              <textarea rows={2} className={`${inputCls} col-span-9 ${c.priority === 'Urgente' ? 'border-red-300 focus:ring-red-300' : ''}`}
                placeholder="Comentario…" value={c.text}
                onChange={e => updateList('comments', l => l.map((x, j) => j === i ? { ...x, text: e.target.value } : x))} />
              <select className={`${inputCls} col-span-2 font-semibold ${c.priority === 'Urgente' ? 'text-red-600 border-red-300' : 'text-slate-600'}`}
                value={c.priority}
                onChange={e => updateList('comments', l => l.map((x, j) => j === i ? { ...x, priority: e.target.value } : x))}>
                {PRIORIDADES.map(p => <option key={p}>{p}</option>)}
              </select>
              <button onClick={() => updateList('comments', l => l.filter((_, j) => j !== i))}
                className="col-span-1 inline-flex items-center justify-center text-red-500 hover:text-red-700 hover:bg-red-50 rounded p-2 transition self-center">
                <Trash2 className="w-4 h-4" />
              </button>
            </div>
          ))}
        </div>
      </Card>
        <Card className="p-5">
          <div className="flex items-center justify-between mb-4">
            <SectionTitle icon={Building2} accent="indigo">Proveedores</SectionTitle>
            <button onClick={() => updateServicios({ proveedores: [...report.servicios.proveedores, { provider: '', task: '' }] })}
              className={`${buttonCls} bg-indigo-50 text-indigo-700 hover:bg-indigo-100`}>
              <Plus className="w-4 h-4" />Agregar
            </button>
          </div>
          {report.servicios.proveedores.length === 0 && <EmptyHint>Sin tareas de proveedores.</EmptyHint>}
          <div className="space-y-2">
            {report.servicios.proveedores.map((p, i) => (
              <div key={i} className="grid grid-cols-12 gap-2">
                <input className={`${inputCls} col-span-3`} placeholder="Proveedor" value={p.provider}
                  onChange={e => updateServicios({ proveedores: report.servicios.proveedores.map((x, j) => j === i ? { ...x, provider: e.target.value } : x) })} />
                <input className={`${inputCls} col-span-8`} placeholder="Tarea realizada" value={p.task}
                  onChange={e => updateServicios({ proveedores: report.servicios.proveedores.map((x, j) => j === i ? { ...x, task: e.target.value } : x) })} />
                <button onClick={() => updateServicios({ proveedores: report.servicios.proveedores.filter((_, j) => j !== i) })}
                  className="col-span-1 inline-flex items-center justify-center text-red-500 hover:text-red-700 hover:bg-red-50 rounded p-2 transition">
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>
            ))}
          </div>
        </Card>
      </div>

      <Card className="p-5 scroll-mt-32" id="form-preventivos">
        <SectionTitle icon={ListChecks} accent="emerald">Resumen Preventivos del Turno</SectionTitle>
        <p className="text-xs text-slate-500 mb-4">
          Estos son los totales globales del turno (los carga el responsable). Si hay realizados &gt; 0,
          la suma del detalle por técnico debe coincidir con "Preventivos realizados".
        </p>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-4 max-w-2xl">
          <Field label="Preventivos asignados en el turno">
            <input type="number" min="0" step="1" className={`${inputCls} num`}
              value={report.preventivosResumen?.asignados ?? ''}
              onChange={e => updateResumen({ asignados: e.target.value })} />
          </Field>
          <Field label="Preventivos realizados en el turno">
            <input type="number" min="0" step="1"
              className={`${inputCls} num ${!validacionCruzadaOK ? 'border-red-400 focus:ring-red-300' : ''}`}
              value={report.preventivosResumen?.realizados ?? ''}
              onChange={e => updateResumen({ realizados: e.target.value })} />
          </Field>
        </div>

        <div className="flex items-center justify-between mb-3">
          <h3 className="text-sm font-semibold text-slate-700 inline-flex items-center gap-2">
            <Users className="w-4 h-4 text-emerald-500" />Detalle por técnico
          </h3>
          <button
            onClick={() => updateResumen({ porTecnico: [...(report.preventivosResumen?.porTecnico || []), { tecnicos: [], cantidad: '' }] })}
            disabled={report.team.length === 0}
            className={`${buttonCls} bg-emerald-50 text-emerald-700 hover:bg-emerald-100 disabled:opacity-50 disabled:cursor-not-allowed`}
            title={report.team.length === 0 ? 'Cargá primero el equipo del turno' : ''}>
            <Plus className="w-4 h-4" />Agregar grupo
          </button>
        </div>

        {/* V2.4 — Opción C: multi-select de técnicos por grupo.
            Si Juan + Pedro hicieron 4 preventivos juntos, se carga UNA fila con
            ambos seleccionados y cantidad 4. La cantidad cuenta UNA VEZ para
            la validación cruzada (4, no 8). En estadísticas individuales,
            tanto Juan como Pedro reciben 4 cada uno. */}
        {report.team.length === 0 ? (
          <EmptyHint>Cargá primero el equipo del turno para asignar preventivos por técnico.</EmptyHint>
        ) : (report.preventivosResumen?.porTecnico || []).length === 0 ? (
          <EmptyHint>Sin detalle por técnico.</EmptyHint>
        ) : (
          <div className="space-y-2">
            {(report.preventivosResumen?.porTecnico || []).map((t, i) => {
              // Compat hacia atrás: si vino del schema viejo con `tecnico` único, lo migramos a array
              const tecnicos = t.tecnicos || (t.tecnico ? [t.tecnico] : []);
              return (
                <div key={i} className="grid grid-cols-12 gap-2">
                  <div className="col-span-7">
                    <MultiSelect
                      options={report.team}
                      value={tecnicos}
                      onChange={vals => updateResumen({
                        porTecnico: report.preventivosResumen.porTecnico.map((x, j) =>
                          j === i ? { tecnicos: vals, cantidad: x.cantidad } : x
                        )
                      })}
                      placeholder="Seleccionar técnico/s del grupo…"
                    />
                  </div>
                  <input
                    type="number" min="1" step="1" placeholder="Cant."
                    className={`${inputCls} col-span-4 num`}
                    value={t.cantidad}
                    onChange={e => updateResumen({
                      porTecnico: report.preventivosResumen.porTecnico.map((x, j) =>
                        j === i ? { ...x, tecnicos, cantidad: e.target.value } : x
                      )
                    })} />
                  <button
                    onClick={() => updateResumen({
                      porTecnico: report.preventivosResumen.porTecnico.filter((_, j) => j !== i)
                    })}
                    className="col-span-1 inline-flex items-center justify-center text-red-500 hover:text-red-700 hover:bg-red-50 rounded p-2 transition">
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              );
            })}
          </div>
        )}

        {/* Feedback de validación cruzada */}
        {realizadosNum > 0 && (
          <div className={`mt-3 text-xs p-2 rounded-lg flex items-center gap-2 ${validacionCruzadaOK
              ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
              : 'bg-red-50 text-red-700 border border-red-200'
            }`}>
            {validacionCruzadaOK
              ? <CheckCircle2 className="w-4 h-4" />
              : <AlertTriangle className="w-4 h-4" />
            }
            Suma del detalle: <span className="num font-bold">{sumaPorTecnico}</span> ·
            Realizados: <span className="num font-bold">{realizadosNum}</span>
            {!validacionCruzadaOK && ' — debe coincidir antes de guardar'}
          </div>
        )}
      </Card>

      <Card className="p-5">
        <SectionTitle icon={Activity} accent="violet">Servicios</SectionTitle>

        {/* PLANTA DE EFLUENTES Y CALDERA — V2.0: schema NUEVO (PTEL + Caldera + Ablandadores) */}
        <div className="mb-5">
          <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
            <Flame className="w-4 h-4 text-orange-500" />Planta de Efluentes y Caldera
          </h3>
          <div className="grid grid-cols-12 gap-3 mb-4">
            <Field label="Técnicos (foguistas)" className="col-span-7">
              <MultiSelect options={FOGUISTAS} value={report.servicios.plantaCaldera.tecnicos || []}
                onChange={vals => updateServicios({ plantaCaldera: { ...report.servicios.plantaCaldera, tecnicos: vals } })}
                placeholder="Seleccionar foguistas…" />
            </Field>
            <Field label="Estado" className="col-span-5">
              <select className={`${inputCls} font-semibold ${report.servicios.plantaCaldera.estado === 'Operativa' ? 'text-emerald-600' : 'text-red-600'}`}
                value={report.servicios.plantaCaldera.estado}
                onChange={e => updateServicios({ plantaCaldera: { ...report.servicios.plantaCaldera, estado: e.target.value } })}>
                {ESTADOS_PLANTA.map(s => <option key={s}>{s}</option>)}
              </select>
            </Field>
          </div>

          {/* PTEL */}
          <div className="border border-slate-200 rounded-lg p-3 bg-orange-50/30 mb-3">
            <div className="text-[11px] font-bold text-orange-700 uppercase tracking-wider mb-2">PTEL</div>
            <div className="grid grid-cols-12 gap-3">
              {[
                ['caudal', 'Caudal (m³/h)'],
                ['vacio', 'Vacío del equipo'],
                ['deltaT', 'ΔT entre torres (°C)'],
                ['tk1', '% Nivel TK1'],
                ['tk2', '% Nivel TK2'],
                ['tk7', '% Nivel TK7']
              ].map(([k, label]) => (
                <Field key={k} label={label} className="col-span-4 lg:col-span-2">
                  <input type="number" step="any" className={`${inputCls} num`}
                    value={report.servicios.plantaCaldera[k] ?? ''}
                    onChange={e => updateServicios({ plantaCaldera: { ...report.servicios.plantaCaldera, [k]: e.target.value } })} />
                </Field>
              ))}
            </div>
          </div>

          {/* CALDERA y AGUA ABLANDADORES */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div className="border border-slate-200 rounded-lg p-3 bg-red-50/30">
              <div className="text-[11px] font-bold text-red-700 uppercase tracking-wider mb-2">Caldera</div>
              <div className="grid grid-cols-2 gap-3 items-end">
                <Field label="Conductividad (mS)">
                  <input type="number" step="any" className={`${inputCls} num`}
                    value={report.servicios.plantaCaldera.conductividadCaldera ?? ''}
                    onChange={e => updateServicios({ plantaCaldera: { ...report.servicios.plantaCaldera, conductividadCaldera: e.target.value } })} />
                </Field>
                <Field label="pH">
                  <input type="number" step="any" className={`${inputCls} num`}
                    value={report.servicios.plantaCaldera.pHCaldera ?? ''}
                    onChange={e => updateServicios({ plantaCaldera: { ...report.servicios.plantaCaldera, pHCaldera: e.target.value } })} />
                </Field>
              </div>
            </div>
            <div className="border border-slate-200 rounded-lg p-3 bg-blue-50/30">
              <div className="text-[11px] font-bold text-blue-700 uppercase tracking-wider mb-2">Agua Ablandadores</div>
              <div className="grid grid-cols-2 gap-3 items-end">
                <Field label="Conductividad (mS)">
                  <input type="number" step="any" className={`${inputCls} num`}
                    value={report.servicios.plantaCaldera.conductividadAblandador ?? ''}
                    onChange={e => updateServicios({ plantaCaldera: { ...report.servicios.plantaCaldera, conductividadAblandador: e.target.value } })} />
                </Field>
                <Field label="pH">
                  <input type="number" step="any" className={`${inputCls} num`}
                    value={report.servicios.plantaCaldera.pHAblandador ?? ''}
                    onChange={e => updateServicios({ plantaCaldera: { ...report.servicios.plantaCaldera, pHAblandador: e.target.value } })} />
                </Field>
              </div>
            </div>
          </div>
        </div>

        {/* COMPRESORES Y GRUPOS ELECTRÓGENOS Y CISTERNAS */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-5">
          <div className="border border-slate-200 rounded-lg p-3">
            <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
              <Cog className="w-4 h-4 text-sky-500" />Compresores
            </h3>
            <div className="space-y-2">
              {report.servicios.compresores.map((c, i) => (
                <div key={c.code} className="flex items-center gap-2">
                  <span className="text-sm flex-1 num">{c.code}</span>
                  <select className={`text-xs px-2 py-1 border border-slate-300 rounded bg-white font-semibold ${c.state === 'Operativo' ? 'text-emerald-600' : 'text-red-600'}`}
                    value={c.state}
                    onChange={e => updateServicios({ compresores: report.servicios.compresores.map((x, j) => j === i ? { ...x, state: e.target.value } : x) })}>
                    {ESTADOS_SERVICIO.map(s => <option key={s}>{s}</option>)}
                  </select>
                </div>
              ))}
            </div>
          </div>

          <div className="border border-slate-200 rounded-lg p-3">
            <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
              <Zap className="w-4 h-4 text-amber-500" />Grupos Electrógenos
            </h3>
            <div className="space-y-2">
              {report.servicios.gruposElectrogenos.map((g, i) => (
                <div key={g.code} className="flex items-center gap-2">
                  <span className="text-sm flex-1 num">{g.code}</span>
                  <select className={`text-xs px-2 py-1 border border-slate-300 rounded bg-white font-semibold ${g.state === 'Operativo' ? 'text-emerald-600' : 'text-red-600'}`}
                    value={g.state}
                    onChange={e => updateServicios({ gruposElectrogenos: report.servicios.gruposElectrogenos.map((x, j) => j === i ? { ...x, state: e.target.value } : x) })}>
                    {ESTADOS_SERVICIO.map(s => <option key={s}>{s}</option>)}
                  </select>
                </div>
              ))}
            </div>
          </div>

          <div className="border border-slate-200 rounded-lg p-3">
            <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
              <Beaker className="w-4 h-4 text-cyan-500" />Cisternas
            </h3>
            <div className="space-y-2">
              <Field label="Nivel">
                <select className={`${inputCls} text-sm font-semibold ${report.servicios.cisternas.nivel === 'Alto' ? 'text-emerald-600' : report.servicios.cisternas.nivel === 'Medio' ? 'text-amber-600' : 'text-red-600'}`}
                  value={report.servicios.cisternas.nivel}
                  onChange={e => updateServicios({ cisternas: { ...report.servicios.cisternas, nivel: e.target.value } })}>
                  {NIVELES_CISTERNAS.map(s => <option key={s}>{s}</option>)}
                </select>
              </Field>
              <Field label="Estado">
                <select className={`${inputCls} text-sm font-semibold ${report.servicios.cisternas.estado === 'Ingreso Normal' ? 'text-emerald-600' : report.servicios.cisternas.estado === 'Ingreso Limitado' ? 'text-amber-600' : 'text-red-600'}`}
                  value={report.servicios.cisternas.estado}
                  onChange={e => updateServicios({ cisternas: { ...report.servicios.cisternas, estado: e.target.value } })}>
                  {ESTADOS_CISTERNAS.map(s => <option key={s}>{s}</option>)}
                </select>
              </Field>
            </div>
          </div>
        </div>

        {/* AGUA DE POZO */}
        <div className="mb-5">
          <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
            <Beaker className="w-4 h-4 text-blue-500" />Agua de Pozo
          </h3>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-w-xl">
            <Field label="Nivel de Cloro Pozo 3">
              <input type="number" step="any" className={`${inputCls} num`}
                value={report.servicios.aguaPozo?.cloroPozo3 ?? ''}
                onChange={e => updateServicios({ aguaPozo: { ...(report.servicios.aguaPozo || {}), cloroPozo3: e.target.value } })} />
            </Field>
            <Field label="Nivel de Cloro Pozo 6">
              <input type="number" step="any" className={`${inputCls} num`}
                value={report.servicios.aguaPozo?.cloroPozo6 ?? ''}
                onChange={e => updateServicios({ aguaPozo: { ...(report.servicios.aguaPozo || {}), cloroPozo6: e.target.value } })} />
            </Field>
          </div>
        </div>
      </Card>

      <Card className="p-5">
        <div className="flex items-center justify-between mb-4">
          <SectionTitle icon={Wrench} accent="orange">Mantenimiento Correctivo</SectionTitle>
          <button onClick={() => updateList('corrective', l => [{ ot: '', equipoCodigo: '', task: '', technicians: [], state: 'Sin Iniciar', createdInShift: `${report.date}-${report.shift}`, timeline: [] }, ...l])}
            className={`${buttonCls} bg-orange-50 text-orange-700 hover:bg-orange-100`}>
            <Plus className="w-4 h-4" />Agregar OT
          </button>
        </div>

        {/* V2.5 — Banner permanente sobre OTs heredadas */}
        <div className="mb-4 flex items-start gap-2 px-3 py-2.5 bg-red-50 border border-red-200 rounded-lg">
          <AlertTriangle className="w-4 h-4 text-red-600 flex-shrink-0 mt-0.5" />
          <div className="text-[13px] text-red-700 leading-snug">
            Para las OTs de turnos previos que aún no están cerradas, actualizá el "Estado de avance" sólo si hay novedades.
          </div>
        </div>

        {(() => {
          // PR-2 (v3.20) — Partición turno / heredadas para las 2 columnas (solo desktop).
          // "Del turno": OT creada en este turno, o heredada con novedad este turno
          // (avance de timeline cargado, o cambio de estado). Reactiva a report.corrective.
          const currentShiftKey = `${report.date}-${report.shift}`;
          const esDelTurno = (c) => {
            if (c.createdInShift === currentShiftKey) return true;
            const tuvoAvance = (c.timeline || []).some(e => e.shiftKey === currentShiftKey);
            if (tuvoAvance) return true;
            const prev = c.ot ? previousStateMap.get(c.ot) : null;
            if (prev != null && prev !== c.state) return true;
            return false;
          };
          const items = report.corrective.map((c, i) => ({ c, i }));
          if (report.corrective.length === 0) {
            return <EmptyHint>Sin órdenes de trabajo correctivas.</EmptyHint>;
          }
          if (!isDesktop) {
            return <div className="space-y-3">{items.map(({ c, i }) => renderOTCard(c, i))}</div>;
          }
          const colTurno = items.filter(({ c }) => esDelTurno(c));
          const colHeredadas = items.filter(({ c }) => !esDelTurno(c));
          const colHeader = (txt, n) => (
            <div className="text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-2 pb-1 border-b border-slate-200">
              {txt} ({n})
            </div>
          );
          return (
            <div className="grid grid-cols-2 gap-4 items-start">
              <div>
                {colHeader('Correctivos del turno', colTurno.length)}
                {colTurno.length === 0
                  ? <EmptyHint>Sin OTs cargadas en este turno.</EmptyHint>
                  : <div className="space-y-3">{colTurno.map(({ c, i }) => renderOTCard(c, i))}</div>}
              </div>
              <div>
                {colHeader('Correctivos heredados', colHeredadas.length)}
                {colHeredadas.length === 0
                  ? <EmptyHint>Sin OTs heredadas pendientes.</EmptyHint>
                  : <div className="space-y-3">{colHeredadas.map(({ c, i }) => renderOTCard(c, i))}</div>}
              </div>
            </div>
          );
        })()}
      </Card>

      {adminMode && (
      <Card className="p-5">
        <div className="flex items-center justify-between mb-4">
          <SectionTitle icon={ListChecks} accent="emerald">Mantenimiento Preventivo</SectionTitle>
          <button onClick={() => updateList('preventive', l => [...l, { codigoTarea: '', equipoCodigo: '', equipoDescripcion: '', task: '', comments: '', otCorrectivaAsociada: '', technicians: [], frequency: 'Diaria' }])}
            className={`${buttonCls} bg-emerald-50 text-emerald-700 hover:bg-emerald-100`}>
            <Plus className="w-4 h-4" />Agregar tarea
          </button>
        </div>
        {report.preventive.length === 0 && <EmptyHint>Sin tareas preventivas.</EmptyHint>}
        <div className="space-y-3">
          {report.preventive.map((p, i) => {
            // V2.5 — Técnico obligatorio también en preventivos
            const missingTechPrev = (!p.technicians || p.technicians.length === 0);
            return (
            <div key={i} className={`border rounded-lg p-3 ${missingTechPrev ? 'border-red-300 bg-red-50/40' : 'border-slate-200 bg-slate-50/40'}`}>
              <div className="grid grid-cols-12 gap-2 mb-2">
                <Field label="Código de tarea" className="col-span-2">
                  <input className={`${inputCls} num`} value={p.codigoTarea}
                    onChange={e => updateList('preventive', l => l.map((x, j) => j === i ? { ...x, codigoTarea: e.target.value } : x))} />
                </Field>
                <Field label="Equipo" className="col-span-2">
                  <input className={inputCls} value={p.equipoCodigo}
                    onChange={e => updateList('preventive', l => l.map((x, j) => j === i ? { ...x, equipoCodigo: e.target.value } : x))} />
                </Field>
                <Field label="Descripción de equipo" className="col-span-3">
                  <input className={inputCls} value={p.equipoDescripcion}
                    onChange={e => updateList('preventive', l => l.map((x, j) => j === i ? { ...x, equipoDescripcion: e.target.value } : x))} />
                </Field>
                <Field label="Frecuencia" className="col-span-2">
                  <select className={inputCls} value={p.frequency}
                    onChange={e => updateList('preventive', l => l.map((x, j) => j === i ? { ...x, frequency: e.target.value } : x))}>
                    {FRECUENCIAS.map(f => <option key={f}>{f}</option>)}
                  </select>
                </Field>
                <Field label="Técnico/s asignado/s *" className="col-span-2">
                  <MultiSelect options={report.team} value={p.technicians}
                    onChange={vals => updateList('preventive', l => l.map((x, j) => j === i ? { ...x, technicians: vals } : x))}
                    placeholder={report.team.length === 0 ? 'Cargá el equipo del turno' : 'Seleccionar…'} />
                </Field>
                <div className="col-span-1 flex items-end justify-end">
                  <button onClick={() => updateList('preventive', l => l.filter((_, j) => j !== i))}
                    className="text-red-500 hover:text-red-700 hover:bg-red-50 rounded p-2 transition">
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </div>
              {missingTechPrev && (
                <div className="text-[11px] text-red-700 mb-1 inline-flex items-center gap-1">
                  <AlertTriangle className="w-3 h-3" />
                  Todos los preventivos deben tener al menos un técnico asignado
                </div>
              )}
              <div className="grid grid-cols-3 gap-2">
                <Field label="Tarea">
                  <textarea rows={2} className={inputCls} value={p.task}
                    onChange={e => updateList('preventive', l => l.map((x, j) => j === i ? { ...x, task: e.target.value } : x))} />
                </Field>
                <Field label="Comentarios">
                  <textarea rows={2} className={inputCls} value={p.comments}
                    onChange={e => updateList('preventive', l => l.map((x, j) => j === i ? { ...x, comments: e.target.value } : x))} />
                </Field>
                <Field label="OT correctiva asociada">
                  <input className={`${inputCls} num`} placeholder="OT-XXXX" value={p.otCorrectivaAsociada}
                    onChange={e => updateList('preventive', l => l.map((x, j) => j === i ? { ...x, otCorrectivaAsociada: e.target.value } : x))} />
                </Field>
              </div>
            </div>
          );
          })}
        </div>
      </Card>
      )}
      </fieldset>

      {/* v3.13 (#12) — FAB admin flotante abajo-izquierda con Limpiar + Eliminar reporte.
          Solo admin (para no-admin no renderiza nada → sin franja ni hueco). Lejos del
          Guardar (que está arriba en el header) para no reintroducir el clic accidental de
          Limpiar que causó el incidente 21/05. Ambas acciones conservan su confirmación:
          Eliminar pasa por onDeleteReport (modal en el padre); Limpiar por cleanForm. */}
      {adminMode && (
        <div className="fixed bottom-6 left-6 z-30 flex flex-col gap-2">
          {history.some(r => r.date === report.date && r.shift === report.shift) && (
            <button onClick={onDeleteReport}
              className="inline-flex items-center gap-2 px-4 py-2.5 bg-red-600 hover:bg-red-700 text-white rounded-full font-medium text-sm shadow-lg ring-1 ring-red-700/30 transition"
              title="Eliminar este reporte completo (admin)">
              <Trash2 className="w-4 h-4" />Eliminar reporte
            </button>
          )}
          <button onClick={cleanForm}
            className="inline-flex items-center gap-2 px-4 py-2.5 bg-slate-600 hover:bg-slate-700 text-white rounded-full font-medium text-sm shadow-lg ring-1 ring-slate-700/30 transition"
            title="Limpiar formulario (admin)">
            <RefreshCw className="w-4 h-4" />Limpiar
          </button>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// V2.5 — SUB-COMPONENTE: SECCIÓN DE CORRECTIVOS EN DASHBOARD
// Renderiza una sub-sección titulada con N OTs. Soporta:
//   - showStateBadge: muestra el StatePill (Sin Iniciar / En Curso)
//   - showAvanceMark: cuando es true, las OTs con avance del turno actual:
//       · van ordenadas primero (arriba)
//       · muestran la línea destacada en verde "↳ Avance del turno: ..."
//     Sin badge "Avance hoy" — la línea verde es suficiente señal visual.
// Cuando count === 0, muestra "Sin novedades" con el subtítulo igual.
// ═══════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════
// V2.5 — SUB-COMPONENTE: SECCIÓN DE CORRECTIVOS EN DASHBOARD
// Renderiza una sub-sección titulada con N OTs. Soporta:
//   - showStateBadge: muestra el StatePill (Sin Iniciar / En Curso)
//   - showAvanceMark: cuando es true, las OTs con avance del turno actual:
//        · van ordenadas primero (arriba)
//        · muestran la línea destacada en verde "↳ Avance del turno: ..."
//     Sin badge "Avance hoy" — la línea verde es suficiente señal visual.
//   - hideAdvance (v3.18 / #36): suprime la línea de "último avance" (#19)
//        para ahorrar espacio en el export. Se usa en la columna de Pendientes.
//        El avance sigue en el timeline de la OT y en la app; solo no se exhibe acá.
// Cuando count === 0, muestra "Sin novedades" con el subtítulo igual.
// ═══════════════════════════════════════════════════════════════════
function CorrectiveSubsection({ title, count, items, showStateBadge, showAvanceMark, currentShiftKey, adminMode, onItemClick, hideAdvance = false }) {
  // #19 (v3.10) — Último avance real de cada OT (filtra ruido ".", "..", vacías).
  // Una sola línea por OT: verde si la última entrada real es del turno actual,
  // gris/heredada si es de un turno anterior. Si todas son ruido → no se muestra.
  const withAdvance = items.map(c => {
    const adv = lastRealAdvance(c.timeline);
    return { ...c, _lastAdvance: adv, _advanceIsCurrent: !!adv && adv.shiftKey === currentShiftKey };
  });

  // V2.5 — Si showAvanceMark, ordenar items con avance del turno actual arriba.
  // Mantiene el orden relativo dentro de cada grupo (con y sin avance del turno).
  const sortedItems = showAvanceMark
    ? [
        ...withAdvance.filter(c => c._advanceIsCurrent),
        ...withAdvance.filter(c => !c._advanceIsCurrent)
      ]
    : withAdvance;

  return (
    <div className="mb-3 last:mb-0">
      <div className="text-[9px] uppercase tracking-wider text-slate-500 font-bold mb-1.5 pb-0.5 border-b border-slate-100">
        {title} ({count})
      </div>
      {count === 0 ? (
        <div className="text-[10px] text-slate-400 italic py-1">Sin novedades</div>
      ) : (
        <div className="divide-y-2 divide-slate-200">
          {sortedItems.map((c, i) => {
            const adv = c._lastAdvance;
            const advIsCurrent = c._advanceIsCurrent;
            const clickable = adminMode && onItemClick;
            return (
              <div key={i}
                className={`py-2 first:pt-0 last:pb-0 ${clickable ? 'cursor-pointer hover:bg-sky-50/60 -mx-1 px-1 rounded transition' : ''}`}
                onClick={clickable ? () => onItemClick(c) : undefined}
                title={clickable ? 'Click para editar en Cargar Reporte' : undefined}>
                <div className="flex items-center justify-between gap-2 mb-1">
                  <div className="flex items-center gap-2 min-w-0 flex-wrap">
                    <span className="num text-[11px] font-bold text-slate-800 whitespace-nowrap">{c.ot || '—'}</span>
                    {c.equipoCodigo && (
                      <span className="text-[10px] text-slate-500 truncate">· {c.equipoCodigo}</span>
                    )}
                  </div>
                  {showStateBadge && <StatePill state={c.state} />}
                </div>
                <div className="text-[12px] text-slate-700 leading-snug whitespace-pre-wrap break-words">{c.task || '—'}</div>
                {!hideAdvance && adv && (
                  advIsCurrent ? (
                    <div className="mt-1 text-[11px] text-emerald-800 bg-emerald-50/60 border-l-2 border-emerald-300 pl-2 py-0.5 leading-snug whitespace-pre-wrap break-words">
                      <span className="text-emerald-600 font-semibold mr-1">↳ Avance del turno:</span>
                      {adv.text || '—'}
                    </div>
                  ) : (
                    <div className="mt-1 text-[11px] text-slate-600 bg-slate-50 border-l-2 border-slate-300 pl-2 py-0.5 leading-snug whitespace-pre-wrap break-words">
                      <span className="text-slate-500 font-semibold mr-1">↳ Último avance ({formatDateShort(adv.date)} · {adv.shift}):</span>
                      {adv.text || '—'}
                    </div>
                  )
                )}
                {(c.technicians || []).length > 0 && (
                  <div className="mt-1 flex flex-wrap gap-1">
                    {c.technicians.map(t => (
                      <span key={t} className="text-[9px] px-1.5 py-0.5 bg-slate-100 text-slate-600 rounded">{t}</span>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// DASHBOARD VIEW — v3.18 (BACKLOG #36)
//   Rediseño compacto para export por mail (aplica a pantalla y export):
//   - Comentarios URGENTES en banner rojo arriba (solo si los hay)
//   - Correctivos a todo el ancho con contadores Realizadas/Pendientes,
//     y debajo DOS columnas en paralelo: Realizadas | Pendientes
//     (cada una junta las del turno + las heredadas)
//   - En Pendientes se oculta la línea de avance (#19) para ahorrar alto:
//     solo N° OT + equipo + estado + descripción (hideAdvance)
//   - Preventivos · Servicios · Proveedores en una fila de 3 columnas abajo
//   - Servicios: estado planta + caudal m³/h + cisternas + cloro pozo 3
//     + compresores/grupos como chips (No Operativo resaltado)
//   - Comentarios NO urgentes en lista compacta al final
//   Mantiene: selector de turnos guardados + export PNG/PDF (fuera del área capturada),
//   filtro de OTs del turno (V2.3/V2.4), particiones V2.5, click-to-edit admin (V2.6).
// ═══════════════════════════════════════════════════════════════════
function DashboardView({ report, history = [], activeReport, dashboardOverride, setDashboardOverride, adminMode, onEditFromDashboard }) {
  const dateLabel = useMemo(() => formatDateLong(report.date), [report.date]);
  const dateShort = useMemo(() => formatDateShort(report.date), [report.date]);
  const p = report.servicios.plantaCaldera;
  const pr = report.preventivosResumen || { asignados: '', realizados: '', porTecnico: [] };

  // PR-1 — Dashboard móvil (viewport < 768px): oculta export y reordena secciones.
  const isMobile = useMediaQuery('(max-width: 767px)');

  // V2.4 — Listado de turnos guardados disponibles para el selector.
  const savedShifts = useMemo(() => {
    return [...history]
      .filter(r => r.date && r.shift)
      .sort((a, b) => {
        const ka = `${a.date}-${shiftOrder(a.shift)}`;
        const kb = `${b.date}-${shiftOrder(b.shift)}`;
        return kb.localeCompare(ka); // desc
      })
      .map(r => ({
        key: `${r.date}|${r.shift}`,
        date: r.date,
        shift: r.shift,
        report: r,
        label: `${formatDateShort(r.date)} · ${r.shift}`,
        responsable: r.responsable || ''
      }));
  }, [history]);

  const isViewerMode = !!dashboardOverride;
  const currentKey = `${report.date}|${report.shift}`;
  const currentIdx = savedShifts.findIndex(s => s.key === currentKey);

  const goToPrev = () => {
    if (currentIdx < 0) return;
    const next = savedShifts[currentIdx + 1];
    if (next) setDashboardOverride(next.report);
  };
  const goToNext = () => {
    if (currentIdx < 0) return;
    const prev = savedShifts[currentIdx - 1];
    if (prev) setDashboardOverride(prev.report);
  };
  const goToCurrent = () => setDashboardOverride(null);
  const goToShift = (key) => {
    if (!key) return;
    const found = savedShifts.find(s => s.key === key);
    if (found) setDashboardOverride(found.report);
  };

  // V2.4 — Filtro de OTs (opción B): reporte guardado → todas; nuevo → solo del turno.
  const currentShiftKey = `${report.date}-${report.shift}`;
  const isExistingReport = useMemo(
    () => history.some(r => r.date === report.date && r.shift === report.shift),
    [history, report.date, report.shift]
  );
  const correctiveActual = useMemo(() => {
    if (isExistingReport) {
      return report.corrective || [];
    }
    return (report.corrective || []).filter(c => {
      if (c.createdInShift === currentShiftKey) return true;
      if (c.lastModifiedInShift === currentShiftKey) return true;
      return false;
    });
  }, [report.corrective, currentShiftKey, isExistingReport]);

  // V2.5 — Particiones de correctivos en 4 categorías.
  const correctivePartitions = useMemo(() => {
    const isCreatedHere = (c) => c.createdInShift === currentShiftKey;
    const isDone = (c) => c.state === 'Realizada';
    const findCurrentShiftEntry = (c) => {
      const tl = c.timeline || [];
      const entries = tl.filter(e => e.shiftKey === currentShiftKey);
      return entries.length > 0 ? entries[entries.length - 1] : null;
    };
    const enriched = correctiveActual.map(c => ({
      ...c,
      _createdHere: isCreatedHere(c),
      _currentShiftEntry: findCurrentShiftEntry(c)
    }));
    return {
      realizadosTurno:     enriched.filter(c => isDone(c) && c._createdHere),
      realizadosHeredados: enriched.filter(c => isDone(c) && !c._createdHere),
      pendientesTurno:     enriched.filter(c => !isDone(c) && c._createdHere),
      pendientesHeredados: enriched.filter(c => !isDone(c) && !c._createdHere)
    };
  }, [correctiveActual, currentShiftKey]);

  // #36 (v3.18) — columna de Realizadas: junta las del turno + las heredadas cerradas acá.
  // (v3.21 — la contraparte `pendientes` se eliminó: ahora se muestran separadas, ver abajo.)
  const realizadas = [...correctivePartitions.realizadosTurno, ...correctivePartitions.realizadosHeredados];

  // v3.21 — Contadores separados: pendientes DEL TURNO vs. HEREDADAS (y estas
  // últimas desagregadas por estado). El criterio turno/heredada es el mismo que
  // usa correctivePartitions: `createdInShift === currentShiftKey` (OT creada en
  // este turno). Una heredada cerrada en este turno cuenta como Realizada, no acá.
  // ESTADOS_OT solo tiene 3 valores, así que toda pendiente heredada cae en
  // "En Curso" o "Sin Iniciar": las dos cajas siempre suman el total de heredadas.
  const pendientesTurno = correctivePartitions.pendientesTurno;
  const pendientesHeredadas = correctivePartitions.pendientesHeredados;
  const heredadasEnCurso = pendientesHeredadas.filter(c => c.state === 'En Curso');
  const heredadasSinIniciar = pendientesHeredadas.filter(c => c.state === 'Sin Iniciar');

  // #36 (v3.18) — separar comentarios urgentes (banner arriba) de normales (lista abajo)
  const urgentComments = (report.comments || []).filter(c => c.priority === 'Urgente' && (c.text || '').trim());
  const normalComments = (report.comments || []).filter(c => c.priority !== 'Urgente' && (c.text || '').trim());

  // V2.2 — Export del Dashboard a PNG/PDF
  const dashboardRef = useRef(null);
  const [exporting, setExporting] = useState('');
  const [exportMsg, setExportMsg] = useState('');

  const loadScript = (src) => new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) return resolve();
    const s = document.createElement('script');
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error(`No se pudo cargar ${src}`));
    document.head.appendChild(s);
  });

  const captureDashboard = async () => {
    if (!window.html2canvas) {
      await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
    }
    const node = dashboardRef.current;
    if (!node) throw new Error('Dashboard no encontrado');

    const scrollables = node.querySelectorAll('*');
    const originalStyles = [];
    scrollables.forEach(el => {
      const cs = window.getComputedStyle(el);
      const needs = cs.overflowY === 'auto' || cs.overflowY === 'scroll' ||
                    cs.overflow === 'auto' || cs.overflow === 'scroll' ||
                    cs.maxHeight !== 'none';
      if (needs) {
        originalStyles.push({
          el,
          overflow: el.style.overflow,
          overflowY: el.style.overflowY,
          maxHeight: el.style.maxHeight,
          height: el.style.height
        });
        el.style.overflow = 'visible';
        el.style.overflowY = 'visible';
        el.style.maxHeight = 'none';
        el.style.height = 'auto';
      }
    });

    const truncatedEls = node.querySelectorAll('.truncate');
    const originalTextStyles = [];
    truncatedEls.forEach(el => {
      originalTextStyles.push({
        el,
        whiteSpace: el.style.whiteSpace,
        overflow: el.style.overflow,
        textOverflow: el.style.textOverflow,
        wordBreak: el.style.wordBreak
      });
      el.style.whiteSpace = 'normal';
      el.style.overflow = 'visible';
      el.style.textOverflow = 'clip';
      el.style.wordBreak = 'break-word';
    });

    await new Promise(r => requestAnimationFrame(r));
    await new Promise(r => setTimeout(r, 50));

    let canvas;
    try {
      canvas = await window.html2canvas(node, {
        backgroundColor: '#f8fafc',
        scale: 2,
        useCORS: true,
        logging: false,
        windowWidth: node.scrollWidth,
        windowHeight: node.scrollHeight
      });
    } finally {
      originalStyles.forEach(({ el, overflow, overflowY, maxHeight, height }) => {
        el.style.overflow = overflow;
        el.style.overflowY = overflowY;
        el.style.maxHeight = maxHeight;
        el.style.height = height;
      });
      originalTextStyles.forEach(({ el, whiteSpace, overflow, textOverflow, wordBreak }) => {
        el.style.whiteSpace = whiteSpace;
        el.style.overflow = overflow;
        el.style.textOverflow = textOverflow;
        el.style.wordBreak = wordBreak;
      });
    }
    return canvas;
  };

  const fileBaseName = () => {
    const d = report.date || 'sin-fecha';
    const t = report.shift || 'sin-turno';
    return `Dashboard_${d}_${t}`;
  };

  const exportPNG = async () => {
    if (exporting) return;
    setExporting('png');
    setExportMsg('Generando imagen…');
    try {
      const canvas = await captureDashboard();
      const link = document.createElement('a');
      link.download = `${fileBaseName()}.png`;
      link.href = canvas.toDataURL('image/png');
      link.click();
      setExportMsg('✓ Imagen descargada');
      setTimeout(() => setExportMsg(''), 2500);
    } catch (e) {
      console.error(e);
      setExportMsg(`Error: ${e.message}`);
      setTimeout(() => setExportMsg(''), 4000);
    }
    setExporting('');
  };

  const exportPDF = async () => {
    if (exporting) return;
    setExporting('pdf');
    setExportMsg('Generando PDF…');
    try {
      if (!window.jspdf) {
        await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
      }
      const canvas = await captureDashboard();
      const imgData = canvas.toDataURL('image/png');

      const pdf = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const margin = 8;
      const availW = pageW - margin * 2;
      const availH = pageH - margin * 2;

      const imgAspect = canvas.width / canvas.height;
      const availAspect = availW / availH;
      let drawW, drawH;
      if (imgAspect > availAspect) {
        drawW = availW;
        drawH = drawW / imgAspect;
      } else {
        drawH = availH;
        drawW = drawH * imgAspect;
      }
      const offsetX = margin + (availW - drawW) / 2;
      const offsetY = margin + (availH - drawH) / 2;

      pdf.addImage(imgData, 'PNG', offsetX, offsetY, drawW, drawH, undefined, 'FAST');
      pdf.save(`${fileBaseName()}.pdf`);
      setExportMsg('✓ PDF descargado');
      setTimeout(() => setExportMsg(''), 2500);
    } catch (e) {
      console.error(e);
      setExportMsg(`Error: ${e.message}`);
      setTimeout(() => setExportMsg(''), 4000);
    }
    setExporting('');
  };

  return (
    <div className="space-y-3">
      {/* V2.4 — Selector de turno guardado + export (FUERA del área capturada) */}
      <div className={`rounded-xl border p-3 ${isViewerMode ? 'bg-amber-50 border-amber-200' : 'bg-white border-slate-200'}`}>
        <div className="flex items-center justify-between gap-3 flex-wrap">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="text-xs uppercase tracking-wide text-slate-500 font-semibold inline-flex items-center gap-1.5">
              <Calendar className="w-3.5 h-3.5" />
              {isViewerMode ? 'Viendo turno guardado · solo lectura' : 'Turno actual'}
            </span>
            <button
              onClick={goToPrev}
              disabled={currentIdx < 0 || currentIdx >= savedShifts.length - 1}
              className="px-2 py-1 text-xs bg-white border border-slate-300 rounded hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed"
              title="Turno anterior (cronológico)"
            >
              ◄
            </button>
            <select
              value={isViewerMode ? currentKey : ''}
              onChange={e => {
                if (e.target.value === '') {
                  goToCurrent();
                } else {
                  goToShift(e.target.value);
                }
              }}
              className="px-2 py-1 text-xs bg-white border border-slate-300 rounded num"
              style={{ minWidth: '180px' }}
            >
              <option value="">— Turno actual —</option>
              {savedShifts.map(s => (
                <option key={s.key} value={s.key}>
                  {s.label}{s.responsable ? ` · ${s.responsable}` : ''}
                </option>
              ))}
            </select>
            <button
              onClick={goToNext}
              disabled={currentIdx <= 0}
              className="px-2 py-1 text-xs bg-white border border-slate-300 rounded hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed"
              title="Turno siguiente (cronológico)"
            >
              ►
            </button>
            {isViewerMode && (
              <button
                onClick={goToCurrent}
                className={`${buttonCls} bg-slate-700 text-white hover:bg-slate-800 text-xs px-2 py-1`}
              >
                Ver turno actual
              </button>
            )}
          </div>

          {/* PR-1 — Export oculto en móvil (viewport). La navegación de turnos queda visible. */}
          <div className="hidden md:flex items-center gap-2">
            {exportMsg && (
              <span className={`text-xs font-medium ${exportMsg.startsWith('Error') ? 'text-red-600' : exportMsg.startsWith('✓') ? 'text-emerald-600' : 'text-slate-500'}`}>
                {exportMsg}
              </span>
            )}
            <button onClick={exportPNG} disabled={!!exporting}
              className={`${buttonCls} bg-sky-50 text-sky-700 hover:bg-sky-100 disabled:opacity-50 disabled:cursor-not-allowed`}>
              <ImageIcon className="w-4 h-4" />
              {exporting === 'png' ? 'Generando…' : 'Exportar PNG'}
            </button>
            <button onClick={exportPDF} disabled={!!exporting}
              className={`${buttonCls} bg-rose-50 text-rose-700 hover:bg-rose-100 disabled:opacity-50 disabled:cursor-not-allowed`}>
              <FileDown className="w-4 h-4" />
              {exporting === 'pdf' ? 'Generando…' : 'Exportar PDF'}
            </button>
          </div>
        </div>
      </div>

      {/* V2.6 — Banner de modo admin */}
      {adminMode && (
        <div className="flex items-center gap-2 px-3 py-2 bg-sky-50 border border-sky-200 rounded-lg text-[12px] text-sky-800">
          <Shield className="w-4 h-4 text-sky-600 flex-shrink-0" />
          <div>
            Modo administrador activo. Hacé click en cualquier correctivo o preventivo para editarlo en "Cargar Reporte".
          </div>
        </div>
      )}

      {/* Wrapper capturado por html2canvas */}
      <div ref={dashboardRef} className="space-y-3">

        {/* HEADER — fecha, turno, responsable + equipo completo en chips */}
        <Card className="p-3">
          <div className="flex items-center justify-between gap-3 flex-wrap mb-2">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-full bg-sky-100 ring-1 ring-sky-200 flex items-center justify-center flex-shrink-0">
                <Wrench className="w-5 h-5 text-sky-600" />
              </div>
              <div>
                <div className="text-[10px] uppercase tracking-widest text-slate-500 font-semibold">Reporte Diario</div>
                <div className="text-sm font-bold text-slate-900 capitalize num">{dateLabel || '—'} · {report.shift}</div>
                <div className="text-[10px] text-slate-400 num">{dateShort}</div>
              </div>
            </div>
            <div className="text-sm">
              <span className="text-[10px] uppercase tracking-widest text-slate-500 font-semibold">Responsable </span>
              <span className="font-medium text-slate-800">{report.responsable || '—'}</span>
            </div>
          </div>
          <div className="border-t border-slate-100 pt-2">
            <div className="text-[10px] uppercase tracking-widest text-slate-500 font-semibold flex items-center gap-1 mb-1">
              <Users className="w-3 h-3" />Equipo (<span className="num">{report.team.length}</span>)
            </div>
            {report.team.length === 0 ? (
              <div className="text-xs text-slate-400 italic">—</div>
            ) : (
              <div className="flex flex-wrap gap-1">
                {report.team.map(t => (
                  <span key={t} className="text-[10px] px-1.5 py-0.5 bg-slate-100 text-slate-700 rounded font-medium">
                    {t}
                  </span>
                ))}
              </div>
            )}
          </div>
        </Card>

        {/* #36 — Banner de comentarios URGENTES (solo si los hay) */}
        {urgentComments.length > 0 && (
          <div className="flex items-start gap-2 px-3 py-2.5 bg-red-50 border border-red-200 rounded-xl">
            <AlertTriangle className="w-4 h-4 text-red-600 flex-shrink-0 mt-0.5" />
            <div className="text-[12px] text-red-700 leading-snug">
              <span className="font-bold">{urgentComments.length} comentario{urgentComments.length === 1 ? '' : 's'} urgente{urgentComments.length === 1 ? '' : 's'}:</span>
              <ul className="mt-1 space-y-0.5">
                {urgentComments.map((c, i) => (
                  <li key={i} className="leading-snug">• {c.text}</li>
                ))}
              </ul>
            </div>
          </div>
        )}

        {/* PR-1 — Secciones del resumen del turno.
            Desktop: Correctivos (ancho completo) + grid[Preventivos·Servicios·Proveedores] (idéntico a #36).
            Móvil (viewport < 768px): Servicios → Proveedores → Correctivos → Preventivos (apilados). */}
        {(() => {
          const correctivosCard = (
            <Card className="p-3">
              {/* v3.21 — Sin sumatoria en el título: el total mezclaba OTs del turno con
                  heredadas y se leía como "carga del turno", que no es. */}
              <h3 className="text-sky-600 font-bold text-sm mb-2 inline-flex items-center gap-2">
                <Wrench className="w-4 h-4" />Correctivos del turno
              </h3>

              {correctiveActual.length === 0 ? (
                <EmptyHint>Sin correctivos en este turno.</EmptyHint>
              ) : (
                // v3.21 — Dos columnas simétricas. Cada una: su(s) contador(es) arriba
                // + su detalle de OTs debajo.
                //   Izquierda: Realizadas (contador) + lista de realizadas (con "qué se hizo").
                //   Derecha:   Pendientes del turno (contador) + Heredadas en curso / sin iniciar
                //              (solo números, slate) + lista de pendientes del turno (con estado).
                //   Heredadas: solo número, sin detalle de OTs (decisión de diseño).
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-5 gap-y-3 items-start">
                  {/* COLUMNA IZQUIERDA — REALIZADAS */}
                  <div>
                    <div className="bg-emerald-50 rounded-lg p-2 text-center mb-3">
                      <div className="text-xl font-bold num text-emerald-700">{realizadas.length}</div>
                      <div className="text-[10px] uppercase tracking-wide text-emerald-600 font-semibold">Realizadas</div>
                    </div>
                    <CorrectiveSubsection
                      title="Realizadas"
                      count={realizadas.length}
                      items={realizadas}
                      showStateBadge={false}
                      showAvanceMark={false}
                      currentShiftKey={currentShiftKey}
                      adminMode={adminMode}
                      onItemClick={adminMode ? (c) => onEditFromDashboard(report, `ot:${c.ot || ""}`) : undefined}
                    />
                  </div>

                  {/* COLUMNA DERECHA — PENDIENTES DEL TURNO + HEREDADAS */}
                  <div>
                    {/* Contador ancho completo: Pendientes del turno */}
                    <div className="bg-amber-50 rounded-lg p-2 text-center mb-2">
                      <div className="text-xl font-bold num text-amber-700">{pendientesTurno.length}</div>
                      <div className="text-[10px] uppercase tracking-wide text-amber-600 font-semibold">Pendientes del turno</div>
                    </div>
                    {/* Dos cajas de heredadas (solo número, slate) */}
                    <div className="flex gap-2 mb-3">
                      <div className="flex-1 bg-slate-100 rounded-lg p-2 text-center">
                        <div className="text-xl font-bold num text-slate-700">{heredadasEnCurso.length}</div>
                        <div className="text-[10px] uppercase tracking-wide text-slate-500 font-semibold">Heredadas en curso</div>
                      </div>
                      <div className="flex-1 bg-slate-100 rounded-lg p-2 text-center">
                        <div className="text-xl font-bold num text-slate-700">{heredadasSinIniciar.length}</div>
                        <div className="text-[10px] uppercase tracking-wide text-slate-500 font-semibold">Heredadas sin iniciar</div>
                      </div>
                    </div>
                    {/* Detalle: solo las Pendientes del turno (heredadas quedan solo como número) */}
                    <CorrectiveSubsection
                      title="Pendientes del turno"
                      count={pendientesTurno.length}
                      items={pendientesTurno}
                      showStateBadge={true}
                      showAvanceMark={false}
                      hideAdvance={true}
                      currentShiftKey={currentShiftKey}
                      adminMode={adminMode}
                      onItemClick={adminMode ? (c) => onEditFromDashboard(report, `ot:${c.ot || ""}`) : undefined}
                    />
                  </div>
                </div>
              )}
            </Card>
          );

          const preventivosCard = (
            <Card
              className={`p-3 ${adminMode ? 'cursor-pointer hover:bg-sky-50/60 hover:ring-2 hover:ring-sky-200 transition' : ''}`}
              onClick={adminMode ? () => onEditFromDashboard(report, 'preventivos') : undefined}
              title={adminMode ? 'Click para editar preventivos del turno' : undefined}
            >
              <div className="flex items-center justify-between mb-2">
                <h3 className="text-sky-600 font-bold text-sm inline-flex items-center gap-2">
                  <ListChecks className="w-4 h-4" />Preventivos
                </h3>
                <div className="text-sm">
                  <span className="num font-bold text-emerald-700">{pr.realizados !== '' && pr.realizados != null ? pr.realizados : '—'}</span>
                  <span className="text-slate-400 text-xs"> / {pr.asignados !== '' && pr.asignados != null ? pr.asignados : '—'} asign.</span>
                </div>
              </div>
              {(() => {
                const grupos = (pr.porTecnico || []).filter(t => {
                  const tecnicos = t.tecnicos || (t.tecnico ? [t.tecnico] : []);
                  return tecnicos.length > 0;
                });
                if (grupos.length === 0) {
                  return <div className="text-[11px] text-slate-400 italic">Sin detalle por técnico</div>;
                }
                return (
                  <div className="flex flex-wrap gap-x-3 gap-y-1 text-[11px] text-slate-600">
                    {grupos.map((t, i) => {
                      const tecnicos = t.tecnicos || (t.tecnico ? [t.tecnico] : []);
                      return (
                        <span key={i}>
                          {tecnicos.join(' · ')} <span className="num font-bold text-slate-800">{t.cantidad || 0}</span>
                        </span>
                      );
                    })}
                  </div>
                );
              })()}
            </Card>
          );

          const serviciosCard = (
            <Card className="p-3">
              <h3 className="text-sky-600 font-bold text-sm mb-2 inline-flex items-center gap-2">
                <Activity className="w-4 h-4" />Servicios
              </h3>
              <div className="grid grid-cols-2 gap-x-3 gap-y-1.5 text-[11px] mb-2">
                <div className="flex justify-between items-center">
                  <span className="text-slate-500 inline-flex items-center gap-1"><Flame className="w-3 h-3 text-orange-500" />Planta</span>
                  <StatePill state={p.estado} />
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-500">Caudal m³/h</span>
                  <span className="num font-bold text-slate-800">{p.caudal !== '' && p.caudal != null ? p.caudal : '—'}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-500 inline-flex items-center gap-1"><Beaker className="w-3 h-3 text-cyan-500" />Cisternas</span>
                  <StatePill state={report.servicios.cisternas.estado} />
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-500">Cloro pozo 3</span>
                  <span className="num font-bold text-slate-800">{report.servicios.aguaPozo?.cloroPozo3 !== '' && report.servicios.aguaPozo?.cloroPozo3 != null ? report.servicios.aguaPozo.cloroPozo3 : '—'}</span>
                </div>
              </div>
              <div className="border-t border-slate-100 pt-2">
                <div className="text-[9px] uppercase tracking-wide text-slate-500 font-semibold mb-1 inline-flex items-center gap-1">
                  <Cog className="w-3 h-3" />Compresores · <Zap className="w-3 h-3" />G. Electrógenos
                </div>
                <div className="flex flex-wrap gap-1">
                  {[...(report.servicios.compresores || []), ...(report.servicios.gruposElectrogenos || [])].map(x => (
                    <span key={x.code}
                      className={`text-[10px] px-1.5 py-0.5 rounded num ${x.state === 'Operativo' ? 'bg-slate-100 text-slate-600' : 'bg-red-100 text-red-700 font-semibold'}`}>
                      {x.code} {x.state === 'Operativo' ? '✓' : '✕'}
                    </span>
                  ))}
                </div>
              </div>
            </Card>
          );

          const proveedoresCard = (
            <Card className="p-3">
              <h3 className="text-sky-600 font-bold text-sm mb-2 inline-flex items-center gap-2">
                <Building2 className="w-4 h-4" />Proveedores
              </h3>
              {(report.servicios.proveedores || []).length === 0 ? (
                <div className="text-[11px] text-slate-400 italic">Sin proveedores en el turno</div>
              ) : (
                <div className="space-y-1">
                  {report.servicios.proveedores.map((pv, i) => (
                    <div key={i} className="text-[11px] text-slate-600 leading-snug">
                      <span className="font-medium text-slate-700">{pv.provider}</span>{pv.task ? <span> · {pv.task}</span> : null}
                    </div>
                  ))}
                </div>
              )}
            </Card>
          );

          const comentariosCard = normalComments.length > 0 ? (
            <Card className="p-3">
              <div className="text-[10px] uppercase tracking-wide text-slate-500 font-semibold mb-1.5 inline-flex items-center gap-1">
                <FileText className="w-3 h-3" />Comentarios
              </div>
              <div className="space-y-1">
                {normalComments.map((c, i) => (
                  <div key={i} className="text-[11px] text-slate-700 bg-slate-50 rounded px-2 py-1 leading-snug">
                    {c.text}
                  </div>
                ))}
              </div>
            </Card>
          ) : null;

          if (isMobile) {
            // Móvil (viewport < 768px): intacto respecto de PR-1 —
            // Servicios → Proveedores → Correctivos → Preventivos.
            // Comentarios normales quedan al final (fuera de este bloque, ver abajo).
            return (
              <>
                {serviciosCard}
                {proveedoresCard}
                {correctivosCard}
                {preventivosCard}
                {comentariosCard}
              </>
            );
          }

          // v3.21 — Desktop: primero el contexto del turno (Preventivos · Servicios ·
          // Proveedores + Comentarios), y ABAJO de todo los Correctivos (que es la
          // sección más alta). Antes Correctivos iba arriba.
          return (
            <>
              {/* FILA SUPERIOR — Preventivos · Servicios · Proveedores (3 columnas) */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-3">
                {preventivosCard}
                {serviciosCard}
                {proveedoresCard}
              </div>
              {comentariosCard}
              {correctivosCard}
            </>
          );
        })()}

      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// STATS VIEW — V2.0
//   - Apartado "Fin de Semana" cerrado (viernes Noche + sábado + domingo)
//   - Apartado "Último Día" (últimos 3 turnos del día más reciente con datos)
// ═══════════════════════════════════════════════════════════════════
function StatsView({ history, adminMode }) {
  const [range, setRange] = useState('week');
  const [customStart, setCustomStart] = useState('');
  const [customEnd, setCustomEnd] = useState('');

  const stats = useMemo(() => computeStats(history, range, customStart, customEnd), [history, range, customStart, customEnd]);
  const finde = useMemo(() => computeWeekendStats(history), [history]);
  const ultimoDia = useMemo(() => computeLastDayStats(history), [history]);
  // V2.7 — Estadísticas admin: OTs pendientes por turno de origen + OTs heredadas cerradas por turno.
  // Solo se calcula cuando adminMode === true. Usa el mismo rango (startStr/endStr) de computeStats.
  const shiftPerformance = useMemo(
    () => adminMode ? computeShiftPerformance(history, stats.startStr, stats.endStr) : null,
    [history, stats.startStr, stats.endStr, adminMode]
  );

  if (history.length === 0) {
    return (
      <Card className="p-12 text-center">
        <BarChart3 className="w-12 h-12 mx-auto text-slate-300 mb-3" />
        <div className="text-slate-500">Sin reportes guardados todavía. Cargá uno desde "Cargar Reporte".</div>
      </Card>
    );
  }

  const RANGES = [
    { id: 'week', label: 'Semana' },
    { id: 'month', label: 'Mes' },
    { id: 'quarter', label: 'Trimestre' },
    { id: 'semester', label: 'Semestre' },
    { id: 'year', label: 'Año' },
    { id: 'custom', label: 'Personalizado' }
  ];

  return (
    <div className="space-y-4">
      {/* APARTADOS RÁPIDOS — V2.0 */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        {/* ÚLTIMO DÍA */}
        <Card className="p-4">
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-sm font-bold text-slate-800 inline-flex items-center gap-2">
              <Clock className="w-4 h-4 text-indigo-500" />Último Día
            </h3>
            <span className="text-[10px] text-slate-500">
              {ultimoDia.turnos.length === 0 ? 'sin datos' :
                `${ultimoDia.turnos.length} de 3 turnos · ${formatDateShort(ultimoDia.fechaBase)}`}
            </span>
          </div>
          {ultimoDia.turnos.length === 0 ? (
            <EmptyHint>
              {ultimoDia.sinDiaAnterior
                ? 'Todavía no hay un día cerrado anterior a hoy'
                : 'Sin reportes recientes'}
            </EmptyHint>
          ) : (
            <>
              {/* v3.24 — La completitud se muestra siempre: un día al que le falta
                  un turno da KPIs que parecen bajos sin que nada lo indique. */}
              {ultimoDia.turnosFaltantes.length > 0 ? (
                <div className="text-[11px] text-amber-700 mb-2 inline-flex items-center gap-1.5">
                  <AlertTriangle className="w-3 h-3 flex-shrink-0" />
                  Día incompleto — falta {ultimoDia.turnosFaltantes.join(', ')}
                </div>
              ) : (
                <div className="text-[11px] text-slate-500 mb-2">
                  Día completo · {ultimoDia.turnos.map(t => t.shift).join(' · ')}
                </div>
              )}
              <div className="grid grid-cols-2 gap-2">
                <MiniKPI label="Correctivos generados" value={ultimoDia.correctivosGenerados} color="orange" />
                <MiniKPI label="Correctivos realizados" value={ultimoDia.correctivosRealizados} color="emerald" />
                <MiniKPI label="Preventivos asignados" value={ultimoDia.preventivosAsignados} color="sky" />
                <MiniKPI label="Preventivos realizados" value={ultimoDia.preventivosRealizados} color="emerald" />
              </div>
              <p className="text-[10px] text-slate-400 mt-2 leading-relaxed">
                Generados = pedidos ese día. Realizados = cerrados ese día, incluidas heredadas.
              </p>
            </>
          )}
        </Card>

        {/* FIN DE SEMANA */}
        <Card className="p-4">
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-sm font-bold text-slate-800 inline-flex items-center gap-2">
              <CalendarDays className="w-4 h-4 text-violet-500" />Fin de Semana
            </h3>
            <span className="text-[10px] text-slate-500">
              {finde.turnos.length === 0 ? 'sin datos' :
                `${finde.turnos.length} de 6 turnos · ${formatDateShort(finde.sabado)} → ${formatDateShort(finde.domingo)}`}
            </span>
          </div>
          {finde.turnos.length === 0 ? (
            <EmptyHint>Sin reportes del último FDS cerrado</EmptyHint>
          ) : (
            <>
              {/* v3.24 — 6 turnos: los etiquetados sábado y domingo.
                  Cobertura real viernes 23:00 → domingo 23:00. */}
              {finde.turnos.length < 6 ? (
                <div className="text-[11px] text-amber-700 mb-2 inline-flex items-center gap-1.5">
                  <AlertTriangle className="w-3 h-3 flex-shrink-0" />
                  FDS incompleto — {finde.turnos.length} de 6 turnos cargados
                </div>
              ) : (
                <div className="text-[11px] text-slate-500 mb-2">
                  FDS completo · viernes 23:00 → domingo 23:00
                </div>
              )}
              <div className="grid grid-cols-2 gap-2">
                <MiniKPI label="Correctivos generados" value={finde.correctivosGenerados} color="orange" />
                <MiniKPI label="Correctivos realizados" value={finde.correctivosRealizados} color="emerald" />
                <MiniKPI label="Preventivos asignados" value={finde.preventivosAsignados} color="sky" />
                <MiniKPI label="Preventivos realizados" value={finde.preventivosRealizados} color="emerald" />
              </div>
              <p className="text-[10px] text-slate-400 mt-2 leading-relaxed">
                Generados = pedidos en el FDS. Realizados = cerrados en el FDS, incluidas heredadas.
              </p>
            </>
          )}
        </Card>
      </div>

      {/* RANGE SELECTOR */}
      <Card className="p-3">
        <div className="flex items-center justify-between flex-wrap gap-3">
          <div className="bg-slate-100 rounded-lg p-1 inline-flex flex-wrap gap-1">
            {RANGES.map(opt => (
              <button key={opt.id} onClick={() => setRange(opt.id)}
                className={`px-3 py-1 text-xs font-medium rounded ${range === opt.id ? 'bg-slate-800 text-white' : 'text-slate-600 hover:bg-slate-200'}`}>
                {opt.label}
              </button>
            ))}
          </div>
          {range === 'custom' && (
            <div className="flex items-center gap-2 text-sm">
              <input type="date" className={`${inputCls} text-xs py-1 w-auto`} value={customStart} onChange={e => setCustomStart(e.target.value)} />
              <span className="text-slate-400">→</span>
              <input type="date" className={`${inputCls} text-xs py-1 w-auto`} value={customEnd} onChange={e => setCustomEnd(e.target.value)} />
            </div>
          )}
          <div className="text-xs text-slate-500">
            Período: <span className="font-medium num">{formatDateShort(stats.startStr)}</span> → <span className="font-medium num">{formatDateShort(stats.endStr)}</span>
          </div>
        </div>
      </Card>

      {/* KPI ROW */}
      <div className="grid grid-cols-3 md:grid-cols-7 gap-2">
        <KPI label="Reportes" value={stats.totalReports} icon={ClipboardList} color="sky" />
        <KPI label="Correctivos" value={stats.totalCorrectives} icon={Wrench} color="orange" />
        <KPI label="Preventivos" value={stats.totalPreventives} icon={ListChecks} color="emerald" />
        <KPI label="Realizados" value={stats.completedCorr} icon={CheckCircle2} color="green" />
        <KPI label="Pendientes" value={stats.pendingCorr} icon={AlertTriangle} color="amber" />
        <KPI label="Urgentes" value={stats.urgent} icon={AlertTriangle} color="red" />
        <KPI label="% Cumpl." value={`${stats.completionRate}%`} icon={TrendingUp} color="indigo" />
      </div>

           {/* CHARTS GRID — orden v3.2:
          Fila 1: Trabajos en el período | Evolución temporal
          Fila 2: Estado al cierre del período | Estado al día de hoy
          Fila 3: OTs pendientes por origen | OTs heredadas cerradas   (admin)
          Fila 4: Carga por técnico | Distribución por turno           (admin)
          Fila 5: Equipos con más correctivos
          Gráfico "estado al cierre" pasó a público (BACKLOG #17 resuelto). */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Card className="p-4">
          <h3 className="text-sm font-semibold text-slate-700 mb-1 inline-flex items-center gap-2">
            <BarChart3 className="w-4 h-4" />Trabajos en el período
          </h3>
          <p className="text-[11px] text-slate-400 mb-3">Correctivos: OTs pasadas a Realizada cada día · Preventivos: realizados por turno</p>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={stats.daily}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="label" stroke="#64748b" style={{ fontSize: '10px' }} />
              <YAxis stroke="#64748b" style={{ fontSize: '10px' }} />
              <Tooltip contentStyle={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: '8px', fontSize: '12px' }} />
              <Legend wrapperStyle={{ fontSize: '11px' }} />
              <Bar dataKey="correctivos" fill="#f97316" name="Correctivos" radius={[3, 3, 0, 0]} />
              <Bar dataKey="preventivos" fill="#10b981" name="Preventivos" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>

        {/* Evolución temporal */}
        <Card className="p-4">
          <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
            <TrendingUp className="w-4 h-4" />Evolución temporal
          </h3>
          <ResponsiveContainer width="100%" height={240}>
            <AreaChart data={stats.daily}>
              <defs>
                <linearGradient id="colorC" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#f97316" stopOpacity={0.7} />
                  <stop offset="95%" stopColor="#f97316" stopOpacity={0} />
                </linearGradient>
                <linearGradient id="colorP" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#10b981" stopOpacity={0.7} />
                  <stop offset="95%" stopColor="#10b981" stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="label" stroke="#64748b" style={{ fontSize: '10px' }} />
              <YAxis stroke="#64748b" style={{ fontSize: '10px' }} />
              <Tooltip contentStyle={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: '8px', fontSize: '12px' }} />
              <Legend wrapperStyle={{ fontSize: '11px' }} />
              <Area type="monotone" dataKey="correctivos" stroke="#f97316" fillOpacity={1} fill="url(#colorC)" name="Correctivos" />
              <Area type="monotone" dataKey="preventivos" stroke="#10b981" fillOpacity={1} fill="url(#colorP)" name="Preventivos" />
            </AreaChart>
          </ResponsiveContainer>
        </Card>

        {/* GRÁFICO 2 — Estado al cierre del período filtrado.
            V3.2 — Público: las OTs huérfanas se limpiaron en Supabase (BACKLOG #17). */}
        <Card className="p-4">
          <h3 className="text-sm font-semibold text-slate-700 mb-1 inline-flex items-center gap-2">
            <Activity className="w-4 h-4" />Correctivos: estado al cierre del período
          </h3>
          <p className="text-[11px] text-slate-400 mb-3">
            Cómo quedaron las OTs al final del período seleccionado ({stats.pendientesPeriodo} pendientes al cierre).
          </p>
          {stats.stateDistPeriodo.length === 0 ? <EmptyHint>Sin datos</EmptyHint> :
            <ResponsiveContainer width="100%" height={240}>
              <PieChart>
                <Pie data={stats.stateDistPeriodo} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80}>
                  {stats.stateDistPeriodo.map((entry, i) => (
                    <Cell key={i} fill={entry.name === 'Realizada' ? '#10b981' : entry.name === 'En Curso' ? '#f59e0b' : '#ef4444'} />
                  ))}
                </Pie>
                <Tooltip />
                <Legend wrapperStyle={{ fontSize: '11px' }}
                  formatter={(value, entry) => `${value}: ${entry.payload.value}`} />
              </PieChart>
            </ResponsiveContainer>
          }
        </Card>

        {/* GRÁFICO 1 — Estado vigente HOY (todo el histórico) */}
        <Card className="p-4">
          <h3 className="text-sm font-semibold text-slate-700 mb-1 inline-flex items-center gap-2">
            <Activity className="w-4 h-4" />Correctivos: estado al día de hoy
          </h3>
          <p className="text-[11px] text-slate-400 mb-3">
            Pendientes abiertas al día de hoy ({stats.pendientesVigentes}). Incluye OTs anteriores al período. No cambia según el rango.
          </p>
          {stats.stateDistVigente.length === 0 ? <EmptyHint>Sin datos</EmptyHint> :
            <ResponsiveContainer width="100%" height={240}>
              <PieChart>
                <Pie data={stats.stateDistVigente} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80}>
                  {stats.stateDistVigente.map((entry, i) => (
                    <Cell key={i} fill={entry.name === 'Realizada' ? '#10b981' : entry.name === 'En Curso' ? '#f59e0b' : '#ef4444'} />
                  ))}
                </Pie>
                <Tooltip />
                <Legend wrapperStyle={{ fontSize: '11px' }}
                  formatter={(value, entry) => `${value}: ${entry.payload.value}`} />
              </PieChart>
            </ResponsiveContainer>
          }
        </Card>

        {/* V2.7 — Bloque admin-only: métricas de performance por turno.
            Orden v3.2: Pendientes por origen, Heredadas cerradas, Carga por técnico, Distribución por turno. */}
        {adminMode && (<>
        {/* V2.7 — OTs dejadas pendientes por turno de origen. */}
        <ShiftRankingCard
          title="OTs dejadas pendientes por turno de origen"
          tooltip="Cuenta OTs creadas en cada turno cuyo último estado global sigue en Sin Iniciar o En Curso. Filtro por fecha aplica al turno de creación. Excluye OTs legacy. Estadísticas desde 20/05/2026."
          data={shiftPerformance?.pendingByOriginShift}
          icon={AlertTriangle}
          colorBar="#ef4444"
        />

        {/* V2.7 — OTs heredadas cerradas por turno. */}
        <ShiftRankingCard
          title="OTs heredadas cerradas por turno"
          tooltip="Cuenta cierres (estado → Realizada) de OTs creadas en turnos anteriores. Filtro por fecha aplica al turno del cierre. Excluye OTs legacy. Estadísticas desde 20/05/2026."
          data={shiftPerformance?.closedByShift}
          icon={CheckCircle2}
          colorBar="#10b981"
        />

        <Card className="p-4">
          <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
            <Users className="w-4 h-4" />Carga por técnico
            <span className="text-[10px] text-slate-500 font-normal ml-2">(todos los técnicos del catálogo)</span>
            <span className="text-[10px] text-amber-600 font-normal ml-1">· (admin)</span>
          </h3>
          {stats.topTechs.length === 0 ? <EmptyHint>Sin datos</EmptyHint> :
            <ResponsiveContainer width="100%" height={Math.max(240, stats.topTechs.length * 22)}>
              <BarChart data={stats.topTechs} layout="vertical" margin={{ top: 5, right: 20, bottom: 5, left: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis type="number" stroke="#64748b" style={{ fontSize: '10px' }} allowDecimals={false} />
                <YAxis dataKey="name" type="category" stroke="#64748b" style={{ fontSize: '10px' }} width={120} interval={0} />
                <Tooltip contentStyle={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: '8px', fontSize: '12px' }} />
                <Legend wrapperStyle={{ fontSize: '11px' }} />
                <Bar dataKey="correctivos" fill="#f97316" name="Correctivos" stackId="a" radius={[0, 0, 0, 0]} />
                <Bar dataKey="preventivos" fill="#10b981" name="Preventivos" stackId="a" radius={[0, 3, 3, 0]} />
              </BarChart>
            </ResponsiveContainer>
          }
        </Card>

        <Card className="p-4">
          <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
            <Calendar className="w-4 h-4" />Distribución por turno
            <span className="text-[10px] text-amber-600 font-normal ml-2">(admin)</span>
          </h3>
          {stats.shiftDist.every(s => s.value === 0) ? <EmptyHint>Sin datos</EmptyHint> :
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={stats.shiftDist}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis dataKey="name" stroke="#64748b" style={{ fontSize: '11px' }} />
                <YAxis stroke="#64748b" style={{ fontSize: '10px' }} />
                <Tooltip contentStyle={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: '8px', fontSize: '12px' }} />
                <Legend wrapperStyle={{ fontSize: '11px' }} />
                <Bar dataKey="correctivos" fill="#f97316" name="Correctivos" />
                <Bar dataKey="preventivos" fill="#10b981" name="Preventivos" />
              </BarChart>
            </ResponsiveContainer>
          }
        </Card>
        </>)}

        {/* Equipos con más correctivos */}
        <Card className="p-4">
          <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
            <Wrench className="w-4 h-4" />Equipos con más correctivos
          </h3>
          {stats.topEquipment.length === 0 ? <EmptyHint>Sin datos</EmptyHint> :
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={stats.topEquipment} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis type="number" stroke="#64748b" style={{ fontSize: '10px' }} />
                <YAxis dataKey="name" type="category" stroke="#64748b" style={{ fontSize: '10px' }} width={110} />
                <Tooltip contentStyle={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: '8px', fontSize: '12px' }} />
                <Bar dataKey="count" fill="#0ea5e9" radius={[0, 3, 3, 0]} />
              </BarChart>
            </ResponsiveContainer>
          }
        </Card>
      </div>

    </div>
  );
}

const MiniKPI = ({ label, value, color }) => {
  const colors = {
    orange: 'bg-orange-50 text-orange-700 border-orange-200',
    emerald: 'bg-emerald-50 text-emerald-700 border-emerald-200',
    sky: 'bg-sky-50 text-sky-700 border-sky-200'
  };
  return (
    <div className={`rounded-lg p-3 border ${colors[color]}`}>
      <div className="text-[10px] uppercase tracking-wide font-semibold opacity-80">{label}</div>
      <div className="text-2xl font-bold num mt-0.5">{value}</div>
    </div>
  );
};

const KPI = ({ label, value, icon: Icon, color }) => {
  const colors = {
    sky: 'bg-sky-100 text-sky-600', orange: 'bg-orange-100 text-orange-600',
    emerald: 'bg-emerald-100 text-emerald-600', green: 'bg-green-100 text-green-600',
    red: 'bg-red-100 text-red-600', amber: 'bg-amber-100 text-amber-600',
    indigo: 'bg-indigo-100 text-indigo-600'
  };
  return (
    <div className="bg-white border border-slate-200 rounded-lg p-2.5 flex items-center gap-2.5">
      <div className={`w-8 h-8 rounded flex items-center justify-center flex-shrink-0 ${colors[color]}`}>
        <Icon className="w-4 h-4" />
      </div>
      <div className="min-w-0">
        <div className="text-[9px] text-slate-500 uppercase tracking-wide font-medium leading-tight truncate">{label}</div>
        <div className="text-base font-bold text-slate-800 num leading-tight">{value}</div>
      </div>
    </div>
  );
};

// ═══════════════════════════════════════════════════════════════════
// V2.7 — Card de ranking por turno (admin)
//   - Muestra ranking 1°/2°/3° arriba + barra horizontal abajo.
//   - data: { Mañana: number|null, Tarde: number|null, Noche: number|null }
//     null = "—" (turno sin datos relevantes en la ventana)
// ═══════════════════════════════════════════════════════════════════
function ShiftRankingCard({ title, tooltip, data, icon: Icon, colorBar }) {
  if (!data) return null;
  const SHIFTS = ['Mañana', 'Tarde', 'Noche'];
  // Ordenar para ranking: nulls al final, ties por orden M/T/N
  const ranked = SHIFTS
    .map(s => ({ shift: s, value: data[s] }))
    .sort((a, b) => {
      if (a.value === null && b.value === null) return 0;
      if (a.value === null) return 1;
      if (b.value === null) return -1;
      return b.value - a.value;
    });
  const allNull = ranked.every(r => r.value === null);
  const barData = SHIFTS.map(s => ({ name: s, value: data[s] === null ? 0 : data[s] }));
  return (
    <Card className="p-4">
      <h3 className="text-sm font-semibold text-slate-700 mb-1 inline-flex items-center gap-2">
        <Icon className="w-4 h-4" />{title}
        <span className="text-[10px] text-amber-600 font-normal ml-1">(admin)</span>
      </h3>
      <p className="text-[10px] text-slate-500 mb-3 leading-snug">{tooltip}</p>
      {allNull ? <EmptyHint>Sin datos en la ventana</EmptyHint> : (
        <>
          <div className="grid grid-cols-3 gap-2 mb-3">
            {ranked.map((r, i) => (
              <div key={r.shift} className={`rounded-lg p-2 border text-center ${
                i === 0 && r.value !== null && r.value > 0 ? 'bg-amber-50 border-amber-200' :
                i === 1 && r.value !== null && r.value > 0 ? 'bg-slate-50 border-slate-200' :
                'bg-white border-slate-200'
              }`}>
                <div className="text-[9px] uppercase tracking-wide text-slate-500 font-semibold">
                  {i === 0 ? '1°' : i === 1 ? '2°' : '3°'} · {r.shift}
                </div>
                <div className="text-xl font-bold num mt-0.5 text-slate-800">
                  {r.value === null ? '—' : r.value}
                </div>
              </div>
            ))}
          </div>
          <ResponsiveContainer width="100%" height={160}>
            <BarChart data={barData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="name" stroke="#64748b" style={{ fontSize: '11px' }} />
              <YAxis stroke="#64748b" style={{ fontSize: '10px' }} allowDecimals={false} />
              <Tooltip
                contentStyle={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: '8px', fontSize: '12px' }}
                formatter={(v, n, p) => {
                  const original = data[p.payload.name];
                  return [original === null ? '—' : original, 'OTs'];
                }}
              />
              <Bar dataKey="value" fill={colorBar} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </>
      )}
    </Card>
  );
}

// ═══════════════════════════════════════════════════════════════════
// STATS COMPUTATION HELPERS
// ═══════════════════════════════════════════════════════════════════

// V2.7 — Estadísticas admin de performance por turno.
//
// Devuelve dos métricas, cada una como { 'Mañana': n|null, 'Tarde': n|null, 'Noche': n|null }:
//
//   1) pendingByOriginShift — OTs creadas en cada turno cuyo último estado GLOBAL
//      sigue en Sin Iniciar o En Curso. Filtro por fecha aplica al turno de origen
//      (createdInShift). Excluye OTs legacy (sin formato XXX-YYYYY válido).
//
//   2) closedByShift — Cierres ocurridos (transición Sin Iniciar/En Curso → Realizada)
//      en turnos distintos al de creación, dentro del rango [startStr, endStr].
//      El cierre se atribuye al turno donde sucedió la transición (no al de origen).
//      Excluye OTs legacy.
//
// IMPORTANTE: "último estado global" se calcula sobre TODO el history, no solo el rango.
// El filtro por fecha solo recorta qué OTs entran al cálculo (por origen o por cierre).
//
// Si una OT no fue tocada por ningún turno en la ventana, su turno cuenta 0.
// Si NINGUNA OT entró al cálculo para un turno (ej: no hubo creaciones de ese turno en la ventana),
// ese turno devuelve null y se muestra como "—".
function computeShiftPerformance(history, startStr, endStr) {
  const SHIFTS = ['Mañana', 'Tarde', 'Noche'];
  // Orden cronológico: Noche es el primer turno de la jornada (arranca la víspera).
  // Coherente con shiftOrder global. NO cambia el orden visual de SHIFTS (presentación).
  const SHIFT_ORDER = { 'Noche': 0, 'Mañana': 1, 'Tarde': 2 };
  // Validar formato XXX-YYYYY (idéntico a isValidOT). Hay que duplicarlo acá porque
  // las funciones de computo viven fuera del componente que importa isValidOT.
  const isValid = (ot) => {
    if (!ot || typeof ot !== 'string') return false;
    const m = ot.trim().match(/^([A-Z0-9]+)-(\d{5})$/);
    if (!m) return false;
    return SECTORES_CODES.includes(m[1]);
  };

  // 1) Construir mapas globales: createdInShift y ÚLTIMO estado global por OT.
  //    También guardamos en qué reportes apareció cada OT (ordenado cronológicamente).
  const sortedHistory = [...history].sort((a, b) => {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return SHIFT_ORDER[a.shift] - SHIFT_ORDER[b.shift];
  });

  // Por OT: { createdInShift, lastState, occurrences: [{date, shift, state}] }
  const otMap = new Map();
  sortedHistory.forEach(r => {
    (r.corrective || []).forEach(c => {
      if (!isValid(c.ot)) return; // Excluye legacy
      const key = c.ot;
      let entry = otMap.get(key);
      if (!entry) {
        entry = {
          createdInShift: c.createdInShift || null, // shiftKey "YYYY-MM-DD-Turno"
          lastState: null,
          occurrences: []
        };
        otMap.set(key, entry);
      }
      // El último que se procese en orden cronológico será el lastState
      entry.lastState = c.state || null;
      entry.occurrences.push({ date: r.date, shift: r.shift, state: c.state || null });
    });
  });

  // Helper: parsear createdInShift "YYYY-MM-DD-Turno" → {date, shift}
  // Cuidado: el formato es DATE-SHIFT donde DATE = "YYYY-MM-DD" (3 partes separadas por "-"),
  // así que split('-').slice(0,3).join('-') da la fecha y .slice(3).join('-') daría el shift.
  // Pero "Mañana"/"Tarde"/"Noche" no tienen guiones, entonces basta con slice(0,3)+slice(3,4).
  const parseShiftKey = (sk) => {
    if (!sk || typeof sk !== 'string') return null;
    const parts = sk.split('-');
    if (parts.length < 4) return null;
    return { date: parts.slice(0, 3).join('-'), shift: parts[3] };
  };

  // 2) Métrica 1: pendientes por turno de origen.
  //    Inicializamos cada turno como null. Cuando una OT del rango entra, lo pasamos a 0.
  //    Si está pendiente, sumamos 1.
  const pending = { 'Mañana': null, 'Tarde': null, 'Noche': null };
  otMap.forEach((entry, ot) => {
    const origin = parseShiftKey(entry.createdInShift);
    if (!origin) return; // sin createdInShift → no podemos atribuir
    if (origin.date < KPI_CUTOFF_DATE) return; // V3.0 — cutoff retroactivo
    if (origin.date < startStr || origin.date > endStr) return;
    if (!SHIFTS.includes(origin.shift)) return;
    // Al menos una OT del turno entró → pasar de null a 0
    if (pending[origin.shift] === null) pending[origin.shift] = 0;
    if (entry.lastState === 'Sin Iniciar' || entry.lastState === 'En Curso') {
      pending[origin.shift] += 1;
    }
  });

  // 3) Métrica 2: cierres heredados por turno donde ocurrió el cierre.
  //    Una OT puede haber cerrado UNA SOLA VEZ (transición → Realizada).
  //    Detectamos esa transición buscando la primera ocurrencia donde state === 'Realizada'
  //    y el state previo era distinto de 'Realizada'.
  //    Solo cuenta si el turno del cierre ≠ turno de creación.
  const closed = { 'Mañana': null, 'Tarde': null, 'Noche': null };
  otMap.forEach((entry, ot) => {
    const origin = parseShiftKey(entry.createdInShift);
    if (!origin) return;
    if (!SHIFTS.includes(origin.shift)) return;
    let prevState = null;
    let closingOcc = null;
    for (const occ of entry.occurrences) {
      if (occ.state === 'Realizada' && prevState !== 'Realizada') {
        closingOcc = occ;
        break;
      }
      prevState = occ.state;
    }
    if (!closingOcc) return;
    // El turno del cierre tiene que ser distinto al de creación (heredada)
    if (closingOcc.date === origin.date && closingOcc.shift === origin.shift) return;
    if (closingOcc.date < KPI_CUTOFF_DATE) return; // V3.0 — cutoff retroactivo
    // El cierre tiene que estar dentro de la ventana
    if (closingOcc.date < startStr || closingOcc.date > endStr) return;
    if (!SHIFTS.includes(closingOcc.shift)) return;
    if (closed[closingOcc.shift] === null) closed[closingOcc.shift] = 0;
    closed[closingOcc.shift] += 1;
  });

  // 4) Para los turnos que tuvieron actividad (creaciones en la ventana), cierre
  //    debería arrancar al menos en 0 si hubo cierres reportables. Pero como
  //    closedByShift se cuenta independiente de creaciones, dejamos null si nadie
  //    cerró nada en ese turno en la ventana. Esto evita confusión con "0".

  return {
    pendingByOriginShift: pending,
    closedByShift: closed
  };
}

// Calcula el último FDS cerrado: viernes Noche + sábado completo + domingo completo,
// donde el domingo ya pasó en el día actual.
function computeWeekendStats(history) {
  if (history.length === 0) {
    return { viernes: '', sabado: '', domingo: '', turnos: [], correctivosGenerados: 0, correctivosRealizados: 0, preventivosAsignados: 0, preventivosRealizados: 0 };
  }
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dow = today.getDay(); // 0=Dom, 1=Lun, ..., 6=Sáb

  // Encontrar el último domingo que ya pasó completo (ayer o antes)
  // Si hoy es domingo, el FDS cerrado fue el domingo anterior (hace 7 días)
  // Si hoy es lunes, el FDS cerrado fue ayer (domingo)
  // Si hoy es martes, el FDS cerrado fue hace 2 días (domingo)
  // ...
  // Si hoy es sábado, el FDS cerrado fue hace 6 días (domingo)
  let daysToLastSunday;
  if (dow === 0) daysToLastSunday = 7;          // domingo: hace 7 días
  else daysToLastSunday = dow;                   // L=1, Mar=2, ... S=6
  const domingo = new Date(today);
  domingo.setDate(today.getDate() - daysToLastSunday);
  const sabado = new Date(domingo);
  sabado.setDate(domingo.getDate() - 1);
  const viernes = new Date(domingo);
  viernes.setDate(domingo.getDate() - 2);

  // v3.24 — fmt local, no toISOString(): toISOString() convierte a UTC y con un
  // offset positivo devolvería el día siguiente. Con UTC-3 no fallaba, pero era
  // una bomba de tiempo. Existe todayLocalISO() para el mismo propósito.
  const fmt = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const viernesStr = fmt(viernes), sabadoStr = fmt(sabado), domingoStr = fmt(domingo);

  // v3.24 — CORRECCIÓN DE VENTANA. Antes tomaba `viernes-Noche` + sábado + domingo
  // = 7 turnos, asumiendo que `viernes-Noche` era el viernes a la noche.
  // NO lo es: el turno Noche se etiqueta con la fecha de CIERRE (arranca la noche
  // anterior, shiftOrder Noche=0), así que `viernes-Noche` corrió de jueves 23:00
  // a viernes 06:00 — un turno de día de semana. Al mismo tiempo faltaba el sábado
  // a la noche, que se etiqueta `domingo-Noche`.
  //
  // La ventana correcta son los 6 turnos etiquetados sábado y domingo:
  //   sabado-Noche    = viernes 23:00 → sábado 06:00   (el "viernes noche" real)
  //   sabado-Mañana/Tarde
  //   domingo-Noche   = sábado 23:00 → domingo 06:00
  //   domingo-Mañana/Tarde
  // Cobertura real: viernes 23:00 → domingo 23:00.
  // El domingo a la noche NO entra: se etiqueta `lunes-Noche` y es del lunes.
  const turnos = history.filter(r => r.date === sabadoStr || r.date === domingoStr);

  let correctivosGenerados = 0, correctivosRealizados = 0;
  let preventivosAsignados = 0, preventivosRealizados = 0;

  // v3.24 — Misma semántica que computeLastDayStats, para que "generados" signifique
  // lo mismo en las dos cards de la pantalla:
  //   generados  = correctivos PEDIDOS en la ventana (createdInShift dentro del FDS)
  //   realizados = correctivos CERRADOS en la ventana, sin importar cuándo se pidieron
  // Antes "generados" contaba OTs distintas vistas, así que toda heredada abierta
  // sumaba como generada. Verificado sobre el FDS 25-26/07: 45 -> 28.
  // Dedup por otKey (decisión 37). Las OTs sin número no se dedupean: clave por posición.
  const sortedTurnos = [...turnos].sort((a, b) => reportSortKey(a).localeCompare(reportSortKey(b)));
  const vistasFds = new Map();
  sortedTurnos.forEach(r => {
    const reportId = `${r.date}-${r.shift}`;
    (r.corrective || []).forEach((c, i) => {
      const k = canonOT(c.ot) ? otKey(c.ot) : `__SINNUM__${reportId}#${i}`;
      vistasFds.set(k, c); // los turnos vienen ordenados: la última escritura deja el estado final
    });
  });
  const enVentana = (c) => typeof c.createdInShift === 'string' &&
    (c.createdInShift.startsWith(`${sabadoStr}-`) || c.createdInShift.startsWith(`${domingoStr}-`));
  correctivosGenerados = [...vistasFds.values()].filter(enVentana).length;
  correctivosRealizados = [...vistasFds.values()].filter(c => c.state === 'Realizada').length;

  turnos.forEach(r => {
    preventivosAsignados += Number(r.preventivosResumen?.asignados) || 0;
    preventivosRealizados += Number(r.preventivosResumen?.realizados) || 0;
  });

  return {
    viernes: viernesStr,
    sabado: sabadoStr,
    domingo: domingoStr,
    turnos: turnos.sort((a, b) => reportSortKey(a).localeCompare(reportSortKey(b))),
    correctivosGenerados, correctivosRealizados,
    preventivosAsignados, preventivosRealizados
  };
}

// Calcula los últimos 3 turnos del último día con datos (Mañana + Tarde + Noche del día previo).
// Caso típico: el responsable carga el reporte por la mañana, y quiere ver los últimos 3 turnos:
// Noche(día previo) + Tarde(día previo) + Mañana(día previo).
function computeLastDayStats(history) {
  const vacio = {
    fechaBase: '', turnos: [], turnosFaltantes: [], sinDiaAnterior: true,
    correctivosGenerados: 0, correctivosRealizados: 0,
    preventivosAsignados: 0, preventivosRealizados: 0
  };
  if (history.length === 0) return vacio;

  // v3.24 — "Último día" = el día CERRADO más reciente, es decir la fecha más
  // reciente ANTERIOR a hoy que tenga datos. Antes tomaba la fecha más reciente
  // a secas, y como el turno Noche se etiqueta con la fecha de cierre (arranca
  // la noche anterior; shiftOrder Noche=0), el Noche de hoy caía en la card y
  // mostraba un único turno en curso en vez del día completo anterior.
  // Se ancla al reloj y no a "la segunda fecha con datos" para que, si hoy
  // todavía no se cargó nada, igual muestre ayer y no anteayer.
  const hoy = todayLocalISO();
  const fechaBase = [...new Set(history.map(r => r.date))]
    .sort((a, b) => b.localeCompare(a))
    .find(f => f < hoy) || '';
  if (!fechaBase) return vacio;

  const turnos = history
    .filter(r => r.date === fechaBase)
    .sort((a, b) => shiftOrder(a.shift).localeCompare(shiftOrder(b.shift)));
  const turnosFaltantes = TURNOS.filter(t => !turnos.some(r => r.shift === t));

  // v3.24 — Los dos KPIs miden poblaciones DISTINTAS a propósito:
  //   generados  = correctivos que se PIDIERON ese día (createdInShift dentro del día)
  //   realizados = correctivos que se HICIERON ese día, sin importar cuándo se pidieron
  // Por eso "realizados" puede superar a "generados" o quedar muy por debajo:
  // no son numerador y denominador de una misma cosa, son demanda y ejecución.
  // Antes "generados" contaba OTs DISTINTAS vistas en el período, así que toda
  // heredada abierta sumaba como generada y el número se inflaba con el carry-over.
  const vistas = new Map();
  turnos.forEach(r => {
    const reportId = `${r.date}-${r.shift}`;
    (r.corrective || []).forEach((c, i) => {
      // Las OTs sin número no se dedupean (misma regla que dedupCorrective):
      // se les da una clave por posición para que no se agrupen entre sí.
      const k = canonOT(c.ot) ? otKey(c.ot) : `__SINNUM__${reportId}#${i}`;
      // Los turnos vienen ordenados, así que la última escritura deja el estado final del día.
      vistas.set(k, c);
    });
  });

  const delDia = (c) => typeof c.createdInShift === 'string' && c.createdInShift.startsWith(`${fechaBase}-`);
  const correctivosGenerados = [...vistas.values()].filter(delDia).length;
  const correctivosRealizados = [...vistas.values()].filter(c => c.state === 'Realizada').length;

  let preventivosAsignados = 0, preventivosRealizados = 0;
  turnos.forEach(r => {
    preventivosAsignados += Number(r.preventivosResumen?.asignados) || 0;
    preventivosRealizados += Number(r.preventivosResumen?.realizados) || 0;
  });

  return {
    fechaBase, turnos, turnosFaltantes, sinDiaAnterior: false,
    correctivosGenerados, correctivosRealizados,
    preventivosAsignados, preventivosRealizados
  };
}

function computeStats(history, range, customStart, customEnd) {
  const now = new Date(); now.setHours(0, 0, 0, 0);
  let start, end = now;
  if (range === 'custom' && customStart && customEnd) {
    start = new Date(customStart);
    end = new Date(customEnd);
  } else {
    start = new Date(now);
    const map = { week: 7, month: 30, quarter: 90, semester: 180, year: 365 };
    start.setDate(start.getDate() - ((map[range] || 7) - 1));
  }
  const startStr = start.toISOString().slice(0, 10);
  const endStr = end.toISOString().slice(0, 10);

  const filtered = history.filter(r => r.date >= startStr && r.date <= endStr);

  // V2.3 — DEDUPLICACIÓN DE OTs CORRECTIVAS
  // Una OT (por número) puede aparecer en N reportes consecutivos mientras está
  // pendiente. Para no inflar las estadísticas, contamos cada OT UNA SOLA VEZ:
  //   - Estado final = el del reporte más reciente en el rango
  //   - Equipo / técnicos asignados / descripción = los del reporte más reciente
  //   - Fecha de "aparición" = la del reporte más antiguo donde aparece la OT
  // Las OTs sin número (vacío) NO se pueden deduplicar; las contamos igual pero
  // cada aparición cuenta (caso borde, no debería pasar en operación normal).
  const sortedReports = [...filtered].sort((a, b) =>
    `${a.date}-${shiftOrder(a.shift)}`.localeCompare(`${b.date}-${shiftOrder(b.shift)}`)
  );
  const uniqueCorr = new Map();          // ot# -> { latest, firstAppearanceDate, firstAppearanceShift }
  const corrWithoutOT = [];              // OTs sin número (no deduplicadas)
  sortedReports.forEach(r => {
    (r.corrective || []).forEach(c => {
      const key = (c.ot || '').trim();
      if (!key) {
        // OT sin número — no se puede deduplicar, cuenta cada aparición
        corrWithoutOT.push({ ...c, _date: r.date, _shift: r.shift });
        return;
      }
      const existing = uniqueCorr.get(key);
      if (!existing) {
        uniqueCorr.set(key, {
          latest: c,
          firstAppearanceDate: r.date,
          firstAppearanceShift: r.shift
        });
      } else {
        // Sobrescribir 'latest' (estamos iterando ordenado, así que esta es más nueva)
        existing.latest = c;
      }
    });
  });

  // Lista plana de "OTs únicas" para usar en stats
  const uniqueCorrEntries = [
    ...[...uniqueCorr.entries()].map(([key, v]) => ({
      ot: key,
      ...v.latest,
      _firstDate: v.firstAppearanceDate,
      _firstShift: v.firstAppearanceShift
    })),
    ...corrWithoutOT.map(c => ({
      ot: '',
      ...c,
      _firstDate: c._date,
      _firstShift: c._shift
    }))
  ];

  // V3.1 — ESTADO VIGENTE HOY (para Gráfico 1 del pie).
  // Definición (opción 1): "lo que tengo abierto hoy" = los correctivos del ÚLTIMO
  // reporte cargado en el histórico (el de fecha+turno más reciente). Ese reporte
  // ya refleja el carry-over real: arrastra solo las pendientes vivas y descarta las
  // realizadas. Las OTs huérfanas (pendientes viejas que dejaron de arrastrarse) NO
  // aparecen acá porque no están en el último reporte. No se modifica ningún dato:
  // solo cambia cómo cuenta el gráfico.
  // El total de Realizadas se sigue calculando sobre todo el histórico (estado más
  // reciente de cada OT) para que la torta tenga sentido proporcional.
  const allReportsSorted = [...history].sort((a, b) =>
    `${a.date}-${shiftOrder(a.shift)}`.localeCompare(`${b.date}-${shiftOrder(b.shift)}`)
  );
  const lastReport = allReportsSorted[allReportsSorted.length - 1] || null;
  const stateDistVigente = { 'Sin Iniciar': 0, 'En Curso': 0, 'Realizada': 0 };
  // Pendientes: solo las del último reporte cargado (carry-over real de hoy)
  if (lastReport) {
    (lastReport.corrective || []).forEach(c => {
      if (c.state === 'Sin Iniciar' || c.state === 'En Curso') stateDistVigente[c.state]++;
    });
  }
  // Realizadas: total de OTs cuya aparición más reciente en TODO el histórico es Realizada
  const latestStateByOT = new Map();
  const noOTRealizadas = [];
  allReportsSorted.forEach(r => {
    (r.corrective || []).forEach(c => {
      const key = (c.ot || '').trim();
      if (!key) { if (c.state === 'Realizada') noOTRealizadas.push(1); return; }
      latestStateByOT.set(key, c.state);
    });
  });
  latestStateByOT.forEach(st => { if (st === 'Realizada') stateDistVigente['Realizada']++; });
  stateDistVigente['Realizada'] += noOTRealizadas.length;

  // Bucketing diario/semanal/mensual
  const buckets = {};
  const dayMs = 86400000;
  const totalDays = Math.max(1, Math.round((end - start) / dayMs) + 1);
  const bucketByMonth = totalDays > 90;
  const bucketByWeek = totalDays > 31 && !bucketByMonth;

  const bucketKeyAndLabel = (dateStr) => {
    const d = new Date(dateStr);
    if (bucketByMonth) {
      return {
        key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`,
        label: d.toLocaleDateString('es-AR', { month: 'short', year: '2-digit' })
      };
    } else if (bucketByWeek) {
      const monday = new Date(d);
      const day = monday.getDay() || 7;
      monday.setDate(monday.getDate() - day + 1);
      return {
        key: monday.toISOString().slice(0, 10),
        label: `S ${monday.getDate()}/${monday.getMonth() + 1}`
      };
    } else {
      return {
        key: dateStr,
        label: d.toLocaleDateString('es-AR', { day: 'numeric', month: 'short' })
      };
    }
  };

  // Trabajos en el período: correctivos contados el día en que pasaron a REALIZADA
  // (primer reporte del rango donde la OT figura "Realizada"), una sola vez por OT.
  // Preventivos por reporte (cada turno suma su detalle individual).
  // Para detectar el día de cierre recorremos los reportes ordenados y, por cada OT,
  // tomamos la fecha del primer reporte donde aparece con estado "Realizada".
  const closedDateByOT = new Map();   // ot# -> fecha del primer reporte (en rango) con estado Realizada
  const closedNoOT = [];              // OTs sin número, realizadas (cuentan cada aparición realizada)
  sortedReports.forEach(r => {
    (r.corrective || []).forEach(c => {
      if (c.state !== 'Realizada') return;
      const key = (c.ot || '').trim();
      if (!key) { closedNoOT.push(r.date); return; }
      if (!closedDateByOT.has(key)) closedDateByOT.set(key, r.date);
    });
  });
  closedDateByOT.forEach(dateStr => {
    const { key, label } = bucketKeyAndLabel(dateStr);
    if (!buckets[key]) buckets[key] = { key, label, correctivos: 0, preventivos: 0 };
    buckets[key].correctivos++;
  });
  closedNoOT.forEach(dateStr => {
    const { key, label } = bucketKeyAndLabel(dateStr);
    if (!buckets[key]) buckets[key] = { key, label, correctivos: 0, preventivos: 0 };
    buckets[key].correctivos++;
  });
  filtered.forEach(r => {
    const { key, label } = bucketKeyAndLabel(r.date);
    if (!buckets[key]) buckets[key] = { key, label, correctivos: 0, preventivos: 0 };
    // #13 (v3.8) — preventivos del período: "realizados" declarados en el Resumen
    // del turno (preventivosResumen.realizados), no el array `preventive` (que no se
    // usa en producción y daba siempre 0). Vacío/no-numérico cuenta como 0.
    buckets[key].preventivos += Number(r.preventivosResumen?.realizados) || 0;
  });
  const daily = Object.values(buckets).sort((a, b) => a.key.localeCompare(b.key));

  // KPIs principales: deduplicados
  let totalCorrectives = uniqueCorrEntries.length;
  let completedCorr = 0, pendingCorr = 0;
  const stateDistPeriodo = { 'Sin Iniciar': 0, 'En Curso': 0, 'Realizada': 0 };
  const equipmentCount = {};
  const techCount = {};
  const shiftCount = {
    Mañana: { correctivos: 0, preventivos: 0 },
    Tarde:  { correctivos: 0, preventivos: 0 },
    Noche:  { correctivos: 0, preventivos: 0 }
  };

  uniqueCorrEntries.forEach(c => {
    if (c.state === 'Realizada') completedCorr++;
    else if (c.state === 'Sin Iniciar' || c.state === 'En Curso') pendingCorr++;  // V2.4 fix
    if (c.state in stateDistPeriodo) stateDistPeriodo[c.state]++;
    const eq = (c.equipoCodigo || '').trim();
    if (eq && eq !== '-') equipmentCount[eq] = (equipmentCount[eq] || 0) + 1;
    (c.technicians || []).forEach(t => {
      techCount[t] = techCount[t] || { correctivos: 0, preventivos: 0 };
      techCount[t].correctivos++;
    });
    // Turno: tomamos el de la primera aparición (cuándo se "abrió" la OT)
    if (shiftCount[c._firstShift]) shiftCount[c._firstShift].correctivos++;
  });

  // #13 (v3.8) — Preventivos: fuente de verdad = "Resumen preventivos del turno".
  //   - Total del período: suma de preventivosResumen.realizados (lo declara el encargado).
  //   - Carga por técnico y distribución por turno: detalle preventivosResumen.porTecnico.
  // El array `preventive` (bloque "Mantenimiento Preventivo") NO se usa para stats:
  // está oculto a no-admin desde v3.3 (#18) y daba siempre 0. Preventivos NO se
  // deduplican (cada turno declara su propio trabajo).
  let totalPreventives = 0, urgent = 0;
  filtered.forEach(r => {
    totalPreventives += Number(r.preventivosResumen?.realizados) || 0;
    (r.comments || []).forEach(c => { if (c.priority === 'Urgente') urgent++; });

    // Carga por técnico (preventivos) + distribución por turno: del detalle por técnico.
    // Opción C: cada grupo {tecnicos:[a,b,c], cantidad:N} suma N a cada técnico individual.
    // Ej: {tecnicos:['Juan','Pedro'], cantidad: 4} → Juan +4 y Pedro +4.
    // Para la distribución por turno se suma la cantidad del grupo una sola vez.
    (r.preventivosResumen?.porTecnico || []).forEach(grupo => {
      const tecnicos = grupo.tecnicos || (grupo.tecnico ? [grupo.tecnico] : []);
      const cantidad = Number(grupo.cantidad) || 0;
      if (cantidad > 0) {
        tecnicos.forEach(t => {
          techCount[t] = techCount[t] || { correctivos: 0, preventivos: 0 };
          techCount[t].preventivos += cantidad;
        });
        if (shiftCount[r.shift]) shiftCount[r.shift].preventivos += cantidad;
      }
    });
  });

  const completionRate = totalCorrectives > 0 ? Math.round((completedCorr / totalCorrectives) * 100) : 0;

  const topEquipment = Object.entries(equipmentCount)
    .sort((a, b) => b[1] - a[1]).slice(0, 8)
    .map(([name, count]) => ({ name, count }));

  // V2.4 — "Carga por técnico" incluye TODOS los técnicos del catálogo (21),
  // aunque no hayan tenido OTs en el período. Orden descendente por carga total.
  // Los técnicos en cero quedan al fondo.
  const topTechs = TECNICO_NAMES.map(fullName => {
    const v = techCount[fullName] || { correctivos: 0, preventivos: 0 };
    return {
      name: fullName.split(' ').slice(0, 2).join(' '),
      correctivos: v.correctivos,
      preventivos: v.preventivos,
      total: v.correctivos + v.preventivos
    };
  }).sort((a, b) => b.total - a.total);

  const shiftDist = Object.entries(shiftCount).map(([name, v]) => ({ name, ...v }));

  return {
    totalReports: filtered.length,
    totalCorrectives, totalPreventives, completedCorr, pendingCorr, urgent,
    completionRate, daily, startStr, endStr,
    // Gráfico 1 (estado vigente HOY, sobre todo el histórico) y su total de pendientes reales
    stateDistVigente: Object.entries(stateDistVigente).map(([name, value]) => ({ name, value })).filter(x => x.value > 0),
    pendientesVigentes: stateDistVigente['Sin Iniciar'] + stateDistVigente['En Curso'],
    // Gráfico 2 (estado al cierre del período filtrado)
    stateDistPeriodo: Object.entries(stateDistPeriodo).map(([name, value]) => ({ name, value })).filter(x => x.value > 0),
    pendientesPeriodo: stateDistPeriodo['Sin Iniciar'] + stateDistPeriodo['En Curso'],
    topEquipment, topTechs, shiftDist
  };
}

// ═══════════════════════════════════════════════════════════════════
// HISTORY VIEW — V2.0
//   - botones nuevos: Solo Comentarios, Solo Proveedores
//   - fechas en formato dd/mmm/aa
// ═══════════════════════════════════════════════════════════════════
function HistoryView({ history, onExportCorrectives, onExportPreventives, onExportComments, onExportProviders, onExportFull, adminMode, onDeleteReport }) {
  const [filter, setFilter] = useState('');
  const filtered = history.filter(r =>
    !filter || r.date.includes(filter) || r.shift.toLowerCase().includes(filter.toLowerCase()) ||
    (r.responsable || '').toLowerCase().includes(filter.toLowerCase())
  );

  const totals = useMemo(() => ({
    reports: history.length,
    correctives: history.reduce((s, r) => s + (r.corrective?.length || 0), 0),
    // #13 (v3.8) — total de preventivos = suma de "realizados" del Resumen del turno
    // (no el array `preventive`, que no se usa y daba 0).
    preventives: history.reduce((s, r) => s + (Number(r.preventivosResumen?.realizados) || 0), 0),
    urgent: history.reduce((s, r) => s + (r.comments?.filter(c => c.priority === 'Urgente').length || 0), 0)
  }), [history]);

  return (
    <div className="space-y-5">
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <KPI label="Reportes guardados" value={totals.reports} icon={ClipboardList} color="sky" />
        <KPI label="Correctivos totales" value={totals.correctives} icon={Wrench} color="orange" />
        <KPI label="Preventivos totales" value={totals.preventives} icon={ListChecks} color="emerald" />
        <KPI label="Comentarios urgentes" value={totals.urgent} icon={AlertTriangle} color="red" />
      </div>

      <Card className="p-5">
        <SectionTitle icon={FileSpreadsheet} accent="emerald">Exportar a Excel</SectionTitle>
        <p className="text-sm text-slate-600 mb-4">
          Los archivos respetan la estructura del template original (mismas hojas y columnas).
          Los datos se acumulan automáticamente con cada turno cargado.
        </p>
        <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-3">
          <button onClick={onExportCorrectives} className="bg-orange-600 text-white hover:bg-orange-700 rounded-lg p-3 transition flex items-start gap-2">
            <Wrench className="w-5 h-5 flex-shrink-0 mt-0.5" />
            <div className="text-left">
              <div className="font-semibold text-sm">Solo Correctivos</div>
              <div className="text-[11px] text-orange-100 mt-0.5">Hoja del template</div>
            </div>
          </button>
          <button onClick={onExportPreventives} className="bg-emerald-600 text-white hover:bg-emerald-700 rounded-lg p-3 transition flex items-start gap-2">
            <ListChecks className="w-5 h-5 flex-shrink-0 mt-0.5" />
            <div className="text-left">
              <div className="font-semibold text-sm">Solo Preventivos</div>
              <div className="text-[11px] text-emerald-100 mt-0.5">Hoja del template</div>
            </div>
          </button>
          <button onClick={onExportComments} className="bg-amber-600 text-white hover:bg-amber-700 rounded-lg p-3 transition flex items-start gap-2">
            <MessageSquare className="w-5 h-5 flex-shrink-0 mt-0.5" />
            <div className="text-left">
              <div className="font-semibold text-sm">Solo Comentarios</div>
              <div className="text-[11px] text-amber-100 mt-0.5">Con fecha y turno</div>
            </div>
          </button>
          <button onClick={onExportProviders} className="bg-indigo-600 text-white hover:bg-indigo-700 rounded-lg p-3 transition flex items-start gap-2">
            <Building2 className="w-5 h-5 flex-shrink-0 mt-0.5" />
            <div className="text-left">
              <div className="font-semibold text-sm">Solo Proveedores</div>
              <div className="text-[11px] text-indigo-100 mt-0.5">Con fecha y turno</div>
            </div>
          </button>
          <button onClick={onExportFull} className="bg-slate-800 text-white hover:bg-slate-700 rounded-lg p-3 transition flex items-start gap-2">
            <FileSpreadsheet className="w-5 h-5 flex-shrink-0 mt-0.5" />
            <div className="text-left">
              <div className="font-semibold text-sm">Reporte completo</div>
              <div className="text-[11px] text-slate-300 mt-0.5">Todas las hojas</div>
            </div>
          </button>
        </div>
      </Card>

      <Card className="p-5">
        <div className="flex items-center justify-between mb-4">
          <SectionTitle icon={Calendar} accent="sky">Reportes guardados</SectionTitle>
          <div className="relative">
            <Filter className="w-4 h-4 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input value={filter} onChange={e => setFilter(e.target.value)} placeholder="Filtrar fecha, turno o responsable…"
              className="pl-8 pr-3 py-2 text-sm border border-slate-300 rounded-lg w-64 focus:outline-none focus:ring-2 focus:ring-sky-500/40" />
          </div>
        </div>
        <p className="text-[11px] text-slate-400 -mt-2 mb-4">
          Correctivos: total de OTs cargadas en el reporte (realizadas + pendientes). Preventivos: realizados / asignados declarados en el "Resumen preventivos del turno" (— si no se cargó el dato).
        </p>
        {filtered.length === 0 ? (
          <EmptyHint>{history.length === 0 ? 'No hay reportes guardados.' : 'Sin coincidencias para el filtro.'}</EmptyHint>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left border-b border-slate-200 text-xs text-slate-500 uppercase tracking-wide">
                  <th className="pb-2 font-medium">Fecha</th>
                  <th className="pb-2 font-medium">Turno</th>
                  <th className="pb-2 font-medium">Responsable</th>
                  <th className="pb-2 font-medium text-right">Equipo</th>
                  <th className="pb-2 font-medium text-right">Correctivos</th>
                  <th className="pb-2 font-medium text-right">Prev. (real/asig)</th>
                  <th className="pb-2 font-medium text-right">Urgentes</th>
                  {adminMode && <th className="pb-2 font-medium text-right">Acciones</th>}
                </tr>
              </thead>
              <tbody>
                {filtered.map((r, i) => (
                  <tr key={i} className="border-b border-slate-100 hover:bg-slate-50">
                    <td className="py-2 num font-medium">{formatDateShort(r.date)}</td>
                    <td className="py-2">{r.shift}</td>
                    <td className="py-2 text-slate-600">{r.responsable || '—'}</td>
                    <td className="py-2 text-right num">{r.team?.length || 0}</td>
                    <td className="py-2 text-right num">{r.corrective?.length || 0}</td>
                    <td className="py-2 text-right num">
                      {(() => {
                        // #13 (v3.8) — Preventivos = realizados / asignados del Resumen del turno.
                        // "—" cuando el campo está vacío/null (distingue "no cargado" de "0 hechos").
                        const pr = r.preventivosResumen || {};
                        const fmt = (v) => (v !== '' && v != null ? v : '—');
                        return `${fmt(pr.realizados)} / ${fmt(pr.asignados)}`;
                      })()}
                    </td>
                    <td className="py-2 text-right num">
                      {r.comments?.filter(c => c.priority === 'Urgente').length || 0}
                    </td>
                    {adminMode && (
                      <td className="py-2 text-right">
                        <button
                          onClick={() => onDeleteReport(r.date, r.shift)}
                          className="inline-flex items-center gap-1 px-2 py-1 rounded text-red-600 hover:bg-red-50 hover:text-red-700 transition text-xs"
                          title="Eliminar reporte"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// XLSX HELPERS
// ═══════════════════════════════════════════════════════════════════
function addSheet(wb, rows, name) {
  const ws = XLSX.utils.json_to_sheet(rows);
  if (rows.length > 0 && rows[0]) {
    const cols = Object.keys(rows[0]).map(k => ({ wch: Math.max(12, Math.min(35, k.length + 4)) }));
    ws['!cols'] = cols;
  }
  XLSX.utils.book_append_sheet(wb, ws, name);
}

function downloadSingle(rows, sheetName, fileName) {
  const wb = XLSX.utils.book_new();
  addSheet(wb, rows, sheetName);
  XLSX.writeFile(wb, fileName);
}

// ── #66 — Export con estilo (solo Horas Extras) ─────────────────────
// Usa las propiedades de estilo que agrega xlsx-js-style. NO se mezcla con
// `addSheet`/`downloadSingle` de arriba porque esas siguen sirviendo a
// Correctivos/Preventivos/Comentarios/Proveedores tal cual estaban — este
// pedido de formato fue puntual para Horas Extras, no para todos los
// exports de la app.
const EXTRAS_EXCEL_HEADER_STYLE = {
  font: { bold: true, color: { rgb: 'FFFFFFFF' } },
  fill: { patternType: 'solid', fgColor: { rgb: 'FF0284C7' } }
};

function addStyledSheet(wb, rows, name) {
  const ws = XLSX.utils.json_to_sheet(rows);
  if (rows.length > 0 && rows[0]) {
    const headers = Object.keys(rows[0]);
    ws['!cols'] = headers.map(k => ({ wch: Math.max(12, Math.min(35, k.length + 4)) }));
    // Fila 1 (encabezados): negrita + fondo celeste de la app, texto blanco.
    headers.forEach((_, i) => {
      const addr = XLSX.utils.encode_cell({ r: 0, c: i });
      if (ws[addr]) ws[addr].s = EXTRAS_EXCEL_HEADER_STYLE;
    });
  }
  XLSX.utils.book_append_sheet(wb, ws, name);
}

// `sheets`: [{ rows, name }, ...] — una entrada por pestaña, en orden.
function downloadStyledExtras(sheets, fileName) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ rows, name }) => addStyledSheet(wb, rows, name));
  XLSX.writeFile(wb, fileName);
}

import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import {
  ClipboardList, BarChart3, Download, Plus, Trash2, Save, Calendar, Users,
  Wrench, Activity, FileSpreadsheet, CheckCircle2, AlertTriangle, Building2,
  HardHat, Beaker, ListChecks, ChevronDown, X, FileText, TrendingUp, Flame,
  Cog, Zap, Filter, Search, Cloud, CloudOff, RefreshCw, Settings, MessageSquare,
  CalendarDays, Clock, Image as ImageIcon, FileDown,
  Lock, LogOut, Edit3, Shield
} from 'lucide-react';
import * as XLSX from 'xlsx';
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
const APP_VERSION = 'v3.1';
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
// V2.6 — MODO ADMINISTRADOR
// Password hardcoded para acceso a edición/eliminación avanzada.
// Sirve como barrera contra clicks accidentales, NO es control de
// acceso real (el password queda visible en GitHub).
// ═══════════════════════════════════════════════════════════════════
const ADMIN_PASSWORD = 'FerringBiomas2026';

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
  { code: 'FAC',  label: 'Facilities' }
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

// Parsea "FOA1-01395" → { sector: 'FOA1', numero: '01395' }
// IMPORTANTE: NO padea el número. Solo lo devuelve tal cual.
// El padding se aplica únicamente al perder foco (handleNumeroBlur).
// Esto permite escribir digito por digito sin que se autocompleten ceros.
const parseOT = (ot) => {
  if (!ot) return { sector: '', numero: '' };
  const match = ot.trim().match(/^([A-Z0-9]+)-(\d{1,5})$/);
  if (match && SECTORES_CODES.includes(match[1])) {
    return { sector: match[1], numero: match[2] };
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
const shiftOrder = (s) => ({ 'Mañana': '1', 'Tarde': '2', 'Noche': '3' }[s] || '9');
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
  date: new Date().toISOString().slice(0, 10),
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

// V2.4 — Hidrata un reporte asegurando estructura completa.
// Esto cubre reportes guardados con schemas anteriores (V1.0 a V2.3):
//   - Correctivos sin `timeline` → se inicializa como []
//   - Grupos del resumen con `tecnico` (singular) → se migran a `tecnicos: [...]`
//   - Servicios y subobjetos faltantes se completan con defaults
const hydrate = (raw) => {
  const base = emptyReport();
  if (!raw) return base;
  return {
    ...base,
    ...raw,
    // V2.4 — asegurar timeline en cada OT correctiva
    corrective: (raw.corrective || []).map(c => ({
      ...c,
      timeline: c.timeline || []
    })),
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
      return rows.map(r => r.data);
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
  }
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
// OT NUMBER INPUT — V2.4
// Input compuesto: dropdown sector + input numérico de 5 dígitos.
// Auto-completa con ceros a la izquierda al perder foco.
// Si la OT viene en formato legacy (sin guión, espacios, etc), muestra
// el valor raw en modo "legacy" con un indicador visual y un tooltip.
// ═══════════════════════════════════════════════════════════════════
function OTNumberInput({ value, onChange, isLegacy, hasError, disabled }) {
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

  // Si es legacy (formato viejo) mostramos input plano con badge "legacy"
  if (isLegacy) {
    return (
      <div className="flex items-center gap-1">
        <input
          className={`${inputCls} num flex-1 ${hasError ? 'border-red-400' : 'border-amber-300 bg-amber-50/30'}`}
          value={value || ''}
          onChange={e => onChange(e.target.value)}
          placeholder="OT legacy"
          disabled={disabled}
          title="OT en formato legacy (anterior a V2.4). Para validar al guardar usá el formato XXX-YYYYY"
        />
        <span className="text-[9px] px-1 py-0.5 bg-amber-100 text-amber-700 rounded font-bold" title="Formato legacy">L</span>
      </div>
    );
  }

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
    <div className="flex items-center gap-1">
      <select
        className={`px-1.5 py-2 text-xs bg-white border rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500/40 transition num font-semibold ${hasError && !sector ? 'border-red-400' : 'border-slate-300'}`}
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
  // V2.6 — Confirmación de eliminación de reporte completo
  const [deleteReportConfirm, setDeleteReportConfirm] = useState(null); // null | { date, shift, source }

  // V2.8 — Conflictos de OT cerrada detectados al guardar.
  // Cuando se intenta guardar un reporte que contiene una OT en estado
  // Sin Iniciar / En Curso pero esa misma OT ya está Realizada en otro reporte
  // posterior en Supabase, se abre este modal para que el usuario decida qué hacer.
  // Cada conflicto: { otNumber, otTask, closedIn: {date, shift, responsable}, currentStateInForm }
  // Estructura: null | { conflicts: [...], onResolve: (decisions) => void }
  const [closedConflicts, setClosedConflicts] = useState(null);

  // V2.9 — Modal de propagación admin.
  // Estructura: null | { diffs, affectedReports, fixedReport }
  //   - diffs: salida de detectChangesForPropagation
  //   - affectedReports: salida de findAffectedLaterReports
  //   - fixedReport: el reporte editado que se va a guardar después de confirmar
  const [propagationModal, setPropagationModal] = useState(null);

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

  const refresh = useCallback(async () => {
    try {
      setConnError('');
      const data = await storage.list();
      // Hidratar todos los reportes para garantizar estructura V2.0
      setHistory(data.map(hydrate));
    } catch (e) {
      setConnError(e.message || 'Error de conexión');
      console.error(e);
    }
  }, []);

  useEffect(() => { (async () => { await refresh(); setLoading(false); })(); }, [refresh]);

  // Validaciones antes de guardar (V2.0)
  // Devuelve string con error o '' si todo OK
  const validateReport = (r) => {
    // V2.9 — Si admin está editando un reporte histórico (tiene snapshot original),
    // saltar todas las validaciones contextuales. Admin asume responsabilidad de
    // lo que guarda. Esto evita que reglas retroactivas (técnico obligatorio,
    // avance de turno cuando hay cambio de estado, formato XXX-YYYYY, etc.) bloqueen
    // la edición de reportes pre-V2.5 / pre-V2.4 / etc.
    if (adminMode && originalReport) return '';
    const currentShiftKey = `${r.date}-${r.shift}`;

    // V2.4 — 1. OTs nuevas (creadas en este turno) deben tener formato XXX-YYYYY válido
    const otsNuevasInvalidas = (r.corrective || []).filter(
      c => c.createdInShift === currentShiftKey && !isValidOT(c.ot)
    );
    if (otsNuevasInvalidas.length > 0) {
      return `${otsNuevasInvalidas.length} OT nueva con formato inválido. Formato requerido: XXX-YYYYY (sector + 5 dígitos). Ej: FOA1-01395`;
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

    const otsSinAvance = (r.corrective || []).filter(c => {
      if (!requiresAdvance(c)) return false;
      const tl = c.timeline || [];
      const hasFromCurrent = tl.some(e => e.shiftKey === currentShiftKey);
      return !hasFromCurrent;
    });
    if (otsSinAvance.length > 0) {
      return `${otsSinAvance.length} OT con cambio de estado en este turno sin entrada de Estado de avance. Cargá el avance antes de guardar.`;
    }

    // V2.5 — 3. TODAS las OTs correctivas deben tener al menos un técnico, sin importar estado.
    // (Antes era sólo para "Realizada"; ahora aplica también a "Sin Iniciar" y "En Curso".)
    const correctivasSinTecnico = (r.corrective || []).filter(
      c => !c.technicians || c.technicians.length === 0
    );
    if (correctivasSinTecnico.length > 0) {
      return `${correctivasSinTecnico.length} OT correctiva sin técnico asignado. Asigná técnicos antes de guardar.`;
    }

    // V2.5 — 4. TODOS los preventivos cargados deben tener al menos un técnico.
    const preventivosSinTecnico = (r.preventive || []).filter(
      p => !p.technicians || p.technicians.length === 0
    );
    if (preventivosSinTecnico.length > 0) {
      return `${preventivosSinTecnico.length} preventivo sin técnico asignado. Asigná técnicos antes de guardar.`;
    }

    // 5. Resumen de preventivos: si hay realizados > 0, la suma del detalle por técnico debe coincidir
    const realizados = Number(r.preventivosResumen?.realizados);
    if (realizados > 0) {
      // V2.4 — opción C: cada fila puede tener N técnicos en grupo. La cantidad cuenta UNA VEZ
      // (no se multiplica por la cantidad de técnicos del grupo).
      const sumaPorGrupo = (r.preventivosResumen?.porTecnico || [])
        .reduce((s, t) => s + (Number(t.cantidad) || 0), 0);
      if (sumaPorGrupo !== realizados) {
        return `Resumen de preventivos: la suma del detalle (${sumaPorGrupo}) no coincide con "Preventivos realizados" (${realizados}).`;
      }
      // Validar que no haya filas sin técnicos o con cantidad <= 0
      const filasMalas = (r.preventivosResumen?.porTecnico || []).filter(t => {
        const tecnicos = t.tecnicos || (t.tecnico ? [t.tecnico] : []);  // compat hacia atrás
        return tecnicos.length === 0 || !t.cantidad || Number(t.cantidad) <= 0;
      });
      if (filasMalas.length > 0) {
        return `Resumen de preventivos: hay filas sin técnicos o con cantidad inválida.`;
      }
    }

    return '';
  };

  // V2.5 — Detecta correctivos y preventivos "vacíos" (3 campos clave en blanco).
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
  const doSaveReport = async (reportToSave) => {
    if (!reportToSave.date || !reportToSave.shift) { setSaveMsg('Falta fecha o turno'); return; }
    const validationError = validateReport(reportToSave);
    if (validationError) {
      setSaveMsg(`Error: ${validationError}`);
      return;
    }
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
    try {
      await storage.save(reportToSave);
      await refresh();
      // V2.5 — Si se limpiaron entradas vacías, persistir el cambio en el state local
      // así el usuario ve el form sin las filas vacías
      if (reportToSave !== report) setReport(reportToSave);
      setOriginalReport(JSON.parse(JSON.stringify(reportToSave)));  // V2.9 — actualizar snapshot al guardado nuevo
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

  return (
    <div className="min-h-screen bg-slate-50" style={{ fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap');
        body { font-family: 'Plus Jakarta Sans', system-ui, sans-serif; }
        .num { font-family: 'JetBrains Mono', monospace; font-feature-settings: 'tnum'; }
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
              {/* V2.6 — Botón Admin debajo del badge de Supabase */}
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
              <div className="hidden md:block text-slate-300">{history.length} {history.length === 1 ? 'reporte' : 'reportes'}</div>
            </div>
          </div>
        </div>
        <div className="max-w-[1600px] mx-auto px-6">
          <nav className="flex gap-1 overflow-x-auto">
            {[
              { id: 'form', label: 'Cargar Reporte', icon: ClipboardList },
              { id: 'dashboard', label: 'Dashboard', icon: BarChart3 },
              { id: 'stats', label: 'Estadísticas', icon: TrendingUp },
              { id: 'history', label: 'Histórico & Excel', icon: FileSpreadsheet }
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
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// V2.6 — MODAL DE LOGIN ADMIN
// Pide el password de administrador antes de activar el modo admin.
// ═══════════════════════════════════════════════════════════════════
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
// V2.5 — MODAL DE CONFIRMACIÓN DE ENTRADAS VACÍAS
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
function FormView({ report, setReport, onSave, saveMsg, setSaveMsg, saving, history, adminMode, originalReport, setOriginalReport, onDeleteReport }) {
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
  const updateCorrectiveItem = (i, patch) => setReport(r => ({
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

  // V2.4 — Timeline de Estado de Avance
  // Cada OT correctiva tiene un array timeline:
  //   [{ shiftKey, date, shift, author, text, timestamp }]
  // - Cargado solo cuando el estado es "En Curso" (y al guardar la app valida
  //   que haya entrada del turno actual)
  // - Las entradas anteriores son read-only (no se pueden editar ni borrar)
  // - V2.6 — En modo admin, las entradas son editables y eliminables
  // - timelineDraft mantiene el texto en redacción por índice de OT
  const [timelineDraft, setTimelineDraft] = useState({});
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

  const [loadInfo, setLoadInfo] = useState('');
  const initialPendingApplied = useRef(false);

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

  return (
    <div className="space-y-5">
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
      {/* TOP ACTION BAR — sticky so the Save button is always visible */}
      <div className="sticky top-0 z-20 -mx-6 px-6 py-3 bg-slate-50/95 backdrop-blur border-b border-slate-200">
        <div className="flex items-center justify-between bg-white border border-slate-200 rounded-xl p-3 shadow-sm">
          <div className="flex items-center gap-2 text-sm text-slate-600 min-w-0">
            <ClipboardList className="w-5 h-5 text-sky-600 flex-shrink-0" />
            <span className="font-medium truncate">
              {loadInfo || 'Cargá los datos del turno y guardá. Se almacenan automáticamente.'}
            </span>
          </div>
          <div className="flex items-center gap-3 flex-shrink-0">
            {saveMsg && <span className={`text-sm font-medium ${saveMsg.startsWith('Error') ? 'text-red-600' : saveMsg.startsWith('✓') ? 'text-emerald-600' : 'text-slate-500'}`}>{saveMsg}</span>}
            {/* V2.6 — Botón eliminar reporte (solo modo admin y si el reporte ya está guardado) */}
            {adminMode && history.some(r => r.date === report.date && r.shift === report.shift) && (
              <button onClick={onDeleteReport}
                className={`${buttonCls} bg-red-50 text-red-700 hover:bg-red-100 ring-1 ring-red-200`}
                title="Eliminar este reporte completo">
                <Trash2 className="w-4 h-4" />Eliminar reporte
              </button>
            )}
            <button onClick={cleanForm} className={`${buttonCls} bg-slate-100 text-slate-600 hover:bg-slate-200`}>
              Limpiar
            </button>
            <button onClick={onSave} disabled={saving} className={`${buttonCls} bg-slate-800 text-white hover:bg-slate-700 px-5 disabled:opacity-50`}>
              <Save className="w-4 h-4" />Guardar reporte
            </button>
          </div>
        </div>
      </div>

      {/* COMPACT TURNO HEADER (single line) */}
      <Card className="p-5">
        <SectionTitle icon={Calendar} accent="sky">Información del Turno</SectionTitle>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <Field label="Fecha">
            <input type="date" className={inputCls} value={report.date} onChange={e => setDateShift(e.target.value, report.shift)} />
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

      {/* CORRECTIVOS — V2.0: SIN botón eliminar; V2.5: técnico obligatorio siempre + banner */}
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

        {report.corrective.length === 0 && <EmptyHint>Sin órdenes de trabajo correctivas.</EmptyHint>}
        <div className="space-y-3">
          {report.corrective.map((c, i) => {
            // V2.5 — Técnico obligatorio en cualquier estado (no solo "Realizada")
            const missingTech = (!c.technicians || c.technicians.length === 0);
            // V2.4 — Determinar si la OT es nueva (creada en este turno) o legacy
            // Las nuevas se editan con OTNumberInput estructurado, las legacy con input plano
            const currentShiftKey = `${report.date}-${report.shift}`;
            const isNewOT = c.createdInShift === currentShiftKey;
            const isLegacyFormat = !isValidOT(c.ot) && !isNewOT;
            const otHasError = isNewOT && !isValidOT(c.ot);
            return (
              <div key={i} id={`form-ot-${c.ot || `idx-${i}`}`} className={`border rounded-lg p-3 relative scroll-mt-32 ${missingTech || otHasError ? 'border-red-300 bg-red-50/40' : 'border-slate-200 bg-slate-50/40'}`}>
                {/* V2.6 — Botón eliminar OT (solo modo admin) */}
                {adminMode && (
                  <button
                    onClick={() => updateList('corrective', l => l.filter((_, j) => j !== i))}
                    className="absolute top-2 right-2 z-10 text-red-500 hover:text-red-700 hover:bg-red-100 rounded p-1.5 transition ring-1 ring-red-200 bg-white"
                    title="Eliminar OT (modo admin)">
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                )}
                <div className="grid grid-cols-12 gap-2 mb-2">
                  <Field label="N° OT *" className="col-span-7 lg:col-span-3">
                    <OTNumberInput
                      value={c.ot}
                      onChange={(newOt) => updateCorrectiveItem(i, { ot: newOt })}
                      isLegacy={isLegacyFormat}
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
                  return (
                    <div className={`mt-2 border rounded-lg p-3 ${requiresEntry && timelineDraft[i] === undefined ? 'border-amber-300 bg-amber-50/40' : 'border-slate-200 bg-slate-50/30'}`}>
                      <div className="flex items-center justify-between mb-2">
                        <div className="text-[11px] font-semibold text-slate-700 uppercase tracking-wide inline-flex items-center gap-1.5">
                          <Activity className="w-3.5 h-3.5 text-amber-500" />
                          Estado de avance {requiresEntry && <span className="text-red-600">*</span>}
                        </div>
                        {requiresEntry && (
                          <span className="text-[10px] text-amber-700 font-medium inline-flex items-center gap-1">
                            <AlertTriangle className="w-3 h-3" />
                            Requiere entrada del turno actual
                          </span>
                        )}
                      </div>

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
                          o si está Realizada y requiere avance (cierre obligatorio del trabajo). */}
                      {(c.state === 'En Curso' || (c.state === 'Realizada' && requiresEntry)) && (
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
          })}
        </div>
      </Card>

      {/* RESUMEN PREVENTIVOS DEL TURNO — V2.1: subido a la 4ta posición (antes de Servicios) */}
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

      {/* SERVICIOS */}
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

        {/* SERVICIOS EXTERNOS */}
        <div>
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-sm font-semibold text-slate-700 inline-flex items-center gap-2">
              <Building2 className="w-4 h-4 text-indigo-500" />Servicios externos / Proveedores
            </h3>
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
        </div>
      </Card>

      {/* COMENTARIOS */}
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

      {/* MANTENIMIENTO PREVENTIVO — V2.1: bajado al final del formulario */}
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
function CorrectiveSubsection({ title, count, items, showStateBadge, showAvanceMark, adminMode, onItemClick }) {
  // V2.5 — Si showAvanceMark, ordenar items con avance del turno arriba.
  // Mantiene el orden relativo dentro de cada grupo (con y sin avance).
  const sortedItems = showAvanceMark
    ? [
        ...items.filter(c => c._currentShiftEntry),
        ...items.filter(c => !c._currentShiftEntry)
      ]
    : items;

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
            const hasAvance = showAvanceMark && c._currentShiftEntry;
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
                {hasAvance && (
                  <div className="mt-1 text-[11px] text-emerald-800 bg-emerald-50/60 border-l-2 border-emerald-300 pl-2 py-0.5 leading-snug whitespace-pre-wrap break-words">
                    <span className="text-emerald-600 font-semibold mr-1">↳ Avance del turno:</span>
                    {c._currentShiftEntry.text || '—'}
                  </div>
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
// DASHBOARD VIEW — V2.3
//   - Solo muestra OTs del turno actual (creadas o modificadas en este turno)
//   - Las del carry-over que nadie tocó NO aparecen acá (sí en Cargar Reporte)
//   - Detalle por técnico de preventivos en grid 2-col compacto
// V2.5 — Correctivos en 4 categorías: Realizados del turno / Heredados realizados /
//        Pendientes del turno / Pendientes heredados
// ═══════════════════════════════════════════════════════════════════
function DashboardView({ report, history = [], activeReport, dashboardOverride, setDashboardOverride, adminMode, onEditFromDashboard }) {
  const dateLabel = useMemo(() => formatDateLong(report.date), [report.date]);
  const dateShort = useMemo(() => formatDateShort(report.date), [report.date]);
  const p = report.servicios.plantaCaldera;
  const pr = report.preventivosResumen || { asignados: '', realizados: '', porTecnico: [] };

  // V2.4 — Listado de turnos guardados disponibles para el selector.
  // Orden cronológico descendente (más nuevos primero).
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

  // V2.4 — Modo visor (cuando hay override). Sirve para no permitir edición y
  // mostrar banner informativo.
  const isViewerMode = !!dashboardOverride;
  const currentKey = `${report.date}|${report.shift}`;
  const currentIdx = savedShifts.findIndex(s => s.key === currentKey);

  // V2.4 — Navegación con flechas (cronológica). savedShifts está orden desc,
  // entonces "anterior" cronológico = índice +1, "siguiente" = índice -1.
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

  // V2.4 — Filtro de OTs (opción B):
  // - Si el reporte está GUARDADO (existe en history con misma fecha+turno),
  //   mostramos TODAS las OTs (representan lo que pasó realmente ese turno).
  // - Si el reporte es NUEVO (no guardado todavía), filtramos:
  //     · OTs creadas en este turno (createdInShift coincide) → SÍ
  //     · OTs modificadas en este turno (lastModifiedInShift coincide) → SÍ
  //     · OTs heredadas del carry-over sin tocar → NO
  // Esto evita que en un turno nuevo aparezcan en el Dashboard todas las
  // OTs heredadas del carry-over sin que el responsable haya hecho nada.
  const currentShiftKey = `${report.date}-${report.shift}`;
  const isExistingReport = useMemo(
    () => history.some(r => r.date === report.date && r.shift === report.shift),
    [history, report.date, report.shift]
  );
  const correctiveActual = useMemo(() => {
    if (isExistingReport) {
      // Reporte guardado → mostrar todo
      return report.corrective || [];
    }
    // Reporte nuevo → solo OTs del turno actual
    return (report.corrective || []).filter(c => {
      if (c.createdInShift === currentShiftKey) return true;
      if (c.lastModifiedInShift === currentShiftKey) return true;
      return false;
    });
  }, [report.corrective, currentShiftKey, isExistingReport]);

  // V2.5 — Particiones de correctivos en 4 categorías para el Dashboard:
  //  - Realizados del turno      → creados en este turno (createdInShift === currentShiftKey) y estado "Realizada"
  //  - Realizados heredados      → creados en turno previo y ahora en "Realizada"
  //  - Pendientes del turno      → creados en este turno y no "Realizada"
  //  - Pendientes heredados      → creados en turno previo y no "Realizada"
  // El criterio es por CREACIÓN de la OT (createdInShift), no por trabajo realizado.
  // Si una OT vieja recibió un avance en este turno, sigue siendo "heredada"
  // pero se marca con badge "Avance hoy" y se muestra el texto del último avance.
  const correctivePartitions = useMemo(() => {
    const isCreatedHere = (c) => c.createdInShift === currentShiftKey;
    const isDone = (c) => c.state === 'Realizada';

    // Para el avance del turno actual: buscar entrada de timeline cuyo shiftKey === currentShiftKey
    const findCurrentShiftEntry = (c) => {
      const tl = c.timeline || [];
      // Si hay varias entradas del turno, tomar la última (más reciente)
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

  // V2.2 — Export del Dashboard a PNG/PDF
  const dashboardRef = useRef(null);
  const [exporting, setExporting] = useState('');
  const [exportMsg, setExportMsg] = useState('');

  // Carga dinámica de las librerías de export desde CDN.
  // Se hace lazy (solo al hacer click) para no agrandar el bundle inicial.
  const loadScript = (src) => new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) return resolve();
    const s = document.createElement('script');
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error(`No se pudo cargar ${src}`));
    document.head.appendChild(s);
  });

  // Captura el dashboard expandiendo todos los scrolls internos para que se vea completo.
  // V2.5 — también des-trunca los textos para que no se corten en el export.
  // Devuelve el canvas resultado.
  const captureDashboard = async () => {
    if (!window.html2canvas) {
      await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
    }
    const node = dashboardRef.current;
    if (!node) throw new Error('Dashboard no encontrado');

    // Antes de capturar: forzar a que TODOS los contenedores con overflow:auto/scroll
    // muestren todo el contenido. Guardar los estilos originales para restaurarlos después.
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

    // V2.5 — Des-truncar textos para que no se corten en el export.
    // Tailwind `truncate` aplica white-space:nowrap + overflow:hidden + text-overflow:ellipsis.
    // Sobreescribimos esos estilos en línea para que el texto fluya y se envuelva normal.
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

    // Esperar un frame para que se aplique
    await new Promise(r => requestAnimationFrame(r));
    await new Promise(r => setTimeout(r, 50));

    let canvas;
    try {
      canvas = await window.html2canvas(node, {
        backgroundColor: '#f8fafc',  // bg-slate-50
        scale: 2,                     // alta resolución
        useCORS: true,
        logging: false,
        windowWidth: node.scrollWidth,
        windowHeight: node.scrollHeight
      });
    } finally {
      // Restaurar estilos originales pase lo que pase
      originalStyles.forEach(({ el, overflow, overflowY, maxHeight, height }) => {
        el.style.overflow = overflow;
        el.style.overflowY = overflowY;
        el.style.maxHeight = maxHeight;
        el.style.height = height;
      });
      // V2.5 — restaurar también los estilos de truncate
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

      // A4 horizontal: 297mm x 210mm
      const pdf = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const margin = 8;
      const availW = pageW - margin * 2;
      const availH = pageH - margin * 2;

      // Escalar la imagen para que entre en la página manteniendo proporción
      const imgAspect = canvas.width / canvas.height;
      const availAspect = availW / availH;
      let drawW, drawH;
      if (imgAspect > availAspect) {
        // imagen más ancha que el área disponible: limitar por ancho
        drawW = availW;
        drawH = drawW / imgAspect;
      } else {
        // imagen más alta: limitar por alto
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
      {/* V2.4 — Selector de turno guardado (modo visor de turnos pasados).
          Si hay override, mostramos banner indicador y permitimos volver al turno actual.
          Si no hay override, mostramos selector con dropdown + flechas anterior/siguiente. */}
      <div className={`rounded-xl border p-3 ${isViewerMode ? 'bg-amber-50 border-amber-200' : 'bg-white border-slate-200'}`}>
        <div className="flex items-center justify-between gap-3 flex-wrap">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="text-xs uppercase tracking-wide text-slate-500 font-semibold inline-flex items-center gap-1.5">
              <Calendar className="w-3.5 h-3.5" />
              {isViewerMode ? 'Viendo turno guardado · solo lectura' : 'Turno actual'}
            </span>

            {/* Botones flechas — habilitados solo si hay turnos guardados */}
            <button
              onClick={goToPrev}
              disabled={currentIdx < 0 || currentIdx >= savedShifts.length - 1}
              className="px-2 py-1 text-xs bg-white border border-slate-300 rounded hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed"
              title="Turno anterior (cronológico)"
            >
              ◄
            </button>

            {/* Dropdown de turnos guardados */}
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

          {/* Botones de export (V2.2) */}
          <div className="flex items-center gap-2">
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

      {/* V2.6 — Banner de modo admin para indicar que se puede editar haciendo click */}
      {adminMode && (
        <div className="flex items-center gap-2 px-3 py-2 bg-sky-50 border border-sky-200 rounded-lg text-[12px] text-sky-800">
          <Shield className="w-4 h-4 text-sky-600 flex-shrink-0" />
          <div>
            Modo administrador activo. Hacé click en cualquier correctivo o preventivo para editarlo en "Cargar Reporte".
          </div>
        </div>
      )}

      {/* Wrapper que va a ser capturado por html2canvas */}
      <div ref={dashboardRef} className="space-y-3">
      {/* HEADER — V2.0: equipo con wrap multi-línea */}
      <Card className="p-3">
        <div className="grid grid-cols-2 lg:grid-cols-12 gap-3 items-start">
          <div className="col-span-2 lg:col-span-3 flex items-center gap-3">
            <div className="w-10 h-10 rounded-full bg-sky-100 ring-1 ring-sky-200 flex items-center justify-center flex-shrink-0">
              <Wrench className="w-5 h-5 text-sky-600" />
            </div>
            <div>
              <div className="text-[10px] uppercase tracking-widest text-slate-500 font-semibold">Reporte Diario</div>
              <div className="text-sm font-bold text-slate-900 capitalize num">{dateLabel || '—'}</div>
              <div className="text-[10px] text-slate-400 num">{dateShort}</div>
            </div>
          </div>
          <div className="col-span-1 lg:col-span-2 pt-1">
            <div className="text-[10px] uppercase tracking-widest text-slate-500 font-semibold">Turno</div>
            <div className="text-sm font-medium">{report.shift}</div>
          </div>
          <div className="col-span-1 lg:col-span-3 pt-1">
            <div className="text-[10px] uppercase tracking-widest text-slate-500 font-semibold">Responsable</div>
            <div className="text-sm font-medium text-slate-800">{report.responsable || '—'}</div>
          </div>
          <div className="col-span-2 lg:col-span-4 pt-1">
            <div className="text-[10px] uppercase tracking-widest text-slate-500 font-semibold flex items-center gap-1 mb-1">
              <Users className="w-3 h-3" />Equipo ({report.team.length})
            </div>
            {/* V2.0: cambio de truncate a wrap con badges para que se vean TODOS los técnicos */}
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
        </div>
      </Card>

      {/* V2.1 LAYOUT OPCIÓN A:
          - Correctivos a 50% izq (a 2 sub-columnas: Realizadas | Pendientes)
          - Preventivos del Turno arriba derecha
          - Servicios abajo derecha
      */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-3" style={{ minHeight: '500px' }}>

        {/* COL IZQ: CORRECTIVOS — V2.5: 4 sub-secciones (Del turno / Heredados, en cada columna) */}
        <Card className="lg:col-span-6 p-3 flex flex-col lg:overflow-hidden">
          <h3 className="text-sky-600 font-bold text-sm mb-2 inline-flex items-center gap-2 flex-shrink-0">
            <Wrench className="w-4 h-4" />Correctivos del turno ({correctiveActual.length})
          </h3>
          <div className="overflow-visible lg:overflow-auto flex-1 grid grid-cols-1 lg:grid-cols-2 gap-3 lg:max-h-[calc(100vh-280px)]">

            {/* SUB-COL: REALIZADAS — V2.5: dividida en "Del turno" y "Heredados" */}
            <div className="border-b lg:border-b-0 lg:border-r border-slate-200 pb-3 lg:pb-0 lg:pr-3 mb-3 lg:mb-0">
              <div className="text-[10px] uppercase tracking-wide text-emerald-700 font-bold mb-2 inline-flex items-center gap-1 sticky top-0 bg-white z-10 pb-1">
                <CheckCircle2 className="w-3 h-3" />
                Realizadas ({correctivePartitions.realizadosTurno.length + correctivePartitions.realizadosHeredados.length})
              </div>
              <CorrectiveSubsection
                title="Del turno"
                count={correctivePartitions.realizadosTurno.length}
                items={correctivePartitions.realizadosTurno}
                showStateBadge={false}
                showAvanceMark={false}
                adminMode={adminMode}
                onItemClick={adminMode ? (c) => onEditFromDashboard(report, `ot:${c.ot || ""}`) : undefined}
              />
              <CorrectiveSubsection
                title="Heredados realizados"
                count={correctivePartitions.realizadosHeredados.length}
                items={correctivePartitions.realizadosHeredados}
                showStateBadge={false}
                showAvanceMark={false}
                adminMode={adminMode}
                onItemClick={adminMode ? (c) => onEditFromDashboard(report, `ot:${c.ot || ""}`) : undefined}
              />
            </div>

            {/* SUB-COL: PENDIENTES — V2.5: dividida en "Del turno" y "Heredados" */}
            <div>
              <div className="text-[10px] uppercase tracking-wide text-amber-700 font-bold mb-2 inline-flex items-center gap-1 sticky top-0 bg-white z-10 pb-1">
                <AlertTriangle className="w-3 h-3" />
                Pendientes ({correctivePartitions.pendientesTurno.length + correctivePartitions.pendientesHeredados.length})
              </div>
              <CorrectiveSubsection
                title="Del turno"
                count={correctivePartitions.pendientesTurno.length}
                items={correctivePartitions.pendientesTurno}
                showStateBadge={true}
                showAvanceMark={false}
                adminMode={adminMode}
                onItemClick={adminMode ? (c) => onEditFromDashboard(report, `ot:${c.ot || ""}`) : undefined}
              />
              <CorrectiveSubsection
                title="Heredados"
                count={correctivePartitions.pendientesHeredados.length}
                items={correctivePartitions.pendientesHeredados}
                showStateBadge={true}
                showAvanceMark={true}
                adminMode={adminMode}
                onItemClick={adminMode ? (c) => onEditFromDashboard(report, `ot:${c.ot || ""}`) : undefined}
              />
            </div>
          </div>
        </Card>

        {/* COL DER: STACK con PREVENTIVOS arriba y SERVICIOS abajo */}
        <div className="lg:col-span-6 flex flex-col gap-3 lg:max-h-[calc(100vh-220px)]">
          {/* PREVENTIVOS DEL TURNO — V2.6: Card entera clickeable en modo admin */}
          <Card
            className={`p-3 flex flex-col lg:overflow-hidden flex-shrink-0 ${
              adminMode ? 'cursor-pointer hover:bg-sky-50/60 hover:ring-2 hover:ring-sky-200 transition' : ''
            }`}
            onClick={adminMode ? () => onEditFromDashboard(report, 'preventivos') : undefined}
            title={adminMode ? 'Click para editar preventivos del turno' : undefined}
          >
            <h3 className="text-sky-600 font-bold text-sm mb-2 inline-flex items-center gap-2 flex-shrink-0">
              <ListChecks className="w-4 h-4" />Preventivos del Turno
            </h3>
            <div className="lg:overflow-auto lg:max-h-[250px]">
              {/* Asignados / Realizados */}
              <div className="grid grid-cols-2 gap-2 mb-3">
                <div className="bg-slate-50 rounded p-2 text-center">
                  <div className="text-[10px] text-slate-500 uppercase tracking-wide">Asignados</div>
                  <div className="text-2xl font-bold num text-slate-800">
                    {pr.asignados !== '' && pr.asignados != null ? pr.asignados : '—'}
                  </div>
                </div>
                <div className="bg-emerald-50 rounded p-2 text-center">
                  <div className="text-[10px] text-emerald-600 uppercase tracking-wide">Realizados</div>
                  <div className="text-2xl font-bold num text-emerald-700">
                    {pr.realizados !== '' && pr.realizados != null ? pr.realizados : '—'}
                  </div>
                </div>
              </div>

              {/* Detalle por técnico — V2.4: muestra grupos multi-técnicos.
                  Si el grupo tiene varios técnicos, muestra "A · B · C → 4". */}
              <div className="text-[10px] uppercase tracking-wide text-slate-500 font-semibold mb-1.5 inline-flex items-center gap-1">
                <Users className="w-3 h-3" />Por técnico
              </div>
              {(() => {
                const grupos = (pr.porTecnico || []).filter(t => {
                  const tecnicos = t.tecnicos || (t.tecnico ? [t.tecnico] : []);
                  return tecnicos.length > 0;
                });
                if (grupos.length === 0) {
                  return <div className="text-[10px] text-slate-400 italic py-1">Sin detalle por técnico</div>;
                }
                return (
                  <div className="grid grid-cols-2 gap-1">
                    {grupos.map((t, i) => {
                      const tecnicos = t.tecnicos || (t.tecnico ? [t.tecnico] : []);
                      const labelGrupo = tecnicos.join(' · ');
                      return (
                        <div key={i} className="flex items-center justify-between bg-slate-50 rounded px-1.5 py-0.5 min-w-0">
                          <span className="text-[10px] text-slate-700 truncate" title={labelGrupo}>{labelGrupo}</span>
                          <span className="num text-[11px] font-bold text-slate-800 bg-white px-1.5 py-0.5 rounded ring-1 ring-slate-200 flex-shrink-0 ml-1">
                            {t.cantidad || 0}
                          </span>
                        </div>
                      );
                    })}
                  </div>
                );
              })()}
            </div>
          </Card>

          {/* SERVICIOS */}
          <Card className="p-3 flex flex-col lg:overflow-hidden lg:flex-1 min-h-0">
            <h3 className="text-sky-600 font-bold text-sm mb-2 inline-flex items-center gap-2 flex-shrink-0">
              <Activity className="w-4 h-4" />Servicios
            </h3>
            <div className="lg:overflow-auto lg:flex-1 space-y-3">
              {/* Planta de Efluentes y Caldera */}
              <div className="pb-3 border-b border-slate-100">
                <div className="flex items-center justify-between mb-1.5">
                  <div className="text-[10px] uppercase tracking-wide text-slate-500 font-semibold inline-flex items-center gap-1">
                    <Flame className="w-3 h-3 text-orange-500" />Planta de Efluentes y Caldera
                  </div>
                  <StatePill state={p.estado} />
                </div>
                <div className="text-xs text-slate-700 mb-2">
                  {(p.tecnicos && p.tecnicos.length > 0) ? p.tecnicos.join(' · ') : '—'}
                </div>

                {/* PTEL */}
                <div className="text-[9px] font-bold text-orange-700 uppercase tracking-wider mb-1">PTEL</div>
                <div className="grid grid-cols-3 gap-1 text-[10px] mb-2">
                  {[
                    ['Caudal m³/h', p.caudal],
                    ['Vacío', p.vacio],
                    ['ΔT °C', p.deltaT],
                    ['% TK1', p.tk1],
                    ['% TK2', p.tk2],
                    ['% TK7', p.tk7]
                  ].map(([l, v]) => (
                    <div key={l} className="bg-orange-50/50 rounded p-1 text-center">
                      <div className="text-slate-500 uppercase">{l}</div>
                      <div className="font-bold num text-slate-800">{v !== '' && v != null ? v : '—'}</div>
                    </div>
                  ))}
                </div>

                {/* CALDERA + ABLANDADORES */}
                <div className="grid grid-cols-2 gap-2">
                  <div className="bg-red-50/50 rounded p-1.5">
                    <div className="text-[9px] font-bold text-red-700 uppercase tracking-wider mb-1">Caldera</div>
                    <div className="grid grid-cols-2 gap-1 text-[10px]">
                      <div className="text-center">
                        <div className="text-slate-500 uppercase">Cond. mS</div>
                        <div className="font-bold num">{p.conductividadCaldera !== '' && p.conductividadCaldera != null ? p.conductividadCaldera : '—'}</div>
                      </div>
                      <div className="text-center">
                        <div className="text-slate-500 uppercase">pH</div>
                        <div className="font-bold num">{p.pHCaldera !== '' && p.pHCaldera != null ? p.pHCaldera : '—'}</div>
                      </div>
                    </div>
                  </div>
                  <div className="bg-blue-50/50 rounded p-1.5">
                    <div className="text-[9px] font-bold text-blue-700 uppercase tracking-wider mb-1">Ablandadores</div>
                    <div className="grid grid-cols-2 gap-1 text-[10px]">
                      <div className="text-center">
                        <div className="text-slate-500 uppercase">Cond. mS</div>
                        <div className="font-bold num">{p.conductividadAblandador !== '' && p.conductividadAblandador != null ? p.conductividadAblandador : '—'}</div>
                      </div>
                      <div className="text-center">
                        <div className="text-slate-500 uppercase">pH</div>
                        <div className="font-bold num">{p.pHAblandador !== '' && p.pHAblandador != null ? p.pHAblandador : '—'}</div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Compresores y Grupos */}
              <div className="grid grid-cols-2 gap-3 pb-3 border-b border-slate-100">
                <div>
                  <div className="text-[10px] uppercase tracking-wide text-slate-500 font-semibold mb-1.5 inline-flex items-center gap-1">
                    <Cog className="w-3 h-3" />Compresores
                  </div>
                  <div className="space-y-0.5">
                    {report.servicios.compresores.map(c => (
                      <div key={c.code} className="flex justify-between items-center text-[10px]">
                        <span className="num text-slate-700">{c.code}</span>
                        <StatePill state={c.state} />
                      </div>
                    ))}
                  </div>
                </div>
                <div>
                  <div className="text-[10px] uppercase tracking-wide text-slate-500 font-semibold mb-1.5 inline-flex items-center gap-1">
                    <Zap className="w-3 h-3" />G. Electrógenos
                  </div>
                  <div className="space-y-0.5">
                    {report.servicios.gruposElectrogenos.map(g => (
                      <div key={g.code} className="flex justify-between items-center text-[10px]">
                        <span className="num text-slate-700">{g.code}</span>
                        <StatePill state={g.state} />
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Cisternas */}
              <div className="pb-3 border-b border-slate-100">
                <div className="text-[10px] uppercase tracking-wide text-slate-500 font-semibold mb-1.5 inline-flex items-center gap-1">
                  <Beaker className="w-3 h-3" />Cisternas
                </div>
                <div className="grid grid-cols-2 gap-2 text-xs">
                  <div className="flex items-center gap-2">
                    <span className="text-slate-500 text-[10px]">Nivel:</span>
                    <span className="font-medium text-slate-800">{report.servicios.cisternas.nivel || '—'}</span>
                  </div>
                  <div className="flex items-center justify-end gap-2">
                    <StatePill state={report.servicios.cisternas.estado} />
                  </div>
                </div>
              </div>

              {/* Agua de Pozo */}
              <div className="pb-3 border-b border-slate-100">
                <div className="text-[10px] uppercase tracking-wide text-slate-500 font-semibold mb-1.5 inline-flex items-center gap-1">
                  <Beaker className="w-3 h-3 text-blue-500" />Agua de Pozo
                </div>
                <div className="grid grid-cols-2 gap-1 text-[10px]">
                  {[
                    ['Cloro Pozo 3', report.servicios.aguaPozo?.cloroPozo3],
                    ['Cloro Pozo 6', report.servicios.aguaPozo?.cloroPozo6]
                  ].map(([l, v]) => (
                    <div key={l} className="bg-slate-50 rounded p-1.5 text-center">
                      <div className="text-slate-500 uppercase">{l}</div>
                      <div className="text-sm font-bold num text-slate-800">{v !== '' && v != null ? v : '—'}</div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Proveedores */}
              {report.servicios.proveedores.length > 0 && (
                <div className="pb-3 border-b border-slate-100">
                  <div className="text-[10px] uppercase tracking-wide text-slate-500 font-semibold mb-1.5 inline-flex items-center gap-1">
                    <Building2 className="w-3 h-3" />Proveedores
                  </div>
                  <div className="space-y-0.5">
                    {report.servicios.proveedores.map((pv, i) => (
                      <div key={i} className="grid grid-cols-3 gap-1 text-[11px]">
                        <div className="font-medium text-slate-700">{pv.provider}</div>
                        <div className="col-span-2 text-slate-600">{pv.task}</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Comentarios */}
              <div>
                <div className="text-[10px] uppercase tracking-wide text-slate-500 font-semibold mb-1.5 inline-flex items-center gap-1">
                  <FileText className="w-3 h-3" />Comentarios
                </div>
                {report.comments.length === 0 ? <div className="text-[10px] text-slate-400 italic">Sin comentarios</div> :
                  <div className="space-y-1">
                    {report.comments.map((c, i) => (
                      <div key={i} className={`text-[11px] p-1.5 rounded ${c.priority === 'Urgente' ? 'bg-red-50 border border-red-200' : 'bg-slate-50'}`}>
                        <span className="text-slate-700">{c.text}</span>
                        {c.priority === 'Urgente' && <span className="ml-1.5 text-[9px] font-bold text-red-700 uppercase">Urgente</span>}
                      </div>
                    ))}
                  </div>}
              </div>
            </div>
          </Card>
        </div>
      </div>
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
                `${ultimoDia.turnos.length} turno${ultimoDia.turnos.length === 1 ? '' : 's'} · ${formatDateShort(ultimoDia.fechaBase)}`}
            </span>
          </div>
          {ultimoDia.turnos.length === 0 ? (
            <EmptyHint>Sin reportes recientes</EmptyHint>
          ) : (
            <>
              <div className="text-[11px] text-slate-500 mb-2">
                {ultimoDia.turnos.length === 1
                  ? `Turno del día: ${ultimoDia.turnos.map(t => t.shift).join(', ')}`
                  : `Turnos del día (${ultimoDia.turnos.length}): ${ultimoDia.turnos.map(t => t.shift).join(', ')}`}
              </div>
              <div className="grid grid-cols-2 gap-2">
                <MiniKPI label="Correctivos generados" value={ultimoDia.correctivosGenerados} color="orange" />
                <MiniKPI label="Correctivos realizados" value={ultimoDia.correctivosRealizados} color="emerald" />
                <MiniKPI label="Preventivos asignados" value={ultimoDia.preventivosAsignados} color="sky" />
                <MiniKPI label="Preventivos realizados" value={ultimoDia.preventivosRealizados} color="emerald" />
              </div>
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
                `${formatDateShort(finde.viernes)} (Noche) → ${formatDateShort(finde.domingo)}`}
            </span>
          </div>
          {finde.turnos.length === 0 ? (
            <EmptyHint>Sin reportes del último FDS cerrado</EmptyHint>
          ) : (
            <>
              <div className="text-[11px] text-slate-500 mb-2">
                Viernes Noche + Sábado completo + Domingo completo · {finde.turnos.length} turno{finde.turnos.length === 1 ? '' : 's'} cargado{finde.turnos.length === 1 ? '' : 's'}
              </div>
              <div className="grid grid-cols-2 gap-2">
                <MiniKPI label="Correctivos generados" value={finde.correctivosGenerados} color="orange" />
                <MiniKPI label="Correctivos realizados" value={finde.correctivosRealizados} color="emerald" />
                <MiniKPI label="Preventivos asignados" value={finde.preventivosAsignados} color="sky" />
                <MiniKPI label="Preventivos realizados" value={finde.preventivosRealizados} color="emerald" />
              </div>
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

      {/* CHARTS GRID */}
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
                <Pie data={stats.stateDistVigente} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80}
                  label={(e) => `${e.name}: ${e.value}`} labelLine={false}>
                  {stats.stateDistVigente.map((entry, i) => (
                    <Cell key={i} fill={entry.name === 'Realizada' ? '#10b981' : entry.name === 'En Curso' ? '#f59e0b' : '#ef4444'} />
                  ))}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          }
        </Card>

        {/* GRÁFICO 2 — Estado al cierre del período filtrado */}
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
                <Pie data={stats.stateDistPeriodo} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80}
                  label={(e) => `${e.name}: ${e.value}`} labelLine={false}>
                  {stats.stateDistPeriodo.map((entry, i) => (
                    <Cell key={i} fill={entry.name === 'Realizada' ? '#10b981' : entry.name === 'En Curso' ? '#f59e0b' : '#ef4444'} />
                  ))}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          }
        </Card>

        {/* V2.7 — Bloque admin-only: métricas de performance por turno.
            Incluye las dos tarjetas existentes (Distribución por turno, Carga por técnico)
            y las dos nuevas de V2.7 (Pendientes por origen, Heredadas cerradas). */}
        {adminMode && (<>
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

        {/* V2.7 — OTs dejadas pendientes por turno de origen.
            Por cada OT en history (excluyendo legacy sin formato XXX-YYYYY),
            se toma el turno donde se creó (createdInShift) y su último estado global.
            Si el último estado es Sin Iniciar o En Curso, suma al turno de origen.
            Solo se consideran OTs creadas dentro del rango [startStr, endStr]. */}
        <ShiftRankingCard
          title="OTs dejadas pendientes por turno de origen"
          tooltip="Cuenta OTs creadas en cada turno cuyo último estado global sigue en Sin Iniciar o En Curso. Filtro por fecha aplica al turno de creación. Excluye OTs legacy. Estadísticas desde 20/05/2026."
          data={shiftPerformance?.pendingByOriginShift}
          icon={AlertTriangle}
          colorBar="#ef4444"
        />

        {/* V2.7 — OTs heredadas cerradas por turno.
            Una transición Sin Iniciar/En Curso → Realizada en un turno distinto
            al de creación cuenta para el turno donde ocurrió el cierre.
            Solo se cuentan cierres ocurridos dentro del rango [startStr, endStr]. */}
        <ShiftRankingCard
          title="OTs heredadas cerradas por turno"
          tooltip="Cuenta cierres (estado → Realizada) de OTs creadas en turnos anteriores. Filtro por fecha aplica al turno del cierre. Excluye OTs legacy. Estadísticas desde 20/05/2026."
          data={shiftPerformance?.closedByShift}
          icon={CheckCircle2}
          colorBar="#10b981"
        />
        </>)}

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
  const SHIFT_ORDER = { 'Mañana': 0, 'Tarde': 1, 'Noche': 2 };
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

  const fmt = (d) => d.toISOString().slice(0, 10);
  const viernesStr = fmt(viernes), sabadoStr = fmt(sabado), domingoStr = fmt(domingo);

  // Filtrar reportes que correspondan:
  //  - viernes Noche
  //  - sábado Mañana, Tarde, Noche
  //  - domingo Mañana, Tarde, Noche
  const turnos = history.filter(r => {
    if (r.date === viernesStr && r.shift === 'Noche') return true;
    if (r.date === sabadoStr) return true;
    if (r.date === domingoStr) return true;
    return false;
  });

  let correctivosGenerados = 0, correctivosRealizados = 0;
  let preventivosAsignados = 0, preventivosRealizados = 0;

  // V2.3 — Deduplicación: cada OT (por número) cuenta una sola vez.
  //   - "correctivosGenerados" = OTs únicas que aparecen en estos turnos
  //   - "correctivosRealizados" = OTs únicas cuyo estado FINAL en estos turnos es 'Realizada'
  // Para OTs sin número, contamos cada aparición (caso borde).
  const sortedTurnos = [...turnos].sort((a, b) => reportSortKey(a).localeCompare(reportSortKey(b)));
  const uniqueByOT = new Map();
  let countWithoutOT = 0, realizedWithoutOT = 0;
  sortedTurnos.forEach(r => {
    (r.corrective || []).forEach(c => {
      const key = (c.ot || '').trim();
      if (!key) {
        countWithoutOT++;
        if (c.state === 'Realizada') realizedWithoutOT++;
        return;
      }
      uniqueByOT.set(key, c); // sobrescribe con la versión más reciente
    });
  });
  correctivosGenerados = uniqueByOT.size + countWithoutOT;
  correctivosRealizados = [...uniqueByOT.values()].filter(c => c.state === 'Realizada').length + realizedWithoutOT;

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
  if (history.length === 0) {
    return { fechaBase: '', turnos: [], correctivosGenerados: 0, correctivosRealizados: 0, preventivosAsignados: 0, preventivosRealizados: 0 };
  }

  // Encontrar la fecha más reciente con al menos un reporte
  const fechasOrdenadas = [...new Set(history.map(r => r.date))].sort((a, b) => b.localeCompare(a));
  if (fechasOrdenadas.length === 0) {
    return { fechaBase: '', turnos: [], correctivosGenerados: 0, correctivosRealizados: 0, preventivosAsignados: 0, preventivosRealizados: 0 };
  }

  // Tomar el último día con datos como "fecha base"
  const fechaBase = fechasOrdenadas[0];

  // Tomar todos los turnos de ese día (Mañana + Tarde + Noche)
  const turnos = history
    .filter(r => r.date === fechaBase)
    .sort((a, b) => shiftOrder(a.shift).localeCompare(shiftOrder(b.shift)));

  let correctivosGenerados = 0, correctivosRealizados = 0;
  let preventivosAsignados = 0, preventivosRealizados = 0;

  // V2.3 — Deduplicación: misma lógica que computeWeekendStats
  const uniqueByOT2 = new Map();
  let countWithoutOT2 = 0, realizedWithoutOT2 = 0;
  turnos.forEach(r => {
    (r.corrective || []).forEach(c => {
      const key = (c.ot || '').trim();
      if (!key) {
        countWithoutOT2++;
        if (c.state === 'Realizada') realizedWithoutOT2++;
        return;
      }
      uniqueByOT2.set(key, c);
    });
  });
  correctivosGenerados = uniqueByOT2.size + countWithoutOT2;
  correctivosRealizados = [...uniqueByOT2.values()].filter(c => c.state === 'Realizada').length + realizedWithoutOT2;

  turnos.forEach(r => {
    preventivosAsignados += Number(r.preventivosResumen?.asignados) || 0;
    preventivosRealizados += Number(r.preventivosResumen?.realizados) || 0;
  });

  return {
    fechaBase,
    turnos,
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
  // Se calcula sobre TODO el histórico (no el rango filtrado): el estado de cada OT
  // según su aparición MÁS RECIENTE en cualquier reporte. Así "pendientes" refleja
  // lo que sigue abierto al día de hoy (coincide con el carry-over del Dashboard),
  // sin importar el período seleccionado en Estadísticas.
  const allReportsSorted = [...history].sort((a, b) =>
    `${a.date}-${shiftOrder(a.shift)}`.localeCompare(`${b.date}-${shiftOrder(b.shift)}`)
  );
  const latestStateByOT = new Map();  // ot# -> estado de la aparición más reciente
  const noOTStatesVigente = [];       // OTs sin número (cada aparición cuenta)
  allReportsSorted.forEach(r => {
    (r.corrective || []).forEach(c => {
      const key = (c.ot || '').trim();
      if (!key) { noOTStatesVigente.push(c.state); return; }
      latestStateByOT.set(key, c.state);  // iteramos ordenado: la última asignación es la más reciente
    });
  });
  const stateDistVigente = { 'Sin Iniciar': 0, 'En Curso': 0, 'Realizada': 0 };
  latestStateByOT.forEach(st => { if (st in stateDistVigente) stateDistVigente[st]++; });
  noOTStatesVigente.forEach(st => { if (st in stateDistVigente) stateDistVigente[st]++; });

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
    buckets[key].preventivos += (r.preventive || []).length;
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

  // Preventivos: NO se deduplican (cada turno hace su propio trabajo preventivo)
  let totalPreventives = 0, urgent = 0;
  filtered.forEach(r => {
    (r.preventive || []).forEach(p => {
      totalPreventives++;
      (p.technicians || []).forEach(t => {
        techCount[t] = techCount[t] || { correctivos: 0, preventivos: 0 };
        techCount[t].preventivos++;
      });
      if (shiftCount[r.shift]) shiftCount[r.shift].preventivos++;
    });
    (r.comments || []).forEach(c => { if (c.priority === 'Urgente') urgent++; });

    // V2.4 — "Carga por técnico" suma también el detalle del Resumen Preventivos.
    // Opción C: cada grupo {tecnicos:[a,b,c], cantidad:N} suma N a cada técnico individual.
    // Ej: {tecnicos:['Juan','Pedro'], cantidad: 4} → Juan +4 y Pedro +4.
    // (La cantidad para validación cruzada con "Realizados" sigue siendo N una sola vez,
    // eso ya está cubierto en validateReport.)
    (r.preventivosResumen?.porTecnico || []).forEach(grupo => {
      const tecnicos = grupo.tecnicos || (grupo.tecnico ? [grupo.tecnico] : []);
      const cantidad = Number(grupo.cantidad) || 0;
      if (cantidad > 0) {
        tecnicos.forEach(t => {
          techCount[t] = techCount[t] || { correctivos: 0, preventivos: 0 };
          techCount[t].preventivos += cantidad;
        });
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
    preventives: history.reduce((s, r) => s + (r.preventive?.length || 0), 0),
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
                  <th className="pb-2 font-medium text-right">Preventivos</th>
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
                    <td className="py-2 text-right num">{r.preventive?.length || 0}</td>
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

import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import {
  ClipboardList, BarChart3, Download, Plus, Trash2, Save, Calendar, Users,
  Wrench, Activity, FileSpreadsheet, CheckCircle2, AlertTriangle, Building2,
  HardHat, Beaker, ListChecks, ChevronDown, X, FileText, TrendingUp, Flame,
  Cog, Zap, Filter, Search, Cloud, CloudOff, RefreshCw, Settings
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
// CATÁLOGOS (matching the Excel template)
// ═══════════════════════════════════════════════════════════════════
const RESPONSABLES = [
  { id: 1, name: 'Juan Martín Alasia' },
  { id: 2, name: 'Luciano Fioretti' },
  { id: 3, name: 'Gustavo Pare' }
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
// EMPTY REPORT
// ═══════════════════════════════════════════════════════════════════
const emptyReport = () => ({
  date: new Date().toISOString().slice(0, 10),
  shift: 'Mañana',
  responsable: '',
  team: [],                    // [string] tech names
  corrective: [],              // [{ot, equipoCodigo, task, technicians:[], state}]
  preventive: [],              // [{codigoTarea, equipoCodigo, equipoDescripcion, task, comments, otCorrectivaAsociada, technicians:[], frequency}]
  servicios: {
    plantaCaldera: {
      tecnicos: [],            // [string] foguistas del turno (multi-select)
      estado: 'Operativa',
      caudal: '', pH: '', ablandador: '', deltaT: '',
      vacio: '', tk1: '', tk2: '', tkEmergencia: '', tk4: ''
    },
    compresores: COMPRESORES.map(c => ({ code: c, state: 'Operativo' })),
    gruposElectrogenos: GRUPOS_ELECTROGENOS.map(g => ({ code: g, state: 'Operativo' })),
    cisternas: { nivel: 'Alto', estado: 'Ingreso Normal' },
    aguaPozo: { cloroPozo3: '', cloroPozo6: '' },
    proveedores: []            // [{provider, task}]
  },
  comments: []                 // [{text, priority}]
});

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
  }
};

// ═══════════════════════════════════════════════════════════════════
// UI PRIMITIVES
// ═══════════════════════════════════════════════════════════════════
const Card = ({ children, className = '' }) => (
  <div className={`bg-white border border-slate-200 rounded-xl shadow-sm ${className}`}>{children}</div>
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

  const refresh = useCallback(async () => {
    try {
      setConnError('');
      const data = await storage.list();
      setHistory(data);
    } catch (e) {
      setConnError(e.message || 'Error de conexión');
      console.error(e);
    }
  }, []);

  useEffect(() => { (async () => { await refresh(); setLoading(false); })(); }, [refresh]);

  const saveReport = async () => {
    if (!report.date || !report.shift) { setSaveMsg('Falta fecha o turno'); return; }
    setSaving(true);
    setSaveMsg('Guardando…');
    try {
      await storage.save(report);
      await refresh();
      setSaveMsg('✓ Reporte guardado');
      setTimeout(() => setSaveMsg(''), 2500);
    } catch (e) {
      setSaveMsg(`Error: ${e.message}`);
    }
    setSaving(false);
  };

  // ── Excel exports (matching template format) ────────────────────
  const exportFull = () => {
    if (!history.length) { alert('No hay reportes guardados.'); return; }
    const wb = XLSX.utils.book_new();

    // 1. Tecnicos presentes
    const tecPres = [];
    let rowId = 1;
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
        Estado: c.state || ''
      });
    }));
    addSheet(wb, corr.length ? corr : [{ OrdenID: '', OrdenNumero: '', SectorID: '', EquipoID: '', EquipoCodigo: '', OrdenTecnicoAsignado: '', Turno: '', FechaRealizacion: '', Descripcion: '', Estado: '' }], 'Correctivos');

    // 3. Preventivos
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

    // 7. Planta de efluentes (one row per foguista per shift to keep the template structure)
    const planta = [];
    let plId = 1;
    history.forEach(r => {
      const p = r.servicios?.plantaCaldera;
      if (!p) return;
      // Backward compatibility: some old reports may have `tecnico` (single)
      const foguistas = (p.tecnicos && p.tecnicos.length > 0)
        ? p.tecnicos
        : (p.tecnico ? [p.tecnico] : ['']);
      // Only emit row if there's some content
      if (foguistas[0] === '' && !p.estado) return;
      foguistas.forEach(name => {
        planta.push({
          CalderaRegistroID: plId++,
          CalderaRegistroTurno: r.shift,
          CalderaTecnicoNombre: name || '',
          TecnicoID: findTecnicoId(name),
          CalderaFechaControl: r.date,
          CalderaEstado: p.estado || '',
          CalderaCaudal: p.caudal !== '' ? Number(p.caudal) : '',
          CalderaPH: p.pH !== '' ? Number(p.pH) : '',
          Ablandador: p.ablandador !== '' ? Number(p.ablandador) : '',
          DeltaTemperatura: p.deltaT !== '' ? Number(p.deltaT) : '',
          NivelVacio: p.vacio !== '' ? Number(p.vacio) : '',
          TK1: p.tk1 !== '' ? Number(p.tk1) : '',
          TK2: p.tk2 !== '' ? Number(p.tk2) : '',
          TKEmergencia: p.tkEmergencia !== '' ? Number(p.tkEmergencia) : '',
          TK4: p.tk4 !== '' ? Number(p.tk4) : ''
        });
      });
    });
    addSheet(wb, planta.length ? planta : [{ CalderaRegistroID: '', CalderaRegistroTurno: '', CalderaTecnicoNombre: '', TecnicoID: '', CalderaFechaControl: '', CalderaEstado: '', CalderaCaudal: '', CalderaPH: '', Ablandador: '', DeltaTemperatura: '', NivelVacio: '', TK1: '', TK2: '', TKEmergencia: '', TK4: '' }], 'Planta de efluentes');

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
      history.forEach(r => (r.corrective || []).forEach(c => rows.push({
        OrdenID: i++,
        OrdenNumero: c.ot || '',
        SectorID: '',
        EquipoID: '',
        EquipoCodigo: c.equipoCodigo || '',
        OrdenTecnicoAsignado: (c.technicians || []).join(', '),
        Turno: r.shift,
        FechaRealizacion: r.date,
        Descripcion: c.task || '',
        Estado: c.state || ''
      })));
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
    }
  };

  return (
    <div className="min-h-screen bg-slate-50" style={{ fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap');
        body { font-family: 'Plus Jakarta Sans', system-ui, sans-serif; }
        .num { font-family: 'JetBrains Mono', monospace; font-feature-settings: 'tnum'; }
      `}</style>

      {/* HEADER */}
      <header className="bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900 text-white border-b border-slate-800">
        <div className="max-w-[1600px] mx-auto px-6 py-4 flex items-center justify-between flex-wrap gap-3">
          <div className="flex items-center gap-4">
            <div className="w-11 h-11 rounded-xl bg-sky-500/20 ring-1 ring-sky-400/40 flex items-center justify-center">
              <HardHat className="w-5 h-5 text-sky-300" />
            </div>
            <div>
              <h1 className="text-lg font-bold tracking-tight">Reporte Diario de Mantenimiento</h1>
              <p className="text-[11px] text-slate-300 mt-0.5">Sistema integral · carga, dashboard, estadísticas y exportación</p>
            </div>
          </div>
          <div className="flex items-center gap-4 text-xs">
            <div className="flex items-center gap-2">
              {supabaseConfigured ? (
                connError
                  ? <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-red-500/20 text-red-200 rounded ring-1 ring-red-400/30">
                    <CloudOff className="w-3.5 h-3.5" />Sin conexión
                  </span>
                  : <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-emerald-500/20 text-emerald-200 rounded ring-1 ring-emerald-400/30">
                    <Cloud className="w-3.5 h-3.5" />Supabase conectado
                  </span>
              ) : (
                <span className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-amber-500/20 text-amber-200 rounded ring-1 ring-amber-400/30">
                  <Settings className="w-3.5 h-3.5" />Modo local
                </span>
              )}
              <button onClick={refresh} className="p-1.5 hover:bg-white/10 rounded transition" title="Refrescar">
                <RefreshCw className="w-4 h-4" />
              </button>
            </div>
            <div className="text-right">
              <div className="num text-sm font-semibold text-white capitalize">
                {new Date().toLocaleDateString('es-AR', { weekday: 'long', day: 'numeric', month: 'long' })}
              </div>
              <div className="text-slate-300">{history.length} {history.length === 1 ? 'reporte' : 'reportes'}</div>
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
        {!loading && tab === 'form' && <FormView report={report} setReport={setReport} onSave={saveReport} saveMsg={saveMsg} saving={saving} history={history} />}
        {!loading && tab === 'dashboard' && <DashboardView report={report} />}
        {!loading && tab === 'stats' && <StatsView history={history} />}
        {!loading && tab === 'history' && <HistoryView history={history}
          onExportCorrectives={() => exportSingleSheet('correctivos')}
          onExportPreventives={() => exportSingleSheet('preventivos')}
          onExportFull={exportFull} />}
      </main>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// FORM VIEW
// ═══════════════════════════════════════════════════════════════════
function FormView({ report, setReport, onSave, saveMsg, saving, history }) {
  const update = (patch) => setReport(r => ({ ...r, ...patch }));
  const updateList = (key, fn) => setReport(r => ({ ...r, [key]: fn(r[key]) }));
  const updateServicios = (patch) => setReport(r => ({ ...r, servicios: { ...r.servicios, ...patch } }));
  const [loadInfo, setLoadInfo] = useState('');
  const initialPendingApplied = useRef(false);

  // Compute correctivos still pending: walk all reports saved up to and including
  // (date, shift), keep the latest version of each OT by number, and return only
  // those whose latest state is "Sin Iniciar" or "En Curso". This is what should
  // carry over to a clean / new turn — pending OTs persist until "Realizada".
  const computePending = (date, shift) => {
    const upTo = history
      .filter(r => (r.date < date) || (r.date === date && shiftOrder(r.shift) <= shiftOrder(shift)))
      .sort((a, b) => (a.date + shiftOrder(a.shift)).localeCompare(b.date + shiftOrder(b.shift)));
    const latestByOT = new Map();
    upTo.forEach(r => (r.corrective || []).forEach(c => {
      if (c.ot) latestByOT.set(c.ot, c);
    }));
    return [...latestByOT.values()]
      .filter(c => c.state === 'Sin Iniciar' || c.state === 'En Curso')
      .map(c => ({ ...c }));
  };

  // When date or shift changes, look up history:
  //  - if a saved report exists for that date+shift -> load it as-is
  //  - otherwise, build a new empty report with pending correctivos pre-loaded
  const setDateShift = (newDate, newShift) => {
    const existing = history.find(r => r.date === newDate && r.shift === newShift);
    if (existing) {
      setReport({ ...emptyReport(), ...existing, servicios: { ...emptyReport().servicios, ...(existing.servicios || {}) } });
      setLoadInfo(`✓ Reporte cargado del histórico (${newDate} - ${newShift})`);
      setTimeout(() => setLoadInfo(''), 4000);
      return;
    }
    const pending = computePending(newDate, newShift);
    setReport({ ...emptyReport(), date: newDate, shift: newShift, corrective: pending });
    if (pending.length > 0) {
      setLoadInfo(`↻ ${pending.length} correctivo${pending.length === 1 ? '' : 's'} pendiente${pending.length === 1 ? '' : 's'} traído${pending.length === 1 ? '' : 's'} del turno anterior`);
      setTimeout(() => setLoadInfo(''), 5000);
    }
  };

  // "Limpiar": wipe the form for the current date+shift but preserve pending OTs
  const cleanForm = () => {
    const pending = computePending(report.date, report.shift);
    setReport({ ...emptyReport(), date: report.date, shift: report.shift, corrective: pending });
    if (pending.length > 0) {
      setLoadInfo(`↻ Formulario limpio. Se mantienen ${pending.length} correctivo${pending.length === 1 ? '' : 's'} pendiente${pending.length === 1 ? '' : 's'}.`);
    } else {
      setLoadInfo('✓ Formulario limpio');
    }
    setTimeout(() => setLoadInfo(''), 4000);
  };

  // On first history-loaded mount, auto-apply pending correctivos if the form
  // is still untouched. This ensures pending OTs are visible immediately.
  useEffect(() => {
    if (initialPendingApplied.current) return;
    if (history.length === 0) return;
    // Only auto-apply if form is empty (no user edits yet)
    const isEmpty = !report.responsable && report.team.length === 0 &&
      report.corrective.length === 0 && report.preventive.length === 0 &&
      report.comments.length === 0;
    if (!isEmpty) { initialPendingApplied.current = true; return; }
    const existing = history.find(r => r.date === report.date && r.shift === report.shift);
    if (existing) {
      // already saved — load it
      setReport({ ...emptyReport(), ...existing, servicios: { ...emptyReport().servicios, ...(existing.servicios || {}) } });
    } else {
      const pending = computePending(report.date, report.shift);
      if (pending.length > 0) {
        setReport(r => ({ ...r, corrective: pending }));
        setLoadInfo(`↻ ${pending.length} correctivo${pending.length === 1 ? '' : 's'} pendiente${pending.length === 1 ? '' : 's'} del turno anterior`);
        setTimeout(() => setLoadInfo(''), 5000);
      }
    }
    initialPendingApplied.current = true;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [history]);

  // Only show technicians selected in Equipo del Turno for assignment to OTs
  const teamOptions = report.team.length > 0 ? report.team : TECNICO_NAMES;

  return (
    <div className="space-y-5">
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

      {/* CORRECTIVOS */}
      <Card className="p-5">
        <div className="flex items-center justify-between mb-4">
          <SectionTitle icon={Wrench} accent="orange">Mantenimiento Correctivo</SectionTitle>
          <button onClick={() => updateList('corrective', l => [...l, { ot: '', equipoCodigo: '', task: '', technicians: [], state: 'Sin Iniciar' }])}
            className={`${buttonCls} bg-orange-50 text-orange-700 hover:bg-orange-100`}>
            <Plus className="w-4 h-4" />Agregar OT
          </button>
        </div>
        {report.corrective.length === 0 && <EmptyHint>Sin órdenes de trabajo correctivas.</EmptyHint>}
        <div className="space-y-3">
          {report.corrective.map((c, i) => (
            <div key={i} className="border border-slate-200 rounded-lg p-3 bg-slate-50/40">
              <div className="grid grid-cols-12 gap-2 mb-2">
                <Field label="N° OT" className="col-span-2">
                  <input className={`${inputCls} num`} placeholder="OT-XXXX" value={c.ot}
                    onChange={e => updateList('corrective', l => l.map((x, j) => j === i ? { ...x, ot: e.target.value } : x))} />
                </Field>
                <Field label="Equipo / Sector" className="col-span-3">
                  <input className={inputCls} value={c.equipoCodigo}
                    onChange={e => updateList('corrective', l => l.map((x, j) => j === i ? { ...x, equipoCodigo: e.target.value } : x))} />
                </Field>
                <Field label="Estado" className="col-span-2">
                  <select className={`${inputCls} font-semibold ${c.state === 'Sin Iniciar' ? 'text-red-600' : c.state === 'En Curso' ? 'text-amber-600' : 'text-emerald-600'}`}
                    value={c.state}
                    onChange={e => updateList('corrective', l => l.map((x, j) => j === i ? { ...x, state: e.target.value } : x))}>
                    {ESTADOS_OT.map(s => <option key={s} value={s}>{s}</option>)}
                  </select>
                </Field>
                <Field label="Técnico/s asignado/s" className="col-span-4">
                  <MultiSelect options={teamOptions} value={c.technicians}
                    onChange={vals => updateList('corrective', l => l.map((x, j) => j === i ? { ...x, technicians: vals } : x))}
                    placeholder={report.team.length === 0 ? 'Cargá primero el equipo del turno' : 'Seleccionar…'} />
                </Field>
                <div className="col-span-1 flex items-end justify-end">
                  <button onClick={() => updateList('corrective', l => l.filter((_, j) => j !== i))}
                    className="text-red-500 hover:text-red-700 hover:bg-red-50 rounded p-2 transition">
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </div>
              <Field label="Tarea / descripción">
                <textarea rows={2} className={inputCls} value={c.task}
                  onChange={e => updateList('corrective', l => l.map((x, j) => j === i ? { ...x, task: e.target.value } : x))} />
              </Field>
            </div>
          ))}
        </div>
      </Card>

      {/* PREVENTIVOS */}
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
          {report.preventive.map((p, i) => (
            <div key={i} className="border border-slate-200 rounded-lg p-3 bg-slate-50/40">
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
                <Field label="Técnico/s asignado/s" className="col-span-2">
                  <MultiSelect options={teamOptions} value={p.technicians}
                    onChange={vals => updateList('preventive', l => l.map((x, j) => j === i ? { ...x, technicians: vals } : x))}
                    placeholder={report.team.length === 0 ? 'Cargá el equipo' : 'Seleccionar…'} />
                </Field>
                <div className="col-span-1 flex items-end justify-end">
                  <button onClick={() => updateList('preventive', l => l.filter((_, j) => j !== i))}
                    className="text-red-500 hover:text-red-700 hover:bg-red-50 rounded p-2 transition">
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </div>
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
          ))}
        </div>
      </Card>

      {/* SERVICIOS */}
      <Card className="p-5">
        <SectionTitle icon={Activity} accent="violet">Servicios</SectionTitle>

        {/* PLANTA DE EFLUENTES Y CALDERA */}
        <div className="mb-5">
          <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
            <Flame className="w-4 h-4 text-orange-500" />Planta de Efluentes y Caldera
          </h3>
          <div className="grid grid-cols-12 gap-3">
            <Field label="Técnicos (foguistas)" className="col-span-4">
              <MultiSelect options={FOGUISTAS} value={report.servicios.plantaCaldera.tecnicos || []}
                onChange={vals => updateServicios({ plantaCaldera: { ...report.servicios.plantaCaldera, tecnicos: vals } })}
                placeholder="Seleccionar foguistas…" />
            </Field>
            <Field label="Estado" className="col-span-2">
              <select className={`${inputCls} font-semibold ${report.servicios.plantaCaldera.estado === 'Operativa' ? 'text-emerald-600' : 'text-red-600'}`}
                value={report.servicios.plantaCaldera.estado}
                onChange={e => updateServicios({ plantaCaldera: { ...report.servicios.plantaCaldera, estado: e.target.value } })}>
                {ESTADOS_PLANTA.map(s => <option key={s}>{s}</option>)}
              </select>
            </Field>
            {[
              ['caudal', 'Caudal'], ['pH', 'pH'], ['ablandador', 'Ablandador'],
              ['deltaT', 'ΔT'], ['vacio', 'Vacío'], ['tk1', 'TK1'],
              ['tk2', 'TK2'], ['tkEmergencia', 'TK Emerg.'], ['tk4', 'TK4']
            ].map(([k, label]) => (
              <Field key={k} label={label} className="col-span-2">
                <input type="number" step="any" className={`${inputCls} num`} value={report.servicios.plantaCaldera[k]}
                  onChange={e => updateServicios({ plantaCaldera: { ...report.servicios.plantaCaldera, [k]: e.target.value } })} />
              </Field>
            ))}
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
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// DASHBOARD VIEW (single screen, no scroll)
// ═══════════════════════════════════════════════════════════════════
function DashboardView({ report }) {
  const dateLabel = useMemo(() => {
    if (!report.date) return '';
    const [y, m, d] = report.date.split('-').map(Number);
    return new Date(y, m - 1, d).toLocaleDateString('es-AR', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  }, [report.date]);

  const p = report.servicios.plantaCaldera;

  return (
    <div className="space-y-3">
      {/* COMPACT HEADER */}
      <Card className="p-3">
        <div className="grid grid-cols-12 gap-3 items-center">
          <div className="col-span-3 flex items-center gap-3">
            <div className="w-10 h-10 rounded-full bg-sky-100 ring-1 ring-sky-200 flex items-center justify-center flex-shrink-0">
              <Wrench className="w-5 h-5 text-sky-600" />
            </div>
            <div>
              <div className="text-[10px] uppercase tracking-widest text-slate-500 font-semibold">Reporte Diario</div>
              <div className="text-sm font-bold text-slate-900 capitalize num">{dateLabel || '—'}</div>
            </div>
          </div>
          <div className="col-span-2">
            <div className="text-[10px] uppercase tracking-widest text-slate-500 font-semibold">Turno</div>
            <div className="text-sm font-medium">{report.shift}</div>
          </div>
          <div className="col-span-3">
            <div className="text-[10px] uppercase tracking-widest text-slate-500 font-semibold">Responsable</div>
            <div className="text-sm font-medium text-slate-800">{report.responsable || '—'}</div>
          </div>
          <div className="col-span-4">
            <div className="text-[10px] uppercase tracking-widest text-slate-500 font-semibold flex items-center gap-1">
              <Users className="w-3 h-3" />Equipo ({report.team.length})
            </div>
            <div className="text-xs text-slate-700 truncate" title={report.team.join(', ')}>
              {report.team.length === 0 ? '—' : report.team.join(' · ')}
            </div>
          </div>
        </div>
      </Card>

      {/* 4-COLUMN BODY */}
      <div className="grid grid-cols-12 gap-3" style={{ height: 'calc(100vh - 240px)', minHeight: '500px' }}>

        {/* COL 1: CORRECTIVOS */}
        <Card className="col-span-4 p-3 flex flex-col overflow-hidden">
          <h3 className="text-sky-600 font-bold text-sm mb-2 inline-flex items-center gap-2 flex-shrink-0">
            <Wrench className="w-4 h-4" />Correctivos ({report.corrective.length})
          </h3>
          <div className="overflow-auto flex-1">
            {report.corrective.length === 0
              ? <EmptyHint>Sin correctivos</EmptyHint>
              : (
                <div className="divide-y-2 divide-slate-200">
                  {report.corrective.map((c, i) => (
                    <div key={i} className="py-2.5 first:pt-0 last:pb-0">
                      <div className="flex items-center justify-between gap-2 mb-1">
                        <div className="flex items-center gap-2 min-w-0">
                          <span className="num text-[11px] font-bold text-slate-800 whitespace-nowrap">{c.ot || '—'}</span>
                          {c.equipoCodigo && (
                            <span className="text-[10px] text-slate-500 truncate">· {c.equipoCodigo}</span>
                          )}
                        </div>
                        <StatePill state={c.state} />
                      </div>
                      <div className="text-[12px] text-slate-700 leading-snug whitespace-pre-wrap break-words">{c.task || '—'}</div>
                      {(c.technicians || []).length > 0 && (
                        <div className="mt-1 flex flex-wrap gap-1">
                          {c.technicians.map(t => (
                            <span key={t} className="text-[9px] px-1.5 py-0.5 bg-slate-100 text-slate-600 rounded">{t}</span>
                          ))}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
          </div>
        </Card>

        {/* COL 2: PREVENTIVOS */}
        <Card className="col-span-4 p-3 flex flex-col overflow-hidden">
          <h3 className="text-sky-600 font-bold text-sm mb-2 inline-flex items-center gap-2 flex-shrink-0">
            <ListChecks className="w-4 h-4" />Preventivos ({report.preventive.length})
          </h3>
          <div className="overflow-auto flex-1">
            {report.preventive.length === 0
              ? <EmptyHint>Sin preventivos</EmptyHint>
              : (
                <div className="divide-y-2 divide-slate-200">
                  {report.preventive.map((p, i) => (
                    <div key={i} className="py-2.5 first:pt-0 last:pb-0">
                      <div className="flex items-center justify-between gap-2 mb-1">
                        <div className="flex items-center gap-2 min-w-0">
                          {p.codigoTarea && <span className="num text-[10px] font-semibold text-slate-700">{p.codigoTarea}</span>}
                          <span className="font-semibold text-slate-800 text-[11px] truncate">{p.equipoCodigo || '—'}</span>
                          {p.equipoDescripcion && <span className="text-[10px] text-slate-500 truncate">· {p.equipoDescripcion}</span>}
                        </div>
                        <span className="text-[9px] px-1.5 py-0.5 bg-emerald-50 text-emerald-700 rounded font-medium whitespace-nowrap">{p.frequency}</span>
                      </div>
                      <div className="text-[12px] text-slate-700 leading-snug whitespace-pre-wrap break-words">{p.task || '—'}</div>
                      {p.comments && (
                        <div className="text-[10px] text-slate-500 italic mt-0.5 leading-snug whitespace-pre-wrap break-words">{p.comments}</div>
                      )}
                      <div className="mt-1 flex flex-wrap items-center gap-1">
                        {(p.technicians || []).map(t => (
                          <span key={t} className="text-[9px] px-1.5 py-0.5 bg-slate-100 text-slate-600 rounded">{t}</span>
                        ))}
                        {p.otCorrectivaAsociada && (
                          <span className="text-[9px] px-1.5 py-0.5 bg-orange-50 text-orange-700 rounded font-medium num">↗ {p.otCorrectivaAsociada}</span>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              )}
          </div>
        </Card>

        {/* COL 3: SERVICIOS */}
        <Card className="col-span-4 p-3 flex flex-col overflow-hidden">
          <h3 className="text-sky-600 font-bold text-sm mb-2 inline-flex items-center gap-2 flex-shrink-0">
            <Activity className="w-4 h-4" />Servicios
          </h3>
          <div className="overflow-auto flex-1 space-y-3">
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
              <div className="grid grid-cols-5 gap-1 text-[10px]">
                {[['Caudal', p.caudal], ['pH', p.pH], ['Abl.', p.ablandador], ['ΔT', p.deltaT], ['Vacío', p.vacio]].map(([l, v]) => (
                  <div key={l} className="bg-slate-50 rounded p-1 text-center">
                    <div className="text-slate-500 uppercase">{l}</div>
                    <div className="font-bold num text-slate-800">{v !== '' && v != null ? v : '—'}</div>
                  </div>
                ))}
                {[['TK1', p.tk1], ['TK2', p.tk2], ['TK Em.', p.tkEmergencia], ['TK4', p.tk4]].map(([l, v]) => (
                  <div key={l} className="bg-slate-50 rounded p-1 text-center">
                    <div className="text-slate-500 uppercase">{l}</div>
                    <div className="font-bold num text-slate-800">{v !== '' && v != null ? v : '—'}</div>
                  </div>
                ))}
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
  );
}

// ═══════════════════════════════════════════════════════════════════
// STATS VIEW
// ═══════════════════════════════════════════════════════════════════
function StatsView({ history }) {
  const [range, setRange] = useState('week');
  const [customStart, setCustomStart] = useState('');
  const [customEnd, setCustomEnd] = useState('');

  const stats = useMemo(() => computeStats(history, range, customStart, customEnd), [history, range, customStart, customEnd]);

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
            Período: <span className="font-medium num">{stats.startStr}</span> → <span className="font-medium num">{stats.endStr}</span>
          </div>
        </div>
      </Card>

      {/* COMPACT KPI ROW */}
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
          <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
            <BarChart3 className="w-4 h-4" />Trabajos en el período
          </h3>
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

        <Card className="p-4">
          <h3 className="text-sm font-semibold text-slate-700 mb-3 inline-flex items-center gap-2">
            <Activity className="w-4 h-4" />Estado de Correctivos
          </h3>
          {stats.stateDist.length === 0 ? <EmptyHint>Sin datos</EmptyHint> :
            <ResponsiveContainer width="100%" height={240}>
              <PieChart>
                <Pie data={stats.stateDist} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80}
                  label={(e) => `${e.name}: ${e.value}`} labelLine={false}>
                  {stats.stateDist.map((entry, i) => (
                    <Cell key={i} fill={entry.name === 'Realizada' ? '#10b981' : entry.name === 'En Curso' ? '#f59e0b' : '#ef4444'} />
                  ))}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          }
        </Card>

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
            <Users className="w-4 h-4" />Carga por técnico
          </h3>
          {stats.topTechs.length === 0 ? <EmptyHint>Sin datos</EmptyHint> :
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={stats.topTechs} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis type="number" stroke="#64748b" style={{ fontSize: '10px' }} />
                <YAxis dataKey="name" type="category" stroke="#64748b" style={{ fontSize: '10px' }} width={120} />
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

  // bucketing by date
  const buckets = {};
  const dayMs = 86400000;
  const totalDays = Math.max(1, Math.round((end - start) / dayMs) + 1);
  const bucketByMonth = totalDays > 90;  // group by month if >3 months
  const bucketByWeek = totalDays > 31 && !bucketByMonth;

  filtered.forEach(r => {
    let key, label;
    const d = new Date(r.date);
    if (bucketByMonth) {
      key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      label = d.toLocaleDateString('es-AR', { month: 'short', year: '2-digit' });
    } else if (bucketByWeek) {
      const monday = new Date(d);
      const day = monday.getDay() || 7;
      monday.setDate(monday.getDate() - day + 1);
      key = monday.toISOString().slice(0, 10);
      label = `S ${monday.getDate()}/${monday.getMonth() + 1}`;
    } else {
      key = r.date;
      label = d.toLocaleDateString('es-AR', { day: 'numeric', month: 'short' });
    }
    if (!buckets[key]) buckets[key] = { key, label, correctivos: 0, preventivos: 0 };
    buckets[key].correctivos += (r.corrective || []).length;
    buckets[key].preventivos += (r.preventive || []).length;
  });
  const daily = Object.values(buckets).sort((a, b) => a.key.localeCompare(b.key));

  let totalCorrectives = 0, totalPreventives = 0, completedCorr = 0, pendingCorr = 0, urgent = 0;
  const stateDist = { 'Sin Iniciar': 0, 'En Curso': 0, 'Realizada': 0 };
  const equipmentCount = {};
  const techCount = {};
  const shiftCount = { Mañana: { correctivos: 0, preventivos: 0 }, Tarde: { correctivos: 0, preventivos: 0 }, Noche: { correctivos: 0, preventivos: 0 } };

  filtered.forEach(r => {
    (r.corrective || []).forEach(c => {
      totalCorrectives++;
      if (c.state === 'Realizada') completedCorr++;
      else if (c.state === 'Sin Iniciar') pendingCorr++;
      if (c.state in stateDist) stateDist[c.state]++;
      const eq = (c.equipoCodigo || '').trim();
      if (eq && eq !== '-') equipmentCount[eq] = (equipmentCount[eq] || 0) + 1;
      (c.technicians || []).forEach(t => {
        techCount[t] = techCount[t] || { correctivos: 0, preventivos: 0 };
        techCount[t].correctivos++;
      });
      if (shiftCount[r.shift]) shiftCount[r.shift].correctivos++;
    });
    (r.preventive || []).forEach(p => {
      totalPreventives++;
      (p.technicians || []).forEach(t => {
        techCount[t] = techCount[t] || { correctivos: 0, preventivos: 0 };
        techCount[t].preventivos++;
      });
      if (shiftCount[r.shift]) shiftCount[r.shift].preventivos++;
    });
    (r.comments || []).forEach(c => { if (c.priority === 'Urgente') urgent++; });
  });

  const completionRate = totalCorrectives > 0 ? Math.round((completedCorr / totalCorrectives) * 100) : 0;

  const topEquipment = Object.entries(equipmentCount)
    .sort((a, b) => b[1] - a[1]).slice(0, 8)
    .map(([name, count]) => ({ name, count }));

  const topTechs = Object.entries(techCount)
    .map(([name, v]) => ({
      name: name.split(' ').slice(0, 2).join(' '),
      correctivos: v.correctivos, preventivos: v.preventivos,
      total: v.correctivos + v.preventivos
    }))
    .sort((a, b) => b.total - a.total).slice(0, 8);

  const shiftDist = Object.entries(shiftCount).map(([name, v]) => ({ name, ...v }));

  return {
    totalReports: filtered.length,
    totalCorrectives, totalPreventives, completedCorr, pendingCorr, urgent,
    completionRate, daily, startStr, endStr,
    stateDist: Object.entries(stateDist).map(([name, value]) => ({ name, value })).filter(x => x.value > 0),
    topEquipment, topTechs, shiftDist
  };
}

// ═══════════════════════════════════════════════════════════════════
// HISTORY VIEW
// ═══════════════════════════════════════════════════════════════════
function HistoryView({ history, onExportCorrectives, onExportPreventives, onExportFull }) {
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
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
          <button onClick={onExportCorrectives} className="bg-orange-600 text-white hover:bg-orange-700 rounded-lg p-4 transition flex items-start gap-3">
            <Wrench className="w-6 h-6 flex-shrink-0 mt-0.5" />
            <div className="text-left">
              <div className="font-semibold">Solo Correctivos</div>
              <div className="text-xs text-orange-100 mt-1">Hoja "Correctivos" con columnas del template</div>
            </div>
          </button>
          <button onClick={onExportPreventives} className="bg-emerald-600 text-white hover:bg-emerald-700 rounded-lg p-4 transition flex items-start gap-3">
            <ListChecks className="w-6 h-6 flex-shrink-0 mt-0.5" />
            <div className="text-left">
              <div className="font-semibold">Solo Preventivos</div>
              <div className="text-xs text-emerald-100 mt-1">Hoja "Preventivos" con columnas del template</div>
            </div>
          </button>
          <button onClick={onExportFull} className="bg-slate-800 text-white hover:bg-slate-700 rounded-lg p-4 transition flex items-start gap-3">
            <FileSpreadsheet className="w-6 h-6 flex-shrink-0 mt-0.5" />
            <div className="text-left">
              <div className="font-semibold">Reporte completo</div>
              <div className="text-xs text-slate-300 mt-1">11 hojas: Téc. presentes, Correctivos, Preventivos, Cisternas, Agua de Pozo, Compresores, Grupos Electrógenos, Planta de Efluentes, Servicios externos, Comentarios, Responsables</div>
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
                </tr>
              </thead>
              <tbody>
                {filtered.map((r, i) => (
                  <tr key={i} className="border-b border-slate-100 hover:bg-slate-50">
                    <td className="py-2 num font-medium">{r.date}</td>
                    <td className="py-2">{r.shift}</td>
                    <td className="py-2 text-slate-600">{r.responsable || '—'}</td>
                    <td className="py-2 text-right num">{r.team?.length || 0}</td>
                    <td className="py-2 text-right num">{r.corrective?.length || 0}</td>
                    <td className="py-2 text-right num">{r.preventive?.length || 0}</td>
                    <td className="py-2 text-right num">
                      {r.comments?.filter(c => c.priority === 'Urgente').length || 0}
                    </td>
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

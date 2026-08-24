import React, { useState, useEffect, useMemo, useCallback } from 'react';
import {
  Menu, Search, Calendar as CalendarIcon, HelpCircle, Plus, X, ChevronLeft, ChevronRight,
  Home as HomeIcon, UtensilsCrossed, Car, HeartPulse, GraduationCap, PartyPopper, Coffee,
  PiggyBank, Wallet, CreditCard, Landmark, ListChecks, LayoutGrid, BarChart3, Trash2, Check,
  ArrowLeft, TrendingUp, Download
} from 'lucide-react';
import * as XLSX from 'xlsx';


// ---------------------------------------------------------------------------
// Datos base (extraídos del Catálogo Maestro de Subcategorías del sistema)
// ---------------------------------------------------------------------------
const CATALOGO = [
['Vivienda y Servicios','Renta / Hipoteca (Vivienda principal)',10500],
['Vivienda y Servicios','Pago terreno / Inversión inmobiliaria',3500],
['Vivienda y Servicios','Luz (CFE)',450],
['Vivienda y Servicios','Agua potable / Suministro municipal',250],
['Vivienda y Servicios','Gas estacionario / natural / cilindro',550],
['Vivienda y Servicios','Internet residencial (Fibra óptica)',650],
['Vivienda y Servicios','Telefonía celular / Plan de datos',450],
['Vivienda y Servicios','Cuota de mantenimiento / Condominio',950],
['Vivienda y Servicios','Artículos de limpieza para el hogar',350],
['Vivienda y Servicios','Herramientas y ferretería',150],
['Vivienda y Servicios','Mantenimiento correctivo (Fontanería, etc.)',200],
['Vivienda y Servicios','Seguro de hogar / Contenidos',100],
['Vivienda y Servicios','Impuesto Predial y Derechos',50],
['Vivienda y Servicios','Servicio doméstico / Limpieza contratada',400],
['Vivienda y Servicios','Mobiliario y decoración menor',200],
['Vivienda y Servicios','Electrodomésticos (Compra o reparación)',150],
['Alimentación','Supermercado / Despensa general',2800],
['Alimentación','Frutas, verduras y legumbres',700],
['Alimentación','Carnes, aves y pescados',900],
['Alimentación','Lácteos y huevos',350],
['Alimentación','Panadería y tortillería',200],
['Alimentación','Agua purificada (Garrafones / Filtros)',150],
['Alimentación','Comida en jornada laboral (Desayuno/Comida)',600],
['Alimentación','Restaurantes (Salidas formales)',500],
['Alimentación','Comida rápida / Antojitos (Calle)',300],
['Alimentación','Delivery (UberEats, Rappi, DidiFood)',350],
['Alimentación','Café de grano / Té (Para preparar en casa)',150],
['Alimentación','Especias, condimentos y salsas',100],
['Alimentación','Postres y repostería (Panaderías/Helados)',150],
['Transporte y Movilidad','Gasolina / Combustible',1900],
['Transporte y Movilidad','Peajes y Casetas (Tag / Efectivo)',350],
['Transporte y Movilidad','Estacionamientos comerciales',100],
['Transporte y Movilidad','Parquímetros y Valet Parking',100],
['Transporte y Movilidad','Mantenimiento preventivo (Afinación/Aceite)',400],
['Transporte y Movilidad','Mantenimiento correctivo (Fallas mecánicas)',200],
['Transporte y Movilidad','Llantas, rines y suspensión',300],
['Transporte y Movilidad','Seguro vehicular (Póliza)',350],
['Transporte y Movilidad','Verificación vehicular y emisiones',80],
['Transporte y Movilidad','Tenencia y Refrendo vehicular',80],
['Transporte y Movilidad','Lavado de auto (Exterior / Interior)',120],
['Transporte y Movilidad','Accesorios para el auto',50],
['Transporte y Movilidad','Transporte público masivo (Metro, Bus)',100],
['Transporte y Movilidad','Taxis y Apps de viaje (Uber, Didi, Cabify)',200],
['Transporte y Movilidad','Multas y recargos de tránsito',50],
['Salud y Bienestar','Suplementos Nootrópicos (L-Tirosina, Glycina)',300],
['Salud y Bienestar','Suplementos Base (Omega-3, Magnesio, D3/K2)',450],
['Salud y Bienestar','Suplementos Deportivos (Creatina, Proteína)',400],
['Salud y Bienestar','Multivitamínicos comerciales (Centrum, etc.)',200],
['Salud y Bienestar','Farmacia (Analgésicos, botiquín básico)',200],
['Salud y Bienestar','Medicamentos de receta / Tratamientos',300],
['Salud y Bienestar','Consultas médicas generales / Especialista',400],
['Salud y Bienestar','Odontología (Limpiezas, resinas, revisión)',250],
['Salud y Bienestar','Oftalmología (Examen visual, lentes, gotas)',150],
['Salud y Bienestar','Laboratorios y análisis clínicos',100],
['Salud y Bienestar','Gimnasio / Membresía deportiva',600],
['Salud y Bienestar','Artículos de Biohacking (Cintas de boca, etc.)',100],
['Salud y Bienestar','Cuidado personal (Shampoo, jabón, desodorante)',150],
['Salud y Bienestar','Corte de cabello / Barbería / Estética',120],
['Salud y Bienestar','Skin care y dermatología estética',200],
['Salud y Bienestar','Seguro de Gastos Médicos Mayores (SGMM)',400],
['Desarrollo y Educación','Mensualidad de Maestría / Posgrado',1500],
['Desarrollo y Educación','Libros físicos (Productividad, Finanzas, Psicología)',400],
['Desarrollo y Educación','Ebooks y audiolibros (Kindle, Audible)',150],
['Desarrollo y Educación','Plataformas de idiomas (Ej. Learnlight)',250],
['Desarrollo y Educación','Cursos online y Bootcamps (Udemy, Coursera)',300],
['Desarrollo y Educación','Certificaciones profesionales (Derechos de examen)',200],
['Desarrollo y Educación','Software de Productividad (MS Project, Notion)',200],
['Desarrollo y Educación','Almacenamiento Cloud (Google One, iCloud, Dropbox)',100],
['Desarrollo y Educación','Material de oficina y papelería (Libretas, plumas)',100],
['Desarrollo y Educación','Hardware (Mouse, teclado, cables, discos duros)',150],
['Desarrollo y Educación','Electrónica mayor (Laptop, tablet, smartphone)',400],
['Desarrollo y Educación','Prácticas y simuladores (Ej. Apps de Ajedrez)',50],
['Desarrollo y Educación','Cursos o talleres de esparcimiento (Arte, cocina)',100],
['Desarrollo y Educación','Suscripciones a medios/noticias (The Economist, etc.)',80],
['Estilo de Vida y Ocio','Salidas y reuniones con amigos/pareja',800],
['Estilo de Vida y Ocio','Cine, teatro y espectáculos',300],
['Estilo de Vida y Ocio','Conciertos, festivales y eventos masivos',400],
['Estilo de Vida y Ocio','Suscripciones de Video (Netflix, Max, Prime)',250],
['Estilo de Vida y Ocio','Suscripciones de Música (Spotify, Apple Music)',130],
['Estilo de Vida y Ocio','Videojuegos y microtransacciones',150],
['Estilo de Vida y Ocio','Ropa casual y formal',400],
['Estilo de Vida y Ocio','Ropa e indumentaria deportiva',200],
['Estilo de Vida y Ocio','Calzado (Zapatos, tenis de correr/casual)',300],
['Estilo de Vida y Ocio','Ropa especializada / Denim crudo (Ej. Sauce Zhan)',150],
['Estilo de Vida y Ocio','Pasatiempos físicos (Tableros, juegos de mesa)',100],
['Estilo de Vida y Ocio','Fondo para viajes y vacaciones',500],
['Gastos Hormiga y Varios','Cafeterías de paso (Starbucks, Cielito, etc.)',300],
['Gastos Hormiga y Varios','Snacks y golosinas en tiendas (Oxxo, 7-Eleven)',200],
['Gastos Hormiga y Varios','Bebidas preparadas y energéticas',150],
['Gastos Hormiga y Varios','Propinas (Valet, empacadores, gasolineros)',150],
['Gastos Hormiga y Varios','Comisiones bancarias (Cajeros ajenos, SPEI)',50],
['Gastos Hormiga y Varios','Intereses por financiamiento (Tarjetas)',50],
['Gastos Hormiga y Varios','Trámites administrativos y legales (Copias, actas)',50],
['Gastos Hormiga y Varios','Regalos de cumpleaños (Familia/Amigos)',200],
['Gastos Hormiga y Varios','Compromisos sociales (Bodas, baby showers)',150],
['Gastos Hormiga y Varios','Apoyo a familiares',300],
['Gastos Hormiga y Varios','Donaciones y caridad',50],
['Gastos Hormiga y Varios','Cuidado de mascotas (Alimento / Arena)',100],
['Gastos Hormiga y Varios','Veterinario y accesorios de mascotas',50],
['Gastos Hormiga y Varios','Imprevistos misceláneos (No categorizables)',100],
['Ahorro e Inversiones','Fondo de Emergencia (Liquidez inmediata)',1500],
['Ahorro e Inversiones','Inversión Renta Fija (CETES, Pagarés bancarios)',2000],
['Ahorro e Inversiones','Inversión Renta Variable (ETFs, Acciones en GBM)',1500],
['Ahorro e Inversiones','FIBRAs / REITs (Inversión inmobiliaria bursátil)',500],
['Ahorro e Inversiones','Plan Personal de Retiro (PPR / AFORE)',500],
['Ahorro e Inversiones','Inversiones alternativas (Criptomonedas, Crowdfunding)',200],
['Ahorro e Inversiones','Abonos a capital de deudas mayores',500],
['Ahorro e Inversiones','Provisión para pago anual de impuestos (ISR)',150]
];

const MACROS = [
  { name: 'Vivienda y Servicios',      color: '#ef4444', Icon: HomeIcon },
  { name: 'Alimentación',              color: '#f97316', Icon: UtensilsCrossed },
  { name: 'Transporte y Movilidad',    color: '#3b82f6', Icon: Car },
  { name: 'Salud y Bienestar',         color: '#ec4899', Icon: HeartPulse },
  { name: 'Desarrollo y Educación',    color: '#a855f7', Icon: GraduationCap },
  { name: 'Estilo de Vida y Ocio',     color: '#06b6d4', Icon: PartyPopper },
  { name: 'Gastos Hormiga y Varios',   color: '#eab308', Icon: Coffee },
  { name: 'Ahorro e Inversiones',      color: '#22c55e', Icon: PiggyBank },
];
const macroMeta = (name) => MACROS.find(m => m.name === name) || MACROS[0];

const ACCOUNTS = [
  { name: 'Efectivo / Billetera',        Icon: Wallet },
  { name: 'Tarjeta de Débito',           Icon: CreditCard },
  { name: 'Tarjeta de Crédito',          Icon: CreditCard },
  { name: 'Transferencia Bancaria',      Icon: Landmark },
  { name: 'Cuenta de Ahorro / Inversión', Icon: PiggyBank },
];

const TOTAL_BUDGET = CATALOGO.reduce((s, c) => s + c[2], 0);
const budgetByMacro = (macro) => CATALOGO.filter(c => c[0] === macro).reduce((s, c) => s + c[2], 0);

const WEEKDAY_ABBR = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
const WEEKDAY_FULL = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
const MONTHS = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

const pad = (n) => String(n).padStart(2, '0');
const toKey = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const fromKey = (k) => { const [y,m,d] = k.split('-').map(Number); return new Date(y, m - 1, d); };
const todayKey = () => toKey(new Date());
const fmt = (n) => new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', maximumFractionDigits: 0 }).format(n || 0);
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

function exportToExcel(transactions) {
  const rows = [...transactions]
    .sort((a, b) => a.date.localeCompare(b.date))
    .map(t => ({
      Fecha: t.date,
      'Macro-Categoría': t.macro,
      Subcategoría: t.sub,
      'Cuenta de Pago': t.account,
      'Monto (MXN)': t.amount,
      Nota: t.note || '',
    }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [{ wch: 12 }, { wch: 24 }, { wch: 34 }, { wch: 22 }, { wch: 12 }, { wch: 28 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Registro de Gastos');
  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `gastos_${fecha}.xlsx`);
}

function weekStrip(centerKey) {
  const center = fromKey(centerKey);
  const days = [];
  for (let i = -3; i <= 3; i++) {
    const d = new Date(center);
    d.setDate(d.getDate() + i);
    days.push(d);
  }
  return days;
}

function monthGrid(year, month) {
  // month: 0-indexed. Devuelve arreglo de 42 fechas (6 semanas, lunes-domingo)
  const first = new Date(year, month, 1);
  const jsWeekday = first.getDay(); // 0=domingo
  const mondayOffset = jsWeekday === 0 ? 6 : jsWeekday - 1;
  const start = new Date(first);
  start.setDate(start.getDate() - mondayOffset);
  const grid = [];
  for (let i = 0; i < 42; i++) {
    const d = new Date(start);
    d.setDate(d.getDate() + i);
    grid.push(d);
  }
  return grid;
}

// ---------------------------------------------------------------------------
// Componente principal
// ---------------------------------------------------------------------------
export default function GastosApp() {
  const [transactions, setTransactions] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [storageOk, setStorageOk] = useState(true);
  const [tab, setTab] = useState('hoy');
  const [selectedDate, setSelectedDate] = useState(todayKey());
  const [calMonth, setCalMonth] = useState(() => { const d = new Date(); return { y: d.getFullYear(), m: d.getMonth() }; });
  const [addOpen, setAddOpen] = useState(false);
  const [categoryDrill, setCategoryDrill] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);

  useEffect(() => {
    let mounted = true;
    (async () => {
      try {
        const res = await window.storage.get('gastos_transactions_v1', false);
        if (mounted && res && res.value) setTransactions(JSON.parse(res.value));
      } catch (e) {
        // sin datos previos todavía
      } finally {
        if (mounted) setLoaded(true);
      }
    })();
    return () => { mounted = false; };
  }, []);

  useEffect(() => {
    if (!loaded) return;
    (async () => {
      try {
        const r = await window.storage.set('gastos_transactions_v1', JSON.stringify(transactions), false);
        setStorageOk(!!r);
      } catch (e) {
        setStorageOk(false);
      }
    })();
  }, [transactions, loaded]);

  const addTransaction = useCallback((tx) => {
    setTransactions(prev => [...prev, { ...tx, id: uid() }]);
  }, []);

  const deleteTransaction = useCallback((id) => {
    setTransactions(prev => prev.filter(t => t.id !== id));
    setConfirmDelete(null);
  }, []);

  const monthKey = (k) => k.slice(0, 7);
  const thisMonthKey = monthKey(selectedDate);
  const monthTx = useMemo(
    () => transactions.filter(t => monthKey(t.date) === thisMonthKey),
    [transactions, thisMonthKey]
  );
  const dayTx = useMemo(
    () => transactions.filter(t => t.date === selectedDate).sort((a,b)=> (b.time||'').localeCompare(a.time||'')),
    [transactions, selectedDate]
  );
  const dayTotal = dayTx.reduce((s, t) => s + t.amount, 0);
  const monthTotal = monthTx.reduce((s, t) => s + t.amount, 0);
  const remaining = TOTAL_BUDGET - monthTotal;

  return (
    <div className="min-h-screen bg-black text-white flex justify-center">
      <div className="w-full max-w-md relative flex flex-col min-h-screen" style={{ fontFamily: "'Segoe UI', system-ui, -apple-system, sans-serif" }}>

        {!storageOk && loaded && (
          <div className="bg-yellow-500/10 text-yellow-400 text-xs text-center py-1.5 px-4">
            No se pudo guardar en este momento. Tus datos son locales a esta sesión.
          </div>
        )}

        {tab === 'hoy' && (
          <HoyView
            selectedDate={selectedDate} setSelectedDate={setSelectedDate}
            dayTx={dayTx} dayTotal={dayTotal}
            onDelete={(t) => setConfirmDelete(t)}
            onAdd={() => setAddOpen(true)}
          />
        )}
        {tab === 'categorias' && (
          <CategoriasView
            monthTx={monthTx} monthLabel={`${MONTHS[fromKey(selectedDate).getMonth()]} ${fromKey(selectedDate).getFullYear()}`}
            onOpenCategory={(m) => setCategoryDrill(m)}
            onAdd={() => setAddOpen(true)}
          />
        )}
        {tab === 'calendario' && (
          <CalendarioView
            calMonth={calMonth} setCalMonth={setCalMonth}
            transactions={transactions}
            onSelectDay={(k) => { setSelectedDate(k); setTab('hoy'); }}
            onAdd={() => setAddOpen(true)}
          />
        )}
        {tab === 'resumen' && (
          <ResumenView
            monthTx={monthTx} monthTotal={monthTotal} remaining={remaining}
            allTransactions={transactions}
            monthLabel={`${MONTHS[fromKey(selectedDate).getMonth()]} ${fromKey(selectedDate).getFullYear()}`}
            onAdd={() => setAddOpen(true)}
          />
        )}

        <BottomNav tab={tab} setTab={setTab} />

        {addOpen && (
          <AddExpenseSheet
            defaultDate={selectedDate}
            onClose={() => setAddOpen(false)}
            onSave={(tx) => { addTransaction(tx); setAddOpen(false); }}
          />
        )}

        {categoryDrill && (
          <CategoryDrillSheet
            macro={categoryDrill}
            monthTx={monthTx.filter(t => t.macro === categoryDrill)}
            onClose={() => setCategoryDrill(null)}
            onDelete={(t) => setConfirmDelete(t)}
          />
        )}

        {confirmDelete && (
          <ConfirmDeleteSheet
            tx={confirmDelete}
            onCancel={() => setConfirmDelete(null)}
            onConfirm={() => deleteTransaction(confirmDelete.id)}
          />
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Header compartido
// ---------------------------------------------------------------------------
function TopBar({ title, onSearch }) {
  return (
    <div className="flex items-center justify-between px-5 pt-6 pb-3">
      <button className="text-red-500"><Menu size={22} /></button>
      <h1 className="text-2xl font-bold flex-1 ml-3">{title}</h1>
      <div className="flex items-center gap-4 text-neutral-400">
        <Search size={20} />
        <HelpCircle size={20} />
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vista: Hoy
// ---------------------------------------------------------------------------
function HoyView({ selectedDate, setSelectedDate, dayTx, dayTotal, onDelete, onAdd }) {
  const days = weekStrip(selectedDate);
  const isToday = selectedDate === todayKey();

  return (
    <div className="flex-1 pb-24">
      <TopBar title={isToday ? 'Hoy' : WEEKDAY_FULL[(fromKey(selectedDate).getDay() + 6) % 7]} />

      <div className="flex gap-2 px-4 overflow-x-auto pb-1" style={{ scrollbarWidth: 'none' }}>
        {days.map((d) => {
          const k = toKey(d);
          const sel = k === selectedDate;
          return (
            <button
              key={k}
              onClick={() => setSelectedDate(k)}
              className="flex flex-col items-center justify-center rounded-2xl px-3 py-2 min-w-[58px] shrink-0 transition-colors"
              style={{ backgroundColor: sel ? '#dc2626' : '#171717' }}
            >
              <span className="text-xs font-medium" style={{ color: sel ? '#fecaca' : '#737373' }}>{WEEKDAY_ABBR[d.getDay()]}</span>
              <span className="text-lg font-bold" style={{ color: sel ? '#ffffff' : '#e5e5e5' }}>{d.getDate()}</span>
            </button>
          );
        })}
      </div>

      <div className="px-5 mt-5 mb-2">
        <div className="text-neutral-500 text-sm">Gasto del día</div>
        <div className="text-3xl font-bold mt-0.5">{fmt(dayTotal)}</div>
      </div>

      <div className="px-4 mt-4 space-y-3">
        {dayTx.length === 0 && (
          <div className="text-center text-neutral-600 text-sm py-16 px-6">
            Sin movimientos este día. Toca el botón rojo para registrar un gasto.
          </div>
        )}
        {dayTx.map((t) => {
          const meta = macroMeta(t.macro);
          const Icon = meta.Icon;
          return (
            <button
              key={t.id}
              onClick={() => onDelete(t)}
              className="w-full flex items-center gap-3 bg-neutral-950 rounded-2xl p-3 text-left active:opacity-70"
            >
              <div className="w-12 h-12 rounded-xl flex items-center justify-center shrink-0" style={{ backgroundColor: meta.color }}>
                <Icon size={22} color="#000" strokeWidth={2.2} />
              </div>
              <div className="flex-1 min-w-0">
                <div className="font-semibold text-[15px] leading-tight truncate">{t.sub}</div>
                <div className="flex items-center gap-2 mt-1.5">
                  <span className="text-xs px-2 py-0.5 rounded-full font-medium" style={{ backgroundColor: meta.color + '26', color: meta.color }}>
                    {t.macro}
                  </span>
                  {t.account && <span className="text-xs text-neutral-500">{t.account}</span>}
                </div>
              </div>
              <div className="font-bold text-[15px] shrink-0">{fmt(t.amount)}</div>
            </button>
          );
        })}
      </div>

      <FloatingAddButton onClick={onAdd} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vista: Categorías
// ---------------------------------------------------------------------------
function CategoriasView({ monthTx, monthLabel, onOpenCategory, onAdd }) {
  return (
    <div className="flex-1 pb-24">
      <TopBar title="Categorías" />
      <div className="px-5 -mt-1 mb-4 text-neutral-500 text-sm">{monthLabel}</div>

      <div className="px-4 grid grid-cols-2 gap-3">
        {MACROS.map((m) => {
          const spent = monthTx.filter(t => t.macro === m.name).reduce((s, t) => s + t.amount, 0);
          const budget = budgetByMacro(m.name);
          const pct = budget > 0 ? Math.min(spent / budget, 1) : 0;
          const over = budget > 0 && spent > budget;
          const Icon = m.Icon;
          return (
            <button
              key={m.name}
              onClick={() => onOpenCategory(m.name)}
              className="bg-neutral-950 rounded-2xl p-3.5 text-left active:opacity-70"
            >
              <div className="w-11 h-11 rounded-xl flex items-center justify-center mb-3" style={{ backgroundColor: m.color }}>
                <Icon size={20} color="#000" strokeWidth={2.2} />
              </div>
              <div className="font-semibold text-sm leading-tight">{m.name}</div>
              <div className="text-xs text-neutral-500 mt-1">{fmt(spent)} de {fmt(budget)}</div>
              <div className="h-1.5 rounded-full bg-neutral-800 mt-2 overflow-hidden">
                <div className="h-full rounded-full" style={{ width: `${pct * 100}%`, backgroundColor: over ? '#ef4444' : m.color }} />
              </div>
            </button>
          );
        })}
      </div>
      <FloatingAddButton onClick={onAdd} />
    </div>
  );
}

function CategoryDrillSheet({ macro, monthTx, onClose, onDelete }) {
  const meta = macroMeta(macro);
  const Icon = meta.Icon;
  const total = monthTx.reduce((s, t) => s + t.amount, 0);
  return (
    <Sheet onClose={onClose}>
      <div className="flex items-center gap-3 mb-5">
        <div className="w-11 h-11 rounded-xl flex items-center justify-center" style={{ backgroundColor: meta.color }}>
          <Icon size={20} color="#000" strokeWidth={2.2} />
        </div>
        <div>
          <div className="font-bold text-lg leading-tight">{macro}</div>
          <div className="text-sm text-neutral-500">{fmt(total)} este mes</div>
        </div>
      </div>
      <div className="space-y-2 max-h-[55vh] overflow-y-auto pb-2">
        {monthTx.length === 0 && <div className="text-neutral-600 text-sm text-center py-10">Sin movimientos este mes.</div>}
        {[...monthTx].sort((a,b)=> b.date.localeCompare(a.date)).map(t => (
          <button key={t.id} onClick={() => onDelete(t)} className="w-full flex items-center justify-between bg-neutral-900 rounded-xl p-3 text-left active:opacity-70">
            <div className="min-w-0">
              <div className="font-medium text-sm truncate">{t.sub}</div>
              <div className="text-xs text-neutral-500">{t.date}</div>
            </div>
            <div className="font-semibold text-sm shrink-0 ml-3">{fmt(t.amount)}</div>
          </button>
        ))}
      </div>
    </Sheet>
  );
}

// ---------------------------------------------------------------------------
// Vista: Calendario
// ---------------------------------------------------------------------------
function CalendarioView({ calMonth, setCalMonth, transactions, onSelectDay, onAdd }) {
  const grid = monthGrid(calMonth.y, calMonth.m);
  const totalsByDay = useMemo(() => {
    const map = {};
    transactions.forEach(t => { map[t.date] = (map[t.date] || 0) + t.amount; });
    return map;
  }, [transactions]);

  const monthTotal = grid
    .filter(d => d.getMonth() === calMonth.m)
    .reduce((s, d) => s + (totalsByDay[toKey(d)] || 0), 0);

  const changeMonth = (delta) => {
    let m = calMonth.m + delta, y = calMonth.y;
    if (m < 0) { m = 11; y -= 1; }
    if (m > 11) { m = 0; y += 1; }
    setCalMonth({ y, m });
  };

  return (
    <div className="flex-1 pb-24">
      <TopBar title="Calendario" />
      <div className="flex items-center justify-between px-5 mb-4">
        <button onClick={() => changeMonth(-1)} className="text-neutral-400 p-1"><ChevronLeft size={22} /></button>
        <div className="text-center">
          <div className="font-bold">{MONTHS[calMonth.m]} {calMonth.y}</div>
          <div className="text-xs text-neutral-500">{fmt(monthTotal)} gastado</div>
        </div>
        <button onClick={() => changeMonth(1)} className="text-neutral-400 p-1"><ChevronRight size={22} /></button>
      </div>

      <div className="px-4">
        <div className="grid grid-cols-7 gap-1 mb-1">
          {['L','M','X','J','V','S','D'].map((d,i) => (
            <div key={i} className="text-center text-xs text-neutral-600 font-medium py-1">{d}</div>
          ))}
        </div>
        <div className="grid grid-cols-7 gap-1">
          {grid.map((d, i) => {
            const inMonth = d.getMonth() === calMonth.m;
            const k = toKey(d);
            const amt = totalsByDay[k] || 0;
            const isToday = k === todayKey();
            return (
              <button
                key={i}
                disabled={!inMonth}
                onClick={() => onSelectDay(k)}
                className="aspect-square rounded-lg flex flex-col items-center justify-center gap-0.5"
                style={{
                  backgroundColor: !inMonth ? 'transparent' : amt > 0 ? '#7f1d1d' : '#171717',
                  opacity: inMonth ? 1 : 0,
                  border: isToday ? '1.5px solid #ef4444' : 'none',
                }}
              >
                {inMonth && (
                  <>
                    <span className="text-xs font-semibold">{d.getDate()}</span>
                    {amt > 0 && <span className="text-[9px] text-red-300 leading-none">{amt >= 1000 ? `${Math.round(amt/1000)}k` : Math.round(amt)}</span>}
                  </>
                )}
              </button>
            );
          })}
        </div>
      </div>
      <FloatingAddButton onClick={onAdd} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vista: Resumen
// ---------------------------------------------------------------------------
function ResumenView({ monthTx, monthTotal, remaining, allTransactions, monthLabel, onAdd }) {
  const pct = TOTAL_BUDGET > 0 ? Math.min(monthTotal / TOTAL_BUDGET, 1) : 0;
  const over = monthTotal > TOTAL_BUDGET;

  const byMacro = MACROS.map(m => ({
    ...m,
    spent: monthTx.filter(t => t.macro === m.name).reduce((s, t) => s + t.amount, 0),
    budget: budgetByMacro(m.name),
  })).sort((a, b) => (b.spent / (b.budget || 1)) - (a.spent / (a.budget || 1)));

  const top5 = [...monthTx].sort((a, b) => b.amount - a.amount).slice(0, 5);

  return (
    <div className="flex-1 pb-24">
      <TopBar title="Resumen" />
      <div className="px-5 -mt-1 mb-4 flex items-center justify-between">
        <span className="text-neutral-500 text-sm">{monthLabel}</span>
        <button
          onClick={() => exportToExcel(allTransactions)}
          disabled={allTransactions.length === 0}
          className="flex items-center gap-1.5 text-xs font-semibold px-3 py-1.5 rounded-full"
          style={{
            backgroundColor: allTransactions.length === 0 ? '#171717' : '#dc262620',
            color: allTransactions.length === 0 ? '#525252' : '#ef4444',
          }}
        >
          <Download size={13} /> Exportar a Excel
        </button>
      </div>

      <div className="mx-4 rounded-2xl p-5 mb-5" style={{ backgroundColor: '#171717' }}>
        <div className="text-neutral-500 text-sm">Presupuesto restante</div>
        <div className="text-3xl font-bold mt-0.5" style={{ color: over ? '#ef4444' : '#ffffff' }}>{fmt(remaining)}</div>
        <div className="h-2 rounded-full bg-neutral-800 mt-4 overflow-hidden">
          <div className="h-full rounded-full" style={{ width: `${pct * 100}%`, backgroundColor: over ? '#ef4444' : '#22c55e' }} />
        </div>
        <div className="flex justify-between text-xs text-neutral-500 mt-2">
          <span>{fmt(monthTotal)} gastado</span>
          <span>{fmt(TOTAL_BUDGET)} presupuesto</span>
        </div>
      </div>

      <div className="px-5 mb-2 flex items-center gap-2 text-neutral-400 text-sm font-semibold">
        <TrendingUp size={15} /> Ejecución por categoría
      </div>
      <div className="px-4 space-y-2.5 mb-6">
        {byMacro.map(m => {
          const p = m.budget > 0 ? Math.min(m.spent / m.budget, 1) : 0;
          const over2 = m.spent > m.budget;
          const Icon = m.Icon;
          return (
            <div key={m.name} className="flex items-center gap-3 bg-neutral-950 rounded-xl p-3">
              <div className="w-9 h-9 rounded-lg flex items-center justify-center shrink-0" style={{ backgroundColor: m.color }}>
                <Icon size={16} color="#000" />
              </div>
              <div className="flex-1 min-w-0">
                <div className="flex justify-between text-sm mb-1">
                  <span className="font-medium truncate">{m.name}</span>
                  <span className="text-neutral-500 shrink-0 ml-2">{Math.round(p * 100)}%</span>
                </div>
                <div className="h-1.5 rounded-full bg-neutral-800 overflow-hidden">
                  <div className="h-full rounded-full" style={{ width: `${p * 100}%`, backgroundColor: over2 ? '#ef4444' : m.color }} />
                </div>
              </div>
            </div>
          );
        })}
      </div>

      {top5.length > 0 && (
        <>
          <div className="px-5 mb-2 text-neutral-400 text-sm font-semibold">Gastos más altos del mes</div>
          <div className="px-4 space-y-2">
            {top5.map(t => {
              const meta = macroMeta(t.macro);
              return (
                <div key={t.id} className="flex items-center justify-between bg-neutral-950 rounded-xl p-3">
                  <div className="min-w-0">
                    <div className="text-sm font-medium truncate">{t.sub}</div>
                    <div className="text-xs text-neutral-500">{t.date}</div>
                  </div>
                  <div className="font-bold text-sm shrink-0 ml-3" style={{ color: meta.color }}>{fmt(t.amount)}</div>
                </div>
              );
            })}
          </div>
        </>
      )}
      <FloatingAddButton onClick={onAdd} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// Registrar gasto (bottom sheet en 2 pasos)
// ---------------------------------------------------------------------------
function AddExpenseSheet({ defaultDate, onClose, onSave }) {
  const [step, setStep] = useState(1);
  const [query, setQuery] = useState('');
  const [activeMacro, setActiveMacro] = useState('Todo');
  const [chosen, setChosen] = useState(null); // [macro, sub, presu]
  const [amount, setAmount] = useState('');
  const [account, setAccount] = useState(ACCOUNTS[0].name);
  const [date, setDate] = useState(defaultDate);
  const [note, setNote] = useState('');

  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    return CATALOGO.filter(c =>
      (activeMacro === 'Todo' || c[0] === activeMacro) &&
      (q === '' || c[1].toLowerCase().includes(q))
    );
  }, [query, activeMacro]);

  const canSave = chosen && parseFloat(amount) > 0 && date;

  const handleSave = () => {
    if (!canSave) return;
    onSave({
      date, macro: chosen[0], sub: chosen[1], account,
      amount: parseFloat(amount), note: note.trim(),
      time: new Date().toTimeString().slice(0, 5),
    });
  };

  return (
    <Sheet onClose={onClose} tall>
      {step === 1 && (
        <>
          <div className="flex items-center justify-between mb-4">
            <h2 className="text-lg font-bold">¿En qué gastaste?</h2>
            <button onClick={onClose} className="text-neutral-500"><X size={22} /></button>
          </div>
          <div className="relative mb-3">
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-500" />
            <input
              autoFocus
              value={query}
              onChange={e => setQuery(e.target.value)}
              placeholder="Buscar subcategoría..."
              className="w-full bg-neutral-900 rounded-xl py-2.5 pl-9 pr-3 text-sm outline-none placeholder-neutral-600"
            />
          </div>
          <div className="flex gap-2 overflow-x-auto pb-3 mb-1" style={{ scrollbarWidth: 'none' }}>
            <button
              onClick={() => setActiveMacro('Todo')}
              className="px-3 py-1.5 rounded-full text-xs font-semibold shrink-0"
              style={{ backgroundColor: activeMacro === 'Todo' ? '#dc2626' : '#171717', color: activeMacro === 'Todo' ? '#fff' : '#a3a3a3' }}
            >Todo</button>
            {MACROS.map(m => (
              <button
                key={m.name}
                onClick={() => setActiveMacro(m.name)}
                className="px-3 py-1.5 rounded-full text-xs font-semibold shrink-0"
                style={{ backgroundColor: activeMacro === m.name ? m.color : '#171717', color: activeMacro === m.name ? '#000' : '#a3a3a3' }}
              >{m.name}</button>
            ))}
          </div>
          <div className="space-y-1.5 max-h-[48vh] overflow-y-auto">
            {results.slice(0, 60).map((c, i) => {
              const meta = macroMeta(c[0]);
              return (
                <button
                  key={i}
                  onClick={() => { setChosen(c); setStep(2); }}
                  className="w-full flex items-center gap-3 bg-neutral-950 rounded-xl p-2.5 text-left active:opacity-70"
                >
                  <div className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: meta.color }} />
                  <div className="flex-1 min-w-0">
                    <div className="text-sm font-medium truncate">{c[1]}</div>
                    <div className="text-xs text-neutral-500">{c[0]}</div>
                  </div>
                  <div className="text-xs text-neutral-600 shrink-0">~{fmt(c[2])}</div>
                </button>
              );
            })}
            {results.length === 0 && <div className="text-center text-neutral-600 text-sm py-8">Sin resultados.</div>}
          </div>
        </>
      )}

      {step === 2 && chosen && (
        <>
          <div className="flex items-center gap-3 mb-5">
            <button onClick={() => setStep(1)} className="text-neutral-400"><ArrowLeft size={20} /></button>
            <div className="min-w-0">
              <div className="font-bold leading-tight truncate">{chosen[1]}</div>
              <div className="text-xs text-neutral-500">{chosen[0]}</div>
            </div>
            <button onClick={onClose} className="ml-auto text-neutral-500"><X size={22} /></button>
          </div>

          <div className="mb-5">
            <div className="text-xs text-neutral-500 mb-1.5">Monto</div>
            <div className="flex items-center bg-neutral-900 rounded-xl px-4 py-3">
              <span className="text-2xl font-bold text-neutral-500 mr-1">$</span>
              <input
                autoFocus
                inputMode="decimal"
                value={amount}
                onChange={e => setAmount(e.target.value.replace(/[^0-9.]/g, ''))}
                placeholder="0"
                className="bg-transparent text-2xl font-bold w-full outline-none placeholder-neutral-700"
              />
            </div>
          </div>

          <div className="mb-5">
            <div className="text-xs text-neutral-500 mb-1.5">Cuenta de pago</div>
            <div className="flex gap-2 overflow-x-auto" style={{ scrollbarWidth: 'none' }}>
              {ACCOUNTS.map(a => (
                <button
                  key={a.name}
                  onClick={() => setAccount(a.name)}
                  className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-medium shrink-0"
                  style={{ backgroundColor: account === a.name ? '#dc2626' : '#171717', color: account === a.name ? '#fff' : '#a3a3a3' }}
                >
                  <a.Icon size={14} /> {a.name}
                </button>
              ))}
            </div>
          </div>

          <div className="mb-5">
            <div className="text-xs text-neutral-500 mb-1.5">Fecha</div>
            <input
              type="date"
              value={date}
              onChange={e => setDate(e.target.value)}
              className="w-full bg-neutral-900 rounded-xl py-2.5 px-3 text-sm outline-none"
            />
          </div>

          <div className="mb-6">
            <div className="text-xs text-neutral-500 mb-1.5">Nota (opcional)</div>
            <input
              value={note}
              onChange={e => setNote(e.target.value)}
              placeholder="Ej. Oxxo, gasolinera..."
              className="w-full bg-neutral-900 rounded-xl py-2.5 px-3 text-sm outline-none placeholder-neutral-600"
            />
          </div>

          <button
            disabled={!canSave}
            onClick={handleSave}
            className="w-full py-3.5 rounded-2xl font-bold flex items-center justify-center gap-2"
            style={{ backgroundColor: canSave ? '#dc2626' : '#262626', color: canSave ? '#fff' : '#525252' }}
          >
            <Check size={18} /> Guardar gasto
          </button>
        </>
      )}
    </Sheet>
  );
}

function ConfirmDeleteSheet({ tx, onCancel, onConfirm }) {
  return (
    <Sheet onClose={onCancel}>
      <div className="text-center py-2">
        <div className="w-12 h-12 rounded-full bg-red-500/10 text-red-500 flex items-center justify-center mx-auto mb-4">
          <Trash2 size={20} />
        </div>
        <div className="font-bold mb-1">Eliminar movimiento</div>
        <div className="text-sm text-neutral-500 mb-6">{tx.sub} · {fmt(tx.amount)}</div>
        <div className="flex gap-3">
          <button onClick={onCancel} className="flex-1 py-3 rounded-xl font-semibold bg-neutral-900 text-neutral-300">Cancelar</button>
          <button onClick={onConfirm} className="flex-1 py-3 rounded-xl font-semibold bg-red-600 text-white">Eliminar</button>
        </div>
      </div>
    </Sheet>
  );
}

// ---------------------------------------------------------------------------
// Utilidades de UI
// ---------------------------------------------------------------------------
function Sheet({ children, onClose, tall }) {
  return (
    <div className="fixed inset-0 z-50 flex justify-center">
      <div className="absolute inset-0 bg-black/70" onClick={onClose} />
      <div className="w-full max-w-md absolute bottom-0">
        <div
          className="bg-black rounded-t-3xl px-5 pt-3 pb-6 border-t border-neutral-800"
          style={{ maxHeight: tall ? '88vh' : '70vh', overflowY: 'auto' }}
          onClick={e => e.stopPropagation()}
        >
          <div className="w-10 h-1 bg-neutral-700 rounded-full mx-auto mb-4" />
          {children}
        </div>
      </div>
    </div>
  );
}

function FloatingAddButton({ onClick }) {
  return (
    <button
      onClick={onClick}
      className="fixed bottom-24 right-6 w-14 h-14 rounded-2xl bg-red-600 flex items-center justify-center shadow-lg active:scale-95 transition-transform"
      style={{ maxWidth: 'calc(28rem - 1.5rem)' }}
    >
      <Plus size={26} color="#fff" strokeWidth={2.5} />
    </button>
  );
}

function BottomNav({ tab, setTab }) {
  const items = [
    { id: 'hoy', label: 'Hoy', Icon: ListChecks },
    { id: 'categorias', label: 'Categorías', Icon: LayoutGrid },
    { id: 'calendario', label: 'Calendario', Icon: CalendarIcon },
    { id: 'resumen', label: 'Resumen', Icon: BarChart3 },
  ];
  return (
    <div className="fixed bottom-0 w-full max-w-md bg-black border-t border-neutral-900 flex justify-around items-center py-2.5 pb-safe">
      {items.map(it => {
        const active = tab === it.id;
        return (
          <button key={it.id} onClick={() => setTab(it.id)} className="flex flex-col items-center gap-1 px-2">
            <it.Icon size={22} color={active ? '#ef4444' : '#737373'} />
            <span className="text-[10px] font-medium" style={{ color: active ? '#ef4444' : '#737373' }}>{it.label}</span>
          </button>
        );
      })}
    </div>
  );
}

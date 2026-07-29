import * as XLSX from 'xlsx';
import type { ChangeEvent } from 'react';
import { supabase } from './supabaseClient';
import type { Client, TableRow, Gestion } from './types';
import {
  getClients as dbGetClients,
  saveClient,
  deleteClientDb,
  clearAllClients,
} from './db';

const formatCurrency = (num: number): string =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(num || 0);

const formatDate = (dateStr: string): string => {
  if (!dateStr || dateStr === '-') return '';
  const parts = dateStr.split('-');
  if (parts.length !== 3) return dateStr;
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
};

let clientsCache: Client[] = [];

function getClients(): Client[] {
  return clientsCache;
}

function persistClients(clients: Client[]): void {
  clientsCache = clients;
  clients.forEach((c) => {
    saveClient(c).catch((err) => console.error('Error guardando cliente', err));
  });
}

function persistClient(client: Client): void {
  const idx = clientsCache.findIndex((c) => c.id === client.id);
  if (idx > -1) clientsCache[idx] = client;
  else clientsCache.push(client);
  saveClient(client).catch((err) => console.error('Error guardando cliente', err));
}

let currentTableData: TableRow[] = [];
let currentMoraClientId: string | null = null;
let parsedExcelData: Record<string, unknown>[] = [];
let exportDataGeneral: Record<string, unknown>[] = [];
let exportDataEjecutivos: Record<string, unknown>[] = [];

let moraParams = [
  { min: 1, max: 15, val: 5.0 },
  { min: 16, max: 30, val: 7.0 },
  { min: 31, max: 60, val: 9.0 },
  { min: 61, max: 9999, val: 10.0 },
];
let cobranzaParams = [
  { min: 0, max: 19.99, val: 3 },
  { min: 20, max: 39.99, val: 5 },
  { min: 40, max: 59.99, val: 9 },
  { min: 60, max: 79.99, val: 12 },
  { min: 80, max: 99.99, val: 15 },
  { min: 100, max: 999999, val: 18 },
];

function showToast(message: string, type: 'success' | 'error' = 'success'): void {
  const container = document.getElementById('toast-container');
  if (!container) return;
  const toast = document.createElement('div');
  toast.className = `px-4 py-3 rounded-md shadow-lg text-white font-medium transition-all duration-300 transform translate-y-10 opacity-0 ${type === 'success' ? 'bg-emerald-600' : 'bg-red-600'}`;
  toast.innerText = message;
  container.appendChild(toast);
  setTimeout(() => toast.classList.remove('translate-y-10', 'opacity-0'), 10);
  setTimeout(() => {
    toast.classList.add('opacity-0');
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}

function showConfirm(message: string, onConfirm: () => void): void {
  const modal = document.getElementById('confirm-modal');
  const msgEl = document.getElementById('confirm-modal-message');
  const confirmBtn = document.getElementById('confirm-modal-yes');
  const cancelBtn = document.getElementById('confirm-modal-no');
  if (!modal || !msgEl || !confirmBtn || !cancelBtn) return;

  msgEl.innerText = message;
  modal.classList.remove('hidden');

  const newConfirmBtn = confirmBtn.cloneNode(true) as HTMLElement;
  confirmBtn.parentNode?.replaceChild(newConfirmBtn, confirmBtn);
  const newCancelBtn = cancelBtn.cloneNode(true) as HTMLElement;
  cancelBtn.parentNode?.replaceChild(newCancelBtn, cancelBtn);

  newConfirmBtn.addEventListener('click', () => {
    modal.classList.add('hidden');
    onConfirm();
  });
  newCancelBtn.addEventListener('click', () => {
    modal.classList.add('hidden');
  });
}

function clearDatabase(): void {
  showConfirm(
    '¿Estás seguro de que quieres BORRAR TODOS los clientes de la base de datos? Esto no se puede deshacer.',
    () => {
      clientsCache = [];
      parsedExcelData = [];
      clearAllClients().catch((err) => console.error('Error limpiando base', err));
      const fileInput = document.getElementById('excel-file-input') as HTMLInputElement | null;
      if (fileInput) fileInput.value = '';
      const previewContainer = document.getElementById('excel-preview-container');
      if (previewContainer) previewContainer.classList.add('hidden');
      const emptyState = document.getElementById('excel-empty-state');
      if (emptyState) emptyState.classList.remove('hidden');
      renderDashboard();
      showToast('Base de datos eliminada exitosamente.');
    },
  );
}

function handleExcelUpload(event: ChangeEvent<HTMLInputElement>): void {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target?.result as ArrayBuffer);
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    parsedExcelData = XLSX.utils.sheet_to_json(worksheet, { raw: false, dateNF: 'yyyy-mm-dd' }) as Record<string, unknown>[];
    if (parsedExcelData.length > 0) {
      renderExcelPreview();
    } else {
      showToast('El archivo Excel está vacío o no se pudo leer.', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function renderExcelPreview(): void {
  const container = document.getElementById('excel-preview-container');
  const emptyState = document.getElementById('excel-empty-state');
  const thead = document.getElementById('excel-preview-head');
  const tbody = document.getElementById('excel-preview-body');
  if (!container || !emptyState || !thead || !tbody) return;

  container.classList.remove('hidden');
  emptyState.classList.add('hidden');
  document.getElementById('excel-count')!.innerText = `(${parsedExcelData.length} registros)`;

  const keys = Object.keys(parsedExcelData[0]);
  thead.innerHTML = '<tr>' + keys.map((k) => `<th class="px-4 py-2 font-semibold text-left">${k}</th>`).join('') + '</tr>';
  tbody.innerHTML =
    parsedExcelData
      .slice(0, 10)
      .map((row, i) => {
        return `<tr class="${i % 2 === 0 ? 'bg-slate-50' : 'bg-white'} hover:bg-blue-50">` + keys.map((k) => `<td class="px-4 py-2">${row[k] !== undefined ? row[k] : ''}</td>`).join('') + '</tr>';
      })
      .join('');
  if (parsedExcelData.length > 10) {
    tbody.innerHTML += `<tr><td colspan="${keys.length}" class="px-4 py-3 text-center text-slate-500 italic">... y ${parsedExcelData.length - 10} filas más.</td></tr>`;
  }
}

function convertExcelDate(dateStr: unknown): string {
  if (!dateStr) return '';
  if (dateStr instanceof Date) {
    return `${dateStr.getFullYear()}-${String(dateStr.getMonth() + 1).padStart(2, '0')}-${String(dateStr.getDate()).padStart(2, '0')}`;
  }
  const str = dateStr.toString().trim();
  const parts = str.split('-');
  if (parts.length === 3) return str;
  const partsSlash = str.split('/');
  if (partsSlash.length === 3) {
    let m = partsSlash[0], d = partsSlash[1], y = partsSlash[2];
    if (y.length === 2) y = '20' + y;
    if (m.length === 1) m = '0' + m;
    if (d.length === 1) d = '0' + d;
    return `${y}-${m}-${d}`;
  }
  return '';
}

function processExcelImport(): void {
  if (parsedExcelData.length === 0) return;
  showConfirm(`¿Importar ${parsedExcelData.length} registros? Los clientes existentes con la misma Cédula NO serán reemplazados.`, () => {
    let clients = getClients();
    let imported = 0;

    const getVal = (row: Record<string, unknown>, keys: string[]): unknown => {
      const rowKeys = Object.keys(row);
      for (const k of keys) {
        const match = rowKeys.find((rk) => rk.trim().toLowerCase() === k.toLowerCase());
        if (match) return row[match];
      }
      return '';
    };

    const parseNum = (val: unknown): number => {
      if (typeof val === 'number') return val;
      if (!val) return 0;
      const p = parseFloat(val.toString().replace(/[^0-9.-]+/g, ''));
      return isNaN(p) ? 0 : p;
    };

    parsedExcelData.forEach((row) => {
      const cedula = getVal(row, ['cedula', 'cédula', 'identificacion', 'documento'])!.toString().trim();
      if (!cedula) return;
      const exist = clients.find((c) => c.docIdentidad === cedula);
      if (exist) return;

      const fAdj = convertExcelDate(getVal(row, ['fecha adjudicacion', 'fecha adjudicación']));
      const fEnt = convertExcelDate(getVal(row, ['fecha entrega']));
      let fInsc = convertExcelDate(getVal(row, ['fecha iinscripción', 'fecha inscripción', 'fecha inscripcion']));
      if (!fInsc) {
        const today = new Date();
        fInsc = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
      }

      const pAdj = getVal(row, ['adj. por', 'adjudicado por']) as string;
      const ePlan = pAdj ? 'Adjudicado' : 'No Adjudicado';

      const vCuota = parseNum(getVal(row, ['cuota', 'valor cuota']));
      const cuotasP = parseInt(String(parseNum(getVal(row, ['cuotas pagadas', 'cuotas canceladas'])))) || 0;
      const vMonto = parseNum(getVal(row, ['monto', 'monto contratado', 'valor contratado', 'monto total', 'capital', 'monto base']));
      const montoFinal = vMonto > 0 ? vMonto : vCuota * 72;

      const cuotasVencidasEx = parseInt(String(parseNum(getVal(row, ['vencidas', 'cuotas vencidas', 'atrasadas', 'cuotas en mora'])))) || 0;
      const valorVencidoEx = parseNum(getVal(row, ['vencidos', 'valor vencido', 'saldo vencido', 'vencido', 'monto vencido', 'total vencido']));

      const grp = getVal(row, ['grupo']) as string;
      const pst = getVal(row, ['puesto']) as string;
      const fullGrupo = grp + (pst ? ' - ' + pst : '');
      const ejecutivo = (getVal(row, ['ejecutivo asignado', 'ejecutivo', 'ejecutivo cliente', 'cobrador']) as string) || 'Sin Asignar';

      const idCodigo = getVal(row, ['idcodigo', 'codigo', 'código', 'id código']) as string;
      let tipoPlanAsignado = 'Compra Planificada';
      if (idCodigo && typeof idCodigo === 'string') {
        if (idCodigo.toUpperCase().includes('ADP')) tipoPlanAsignado = 'Adjudicación Planificada';
        else if (idCodigo.toUpperCase().includes('ACV')) tipoPlanAsignado = 'Compra Planificada';
      }

      const newClient: Client = {
        id: Date.now().toString() + Math.random().toString(36).substr(2, 9),
        nombres: (getVal(row, ['cliente', 'nombres', 'nombre']) as string) || 'Sin Nombre',
        docIdentidad: cedula,
        grupoCodigo: fullGrupo,
        ejecutivoCartera: ejecutivo,
        tipoPlan: tipoPlanAsignado,
        estadoActivo: 'Activo',
        montoContratado: montoFinal,
        plazoPlan: 72,
        valorCuota: vCuota,
        valorInscripcion: 0,
        estadoPlan: ePlan,
        formaAdjudicacion: pAdj || '',
        fechaAdjudicacion: fAdj,
        numeroAsamblea: (getVal(row, ['asamblea adjudicación', 'asamblea adjudicacion']) as string) || '',
        fechaEntrega: fEnt,
        cuotasPagadas: cuotasP,
        fechaPrimerPago: fInsc,
        valorEntrada: 0,
        porcentajeEntrada: 0,
        fechaPagoEntrada: '',
        gestiones: [],
        cuotasVencidasExcel: cuotasVencidasEx,
        valorVencidoExcel: valorVencidoEx,
        totalPlan: 0,
        tableData: [],
      };
      newClient.totalPlan = newClient.valorCuota * newClient.plazoPlan + newClient.valorEntrada;
      newClient.tableData = generateInitialTableData(newClient);
      clients.push(newClient);
      imported++;
    });

    persistClients(clients);
    renderDashboard();
    showToast(`Se importaron exitosamente ${imported} clientes nuevos.`);
    switchTab('dashboard');
  });
}

function renderDashboard(): void {
  const clients = getClients();
  const tbody = document.getElementById('dashboard-tbody');
  const emptyState = document.getElementById('dashboard-empty');
  if (!tbody || !emptyState) return;

  tbody.innerHTML = '';
  if (clients.length === 0) {
    emptyState.classList.remove('hidden');
  } else {
    emptyState.classList.add('hidden');
    clients.forEach((client) => {
      const ejecutivoStr = client.ejecutivoCartera || 'Sin Asignar';
      const tr = document.createElement('tr');
      tr.className = 'hover:bg-slate-50 transition-colors';
      tr.innerHTML = `
        <td class="px-4 py-3 text-sm font-medium text-slate-900">${client.nombres}</td>
        <td class="px-4 py-3 text-sm text-slate-500">${client.docIdentidad}</td>
        <td class="px-4 py-3 text-sm text-slate-500"><div>${client.tipoPlan}</div><div class="text-[10px] text-slate-400 font-bold uppercase mt-0.5">Grupo: ${client.grupoCodigo}</div></td>
        <td class="px-4 py-3 text-sm font-semibold text-slate-700">${formatCurrency(client.montoContratado)}</td>
        <td class="px-4 py-3 text-sm text-slate-600">${ejecutivoStr}</td>
        <td class="px-4 py-3 text-center text-sm font-medium space-x-2">
          <button type="button" data-action="view" data-id="${client.id}" class="inline-flex text-emerald-700 bg-emerald-100 hover:bg-emerald-200 px-2.5 py-1.5 rounded-md transition-colors shadow-sm">Tabla</button>
          <button type="button" data-action="edit" data-id="${client.id}" class="inline-flex text-blue-700 bg-blue-100 hover:bg-blue-200 px-2.5 py-1.5 rounded-md transition-colors shadow-sm">Editar</button>
          <button type="button" data-action="delete" data-id="${client.id}" class="inline-flex text-red-700 bg-red-100 hover:bg-red-200 px-2.5 py-1.5 rounded-md transition-colors shadow-sm">Borrar</button>
        </td>
      `;
      tbody.appendChild(tr);
    });
  }
}

function filterDashboard(): void {
  const input = document.getElementById('searchInput') as HTMLInputElement | null;
  if (!input) return;
  const text = input.value.toLowerCase();
  const rows = document.getElementById('dashboard-tbody')!.getElementsByTagName('tr');
  for (let i = 0; i < rows.length; i++) {
    const nameCol = rows[i].getElementsByTagName('td')[0];
    const docCol = rows[i].getElementsByTagName('td')[1];
    const ejeCol = rows[i].getElementsByTagName('td')[4];
    if (nameCol || docCol || ejeCol) {
      if (
        nameCol.innerText.toLowerCase().indexOf(text) > -1 ||
        docCol.innerText.toLowerCase().indexOf(text) > -1 ||
        (ejeCol && ejeCol.innerText.toLowerCase().indexOf(text) > -1)
      ) {
        rows[i].style.display = '';
      } else {
        rows[i].style.display = 'none';
      }
    }
  }
}

function clearForm(): void {
  const form = document.getElementById('client-form') as HTMLFormElement | null;
  if (form) form.reset();
  const dp = document.getElementById('display-total-plan');
  if (dp) dp.innerText = '$0.00';
  const vtp = document.getElementById('valorTotalPagado') as HTMLInputElement | null;
  if (vtp) vtp.value = '';
  const fpe = document.getElementById('fechaPagoEntrada') as HTMLInputElement | null;
  if (fpe) fpe.value = '';
  const ea = document.getElementById('estadoActivo') as HTMLSelectElement | null;
  if (ea) ea.value = 'Activo';
  const ej = document.getElementById('ejecutivoCartera') as HTMLInputElement | null;
  if (ej) ej.value = '';
  handlePlanTypeChange();
  handleEstadoPlanChange();
}

function createNewClient(): void {
  clearForm();
  const cid = document.getElementById('clientId') as HTMLInputElement | null;
  if (cid) cid.value = '';
  document.getElementById('empty-state')?.classList.remove('hidden');
  document.getElementById('statement-view')?.classList.add('hidden');
  switchTab('client-info');
}

function handlePlanTypeChange(): void {
  const sel = document.getElementById('tipoPlan') as HTMLSelectElement | null;
  if (!sel) return;
  const isAdjPlan = sel.value === 'Adjudicación Planificada';
  document.getElementById('col-porcentaje')!.classList.toggle('opacity-40', !isAdjPlan);
  document.getElementById('col-valor-entrada')!.classList.toggle('opacity-40', !isAdjPlan);
  document.getElementById('col-fecha-pago-entrada')!.classList.toggle('opacity-40', !isAdjPlan);
  (document.getElementById('porcentajeEntrada') as HTMLInputElement).disabled = !isAdjPlan;
  (document.getElementById('fechaPagoEntrada') as HTMLInputElement).disabled = !isAdjPlan;
  if (!isAdjPlan) {
    (document.getElementById('porcentajeEntrada') as HTMLInputElement).value = '';
    (document.getElementById('fechaPagoEntrada') as HTMLInputElement).value = '';
    (document.getElementById('valorEntrada') as HTMLInputElement).value = '';
  }
  calculateValues();
}

function handleEstadoPlanChange(): void {
  const sel = document.getElementById('estadoPlan') as HTMLSelectElement | null;
  if (!sel) return;
  const isAdj = sel.value === 'Adjudicado';
  document.getElementById('col-forma-adjudicacion')!.classList.toggle('hidden', !isAdj);
  document.getElementById('col-fecha-adjudicacion')!.classList.toggle('hidden', !isAdj);
  document.getElementById('col-num-asamblea')!.classList.toggle('hidden', !isAdj);
  document.getElementById('col-fecha-entrega')!.classList.toggle('hidden', !isAdj);
  if (!isAdj) {
    (document.getElementById('formaAdjudicacion') as HTMLInputElement).value = '';
    (document.getElementById('fechaAdjudicacion') as HTMLInputElement).value = '';
    (document.getElementById('numeroAsamblea') as HTMLInputElement).value = '';
    (document.getElementById('fechaEntrega') as HTMLInputElement).value = '';
  }
}

function calculateValues(): void {
  const m = parseFloat((document.getElementById('montoContratado') as HTMLInputElement).value) || 0;
  const p = parseInt((document.getElementById('plazoPlan') as HTMLInputElement).value) || 0;
  const v = parseFloat((document.getElementById('valorCuota') as HTMLInputElement).value) || 0;
  const c = parseInt((document.getElementById('cuotasPagadas') as HTMLInputElement).value) || 0;
  const t = (document.getElementById('tipoPlan') as HTMLSelectElement).value;

  let valEntrada = 0;
  if (t === 'Adjudicación Planificada') {
    const por = parseFloat((document.getElementById('porcentajeEntrada') as HTMLInputElement).value) || 0;
    valEntrada = m * (por / 100);
    (document.getElementById('valorEntrada') as HTMLInputElement).value = valEntrada.toFixed(2);
  }

  const total = v * p + valEntrada;
  document.getElementById('display-total-plan')!.innerText = formatCurrency(total);
  (document.getElementById('valorTotalPagado') as HTMLInputElement).value = (c * v).toFixed(2);
}

function editClient(id: string): void {
  const c = getClients().find((x) => x.id === id);
  if (!c) return;

  (document.getElementById('clientId') as HTMLInputElement).value = c.id;
  (document.getElementById('nombres') as HTMLInputElement).value = c.nombres;
  (document.getElementById('docIdentidad') as HTMLInputElement).value = c.docIdentidad;
  (document.getElementById('ejecutivoCartera') as HTMLInputElement).value = c.ejecutivoCartera || '';
  (document.getElementById('tipoPlan') as HTMLSelectElement).value = c.tipoPlan || 'Compra Planificada';
  (document.getElementById('estadoActivo') as HTMLSelectElement).value = c.estadoActivo || 'Activo';
  (document.getElementById('grupoCodigo') as HTMLInputElement).value = c.grupoCodigo;
  (document.getElementById('montoContratado') as HTMLInputElement).value = String(c.montoContratado);
  (document.getElementById('valorInscripcion') as HTMLInputElement).value = String(c.valorInscripcion);
  (document.getElementById('plazoPlan') as HTMLInputElement).value = String(c.plazoPlan);
  (document.getElementById('valorCuota') as HTMLInputElement).value = String(c.valorCuota);
  (document.getElementById('estadoPlan') as HTMLSelectElement).value = c.estadoPlan || 'No Adjudicado';
  (document.getElementById('cuotasPagadas') as HTMLInputElement).value = String(c.cuotasPagadas);
  (document.getElementById('fechaPrimerPago') as HTMLInputElement).value = c.fechaPrimerPago;
  (document.getElementById('fechaPagoEntrada') as HTMLInputElement).value = c.fechaPagoEntrada || '';

  handlePlanTypeChange();
  handleEstadoPlanChange();

  (document.getElementById('formaAdjudicacion') as HTMLInputElement).value = c.formaAdjudicacion || '';
  (document.getElementById('fechaAdjudicacion') as HTMLInputElement).value = c.fechaAdjudicacion || '';
  (document.getElementById('numeroAsamblea') as HTMLInputElement).value = c.numeroAsamblea || '';
  (document.getElementById('fechaEntrega') as HTMLInputElement).value = c.fechaEntrega || '';

  if (c.tipoPlan === 'Adjudicación Planificada') {
    (document.getElementById('porcentajeEntrada') as HTMLInputElement).value = ((c.valorEntrada / c.montoContratado) * 100).toFixed(1);
  }
  calculateValues();
  switchTab('client-info');
}

function deleteClient(id: string): void {
  showConfirm('¿Eliminar este cliente y todos sus pagos?', () => {
    clientsCache = getClients().filter((c) => c.id !== id);
    deleteClientDb(id).catch((err) => console.error('Error eliminando cliente', err));
    renderDashboard();
    const cid = document.getElementById('clientId') as HTMLInputElement | null;
    if (cid && cid.value === id) {
      createNewClient();
      switchTab('dashboard');
    }
    showToast('Cliente eliminado.');
  });
}

function saveData(goToTable: boolean): void {
  const form = document.getElementById('client-form') as HTMLFormElement | null;
  if (!form || !form.checkValidity()) {
    form?.reportValidity();
    return;
  }

  const isNew = !(document.getElementById('clientId') as HTMLInputElement).value;
  const cId = isNew ? Date.now().toString() : (document.getElementById('clientId') as HTMLInputElement).value;
  const existingClient = getClients().find((c) => c.id === cId);

  const data: Client = {
    id: cId,
    nombres: (document.getElementById('nombres') as HTMLInputElement).value,
    docIdentidad: (document.getElementById('docIdentidad') as HTMLInputElement).value,
    ejecutivoCartera: (document.getElementById('ejecutivoCartera') as HTMLInputElement).value || 'Sin Asignar',
    tipoPlan: (document.getElementById('tipoPlan') as HTMLSelectElement).value,
    estadoActivo: (document.getElementById('estadoActivo') as HTMLSelectElement).value,
    grupoCodigo: (document.getElementById('grupoCodigo') as HTMLInputElement).value,
    montoContratado: parseFloat((document.getElementById('montoContratado') as HTMLInputElement).value) || 0,
    valorInscripcion: parseFloat((document.getElementById('valorInscripcion') as HTMLInputElement).value) || 0,
    plazoPlan: parseInt((document.getElementById('plazoPlan') as HTMLInputElement).value) || 0,
    valorCuota: parseFloat((document.getElementById('valorCuota') as HTMLInputElement).value) || 0,
    estadoPlan: (document.getElementById('estadoPlan') as HTMLSelectElement).value,
    formaAdjudicacion: (document.getElementById('formaAdjudicacion') as HTMLInputElement).value,
    fechaAdjudicacion: (document.getElementById('fechaAdjudicacion') as HTMLInputElement).value,
    numeroAsamblea: (document.getElementById('numeroAsamblea') as HTMLInputElement).value,
    fechaEntrega: (document.getElementById('fechaEntrega') as HTMLInputElement).value,
    cuotasPagadas: parseInt((document.getElementById('cuotasPagadas') as HTMLInputElement).value) || 0,
    fechaPrimerPago: (document.getElementById('fechaPrimerPago') as HTMLInputElement).value,
    valorEntrada: parseFloat((document.getElementById('valorEntrada') as HTMLInputElement).value) || 0,
    porcentajeEntrada: parseFloat((document.getElementById('porcentajeEntrada') as HTMLInputElement).value) || 0,
    fechaPagoEntrada: (document.getElementById('fechaPagoEntrada') as HTMLInputElement).value,
    gestiones: existingClient?.gestiones || [],
    cuotasVencidasExcel: existingClient?.cuotasVencidasExcel || 0,
    valorVencidoExcel: existingClient?.valorVencidoExcel || 0,
    totalPlan: 0,
    tableData: [],
  };
  data.totalPlan = data.valorCuota * data.plazoPlan + data.valorEntrada;

  const clients = getClients();
  const idx = clients.findIndex((c) => c.id === cId);

  if (isNew || idx === -1) {
    currentTableData = generateInitialTableData(data);
    finalizeSave(data, idx, goToTable, true);
  } else {
    const syncedTable = syncTableData(data, clients[idx].tableData || generateInitialTableData(data));
    if (clients[idx].plazoPlan !== data.plazoPlan) {
      showConfirm('El plazo cambió. Se ajustará la tabla conservando los abonos. ¿Continuar?', () => {
        currentTableData = syncedTable;
        finalizeSave(data, idx, goToTable, false);
      });
    } else {
      currentTableData = syncedTable;
      finalizeSave(data, idx, goToTable, false);
    }
  }
}

function finalizeSave(data: Client, idx: number, goToTable: boolean, isNew: boolean): void {
  data.tableData = currentTableData;
  const clients = getClients();
  if (isNew) {
    clients.push(data);
    (document.getElementById('clientId') as HTMLInputElement).value = data.id;
  } else {
    clients[idx] = data;
  }
  persistClient(data);
  renderDashboard();
  populateStatement(data);
  renderTable(data);

  document.getElementById('empty-state')?.classList.add('hidden');
  document.getElementById('statement-view')?.classList.remove('hidden');

  if (goToTable) {
    showToast('Guardado. Visualizando tabla...');
    switchTab('payment-table');
  } else {
    showToast('Información y Estado de Cuenta actualizados exitosamente.');
  }
}

function saveFromTable(): void {
  const id = (document.getElementById('clientId') as HTMLInputElement).value;
  if (!id) {
    showToast('Error: Cliente no activo.', 'error');
    return;
  }
  const clients = getClients();
  const idx = clients.findIndex((c) => c.id === id);
  if (idx > -1) {
    clients[idx].tableData = currentTableData;
    let cCanceladas = 0;
    currentTableData.forEach((r) => {
      if ((parseFloat(String(r.abono)) || 0) >= r.cuotaMensual && r.cuotaMensual > 0) cCanceladas++;
    });
    clients[idx].cuotasPagadas = cCanceladas;
    persistClient(clients[idx]);

    (document.getElementById('cuotasPagadas') as HTMLInputElement).value = String(cCanceladas);
    calculateValues();
    renderDashboard();
    populateStatement(clients[idx]);
    renderTable(clients[idx]);
    showToast('Tabla guardada exitosamente.');
  }
}

function viewClient(id: string): void {
  const c = getClients().find((x) => x.id === id);
  if (!c) return;
  editClient(id);
  currentTableData = c.tableData || [];
  populateStatement(c);
  renderTable(c);
  document.getElementById('empty-state')?.classList.add('hidden');
  document.getElementById('statement-view')?.classList.remove('hidden');
  document.getElementById('fecha-generacion')!.innerText = new Date().toLocaleDateString('es-ES', { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' });
  switchTab('payment-table');
}

function generateInitialTableData(data: Client): TableRow[] {
  const table: TableRow[] = [];
  if (!data.fechaPrimerPago) return table;

  const parts = data.fechaPrimerPago.split('-');
  let fYear = parseInt(parts[0]), fMonth = parseInt(parts[1]), fDay = parseInt(parts[2]);
  let vCuota = parseFloat(String(data.valorCuota));
  if (isNaN(vCuota) || vCuota < 0) vCuota = 0;

  const totalPagadas = parseInt(String(data.cuotasPagadas)) || 0;

  for (let i = 1; i <= data.plazoPlan; i++) {
    let vStr = '';
    if (i === 1) {
      vStr = `${fYear}-${fMonth < 10 ? '0' + fMonth : fMonth}-${fDay < 10 ? '0' + fDay : fDay}`;
    } else {
      fMonth++;
      if (fMonth > 12) {
        fMonth = 1;
        fYear++;
      }
      vStr = `${fYear}-${fMonth < 10 ? '0' + fMonth : fMonth}-05`;
    }
    const isPagado = i <= totalPagadas;
    table.push({ numero: i, vencimiento: vStr, cuotaMensual: vCuota, abono: isPagado ? vCuota : 0, fechaPago: isPagado ? vStr : '', descMora: 0, descCobranza: 0 });
  }
  return table;
}

function syncTableData(data: Client, oldData: TableRow[]): TableRow[] {
  const table: TableRow[] = [];
  if (!data.fechaPrimerPago) return oldData;

  const parts = data.fechaPrimerPago.split('-');
  let fYear = parseInt(parts[0]), fMonth = parseInt(parts[1]), fDay = parseInt(parts[2]);
  let vCuota = parseFloat(String(data.valorCuota));
  if (isNaN(vCuota) || vCuota < 0) vCuota = 0;

  const totalPagadas = parseInt(String(data.cuotasPagadas)) || 0;

  for (let i = 1; i <= data.plazoPlan; i++) {
    let vStr = '';
    if (i === 1) {
      vStr = `${fYear}-${fMonth < 10 ? '0' + fMonth : fMonth}-${fDay < 10 ? '0' + fDay : fDay}`;
    } else {
      fMonth++;
      if (fMonth > 12) {
        fMonth = 1;
        fYear++;
      }
      vStr = `${fYear}-${fMonth < 10 ? '0' + fMonth : fMonth}-05`;
    }

    const ex = oldData.find((r) => r.numero === i);
    const isPagada = i <= totalPagadas;

    if (ex) {
      let nuevoAbono = isPagada ? vCuota : 0;
      if (!isPagada) {
        const viejoAbono = parseFloat(String(ex.abono));
        if (!isNaN(viejoAbono) && viejoAbono > 0) {
          nuevoAbono = viejoAbono > vCuota ? vCuota : viejoAbono;
        }
      }
      table.push({
        numero: i,
        vencimiento: ex.vencimiento || vStr,
        cuotaMensual: vCuota,
        abono: nuevoAbono,
        fechaPago: isPagada ? ex.fechaPago || vStr : '',
        descMora: ex.descMora || 0,
        descCobranza: ex.descCobranza || 0,
      });
    } else {
      table.push({ numero: i, vencimiento: vStr, cuotaMensual: vCuota, abono: isPagada ? vCuota : 0, fechaPago: isPagada ? vStr : '', descMora: 0, descCobranza: 0 });
    }
  }
  return table;
}

function populateStatement(data: Client): void {
  document.getElementById('lbl-nombre')!.innerText = data.nombres;
  document.getElementById('lbl-doc')!.innerText = data.docIdentidad;
  document.getElementById('lbl-grupo')!.innerText = data.grupoCodigo;
  document.getElementById('lbl-tipo-plan')!.innerText = data.tipoPlan;
  document.getElementById('lbl-ejecutivo')!.innerText = data.ejecutivoCartera || 'Sin Asignar';

  const eEl = document.getElementById('lbl-estado-plan')!;
  if (data.estadoPlan === 'Adjudicado') {
    eEl.innerText = data.formaAdjudicacion ? `${data.estadoPlan} (${data.formaAdjudicacion})` : data.estadoPlan;
    document.querySelectorAll('.lbl-adjudicacion').forEach((el) => el.classList.remove('hidden'));
    document.getElementById('lbl-fecha-adj')!.innerText = formatDate(data.fechaAdjudicacion) || '-';
    document.getElementById('lbl-num-asamblea')!.innerText = data.numeroAsamblea || '-';
  } else {
    eEl.innerText = data.estadoPlan || '-';
    document.querySelectorAll('.lbl-adjudicacion').forEach((el) => el.classList.add('hidden'));
  }
  eEl.className = data.estadoPlan === 'Adjudicado' ? 'font-bold text-emerald-600' : 'font-bold text-amber-500';

  document.getElementById('lbl-monto')!.innerText = formatCurrency(data.montoContratado);
  document.getElementById('lbl-plazo')!.innerText = `${data.plazoPlan} Meses`;
  document.getElementById('lbl-cuota')!.innerText = formatCurrency(data.valorCuota);
  document.getElementById('lbl-inscripcion')!.innerText = formatCurrency(data.valorInscripcion);

  if (data.tipoPlan === 'Adjudicación Planificada') {
    document.getElementById('lbl-entrada')!.innerText = formatCurrency(data.valorEntrada);
    (document.getElementById('box-entrada') as HTMLElement).style.display = 'block';
    document.querySelectorAll('.lbl-adj-planificada').forEach((el) => el.classList.remove('hidden'));
    document.getElementById('lbl-porcentaje-entrada')!.innerText = (data.porcentajeEntrada || 0) + '%';
  } else {
    document.getElementById('lbl-entrada')!.innerText = '-';
    (document.getElementById('box-entrada') as HTMLElement).style.display = 'none';
    document.querySelectorAll('.lbl-adj-planificada').forEach((el) => el.classList.add('hidden'));
  }
  document.getElementById('lbl-total-plan')!.innerText = formatCurrency(data.totalPlan);
}

function calculateDaysLate(vencimiento: string, fechaPago: string): string {
  if (!fechaPago) return '-';
  const vDate = new Date(vencimiento + 'T00:00:00');
  const pDate = new Date(fechaPago + 'T00:00:00');
  const diffDays = Math.ceil((pDate.getTime() - vDate.getTime()) / (1000 * 60 * 60 * 24));
  return diffDays > 0 ? String(diffDays) : '0';
}

function renderTable(data: Client): void {
  const tbody = document.getElementById('table-body');
  if (!tbody) return;
  tbody.innerHTML = '';

  let saldoPlan = data.totalPlan;
  let abonoEntrada = 0;
  let totalP = 0;
  let cCanceladas = 0;

  if (data.tipoPlan === 'Adjudicación Planificada') {
    const salIniE = saldoPlan;
    abonoEntrada = data.fechaPagoEntrada ? data.valorEntrada : 0;
    const salCuoE = data.valorEntrada - abonoEntrada;
    saldoPlan -= abonoEntrada;
    const bClass = data.fechaPagoEntrada ? 'bg-blue-100 text-blue-800' : 'bg-emerald-100 text-emerald-800';
    tbody.innerHTML += `
      <tr class="bg-blue-50 border-b-2 border-blue-200">
        <td class="px-2 py-2 text-center text-sm font-bold text-blue-700">Ent</td>
        <td class="px-3 py-2 text-right text-sm text-slate-500">${formatCurrency(salIniE)}</td>
        <td class="px-3 py-2 text-right text-sm font-semibold text-slate-800">${formatCurrency(data.valorEntrada)}</td>
        <td class="px-3 py-2 text-right text-sm font-bold text-emerald-600">${formatCurrency(abonoEntrada)}</td>
        <td class="px-3 py-2 text-right text-sm font-semibold ${salCuoE > 0 ? 'text-red-500' : 'text-slate-400'}">${formatCurrency(salCuoE)}</td>
        <td class="px-3 py-2 text-right text-sm font-bold text-blue-800">${formatCurrency(saldoPlan)}</td>
        <td class="px-3 py-2 text-center text-sm text-slate-700">-</td>
        <td class="px-3 py-2 text-center text-sm text-slate-700">${formatDate(data.fechaPagoEntrada) || '-'}</td>
        <td class="px-2 py-2 text-center text-sm text-slate-500">-</td>
        <td class="px-3 py-2 text-center"><span class="px-2 py-1 inline-flex text-[10px] font-bold rounded-full uppercase ${bClass}">${data.fechaPagoEntrada ? 'Cancelada' : 'Pendiente'}</span></td>
      </tr>
    `;
  }

  const today = new Date();
  const todayStr = today.getFullYear() + '-' + String(today.getMonth() + 1).padStart(2, '0') + '-' + String(today.getDate()).padStart(2, '0');

  for (let i = 0; i < currentTableData.length; i++) {
    const row = currentTableData[i];
    const saldoIni = saldoPlan;
    const cMensual = parseFloat(String(row.cuotaMensual)) || 0;
    const abono = parseFloat(String(row.abono)) || 0;
    const salCuota = cMensual - abono;

    saldoPlan -= abono;
    if (saldoPlan < 0) saldoPlan = 0;

    let estCuota = '';
    let bClass = '';
    let onClickStr = '';
    const isOverdue = salCuota > 0 && row.vencimiento < todayStr;

    if (abono >= cMensual && cMensual > 0) {
      estCuota = 'Cancelada';
      bClass = 'bg-blue-100 text-blue-800';
      cCanceladas++;
    } else if (isOverdue) {
      const hasEntrega = data.estadoPlan === 'Adjudicado' && data.fechaEntrega;
      estCuota = hasEntrega ? (abono > 0 ? 'Parcial Venc.' : 'Vencido') : abono > 0 ? 'Parcial P.V.' : 'Pendiente V.';
      bClass = 'bg-red-100 text-red-800 shadow-sm ring-1 ring-red-300 cursor-pointer hover:bg-red-200 transition-colors';
      onClickStr = `data-mora-action="open" data-mora-id="${data.id}" title="Calcular Mora y Cobranzas"`;
    } else {
      estCuota = abono > 0 ? 'Parcial' : 'Pendiente';
      bClass = 'bg-emerald-100 text-emerald-800';
    }

    totalP += abono;
    const diasPago = calculateDaysLate(row.vencimiento, row.fechaPago);

    tbody.innerHTML += `
      <tr class="${i % 2 === 0 ? 'bg-slate-50' : 'bg-white'} hover:bg-blue-50">
        <td class="px-2 py-2 text-center text-sm font-bold text-slate-700">${row.numero}</td>
        <td class="px-3 py-2 text-right text-sm text-slate-500">${formatCurrency(saldoIni)}</td>
        <td class="px-3 py-2 text-right text-sm">
          <input type="number" step="0.01" value="${cMensual.toFixed(2)}" class="w-24 text-right rounded-md border-slate-300 shadow-sm sm:text-sm p-1 border no-print focus:ring-blue-500 focus:border-blue-500" data-row-idx="${i}" data-row-field="cuotaMensual" data-client-id="${data.id}">
          <span class="print-only hidden font-semibold text-slate-800">${formatCurrency(cMensual)}</span>
        </td>
        <td class="px-3 py-2 text-right text-sm">
          <input type="number" step="0.01" value="${abono.toFixed(2)}" class="w-24 text-right rounded-md border-slate-300 shadow-sm sm:text-sm p-1 border no-print focus:ring-blue-500 focus:border-blue-500" data-row-idx="${i}" data-row-field="abono" data-client-id="${data.id}">
          <span class="print-only hidden">${formatCurrency(abono)}</span>
        </td>
        <td class="px-3 py-2 text-right text-sm font-semibold ${salCuota > 0 ? 'text-red-500' : 'text-slate-400'}">${formatCurrency(salCuota)}</td>
        <td class="px-3 py-2 text-right text-sm font-bold text-blue-800">${formatCurrency(saldoPlan)}</td>
        <td class="px-3 py-2 text-center text-sm">
          <input type="date" value="${row.vencimiento}" class="w-32 text-center rounded-md border-slate-300 shadow-sm sm:text-sm p-1 border no-print focus:ring-blue-500 focus:border-blue-500" data-row-idx="${i}" data-row-field="vencimiento" data-client-id="${data.id}">
          <span class="print-only hidden">${formatDate(row.vencimiento) || '-'}</span>
        </td>
        <td class="px-3 py-2 text-center text-sm">
          <input type="date" value="${row.fechaPago}" class="w-32 text-center rounded-md border-slate-300 shadow-sm sm:text-sm p-1 border no-print focus:ring-blue-500 focus:border-blue-500" data-row-idx="${i}" data-row-field="fechaPago" data-client-id="${data.id}">
          <span class="print-only hidden">${formatDate(row.fechaPago) || '-'}</span>
        </td>
        <td class="px-2 py-2 text-center text-sm text-slate-500">${diasPago}</td>
        <td class="px-3 py-2 text-center">
          <span ${onClickStr} class="px-2 py-1 inline-flex text-[10px] font-bold rounded-full uppercase ${bClass}">${estCuota}</span>
        </td>
      </tr>
    `;
  }

  document.getElementById('sum-cuotas-pagadas')!.innerText = String(cCanceladas);
  document.getElementById('sum-valor-pagado')!.innerText = formatCurrency(totalP + abonoEntrada);
  document.getElementById('sum-cuotas-pendientes')!.innerText = String(Math.max(0, data.plazoPlan - cCanceladas));
  document.getElementById('sum-valor-pendiente')!.innerText = formatCurrency(Math.max(0, data.totalPlan - (totalP + abonoEntrada)));
}

function updateRowData(idx: number, field: keyof TableRow, value: string, clientId: string): void {
  if (field === 'abono') {
    let val = parseFloat(value);
    if (isNaN(val) || val < 0) val = 0;
    if (val > currentTableData[idx].cuotaMensual) val = currentTableData[idx].cuotaMensual;
    currentTableData[idx].abono = val;
  } else if (field === 'cuotaMensual') {
    let val = parseFloat(value);
    if (isNaN(val) || val < 0) val = 0;
    currentTableData[idx].cuotaMensual = val;
    currentTableData[idx].abono = val;
  } else if (field === 'fechaPago') {
    currentTableData[idx].fechaPago = value;
  } else if (field === 'vencimiento') {
    currentTableData[idx].vencimiento = value;
    if (value) {
      const parts = value.split('-');
      let fY = parseInt(parts[0], 10), fM = parseInt(parts[1], 10);
      for (let k = idx + 1; k < currentTableData.length; k++) {
        fM++;
        if (fM > 12) {
          fM = 1;
          fY++;
        }
        currentTableData[k].vencimiento = `${fY}-${fM < 10 ? '0' + fM : fM}-05`;
      }
    }
  }

  const clients = getClients();
  const cIdx = clients.findIndex((c) => c.id === clientId);
  if (cIdx > -1) {
    clients[cIdx].tableData = currentTableData;
    let cCan = 0;
    currentTableData.forEach((r) => {
      if ((parseFloat(String(r.abono)) || 0) >= r.cuotaMensual && r.cuotaMensual > 0) cCan++;
    });
    clients[cIdx].cuotasPagadas = cCan;
    persistClient(clients[cIdx]);

    const cidEl = document.getElementById('clientId') as HTMLInputElement | null;
    if (cidEl && cidEl.value === clientId) {
      (document.getElementById('cuotasPagadas') as HTMLInputElement).value = String(cCan);
      calculateValues();
    }
    renderTable(clients[cIdx]);
  }
}

function calcularTasaAnual(client: Client): number {
  if (!client || client.plazoPlan <= 0 || client.montoContratado <= 0) return 0;
  const monto = client.montoContratado;
  const totalCuotasVal = client.valorCuota * client.plazoPlan;
  let diferencia = 0;
  if (client.tipoPlan === 'Adjudicación Planificada') {
    diferencia = totalCuotasVal - (monto - (client.valorEntrada || 0));
  } else {
    diferencia = totalCuotasVal - monto;
  }
  const anios = client.plazoPlan / 12;
  return (diferencia / monto / anios) * 100;
}

function openMoraTab(id: string): void {
  currentMoraClientId = id;
  document.getElementById('tab-btn-mora-cobranzas')?.classList.remove('hidden');
  switchTab('mora-cobranzas');

  const dateInput = document.getElementById('fecha-calculo-mora') as HTMLInputElement | null;
  if (dateInput && !dateInput.value) {
    const today = new Date();
    dateInput.value = today.getFullYear() + '-' + String(today.getMonth() + 1).padStart(2, '0') + '-' + String(today.getDate()).padStart(2, '0');
  }

  renderMoraParamsUI();
  calculateAndRenderMora();
  document.getElementById('mora-fecha-generacion')!.innerText = new Date().toLocaleDateString('es-ES', { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' });

  const client = getClients().find((c) => c.id === id);
  renderGestiones(client);
}

function guardarGestion(): void {
  if (!currentMoraClientId) return;
  const texto = (document.getElementById('nueva-gestion-texto') as HTMLTextAreaElement).value.trim();
  if (!texto) {
    showToast('Ingrese el detalle de la gestión.', 'error');
    return;
  }
  const clients = getClients();
  const cIdx = clients.findIndex((c) => c.id === currentMoraClientId);
  if (cIdx > -1) {
    if (!clients[cIdx].gestiones) clients[cIdx].gestiones = [];
    const gestion: Gestion = { fecha: new Date().toISOString(), texto };
    clients[cIdx].gestiones.push(gestion);
    persistClient(clients[cIdx]);
    (document.getElementById('nueva-gestion-texto') as HTMLTextAreaElement).value = '';
    showToast('Gestión guardada exitosamente.');
    renderGestiones(clients[cIdx]);
  }
}

function renderGestiones(client: Client | undefined): void {
  const list = document.getElementById('historial-gestiones-list');
  if (!list) return;
  list.innerHTML = '';
  if (!client || !client.gestiones || client.gestiones.length === 0) {
    list.innerHTML = '<p class="text-sm text-slate-500 italic">No hay gestiones registradas para este cliente.</p>';
    return;
  }
  const sortedGestiones = [...client.gestiones].sort((a, b) => new Date(b.fecha).getTime() - new Date(a.fecha).getTime());
  sortedGestiones.forEach((g) => {
    const f = new Date(g.fecha);
    const options: Intl.DateTimeFormatOptions = { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' };
    const fechaStr = f.toLocaleDateString('es-ES', options);
    list.innerHTML += `
      <div class="bg-white p-3 rounded-md border border-slate-200 shadow-sm">
        <div class="text-[10px] text-slate-500 font-bold mb-1 uppercase tracking-wider">${fechaStr}</div>
        <div class="text-sm text-slate-800 whitespace-pre-wrap">${g.texto}</div>
      </div>
    `;
  });
}

function updateCondonacion(clientId: string, rowIdx: number, field: 'descMora' | 'descCobranza', value: string): void {
  const clients = getClients();
  const cIdx = clients.findIndex((c) => c.id === clientId);
  if (cIdx > -1) {
    let val = parseFloat(value) || 0;
    if (val < 0) val = 0;
    if (val > 100) val = 100;
    clients[cIdx].tableData[rowIdx][field] = val;
    persistClient(clients[cIdx]);
    if (currentMoraClientId === clientId) calculateAndRenderMora();
  }
}

function renderMoraParamsUI(): void {
  let tasaAdminAnual = 0;
  if (currentMoraClientId) {
    const client = getClients().find((c) => c.id === currentMoraClientId);
    tasaAdminAnual = calcularTasaAnual(client!);
  }

  document.getElementById('mora-params-body')!.innerHTML = moraParams
    .map((p, i) => {
      const recargo = tasaAdminAnual * (p.val / 100);
      const nuevaTasa = tasaAdminAnual + recargo;
      const tasaDiaria = (nuevaTasa / 365).toFixed(4);
      return `
      <tr class="hover:bg-slate-100 transition-colors">
        <td class="p-1"><input type="number" class="w-full text-center border-slate-300 rounded shadow-sm p-1" value="${p.min}" data-param-type="mora" data-param-idx="${i}" data-param-field="min"></td>
        <td class="p-1"><input type="number" class="w-full text-center border-slate-300 rounded shadow-sm p-1" value="${p.max}" data-param-type="mora" data-param-idx="${i}" data-param-field="max"></td>
        <td class="p-1"><input type="number" step="0.01" class="w-full text-center border-slate-300 rounded shadow-sm p-1" value="${p.val}" data-param-type="mora" data-param-idx="${i}" data-param-field="val"></td>
        <td class="p-1 text-center font-semibold text-slate-600 bg-slate-50">${tasaDiaria}%</td>
        <td class="p-1 text-center"><button class="text-red-400 hover:text-red-600" data-param-remove="mora" data-param-idx="${i}">X</button></td>
      </tr>
    `;
    })
    .join('');

  document.getElementById('cobranza-params-body')!.innerHTML = cobranzaParams
    .map((p, i) => `
      <tr class="hover:bg-slate-100 transition-colors">
        <td class="p-1"><input type="number" step="0.01" class="w-full text-center border-slate-300 rounded shadow-sm p-1" value="${p.min}" data-param-type="cobranza" data-param-idx="${i}" data-param-field="min"></td>
        <td class="p-1"><input type="number" step="0.01" class="w-full text-center border-slate-300 rounded shadow-sm p-1" value="${p.max}" data-param-type="cobranza" data-param-idx="${i}" data-param-field="max"></td>
        <td class="p-1"><input type="number" step="0.01" class="w-full text-center border-slate-300 rounded shadow-sm p-1" value="${p.val}" data-param-type="cobranza" data-param-idx="${i}" data-param-field="val"></td>
        <td class="p-1 text-center"><button class="text-red-400 hover:text-red-600" data-param-remove="cobranza" data-param-idx="${i}">X</button></td>
      </tr>
    `)
    .join('');
}

function updateParam(type: 'mora' | 'cobranza', index: number, field: string, value: string): void {
  const val = parseFloat(value) || 0;
  if (type === 'mora') (moraParams[index] as Record<string, number>)[field] = val;
  else (cobranzaParams[index] as Record<string, number>)[field] = val;
  renderMoraParamsUI();
  calculateAndRenderMora();
}

function addParam(type: 'mora' | 'cobranza'): void {
  if (type === 'mora') moraParams.push({ min: 0, max: 0, val: 0 });
  else cobranzaParams.push({ min: 0, max: 0, val: 0 });
  renderMoraParamsUI();
}

function removeParam(type: 'mora' | 'cobranza', index: number): void {
  if (type === 'mora') moraParams.splice(index, 1);
  else cobranzaParams.splice(index, 1);
  renderMoraParamsUI();
  calculateAndRenderMora();
}

function calculateAndRenderMora(_event?: ChangeEvent<HTMLInputElement> | Event): void {
  if (!currentMoraClientId) return;
  const client = getClients().find((c) => c.id === currentMoraClientId);
  if (!client) return;

  document.getElementById('mora-client-name')!.innerText = client.nombres;
  document.getElementById('mora-client-plan')!.innerText = `${client.tipoPlan} - ${client.estadoPlan}`;
  document.getElementById('mora-client-fecha')!.innerText = client.fechaEntrega ? `Entrega Vehículo: ${formatDate(client.fechaEntrega)}` : 'Sin Entrega de Vehículo (Solo Info)';

  const estadoActivo = client.estadoActivo || 'Activo';
  const spanEstado = document.getElementById('mora-client-estado-activo')!;
  spanEstado.innerText = estadoActivo;
  spanEstado.className = estadoActivo === 'Activo' ? 'px-2 py-1 rounded-md text-sm font-bold uppercase bg-emerald-100 text-emerald-700' : 'px-2 py-1 rounded-md text-sm font-bold uppercase bg-red-100 text-red-700';

  document.getElementById('mora-frm-grupo')!.innerText = client.grupoCodigo || '-';
  document.getElementById('mora-frm-plazo')!.innerText = `${client.plazoPlan || 0} Meses`;

  const tasaAdminAnual = calcularTasaAnual(client);
  let totalCuotasVal = client.valorCuota * client.plazoPlan;
  if (client.tipoPlan === 'Adjudicación Planificada') totalCuotasVal += client.valorEntrada || 0;

  document.getElementById('mora-frm-monto')!.innerText = formatCurrency(client.montoContratado);
  document.getElementById('mora-frm-total-cuotas')!.innerText = formatCurrency(totalCuotasVal);
  document.getElementById('mora-frm-anios')!.innerText = (client.plazoPlan / 12).toFixed(2) + ' Años';
  document.getElementById('mora-frm-tasa-admin')!.innerText = tasaAdminAnual.toFixed(2) + '%';

  const hasEntrega = client.estadoPlan === 'Adjudicado' && !!client.fechaEntrega;
  const tbody = document.getElementById('mora-results-body')!;
  tbody.innerHTML = '';

  let tMora = 0, tCob = 0, tCuotas = 0;

  const fechaCalculoInput = (document.getElementById('fecha-calculo-mora') as HTMLInputElement).value;
  let todayStr = fechaCalculoInput;
  if (!todayStr) {
    const today = new Date();
    todayStr = today.getFullYear() + '-' + String(today.getMonth() + 1).padStart(2, '0') + '-' + String(today.getDate()).padStart(2, '0');
  }

  client.tableData.forEach((row, rowIdx) => {
    const salCuota = row.cuotaMensual - (parseFloat(String(row.abono)) || 0);
    if (salCuota > 0 && row.vencimiento < todayStr) {
      const vDate = new Date(row.vencimiento + 'T00:00:00');
      const tDate = new Date(todayStr + 'T00:00:00');
      const diffDays = Math.ceil((tDate.getTime() - vDate.getTime()) / (1000 * 60 * 60 * 24));
      if (diffDays >= 1) {
        let moraVal = 0, cobVal = 0;
        tCuotas += salCuota;
        const descM = parseFloat(String(row.descMora)) || 0;
        const descC = parseFloat(String(row.descCobranza)) || 0;

        if (hasEntrega) {
          let rateTabla = 0;
          moraParams.forEach((p) => {
            if (diffDays >= p.min && diffDays <= p.max) rateTabla = p.val;
          });
          const recargo = tasaAdminAnual * (rateTabla / 100);
          const nuevaTasa = tasaAdminAnual + recargo;
          const rateDiaria = nuevaTasa / 365;
          const originalMora = salCuota * (rateDiaria / 100) * diffDays;
          let originalCob = 0;
          if (diffDays >= 16) {
            cobranzaParams.forEach((p) => {
              if (salCuota >= p.min && salCuota <= p.max) originalCob = p.val;
            });
          }
          moraVal = originalMora * (1 - descM / 100);
          cobVal = originalCob * (1 - descC / 100);
        }

        tMora += moraVal;
        tCob += cobVal;

        tbody.innerHTML += `
          <tr class="hover:bg-slate-50 transition-colors border-b border-slate-100">
            <td class="px-2 py-2 text-center text-sm font-bold text-slate-700">${row.numero}</td>
            <td class="px-2 py-2 text-center text-sm text-slate-500">${formatDate(row.vencimiento)}</td>
            <td class="px-2 py-2 text-center text-sm font-semibold text-red-500">${diffDays}</td>
            <td class="px-2 py-2 text-right text-sm font-bold text-slate-800">${formatCurrency(salCuota)}</td>
            <td class="px-2 py-2 text-right text-sm text-amber-600 font-semibold">${hasEntrega ? formatCurrency(moraVal) : '-'}</td>
            <td class="px-1 py-2 text-center no-print">
              ${hasEntrega ? `<input type="number" min="0" max="100" value="${descM}" data-condonacion-id="${client.id}" data-condonacion-idx="${rowIdx}" data-condonacion-field="descMora" class="w-12 text-center rounded border-emerald-300 text-emerald-700 text-xs p-1 focus:ring-emerald-500 bg-white" title="Descuento Mora %">` : '-'}
            </td>
            <td class="px-2 py-2 text-right text-sm text-rose-600 font-semibold">${hasEntrega ? formatCurrency(cobVal) : '-'}</td>
            <td class="px-1 py-2 text-center no-print">
              ${hasEntrega ? `<input type="number" min="0" max="100" value="${descC}" data-condonacion-id="${client.id}" data-condonacion-idx="${rowIdx}" data-condonacion-field="descCobranza" class="w-12 text-center rounded border-emerald-300 text-emerald-700 text-xs p-1 focus:ring-emerald-500 bg-white" title="Descuento Cobranza %">` : '-'}
            </td>
            <td class="px-2 py-2 text-right text-sm font-black text-blue-900">${formatCurrency(salCuota + moraVal + cobVal)}</td>
          </tr>
        `;
      }
    }
  });

  if (tbody.innerHTML === '') {
    tbody.innerHTML = `<tr><td colspan="9" class="px-3 py-6 text-center text-sm text-slate-500">No hay cuotas vencidas para calcular.</td></tr>`;
  }

  document.getElementById('sum-mora-cuotas')!.innerText = formatCurrency(tCuotas);
  document.getElementById('sum-mora-recargos')!.innerText = formatCurrency(tMora);
  document.getElementById('sum-mora-cobranzas')!.innerText = formatCurrency(tCob);
  document.getElementById('sum-mora-total')!.innerText = formatCurrency(tCuotas + tMora + tCob);
}

function openReportesTab(): void {
  switchTab('reportes');
  populateReportFilters();
  generateReports();
}

function populateReportFilters(): void {
  const clients = getClients();
  const filterEjecutivo = document.getElementById('report-filter-ejecutivo') as HTMLSelectElement | null;
  if (!filterEjecutivo) return;
  const currentVal = filterEjecutivo.value;
  const ejecutivos = [...new Set(clients.map((c) => c.ejecutivoCartera).filter((e) => e && e !== 'Sin Asignar'))].sort();
  filterEjecutivo.innerHTML = '<option value="Todos">Todos</option>';
  ejecutivos.forEach((ej) => {
    filterEjecutivo.innerHTML += `<option value="${ej}">${ej}</option>`;
  });
  if (ejecutivos.includes(currentVal) || currentVal === 'Todos') {
    filterEjecutivo.value = currentVal;
  }
}

function generateReports(): void {
  const clients = getClients();
  const searchInput = (document.getElementById('report-search') as HTMLInputElement | null)?.value.toLowerCase() || '';
  const filterEstado = (document.getElementById('report-filter-estado') as HTMLSelectElement | null)?.value || 'Todos';
  const filterEjecutivo = (document.getElementById('report-filter-ejecutivo') as HTMLSelectElement | null)?.value || 'Todos';
  const filterVencidas = (document.getElementById('report-filter-vencidas') as HTMLInputElement | null)?.value || '';

  const filteredClients = clients.filter((c) => {
    let matchSearch = true;
    if (searchInput) {
      const searchStr = `${c.nombres} ${c.docIdentidad} ${c.grupoCodigo}`.toLowerCase();
      matchSearch = searchStr.includes(searchInput);
    }
    const matchEstado = filterEstado === 'Todos' || c.estadoPlan === filterEstado;
    const ejecutivoStr = c.ejecutivoCartera || 'Sin Asignar';
    const matchEjecutivo = filterEjecutivo === 'Todos' || ejecutivoStr === filterEjecutivo;
    let matchVencidas = true;
    if (filterVencidas !== '') {
      const cuotasVencidasEx = parseInt(String(c.cuotasVencidasExcel)) || 0;
      matchVencidas = cuotasVencidasEx === parseInt(filterVencidas);
    }
    return matchSearch && matchEstado && matchEjecutivo && matchVencidas;
  });

  const today = new Date();
  const currentY = today.getFullYear();
  const currentM = today.getMonth() + 1;
  const meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
  const monthName = meses[today.getMonth()];

  const genTitle = document.getElementById('report-general-title');
  if (genTitle) genTitle.innerText = `Reporte General de Clientes (${monthName} ${currentY})`;
  const ejTitle = document.getElementById('report-ejecutivo-title');
  if (ejTitle) ejTitle.innerText = `Recaudación por Ejecutivo (${monthName} ${currentY})`;

  exportDataGeneral = [];
  const reportEjecutivosMap: Record<string, { clientes: number; recaudadoMes: number; vencido: number }> = {};
  const tbodyGen = document.getElementById('report-general-body')!;
  tbodyGen.innerHTML = '';

  if (filteredClients.length === 0) {
    tbodyGen.innerHTML = `<tr><td colspan="13" class="px-3 py-6 text-center text-sm text-slate-500">No hay clientes que coincidan con los filtros de búsqueda.</td></tr>`;
    document.getElementById('report-ejecutivos-body')!.innerHTML = `<tr><td colspan="4" class="px-3 py-6 text-center text-sm text-slate-500">No hay datos para mostrar.</td></tr>`;
    return;
  }

  filteredClients.forEach((c, idx) => {
    const cuotasVencidas = parseInt(String(c.cuotasVencidasExcel)) || 0;
    const valorVencido = parseFloat(String(c.valorVencidoExcel)) || 0;
    let cuotasCobradasMes = 0;
    let valorCobradoMes = 0;

    if (c.tipoPlan === 'Adjudicación Planificada' && c.fechaPagoEntrada) {
      const pParts = c.fechaPagoEntrada.split('-');
      if (pParts.length >= 2 && parseInt(pParts[0]) === currentY && parseInt(pParts[1]) === currentM) {
        valorCobradoMes += parseFloat(String(c.valorEntrada)) || 0;
        cuotasCobradasMes += 1;
      }
    }
    if (c.tableData) {
      c.tableData.forEach((row) => {
        const abono = parseFloat(String(row.abono)) || 0;
        if (row.fechaPago && abono > 0) {
          const pParts = row.fechaPago.split('-');
          if (pParts.length >= 2 && parseInt(pParts[0]) === currentY && parseInt(pParts[1]) === currentM) {
            valorCobradoMes += abono;
            cuotasCobradasMes += 1;
          }
        }
      });
    }

    const cuotasPendientes = cuotasVencidas - cuotasCobradasMes;
    const valorPendiente = valorVencido - valorCobradoMes;
    const ejecutivo = c.ejecutivoCartera || 'Sin Asignar';

    exportDataGeneral.push({
      'Nombre Completo': c.nombres,
      Identificación: c.docIdentidad,
      'Grupo/Código': c.grupoCodigo,
      'Monto Contratado': parseFloat(String(c.montoContratado)) || 0,
      'Tipo de Plan': c.tipoPlan,
      'Estado del Plan': c.estadoPlan,
      'Estado Activo': c.estadoActivo,
      'Cuota Mensual': parseFloat(String(c.valorCuota)) || 0,
      'Cuotas Vencidas (Excel)': cuotasVencidas,
      'Valor Total Vencido (Excel)': parseFloat(valorVencido.toFixed(2)),
      'Cuotas Cobradas (Mes Actual)': cuotasCobradasMes,
      'Valor Cobrado (Mes Actual)': parseFloat(valorCobradoMes.toFixed(2)),
      'Cuotas Pendientes': cuotasPendientes,
      'Valor Pendiente': parseFloat(valorPendiente.toFixed(2)),
      'Ejecutivo de Cartera': ejecutivo,
    });

    tbodyGen.innerHTML += `
      <tr class="${idx % 2 === 0 ? 'bg-slate-50' : 'bg-white'} hover:bg-blue-50 transition-colors">
        <td class="px-3 py-2 text-left font-medium text-slate-800">${c.nombres}</td>
        <td class="px-3 py-2 text-left text-slate-600">${c.docIdentidad}</td>
        <td class="px-3 py-2 text-left text-slate-600">${c.grupoCodigo}</td>
        <td class="px-3 py-2 text-right font-semibold text-slate-700">${formatCurrency(c.montoContratado)}</td>
        <td class="px-3 py-2 text-center text-[10px] uppercase font-bold text-slate-500">${c.estadoPlan}</td>
        <td class="px-3 py-2 text-right font-semibold text-slate-700">${formatCurrency(c.valorCuota)}</td>
        <td class="px-3 py-2 text-center font-bold ${cuotasVencidas > 0 ? 'text-red-500' : 'text-slate-400'}">${cuotasVencidas}</td>
        <td class="px-3 py-2 text-right font-bold ${valorVencido > 0 ? 'text-red-600' : 'text-slate-400'}">${formatCurrency(valorVencido)}</td>
        <td class="px-3 py-2 text-center font-bold text-emerald-600 bg-emerald-50">${cuotasCobradasMes}</td>
        <td class="px-3 py-2 text-right font-bold text-emerald-700 bg-emerald-50">${formatCurrency(valorCobradoMes)}</td>
        <td class="px-3 py-2 text-center font-bold text-amber-600 bg-amber-50">${cuotasPendientes}</td>
        <td class="px-3 py-2 text-right font-bold text-amber-700 bg-amber-50">${formatCurrency(valorPendiente)}</td>
        <td class="px-3 py-2 text-left font-bold text-blue-900 bg-blue-50">${ejecutivo}</td>
      </tr>
    `;

    if (!reportEjecutivosMap[ejecutivo]) {
      reportEjecutivosMap[ejecutivo] = { clientes: 0, recaudadoMes: 0, vencido: 0 };
    }
    reportEjecutivosMap[ejecutivo].clientes++;
    reportEjecutivosMap[ejecutivo].recaudadoMes += valorCobradoMes;
    reportEjecutivosMap[ejecutivo].vencido += valorVencido;
  });

  const tbodyEj = document.getElementById('report-ejecutivos-body')!;
  tbodyEj.innerHTML = '';
  exportDataEjecutivos = [];

  Object.keys(reportEjecutivosMap)
    .sort()
    .forEach((ej) => {
      const data = reportEjecutivosMap[ej];
      exportDataEjecutivos.push({
        'Ejecutivo de Cartera': ej,
        'Total Clientes Asignados': data.clientes,
        'Recaudado en el Mes Actual': parseFloat(data.recaudadoMes.toFixed(2)),
        'Total en Mora (Vencido)': parseFloat(data.vencido.toFixed(2)),
      });
      tbodyEj.innerHTML += `
        <tr class="hover:bg-slate-50 transition-colors">
          <td class="px-4 py-3 text-left font-bold text-slate-800">${ej}</td>
          <td class="px-4 py-3 text-center font-semibold text-slate-600">${data.clientes}</td>
          <td class="px-4 py-3 text-right font-black text-emerald-600 bg-emerald-50">${formatCurrency(data.recaudadoMes)}</td>
          <td class="px-4 py-3 text-right font-bold text-red-500">${formatCurrency(data.vencido)}</td>
        </tr>
      `;
    });
}

function exportToExcel(type: 'general' | 'ejecutivos'): void {
  const dataToExport = type === 'general' ? exportDataGeneral : exportDataEjecutivos;
  const fileName = type === 'general' ? 'Reporte_General_Clientes.xlsx' : 'Reporte_Recaudacion_Ejecutivos.xlsx';
  if (dataToExport.length === 0) {
    showToast('No hay datos para exportar.', 'error');
    return;
  }
  const ws = XLSX.utils.json_to_sheet(dataToExport);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Reportes');
  XLSX.writeFile(wb, fileName);
  showToast('Reporte descargado exitosamente.');
}

function switchTab(tabId: string): void {
  document.querySelectorAll('.tab-content').forEach((el) => {
    el.classList.remove('active');
    setTimeout(() => ((el as HTMLElement).style.display = 'none'), 150);
  });
  document.querySelectorAll('.tab-btn').forEach((el) => {
    el.classList.remove('border-blue-600', 'text-blue-700');
    el.classList.add('border-transparent', 'text-gray-500');
  });

  setTimeout(() => {
    const activeContent = document.getElementById(`tab-${tabId}`);
    if (activeContent) {
      activeContent.style.display = 'block';
      setTimeout(() => activeContent.classList.add('active'), 50);
    }
  }, 150);

  const activeBtn = document.getElementById(`tab-btn-${tabId}`);
  if (activeBtn) {
    activeBtn.classList.remove('border-transparent', 'text-gray-500');
    activeBtn.classList.add('border-blue-600', 'text-blue-700');
  }
}

function handleImageUpload(e: ChangeEvent<HTMLInputElement>): void {
  const input = e.target as HTMLInputElement;
  if (input.files?.[0]) {
    const reader = new FileReader();
    reader.onload = (evt) => {
      const dataUrl = evt.target?.result as string;
      const logoEl = document.getElementById('statement-logo') as HTMLImageElement | null;
      const logoContainer = document.getElementById('logo-container');
      if (logoEl) logoEl.src = dataUrl;
      logoContainer?.classList.remove('hidden');
      const moraLogoEl = document.getElementById('mora-statement-logo') as HTMLImageElement | null;
      const moraContainerEl = document.getElementById('mora-logo-container');
      if (moraLogoEl) moraLogoEl.src = dataUrl;
      moraContainerEl?.classList.remove('hidden');
      localStorage.setItem('company_logo', dataUrl);
    };
    reader.readAsDataURL(input.files[0]);
  }
}

function attachEventListeners(): void {
  const dashboardBody = document.getElementById('dashboard-tbody');
  if (dashboardBody) {
    dashboardBody.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest<HTMLElement>('[data-action]');
      if (!btn) return;
      const action = btn.dataset.action;
      const id = btn.dataset.id;
      if (action === 'view' && id) viewClient(id);
      else if (action === 'edit' && id) editClient(id);
      else if (action === 'delete' && id) deleteClient(id);
    });
  }

  const tableBody = document.getElementById('table-body');
  if (tableBody) {
    tableBody.addEventListener('blur', (e) => {
      const inp = (e.target as HTMLElement).closest<HTMLInputElement>('input[data-row-idx]');
      if (!inp) return;
      const idx = parseInt(inp.dataset.rowIdx!);
      const field = inp.dataset.rowField as keyof TableRow;
      const clientId = inp.dataset.clientId!;
      updateRowData(idx, field, inp.value, clientId);
    }, true);
    tableBody.addEventListener('click', (e) => {
      const el = (e.target as HTMLElement).closest<HTMLElement>('[data-mora-action="open"]');
      if (!el) return;
      const id = el.dataset.moraId;
      if (id) openMoraTab(id);
    });
  }

  const moraBody = document.getElementById('mora-results-body');
  if (moraBody) {
    moraBody.addEventListener('blur', (e) => {
      const inp = (e.target as HTMLElement).closest<HTMLInputElement>('[data-condonacion-id]');
      if (!inp) return;
      const id = inp.dataset.condonacionId!;
      const idx = parseInt(inp.dataset.condonacionIdx!);
      const field = inp.dataset.condonacionField as 'descMora' | 'descCobranza';
      updateCondonacion(id, idx, field, inp.value);
    }, true);
  }

  const moraParamsBody = document.getElementById('mora-params-body');
  if (moraParamsBody) {
    moraParamsBody.addEventListener('change', (e) => {
      const inp = (e.target as HTMLElement).closest<HTMLInputElement>('[data-param-type]');
      if (!inp) return;
      const type = inp.dataset.paramType as 'mora' | 'cobranza';
      const idx = parseInt(inp.dataset.paramIdx!);
      const field = inp.dataset.paramField!;
      updateParam(type, idx, field, inp.value);
    });
    moraParamsBody.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest<HTMLElement>('[data-param-remove]');
      if (!btn) return;
      const type = btn.dataset.paramRemove as 'mora' | 'cobranza';
      const idx = parseInt(btn.dataset.paramIdx!);
      removeParam(type, idx);
    });
  }

  const cobranzaParamsBody = document.getElementById('cobranza-params-body');
  if (cobranzaParamsBody) {
    cobranzaParamsBody.addEventListener('change', (e) => {
      const inp = (e.target as HTMLElement).closest<HTMLInputElement>('[data-param-type]');
      if (!inp) return;
      const type = inp.dataset.paramType as 'mora' | 'cobranza';
      const idx = parseInt(inp.dataset.paramIdx!);
      const field = inp.dataset.paramField!;
      updateParam(type, idx, field, inp.value);
    });
    cobranzaParamsBody.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest<HTMLElement>('[data-param-remove]');
      if (!btn) return;
      const type = btn.dataset.paramRemove as 'mora' | 'cobranza';
      const idx = parseInt(btn.dataset.paramIdx!);
      removeParam(type, idx);
    });
  }
}

export async function initApp(): Promise<void> {
  try {
    clientsCache = await dbGetClients();
  } catch (err) {
    console.error('Error cargando clientes desde Supabase', err);
    clientsCache = [];
  }

  renderDashboard();
  handlePlanTypeChange();
  handleEstadoPlanChange();
  switchTab('base');

  const savedLogo = localStorage.getItem('company_logo');
  if (savedLogo) {
    const logoEl = document.getElementById('statement-logo') as HTMLImageElement | null;
    const logoContainer = document.getElementById('logo-container');
    if (logoEl) logoEl.src = savedLogo;
    logoContainer?.classList.remove('hidden');
    const moraLogoEl = document.getElementById('mora-statement-logo') as HTMLImageElement | null;
    const moraContainerEl = document.getElementById('mora-logo-container');
    if (moraLogoEl) moraLogoEl.src = savedLogo;
    moraContainerEl?.classList.remove('hidden');
  }

  attachEventListeners();

  supabase.channel('clients-changes')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'clients' }, async () => {
      try {
        clientsCache = await dbGetClients();
        renderDashboard();
        if (document.getElementById('tab-reportes')?.classList.contains('active')) {
          populateReportFilters();
          generateReports();
        }
      } catch (err) {
        console.error('Error sincronizando cambios', err);
      }
    })
    .subscribe();
}

export {
  handleExcelUpload,
  processExcelImport,
  clearDatabase,
  filterDashboard,
  createNewClient,
  clearForm,
  saveData,
  saveFromTable,
  switchTab,
  openReportesTab,
  openMoraTab,
  guardarGestion,
  addParam,
  handleImageUpload,
  handlePlanTypeChange,
  handleEstadoPlanChange,
  calculateValues,
  generateReports,
  exportToExcel,
  calculateAndRenderMora,
  attachEventListeners,
};

import { useEffect, useRef } from 'react';
import {
  initApp,
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
} from './appLogic';

function App() {
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    initApp();
  }, []);

  return (
    <div ref={rootRef}>
      <header className="bg-blue-900 text-white shadow-md no-print">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between h-16 items-center">
            <div className="flex items-center">
              <svg className="h-8 w-8 mr-3 text-blue-300" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <h1 className="text-xl font-bold">Sistema de Aportes Mensuales</h1>
            </div>
          </div>
        </div>
      </header>

      <main className="flex-grow max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <div className="mb-6 border-b border-gray-200 no-print">
          <nav className="-mb-px flex space-x-8 overflow-x-auto" aria-label="Tabs">
            <button onClick={() => switchTab('base')} id="tab-btn-base" className="tab-btn whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm border-blue-600 text-blue-700">
              0. BASE (Importar)
            </button>
            <button onClick={() => switchTab('dashboard')} id="tab-btn-dashboard" className="tab-btn whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300">
              1. Bandeja de Gestión
            </button>
            <button onClick={() => switchTab('client-info')} id="tab-btn-client-info" className="tab-btn whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300">
              2. Información del Cliente
            </button>
            <button onClick={() => switchTab('payment-table')} id="tab-btn-payment-table" className="tab-btn whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300">
              3. Estado de Cuenta y Pagos
            </button>
            <button onClick={() => openMoraTab('')} id="tab-btn-mora-cobranzas" className="tab-btn whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300 hidden">
              4. Mora y Cobranzas
            </button>
            <button onClick={() => openReportesTab()} id="tab-btn-reportes" className="tab-btn whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300">
              5. Reportes y Productividad
            </button>
          </nav>
        </div>

        {/* 0. BASE IMPORTAR EXCEL */}
        <div id="tab-base" className="tab-content active bg-white shadow-lg rounded-xl border border-slate-100 p-6 no-print">
          <div className="mb-6 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <div>
              <h2 className="text-2xl font-semibold text-slate-800">Carga de Base de Datos</h2>
              <p className="text-sm text-slate-500 mt-1">Importa clientes desde un archivo Excel. Se vinculará automáticamente la información.</p>
            </div>
            <div className="flex flex-col sm:flex-row gap-3 w-full md:w-auto items-center">
              <input type="file" id="excel-file-input" accept=".xlsx, .xls, .csv" className="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 transition-colors" onChange={handleExcelUpload} />
              <button type="button" onClick={clearDatabase} className="px-4 py-2 bg-red-100 text-red-700 rounded-md hover:bg-red-200 transition-colors shadow-sm text-sm font-medium w-full sm:w-auto whitespace-nowrap">
                <svg className="w-4 h-4 inline-block mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg>
                Borrar Base Actual
              </button>
            </div>
          </div>

          <div id="excel-preview-container" className="hidden">
            <div className="flex justify-between items-center mb-3">
              <h3 className="text-lg font-medium text-slate-700">Vista Previa de Datos <span id="excel-count" className="text-sm font-normal text-slate-500 ml-2"></span></h3>
              <button onClick={processExcelImport} className="px-4 py-2 bg-emerald-600 text-white rounded-md hover:bg-emerald-700 transition-colors shadow-sm font-medium flex items-center">
                <svg className="w-5 h-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" /></svg>
                Importar a Bandeja
              </button>
            </div>
            <div className="table-container overflow-x-auto border border-slate-200 rounded-lg max-h-96">
              <table className="min-w-full divide-y divide-slate-200 text-sm whitespace-nowrap">
                <thead id="excel-preview-head" className="bg-slate-800 text-white text-xs uppercase tracking-wider sticky top-0"></thead>
                <tbody id="excel-preview-body" className="bg-white divide-y divide-slate-200"></tbody>
              </table>
            </div>
          </div>

          <div id="excel-empty-state" className="text-center py-16 px-4 border-2 border-dashed border-slate-300 rounded-lg bg-slate-50 mt-4">
            <svg className="mx-auto h-12 w-12 text-slate-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            <h3 className="mt-2 text-sm font-medium text-slate-900">Ningún archivo cargado</h3>
            <p className="mt-1 text-xs text-slate-500">Sube un archivo Excel (.xlsx) para previsualizar e importar la data.</p>
          </div>
        </div>

        {/* 1. BANDEJA DE GESTION */}
        <div id="tab-dashboard" className="tab-content bg-white shadow-lg rounded-xl border border-slate-100 p-6 no-print">
          <div className="flex justify-between items-center mb-6">
            <h2 className="text-2xl font-semibold text-slate-800">Bandeja de Gestión de Clientes</h2>
            <div className="space-x-2">
              <button type="button" onClick={createNewClient} className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors shadow-sm text-sm font-medium">+ Nuevo Cliente</button>
            </div>
          </div>

          <div className="mb-4">
            <input type="text" id="searchInput" onKeyUp={filterDashboard} placeholder="Buscar por nombre o documento..." className="w-full md:w-1/3 rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" />
          </div>

          <div className="table-container overflow-x-auto border border-slate-200 rounded-lg">
            <table className="min-w-full divide-y divide-slate-200 text-sm">
              <thead className="bg-slate-50 text-slate-700">
                <tr>
                  <th className="px-4 py-3 text-left font-semibold text-xs uppercase">Cliente</th>
                  <th className="px-4 py-3 text-left font-semibold text-xs uppercase">Documento</th>
                  <th className="px-4 py-3 text-left font-semibold text-xs uppercase">Plan / Grupo</th>
                  <th className="px-4 py-3 text-left font-semibold text-xs uppercase">Monto</th>
                  <th className="px-4 py-3 text-left font-semibold text-xs uppercase">Ejecutivo</th>
                  <th className="px-4 py-3 text-center font-semibold text-xs uppercase">Acciones</th>
                </tr>
              </thead>
              <tbody id="dashboard-tbody" className="bg-white divide-y divide-slate-200"></tbody>
            </table>
            <div id="dashboard-empty" className="hidden text-center py-8 text-slate-500 text-sm">
              No hay clientes registrados. Crea un nuevo cliente o importa una Base de Datos para empezar.
            </div>
          </div>
        </div>

        {/* 2. FORMULARIO DEL CLIENTE */}
        <div id="tab-client-info" className="tab-content bg-white shadow-lg rounded-xl border border-slate-100 p-6 no-print">
          <div className="flex justify-between items-center mb-6">
            <h2 className="text-2xl font-semibold text-slate-800">Datos del Plan y Cliente</h2>
            <div className="space-x-2 flex items-center">
              <input type="hidden" id="clientId" value="" />
              <button type="button" onClick={clearForm} className="px-3 py-2 bg-slate-100 text-slate-600 border border-slate-300 rounded-md hover:bg-slate-200 transition-colors text-sm font-medium hidden md:block">Limpiar Datos</button>
              <button type="button" onClick={() => saveData(false)} className="px-4 py-2 bg-emerald-600 text-white rounded-md hover:bg-emerald-700 transition-colors shadow-sm text-sm font-medium">Solo Guardar</button>
              <button type="button" onClick={() => saveData(true)} className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors shadow-sm text-sm font-medium">Guardar y Ver Tabla</button>
            </div>
          </div>

          <form id="client-form" className="space-y-8">
            <div className="bg-slate-50 p-5 rounded-lg border border-slate-200">
              <h3 className="text-xs font-bold text-slate-500 uppercase tracking-widest mb-4 border-b border-slate-200 pb-2">Identificación</h3>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div>
                  <label htmlFor="nombres" className="block text-sm font-medium text-slate-700 mb-1">1.1 Nombre y Apellidos</label>
                  <input type="text" id="nombres" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" required />
                </div>
                <div>
                  <label htmlFor="docIdentidad" className="block text-sm font-medium text-slate-700 mb-1">1.2 # Doc. Identidad</label>
                  <input type="text" id="docIdentidad" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" required />
                </div>
                <div>
                  <label htmlFor="ejecutivoCartera" className="block text-sm font-medium text-slate-700 mb-1">1.2.1 Ejecutivo Asignado</label>
                  <input type="text" id="ejecutivoCartera" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" placeholder="Nombre del Ejecutivo" />
                </div>
              </div>
            </div>

            <div className="bg-slate-50 p-5 rounded-lg border border-slate-200">
              <h3 className="text-xs font-bold text-slate-500 uppercase tracking-widest mb-4 border-b border-slate-200 pb-2">Detalles del Plan</h3>
              <div className="grid grid-cols-1 md:grid-cols-6 gap-6">
                <div>
                  <label htmlFor="tipoPlan" className="block text-sm font-medium text-slate-700 mb-1">1.3 Tipo de Plan</label>
                  <select id="tipoPlan" onChange={handlePlanTypeChange} className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border bg-white">
                    <option value="Compra Planificada">Compra Planificada</option>
                    <option value="Adjudicación Planificada">Adjudicación Planificada</option>
                  </select>
                </div>
                <div>
                  <label htmlFor="estadoActivo" className="block text-sm font-medium text-slate-700 mb-1">1.3.1 Estado Plan 1</label>
                  <select id="estadoActivo" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border bg-white">
                    <option value="Activo">Activo</option>
                    <option value="Desactivado">Desactivado</option>
                  </select>
                </div>
                <div>
                  <label htmlFor="grupoCodigo" className="block text-sm font-medium text-slate-700 mb-1">1.4 Grupo/Código</label>
                  <input type="text" id="grupoCodigo" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" required />
                </div>
                <div>
                  <label htmlFor="estadoPlan" className="block text-sm font-medium text-slate-700 mb-1">1.14 Estado del Plan</label>
                  <select id="estadoPlan" onChange={handleEstadoPlanChange} className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border bg-white">
                    <option value="No Adjudicado">No Adjudicado</option>
                    <option value="Adjudicado">Adjudicado</option>
                  </select>
                </div>
                <div id="col-forma-adjudicacion" className="hidden">
                  <label htmlFor="formaAdjudicacion" className="block text-sm font-medium text-slate-700 mb-1">1.15 Forma Adj.</label>
                  <input type="text" id="formaAdjudicacion" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border bg-white" placeholder="Ej: Sorteo, Oferta..." />
                </div>
                <div id="col-fecha-adjudicacion" className="hidden">
                  <label htmlFor="fechaAdjudicacion" className="block text-sm font-medium text-slate-700 mb-1">1.16 Fecha Adj.</label>
                  <input type="date" id="fechaAdjudicacion" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" />
                </div>
                <div id="col-num-asamblea" className="hidden">
                  <label htmlFor="numeroAsamblea" className="block text-sm font-medium text-slate-700 mb-1">1.17 # Asamblea</label>
                  <input type="text" id="numeroAsamblea" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" />
                </div>
                <div id="col-fecha-entrega" className="hidden">
                  <label htmlFor="fechaEntrega" className="block text-sm font-bold text-red-600 mb-1">1.18 Fecha Entrega</label>
                  <input type="date" id="fechaEntrega" className="w-full rounded-md border-red-300 shadow-sm focus:border-red-500 focus:ring-red-500 sm:text-sm p-2 border" />
                </div>
              </div>
            </div>

            <div className="bg-slate-50 p-5 rounded-lg border border-slate-200">
              <h3 className="text-xs font-bold text-slate-500 uppercase tracking-widest mb-4 border-b border-slate-200 pb-2">Valores y Financiamiento</h3>
              <div className="grid grid-cols-1 md:grid-cols-4 gap-6 mb-6">
                <div>
                  <label htmlFor="montoContratado" className="block text-sm font-medium text-slate-700 mb-1">1.5 Monto</label>
                  <div className="relative">
                    <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-500">$</span>
                    <input type="number" id="montoContratado" onInput={calculateValues} step="0.01" className="w-full pl-8 rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" required />
                  </div>
                </div>
                <div>
                  <label htmlFor="valorInscripcion" className="block text-sm font-medium text-slate-700 mb-1">1.6 Valor Inscripción</label>
                  <div className="relative">
                    <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-500">$</span>
                    <input type="number" id="valorInscripcion" step="0.01" className="w-full pl-8 rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" />
                  </div>
                </div>
                <div>
                  <label htmlFor="plazoPlan" className="block text-sm font-medium text-slate-700 mb-1">1.7 Plazo (Meses)</label>
                  <input type="number" id="plazoPlan" onInput={calculateValues} className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" required />
                </div>
                <div>
                  <label htmlFor="valorCuota" className="block text-sm font-medium text-slate-700 mb-1">1.8 Cuota Mensual</label>
                  <div className="relative">
                    <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-500">$</span>
                    <input type="number" id="valorCuota" onInput={calculateValues} step="0.01" className="w-full pl-8 rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" required />
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-4 gap-6 mb-6">
                <div id="col-porcentaje" className="opacity-40 transition-opacity">
                  <label htmlFor="porcentajeEntrada" className="block text-sm font-medium text-slate-700 mb-1">1.9 % de Entrada</label>
                  <div className="relative">
                    <input type="number" id="porcentajeEntrada" onInput={calculateValues} step="0.1" className="w-full pr-8 rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" disabled />
                    <span className="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-500">%</span>
                  </div>
                </div>
                <div id="col-valor-entrada" className="opacity-40 transition-opacity">
                  <label htmlFor="valorEntrada" className="block text-sm font-medium text-slate-700 mb-1">1.10 Valor Entrada</label>
                  <div className="relative">
                    <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-500">$</span>
                    <input type="number" id="valorEntrada" className="w-full pl-8 rounded-md border-slate-200 bg-slate-100 shadow-sm sm:text-sm p-2 border text-slate-700 font-semibold" readOnly />
                  </div>
                </div>
                <div id="col-fecha-pago-entrada" className="opacity-40 transition-opacity">
                  <label htmlFor="fechaPagoEntrada" className="block text-sm font-medium text-slate-700 mb-1">1.10.1 Fecha Pago Ent.</label>
                  <input type="date" id="fechaPagoEntrada" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" disabled />
                </div>
                <div className="md:col-span-1 bg-blue-50 p-3 rounded-lg border border-blue-200 flex flex-col items-center justify-center">
                  <label className="block text-[10px] font-bold text-blue-800 uppercase tracking-wide text-center">1.11 Total Plan</label>
                  <div className="text-xl font-black text-blue-900 mt-1" id="display-total-plan">$0.00</div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 pt-4 border-t border-slate-200">
                <div>
                  <label htmlFor="cuotasPagadas" className="block text-sm font-medium text-slate-700 mb-1">1.12 # Cuotas Pagadas</label>
                  <input type="number" id="cuotasPagadas" onInput={calculateValues} min="0" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" required />
                </div>
                <div>
                  <label htmlFor="valorTotalPagado" className="block text-sm font-medium text-slate-700 mb-1">1.13 Valor Pagado</label>
                  <div className="relative">
                    <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-500">$</span>
                    <input type="number" id="valorTotalPagado" className="w-full pl-8 rounded-md border-slate-200 bg-slate-100 shadow-sm sm:text-sm p-2 border text-slate-700 font-semibold" readOnly />
                  </div>
                </div>
                <div>
                  <label htmlFor="fechaPrimerPago" className="block text-sm font-medium text-slate-700 mb-1">1.14 Fecha Primer Pago</label>
                  <input type="date" id="fechaPrimerPago" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" required />
                </div>
              </div>
            </div>
          </form>
        </div>

        {/* 3. ESTADO DE CUENTA */}
        <div id="tab-payment-table" className="tab-content bg-white shadow-lg rounded-xl border border-slate-100 p-0 sm:p-6 print:shadow-none print:border-none print:p-0">
          <div className="mb-6 flex flex-col sm:flex-row justify-between items-start sm:items-end gap-4 p-6 sm:p-0 border-b border-slate-200 pb-4 no-print">
            <div className="w-full sm:w-1/2">
              <label className="block text-sm font-medium text-slate-700 mb-2">Agregar Logo (Se guardará automáticamente)</label>
              <input type="file" id="imageUpload" accept="image/*" className="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 transition-colors" onChange={handleImageUpload} />
            </div>
            <div className="flex gap-2 w-full sm:w-auto">
              <button onClick={saveFromTable} className="w-full sm:w-auto flex justify-center items-center px-4 py-2.5 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors shadow-sm font-medium">Guardar Tabla</button>
              <button onClick={() => window.print()} className="w-full sm:w-auto flex justify-center items-center px-4 py-2.5 bg-emerald-600 text-white rounded-md hover:bg-emerald-700 transition-colors shadow-sm font-medium">Imprimir Reporte</button>
            </div>
          </div>

          <div id="statement-view" className="hidden p-4 sm:p-0">
            <div className="flex flex-col md:flex-row justify-between items-center mb-8 border-b-2 border-slate-800 pb-4">
              <div id="logo-container" className="mb-4 md:mb-0 hidden">
                <img id="statement-logo" src="" alt="Logo" className="max-h-20 object-contain" />
              </div>
              <div className="text-center md:text-right flex-grow">
                <h2 className="text-3xl font-black text-slate-900 uppercase tracking-tight">Estado de Cuenta</h2>
                <p className="text-slate-500 font-medium">Reporte de Aportes Mensuales</p>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6 text-sm">
              <div className="bg-slate-50 p-4 rounded-lg border border-slate-200">
                <h4 className="font-bold text-slate-800 border-b border-slate-200 pb-2 mb-3">Datos del Cliente</h4>
                <div className="grid grid-cols-[130px_1fr] gap-y-2">
                  <span className="text-slate-500 font-medium">Cliente:</span> <span id="lbl-nombre" className="font-bold text-slate-900">-</span>
                  <span className="text-slate-500 font-medium">Identificación:</span> <span id="lbl-doc" className="font-semibold text-slate-700">-</span>
                  <span className="text-slate-500 font-medium">Grupo / Código:</span> <span id="lbl-grupo" className="font-semibold text-slate-700">-</span>
                  <span className="text-slate-500 font-medium">Ejecutivo Asignado:</span> <span id="lbl-ejecutivo" className="font-semibold text-slate-700">-</span>
                </div>
              </div>
              <div className="bg-slate-50 p-4 rounded-lg border border-slate-200">
                <h4 className="font-bold text-slate-800 border-b border-slate-200 pb-2 mb-3">Información del Plan</h4>
                <div className="grid grid-cols-[130px_1fr] gap-y-2">
                  <span className="text-slate-500 font-medium">Tipo de Plan:</span> <span id="lbl-tipo-plan" className="font-semibold text-slate-700">-</span>
                  <span className="text-slate-500 font-medium">Estado:</span> <span id="lbl-estado-plan" className="font-bold">-</span>
                  <span className="text-slate-500 font-medium">Plazo Contrato:</span> <span id="lbl-plazo" className="font-semibold text-slate-700">-</span>
                  <span className="text-slate-500 font-medium lbl-adjudicacion hidden">F. Adjudicación:</span> <span id="lbl-fecha-adj" className="font-semibold text-slate-700 lbl-adjudicacion hidden">-</span>
                  <span className="text-slate-500 font-medium lbl-adjudicacion hidden"># Asamblea:</span> <span id="lbl-num-asamblea" className="font-semibold text-slate-700 lbl-adjudicacion hidden">-</span>
                  <span className="text-slate-500 font-medium lbl-adj-planificada hidden">% de Entrada:</span> <span id="lbl-porcentaje-entrada" className="font-semibold text-slate-700 lbl-adj-planificada hidden">-</span>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-5 gap-0 bg-blue-50 border border-blue-200 rounded-lg mb-8 overflow-hidden">
              <div className="p-3 text-center border-r border-b md:border-b-0 border-blue-200">
                <p className="text-[10px] font-bold text-blue-600 uppercase tracking-wider">Monto Base</p>
                <p className="text-lg font-bold text-slate-800" id="lbl-monto">-</p>
              </div>
              <div className="p-3 text-center border-r border-b md:border-b-0 border-blue-200">
                <p className="text-[10px] font-bold text-blue-600 uppercase tracking-wider">Cuota Mensual</p>
                <p className="text-lg font-bold text-slate-800" id="lbl-cuota">-</p>
              </div>
              <div className="p-3 text-center border-r border-blue-200">
                <p className="text-[10px] font-bold text-blue-600 uppercase tracking-wider">Inscripción</p>
                <p className="text-lg font-bold text-slate-800" id="lbl-inscripcion">-</p>
              </div>
              <div className="p-3 text-center border-r border-blue-200" id="box-entrada">
                <p className="text-[10px] font-bold text-blue-600 uppercase tracking-wider">Valor Entrada</p>
                <p className="text-lg font-bold text-slate-800" id="lbl-entrada">-</p>
              </div>
              <div className="p-3 text-center bg-blue-600 text-white col-span-2 md:col-span-1 flex flex-col justify-center">
                <p className="text-[10px] font-bold text-blue-200 uppercase tracking-wider">Total Plan</p>
                <p className="text-xl font-black" id="lbl-total-plan">-</p>
              </div>
            </div>

            <div className="table-container overflow-x-auto border border-slate-200 rounded-lg">
              <table className="min-w-full divide-y divide-slate-200 text-sm">
                <thead className="bg-slate-800 text-white">
                  <tr>
                    <th className="px-2 py-3 text-center font-semibold text-xs uppercase tracking-wider">#</th>
                    <th className="px-3 py-3 text-right font-semibold text-xs uppercase tracking-wider">Saldo Inicial</th>
                    <th className="px-3 py-3 text-right font-semibold text-xs uppercase tracking-wider">Cuota Mensual</th>
                    <th className="px-3 py-3 text-right font-semibold text-xs uppercase tracking-wider">Abono Mensual</th>
                    <th className="px-3 py-3 text-right font-semibold text-xs uppercase tracking-wider">Saldo Cuota</th>
                    <th className="px-3 py-3 text-right font-semibold text-xs uppercase tracking-wider">Saldo Plan</th>
                    <th className="px-3 py-3 text-center font-semibold text-xs uppercase tracking-wider">Vencimiento</th>
                    <th className="px-3 py-3 text-center font-semibold text-xs uppercase tracking-wider">F. Pago</th>
                    <th className="px-2 py-3 text-center font-semibold text-xs uppercase tracking-wider">Días</th>
                    <th className="px-3 py-3 text-center font-semibold text-xs uppercase tracking-wider">Estado</th>
                  </tr>
                </thead>
                <tbody id="table-body" className="bg-white divide-y divide-slate-200"></tbody>
              </table>
            </div>

            <div className="mt-8 grid grid-cols-1 md:grid-cols-2 gap-8 break-inside-avoid">
              <div className="bg-slate-50 p-6 rounded-lg border border-slate-200">
                <h4 className="font-bold text-slate-800 border-b border-slate-200 pb-2 mb-4 text-lg">Resumen Final</h4>
                <div className="flex justify-between py-2 border-b border-slate-200">
                  <span className="text-slate-600 font-medium">Cuotas Canceladas:</span>
                  <span className="font-bold text-emerald-600 text-lg" id="sum-cuotas-pagadas">0</span>
                </div>
                <div className="flex justify-between py-2 border-b border-slate-200 mb-2">
                  <span className="text-slate-600 font-medium">Total Cancelado:</span>
                  <span className="font-bold text-emerald-600 text-lg" id="sum-valor-pagado">$0.00</span>
                </div>
                <div className="flex justify-between py-2 border-b border-slate-200">
                  <span className="text-slate-600 font-medium">Cuotas Pendientes:</span>
                  <span className="font-bold text-red-500 text-lg" id="sum-cuotas-pendientes">0</span>
                </div>
                <div className="flex justify-between py-2">
                  <span className="text-slate-600 font-medium">Total Pendiente:</span>
                  <span className="font-bold text-red-500 text-lg" id="sum-valor-pendiente">$0.00</span>
                </div>
              </div>
              <div className="flex flex-col justify-end text-sm text-slate-500">
                <p>Generado el: <span id="fecha-generacion" className="font-semibold text-slate-700"></span></p>
              </div>
            </div>
          </div>

          <div id="empty-state" className="text-center py-24 px-4">
            <svg className="mx-auto h-16 w-16 text-slate-300" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            <h3 className="mt-4 text-lg font-medium text-slate-900">Tabla no generada</h3>
            <button type="button" onClick={() => switchTab('client-info')} className="mt-6 inline-flex items-center px-4 py-2 bg-blue-100 text-blue-700 hover:bg-blue-200 rounded-md font-medium">Ir al Formulario</button>
          </div>
        </div>

        {/* 4. MORA Y COBRANZAS */}
        <div id="tab-mora-cobranzas" className="tab-content bg-white shadow-lg rounded-xl border border-slate-100 p-0 sm:p-6 print:shadow-none print:border-none print:p-0">
          <div className="mb-6 flex flex-col sm:flex-row justify-between items-start sm:items-end gap-4 p-6 sm:p-0 border-b border-slate-200 pb-4 no-print">
            <h2 className="text-2xl font-semibold text-slate-800">Cálculo de Mora y Cobranzas</h2>
            <div className="flex flex-col sm:flex-row items-center gap-4 w-full sm:w-auto">
              <div className="flex items-center gap-2 bg-blue-50 p-2 rounded-md border border-blue-200 w-full sm:w-auto">
                <label htmlFor="fecha-calculo-mora" className="text-sm font-bold text-blue-800 whitespace-nowrap">Fecha de Cálculo:</label>
                <input type="date" id="fecha-calculo-mora" onChange={calculateAndRenderMora} className="rounded-md border-blue-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-1.5 border bg-white text-blue-900 font-semibold w-full sm:w-auto" />
              </div>
              <button onClick={() => window.print()} className="w-full sm:w-auto flex justify-center items-center px-4 py-2 bg-emerald-600 text-white rounded-md hover:bg-emerald-700 transition-colors shadow-sm font-medium">
                <svg className="w-5 h-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z" /></svg>
                Imprimir Reporte
              </button>
            </div>
          </div>

          <div className="p-4 sm:p-0">
            <div className="flex flex-col md:flex-row justify-between items-center mb-8 border-b-2 border-slate-800 pb-4">
              <div id="mora-logo-container" className="mb-4 md:mb-0 hidden">
                <img id="mora-statement-logo" src="" alt="Logo" className="max-h-20 object-contain" />
              </div>
              <div className="text-center md:text-right flex-grow">
                <h2 className="text-3xl font-black text-slate-900 uppercase tracking-tight">Mora y Cobranzas</h2>
                <p className="text-slate-500 font-medium">Reporte de Atrasos y Recargos</p>
              </div>
            </div>

            <div className="bg-slate-50 p-4 rounded-lg border border-slate-200 mb-6">
              <h4 className="font-bold text-slate-800 border-b border-slate-200 pb-2 mb-3">Datos del Cliente y Plan</h4>
              <div className="flex flex-wrap justify-between items-center mb-4 gap-4">
                <div><span className="text-slate-500 font-medium">Cliente:</span> <span id="mora-client-name" className="font-bold text-slate-900">-</span></div>
                <div><span className="text-slate-500 font-medium">Plan:</span> <span id="mora-client-plan" className="font-semibold text-slate-700">-</span></div>
                <div><span className="text-slate-500 font-medium">Estado:</span> <span id="mora-client-estado-activo" className="px-2 py-1 rounded-md text-sm font-bold uppercase">-</span></div>
                <div><span id="mora-client-fecha" className="font-semibold text-blue-700 bg-blue-100 px-2 py-1 rounded-md text-sm">-</span></div>
              </div>
              <h5 className="text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Información del Contrato</h5>
              <div className="grid grid-cols-1 md:grid-cols-4 gap-4 text-sm text-slate-800">
                <div className="bg-white p-2 rounded border border-slate-200 shadow-sm"><span className="block text-[10px] uppercase font-bold text-slate-500">Grupo / Código</span><span id="mora-frm-grupo" className="font-semibold">-</span></div>
                <div className="bg-white p-2 rounded border border-slate-200 shadow-sm"><span className="block text-[10px] uppercase font-bold text-slate-500">Monto Contratado</span><span id="mora-frm-monto" className="font-semibold">-</span></div>
                <div className="bg-white p-2 rounded border border-slate-200 shadow-sm"><span className="block text-[10px] uppercase font-bold text-slate-500">Plazo Contrato</span><span id="mora-frm-plazo" className="font-semibold">-</span></div>
                <div className="bg-white p-2 rounded border border-slate-200 shadow-sm"><span className="block text-[10px] uppercase font-bold text-slate-500">Total Cuotas</span><span id="mora-frm-total-cuotas" className="font-semibold">-</span></div>
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
              <div className="lg:col-span-1 space-y-6 no-print">
                <div className="bg-slate-50 p-4 rounded-lg border border-slate-200">
                  <h4 className="font-bold text-slate-800 text-sm border-b border-slate-200 pb-2 mb-3">Tasa Anual del Plan (Interno)</h4>
                  <div className="grid grid-cols-2 gap-4 text-sm text-slate-800">
                    <div className="bg-white p-2 rounded border border-slate-200 shadow-sm"><span className="block text-[10px] uppercase font-bold text-slate-500">Años del Plan</span><span id="mora-frm-anios" className="font-semibold">-</span></div>
                    <div className="bg-white p-2 rounded border border-blue-200 shadow-sm bg-blue-50"><span className="block text-[10px] uppercase font-bold text-blue-700">Tasa Anual</span><span id="mora-frm-tasa-admin" className="font-bold text-blue-800">-</span></div>
                  </div>
                </div>

                <div className="bg-slate-50 p-4 rounded-lg border border-slate-200">
                  <h4 className="font-bold text-slate-800 text-sm border-b border-slate-200 pb-2 mb-3 flex justify-between">
                    Parámetros de MORA
                    <button onClick={() => addParam('mora')} className="text-blue-600 hover:text-blue-800 text-xs">+ Fila</button>
                  </h4>
                  <p className="text-[10px] text-slate-500 mb-2">Aplica en el día 6 de calendario posterior al vencimiento (Día 1 de atraso).</p>
                  <table className="w-full text-xs">
                    <thead className="bg-slate-200 text-slate-600">
                      <tr><th className="p-1">Días Min</th><th className="p-1">Días Max</th><th className="p-1" title="Tasa Anual %">T. Anual %</th><th className="p-1" title="Tasa Diaria %">T. Diaria %</th><th className="p-1"></th></tr>
                    </thead>
                    <tbody id="mora-params-body"></tbody>
                  </table>
                </div>

                <div className="bg-slate-50 p-4 rounded-lg border border-slate-200">
                  <h4 className="font-bold text-slate-800 text-sm border-b border-slate-200 pb-2 mb-3 flex justify-between">
                    Parámetros de COBRANZAS
                    <button onClick={() => addParam('cobranza')} className="text-blue-600 hover:text-blue-800 text-xs">+ Fila</button>
                  </h4>
                  <p className="text-[10px] text-slate-500 mb-2">Aplica en el día 21 de calendario (16 días de atraso). Por Saldo Cuota.</p>
                  <table className="w-full text-xs">
                    <thead className="bg-slate-200 text-slate-600">
                      <tr><th className="p-1">Saldo Min $</th><th className="p-1">Saldo Max $</th><th className="p-1">Valor $</th><th className="p-1"></th></tr>
                    </thead>
                    <tbody id="cobranza-params-body"></tbody>
                  </table>
                </div>
              </div>

              <div className="lg:col-span-2">
                <div className="table-container overflow-x-auto border border-slate-200 rounded-lg mb-6">
                  <table className="min-w-full divide-y divide-slate-200 text-sm">
                    <thead className="bg-slate-800 text-white">
                      <tr>
                        <th className="px-2 py-3 text-center font-semibold text-[10px] uppercase tracking-wider">Cuota</th>
                        <th className="px-2 py-3 text-center font-semibold text-[10px] uppercase tracking-wider">Vence</th>
                        <th className="px-2 py-3 text-center font-semibold text-[10px] uppercase tracking-wider">Días</th>
                        <th className="px-2 py-3 text-right font-semibold text-[10px] uppercase tracking-wider">Saldo</th>
                        <th className="px-2 py-3 text-right font-semibold text-[10px] uppercase tracking-wider text-amber-300">Mora</th>
                        <th className="px-1 py-3 text-center font-semibold text-[10px] uppercase tracking-wider text-emerald-300 no-print" title="Descuento Condonación Mora">% Desc M.</th>
                        <th className="px-2 py-3 text-right font-semibold text-[10px] uppercase tracking-wider text-rose-300">Cobranza</th>
                        <th className="px-1 py-3 text-center font-semibold text-[10px] uppercase tracking-wider text-emerald-300 no-print" title="Descuento Condonación Cobranza">% Desc C.</th>
                        <th className="px-2 py-3 text-right font-semibold text-[10px] uppercase tracking-wider text-blue-300">Total</th>
                      </tr>
                    </thead>
                    <tbody id="mora-results-body" className="bg-white divide-y divide-slate-200"></tbody>
                  </table>
                </div>

                <div className="bg-slate-50 p-6 rounded-lg border border-slate-200 break-inside-avoid">
                  <h4 className="font-bold text-slate-800 border-b border-slate-200 pb-2 mb-4 text-lg">Resumen a Pagar</h4>
                  <div className="flex justify-between py-2 border-b border-slate-200">
                    <span className="text-slate-600 font-medium">Subtotal Cuotas Vencidas:</span>
                    <span className="font-bold text-slate-800 text-lg" id="sum-mora-cuotas">$0.00</span>
                  </div>
                  <div className="flex justify-between py-2 border-b border-slate-200">
                    <span className="text-slate-600 font-medium">Subtotal Mora (Con desc):</span>
                    <span className="font-bold text-amber-600 text-lg" id="sum-mora-recargos">$0.00</span>
                  </div>
                  <div className="flex justify-between py-2 border-b border-slate-200 mb-2">
                    <span className="text-slate-600 font-medium">Subtotal Cobranzas (Con desc):</span>
                    <span className="font-bold text-rose-600 text-lg" id="sum-mora-cobranzas">$0.00</span>
                  </div>
                  <div className="flex justify-between py-2">
                    <span className="text-slate-800 font-black text-xl">TOTAL GENERAL:</span>
                    <span className="font-black text-blue-700 text-2xl" id="sum-mora-total">$0.00</span>
                  </div>
                </div>

                <p className="text-right text-sm text-slate-500 mt-4">Generado el: <span id="mora-fecha-generacion" className="font-semibold text-slate-700"></span></p>

                <div className="mt-8 bg-slate-50 p-6 rounded-lg border border-slate-200 no-print break-inside-avoid">
                  <h4 className="font-bold text-slate-800 border-b border-slate-200 pb-2 mb-4 text-lg">Historial de Gestiones</h4>
                  <div className="mb-4">
                    <textarea id="nueva-gestion-texto" rows={3} className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" placeholder="Ingrese los detalles de la gestión, acuerdos o llamadas realizadas con el cliente..."></textarea>
                    <div className="mt-2 flex justify-end">
                      <button type="button" onClick={guardarGestion} className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors shadow-sm text-sm font-medium">+ Guardar Gestión</button>
                    </div>
                  </div>
                  <div className="space-y-3 max-h-96 overflow-y-auto pr-2" id="historial-gestiones-list"></div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* 5. REPORTES Y PRODUCTIVIDAD */}
        <div id="tab-reportes" className="tab-content bg-white shadow-lg rounded-xl border border-slate-100 p-6 no-print">
          <div className="flex justify-between items-center mb-6 border-b border-slate-200 pb-4">
            <h2 className="text-2xl font-semibold text-slate-800">Reportes y Productividad</h2>
          </div>

          <div className="mb-6 grid grid-cols-1 md:grid-cols-4 gap-4 bg-slate-50 p-4 rounded-lg border border-slate-200">
            <div>
              <label htmlFor="report-search" className="block text-xs font-medium text-slate-700 mb-1">Buscar Cliente / Documento</label>
              <input type="text" id="report-search" onKeyUp={generateReports} placeholder="Escriba para buscar..." className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" />
            </div>
            <div>
              <label htmlFor="report-filter-estado" className="block text-xs font-medium text-slate-700 mb-1">Estado del Plan</label>
              <select id="report-filter-estado" onChange={generateReports} className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border bg-white">
                <option value="Todos">Todos</option>
                <option value="Adjudicado">Adjudicado</option>
                <option value="No Adjudicado">No Adjudicado</option>
              </select>
            </div>
            <div>
              <label htmlFor="report-filter-ejecutivo" className="block text-xs font-medium text-slate-700 mb-1">Ejecutivo</label>
              <select id="report-filter-ejecutivo" onChange={generateReports} className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border bg-white">
                <option value="Todos">Todos</option>
              </select>
            </div>
            <div>
              <label htmlFor="report-filter-vencidas" className="block text-xs font-medium text-slate-700 mb-1">Vencidas (Excel)</label>
              <input type="number" id="report-filter-vencidas" onKeyUp={generateReports} onChange={generateReports} min="0" placeholder="Ej: 3" className="w-full rounded-md border-slate-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm p-2 border" />
            </div>
          </div>

          <div className="mb-8">
            <div className="flex justify-between items-center mb-4">
              <h3 id="report-general-title" className="text-lg font-bold text-slate-700">Reporte General de Clientes</h3>
              <button onClick={() => exportToExcel('general')} className="px-4 py-2 bg-emerald-600 text-white rounded-md hover:bg-emerald-700 transition-colors shadow-sm font-medium flex items-center text-sm">
                <svg className="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
                Descargar Excel
              </button>
            </div>
            <div className="table-container overflow-x-auto border border-slate-200 rounded-lg max-h-[500px]">
              <table className="min-w-full divide-y divide-slate-200 text-sm whitespace-nowrap">
                <thead className="bg-slate-800 text-white text-xs uppercase tracking-wider sticky top-0">
                  <tr>
                    <th className="px-3 py-2 text-left">Cliente</th>
                    <th className="px-3 py-2 text-left">Identificación</th>
                    <th className="px-3 py-2 text-left">Grupo/Plan</th>
                    <th className="px-3 py-2 text-right">Monto</th>
                    <th className="px-3 py-2 text-center">Estado</th>
                    <th className="px-3 py-2 text-right">Cuota Mes</th>
                    <th className="px-3 py-2 text-center font-bold text-red-300">Vencidas (Excel)</th>
                    <th className="px-3 py-2 text-right font-bold text-red-300">Valor Vencido</th>
                    <th className="px-3 py-2 text-center text-emerald-300">Cobradas (Mes)</th>
                    <th className="px-3 py-2 text-right text-emerald-300">Recaudo (Mes)</th>
                    <th className="px-3 py-2 text-center text-amber-300">Pendientes</th>
                    <th className="px-3 py-2 text-right text-amber-300">Valor Pendiente</th>
                    <th className="px-3 py-2 text-left">Ejecutivo</th>
                  </tr>
                </thead>
                <tbody id="report-general-body" className="bg-white divide-y divide-slate-200"></tbody>
              </table>
            </div>
          </div>

          <div>
            <div className="flex justify-between items-center mb-4">
              <h3 id="report-ejecutivo-title" className="text-lg font-bold text-slate-700">Recaudación por Ejecutivo</h3>
              <button onClick={() => exportToExcel('ejecutivos')} className="px-4 py-2 bg-emerald-600 text-white rounded-md hover:bg-emerald-700 transition-colors shadow-sm font-medium flex items-center text-sm">
                <svg className="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
                Descargar Excel
              </button>
            </div>
            <div className="table-container overflow-x-auto border border-slate-200 rounded-lg max-w-3xl">
              <table className="min-w-full divide-y divide-slate-200 text-sm whitespace-nowrap">
                <thead className="bg-blue-900 text-white text-xs uppercase tracking-wider">
                  <tr>
                    <th className="px-4 py-3 text-left">Ejecutivo de Cartera</th>
                    <th className="px-4 py-3 text-center">Total Clientes</th>
                    <th className="px-4 py-3 text-right text-emerald-300">Recaudo (Mes)</th>
                    <th className="px-4 py-3 text-right text-red-300">Total Vencido (Excel)</th>
                  </tr>
                </thead>
                <tbody id="report-ejecutivos-body" className="bg-white divide-y divide-slate-200"></tbody>
              </table>
            </div>
          </div>
        </div>
      </main>

      <div id="toast-container" className="fixed bottom-5 right-5 z-50 flex flex-col gap-2 no-print"></div>

      <div id="confirm-modal" className="fixed inset-0 bg-slate-900 bg-opacity-50 z-50 hidden flex items-center justify-center no-print backdrop-blur-sm transition-opacity">
        <div className="bg-white rounded-xl shadow-2xl p-6 max-w-sm w-full mx-4 transform transition-transform">
          <div className="flex items-center justify-center w-12 h-12 rounded-full bg-amber-100 mb-4 mx-auto">
            <svg className="h-6 w-6 text-amber-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
            </svg>
          </div>
          <h3 className="text-lg font-bold text-center text-slate-900 mb-2">Confirmar Acción</h3>
          <p id="confirm-modal-message" className="text-sm text-slate-500 text-center mb-6">¿Estás seguro?</p>
          <div className="flex justify-center gap-3">
            <button id="confirm-modal-no" className="px-4 py-2 bg-slate-100 text-slate-700 rounded-md hover:bg-slate-200 transition-colors font-medium text-sm">Cancelar</button>
            <button id="confirm-modal-yes" className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors font-medium text-sm">Aceptar</button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;

export interface TableRow {
  numero: number;
  vencimiento: string;
  cuotaMensual: number;
  abono: number;
  fechaPago: string;
  descMora: number;
  descCobranza: number;
}

export interface Gestion {
  fecha: string;
  texto: string;
}

export interface Client {
  id: string;
  nombres: string;
  docIdentidad: string;
  ejecutivoCartera: string;
  tipoPlan: string;
  estadoActivo: string;
  grupoCodigo: string;
  montoContratado: number;
  valorInscripcion: number;
  plazoPlan: number;
  valorCuota: number;
  estadoPlan: string;
  formaAdjudicacion: string;
  fechaAdjudicacion: string;
  numeroAsamblea: string;
  fechaEntrega: string;
  cuotasPagadas: number;
  fechaPrimerPago: string;
  valorEntrada: number;
  porcentajeEntrada: number;
  fechaPagoEntrada: string;
  totalPlan: number;
  tableData: TableRow[];
  gestiones: Gestion[];
  cuotasVencidasExcel: number;
  valorVencidoExcel: number;
}

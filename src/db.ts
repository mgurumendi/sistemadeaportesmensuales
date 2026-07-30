import { 
  collection, 
  doc, 
  getDocs, 
  setDoc, 
  deleteDoc, 
  writeBatch 
} from 'firebase/firestore';
import { db } from './firebase';
import type { Client } from './types';

const TABLE = 'clients';

type DbRow = {
  id: string;
  nombres: string;
  doc_identidad: string;
  ejecutivo_cartera: string;
  tipo_plan: string;
  estado_activo: string;
  grupo_codigo: string;
  monto_contratado: number;
  valor_inscripcion: number;
  plazo_plan: number;
  valor_cuota: number;
  estado_plan: string;
  forma_adjudicacion: string;
  fecha_adjudicacion: string | null;
  numero_asamblea: string;
  fecha_entrega: string | null;
  cuotas_pagadas: number;
  fecha_primer_pago: string | null;
  valor_entrada: number;
  porcentaje_entrada: number;
  fecha_pago_entrada: string | null;
  total_plan: number;
  table_data: Client['tableData'];
  gestiones: Client['gestiones'];
  cuotas_vencidas_excel: number;
  valor_vencido_excel: number;
};

function toDb(c: Client): Record<string, unknown> {
  return {
    id: c.id,
    nombres: c.nombres,
    doc_identidad: c.docIdentidad,
    ejecutivo_cartera: c.ejecutivoCartera,
    tipo_plan: c.tipoPlan,
    estado_activo: c.estadoActivo,
    grupo_codigo: c.grupoCodigo,
    monto_contratado: c.montoContratado,
    valor_inscripcion: c.valorInscripcion,
    plazo_plan: c.plazoPlan,
    valor_cuota: c.valorCuota,
    estado_plan: c.estadoPlan,
    forma_adjudicacion: c.formaAdjudicacion,
    fecha_adjudicacion: c.fechaAdjudicacion || null,
    numero_asamblea: c.numeroAsamblea,
    fecha_entrega: c.fechaEntrega || null,
    cuotas_pagadas: c.cuotasPagadas,
    fecha_primer_pago: c.fechaPrimerPago || null,
    valor_entrada: c.valorEntrada,
    porcentaje_entrada: c.porcentajeEntrada,
    fecha_pago_entrada: c.fechaPagoEntrada || null,
    total_plan: c.totalPlan,
    table_data: c.tableData,
    gestiones: c.gestiones,
    cuotas_vencidas_excel: c.cuotasVencidasExcel,
    valor_vencido_excel: c.valorVencidoExcel,
  };
}

function fromDb(r: DbRow): Client {
  return {
    id: r.id,
    nombres: r.nombres,
    docIdentidad: r.doc_identidad,
    ejecutivoCartera: r.ejecutivo_cartera,
    tipoPlan: r.tipo_plan,
    estadoActivo: r.estado_activo,
    grupoCodigo: r.grupo_codigo,
    montoContratado: Number(r.monto_contratado) || 0,
    valorInscripcion: Number(r.valor_inscripcion) || 0,
    plazoPlan: r.plazo_plan,
    valorCuota: Number(r.valor_cuota) || 0,
    estadoPlan: r.estado_plan,
    formaAdjudicacion: r.forma_adjudicacion,
    fechaAdjudicacion: r.fecha_adjudicacion || '',
    numeroAsamblea: r.numero_asamblea,
    fechaEntrega: r.fecha_entrega || '',
    cuotasPagadas: r.cuotas_pagadas,
    fechaPrimerPago: r.fecha_primer_pago || '',
    valorEntrada: Number(r.valor_entrada) || 0,
    porcentajeEntrada: Number(r.porcentaje_entrada) || 0,
    fechaPagoEntrada: r.fecha_pago_entrada || '',
    totalPlan: Number(r.total_plan) || 0,
    tableData: r.table_data || [],
    gestiones: r.gestiones || [],
    cuotasVencidasExcel: r.cuotas_vencidas_excel || 0,
    valorVencidoExcel: Number(r.valor_vencido_excel) || 0,
  };
}

export async function getClients(): Promise<Client[]> {
  const querySnapshot = await getDocs(collection(db, TABLE));
  const clients: Client[] = [];
  querySnapshot.forEach((docSnap) => {
    // Al recuperar, forzamos el tipo a DbRow para que fromDb lo entienda
    const data = docSnap.data() as DbRow;
    clients.push(fromDb(data));
  });
  return clients;
}

export async function saveClient(client: Client): Promise<void> {
  const clientData = toDb(client);
  const docRef = doc(db, TABLE, client.id);
  // setDoc hace la función de 'upsert': crea o actualiza
  await setDoc(docRef, clientData);
}

export async function saveClients(clients: Client[]): Promise<void> {
  if (clients.length === 0) return;

  // Firebase permite hasta 500 operaciones por lote (batch).
  // Procesamos el arreglo en fragmentos para evitar errores al subir Excel grandes.
  const CHUNK_SIZE = 500;
  for (let i = 0; i < clients.length; i += CHUNK_SIZE) {
    const chunk = clients.slice(i, i + CHUNK_SIZE);
    const batch = writeBatch(db);

    chunk.forEach((client) => {
      const docRef = doc(db, TABLE, client.id);
      batch.set(docRef, toDb(client));
    });

    await batch.commit();
  }
}

export async function deleteClientDb(id: string): Promise<void> {
  const docRef = doc(db, TABLE, id);
  await deleteDoc(docRef);
}

export async function clearAllClients(): Promise<void> {
  const querySnapshot = await getDocs(collection(db, TABLE));
  if (querySnapshot.empty) return;

  const CHUNK_SIZE = 500;
  const docs = querySnapshot.docs;

  for (let i = 0; i < docs.length; i += CHUNK_SIZE) {
    const chunk = docs.slice(i, i + CHUNK_SIZE);
    const batch = writeBatch(db);

    chunk.forEach((docSnap) => {
      batch.delete(docSnap.ref);
    });

    await batch.commit();
  }
}
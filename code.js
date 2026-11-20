import { createClient } from 'https://esm.sh';

// ใส่ค่าจาก Supabase > Project Settings > API
const SUPABASE_URL = 'https://ydigapgjnalaqfdjrhgd.supabase.co/rest/v1/';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlkaWdhcGdqbmFsYXFmZGpyaGdkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTUxMzUzODgsImV4cCI6MjExMDcxMTM4OH0.VvT03A_dql4QTmgiwrCtZi-DKkJgXDFJIrG4T2bPvJ8';

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const FUEL_CAT = 'ซื้อน้ำมัน';
const r2_ = n => Math.round(n * 100) / 100;

/* ---------- 1. ระบบบัญชีรายรับ-รายจ่าย ---------- */
async function readTx_() {
  const { data, error } = await supabase
    .from('Transactions')
    .select('*')
    .order('date', { ascending: true })
    .order('createdAt', { ascending: true });
  if (error) throw new Error(error.message);

  return (data || []).map(o => ({
    ...o,
    id: String(o.id),
    date: String(o.date || '').slice(0, 10),
    qty: Number(o.qty) || 0,
    amount: Number(o.amount) || 0,
    costPerUnit: Number(o.costPerUnit) || 0,
    sellPrice: Number(o.sellPrice) || 0,
    profitPerUnit: Number(o.profitPerUnit) || 0,
    createdAt: o.createdAt || ''
  }));
}

async function readItems_() {
  const { data, error } = await supabase.from('Items').select('*');
  if (error) throw new Error(error.message);

  return (data || []).map(r => ({
    name: String(r.name),
    category: String(r.category || ''),
    sellPrice: Number(r.sellPrice) || 0
  }));
}

export async function getInitData() {
  const [transactions, items] = await Promise.all([readTx_(), readItems_()]);
  return { transactions, items };
}

export async function saveTransaction(t) {
  if (!t || !/^\d{4}-\d{2}-\d{2}$/.test(t.date)) throw new Error('กรุณาระบุวันที่');
  if (t.type !== 'income' && t.type !== 'expense') throw new Error('ประเภทรายการไม่ถูกต้อง');
  const amount = Number(t.amount) || 0;
  if (amount <= 0) throw new Error('กรุณาระบุจำนวนเงิน');

  const qty = Number(t.qty) || 0;
  const sell = Number(t.sellPrice) || 0;
  let cost = 0, profit = 0;

  if (t.type === 'expense' && t.category === FUEL_CAT && qty > 0) {
    cost = r2_(amount / qty);
    if (sell > 0) profit = r2_(sell - cost);
  }

  const rowData = {
    date: t.date,
    type: t.type,
    category: String(t.category || ''),
    item: String(t.item || '').trim(),
    qty,
    amount,
    costPerUnit: cost,
    sellPrice: sell,
    profitPerUnit: profit,
    note: String(t.note || '').trim()
  };

  if (t.id) {
    const { error } = await supabase.from('Transactions').update(rowData).eq('id', String(t.id));
    if (error) throw new Error(error.message);
  } else {
    rowData.id = crypto.randomUUID();
    rowData.createdAt = new Date().toISOString();
    const { error } = await supabase.from('Transactions').insert([rowData]);
    if (error) throw new Error(error.message);
  }
  return await readTx_();
}

export async function deleteTransaction(id) {
  const { error } = await supabase.from('Transactions').delete().eq('id', String(id));
  if (error) throw new Error(error.message);
  return await readTx_();
}

export async function saveItems(list) {
  const clean = (list || [])
    .filter(x => x && String(x.name).trim() !== '')
    .map(x => ({
      name: String(x.name).trim(),
      category: String(x.category || ''),
      sellPrice: Number(x.sellPrice) || 0
    }));

  const { error: delError } = await supabase.from('Items').delete().neq('name', '');
  if (delError) throw new Error(delError.message);

  if (clean.length) {
    const { error: insError } = await supabase.from('Items').insert(clean);
    if (insError) throw new Error(insError.message);
  }
  return await readItems_();
}

/* ---------- 2. ข้อมูลเครื่องลูก (QR_Daily / QR_Coin) ---------- */
async function readCoin_() {
  const { data, error } = await supabase
    .from('QR_Coin')
    .select('*')
    .order('txDate', { ascending: false });
  if (error) throw new Error(error.message);
  return (data || []).map(r => ({
    machine: String(r.machine),
    channel: String(r.channel),
    amount: Number(r.amount) || 0,
    txDate: String(r.txDate)
  }));
}

async function readDaily_() {
  const { data, error } = await supabase
    .from('QR_Daily')
    .select('*')
    .order('date', { ascending: false });
  if (error) throw new Error(error.message);
  return (data || []).map(r => ({
    machine: String(r.machine),
    date: String(r.date),
    transfer: Number(r.transfer) || 0,
    send: Number(r.send) || 0,
    transferSend: Number(r.transferSend) || 0,
    ch1: Number(r.ch1) || 0,
    ch2: Number(r.ch2) || 0,
    cash: Number(r.cash) || 0
  }));
}

async function qrPayload_(addedCoin = 0) {
  const [coins, daily, metaRes] = await Promise.all([
    readCoin_(),
    readDaily_(),
    supabase.from('app_meta').select('*').in('key', ['QR_META', 'QR_CHILD_META'])
  ]);
  const metaMap = {};
  (metaRes.data || []).forEach(m => { metaMap[m.key] = m.value; });

  return {
    meta: metaMap['QR_META'] || {},
    coins,
    daily,
    child: metaMap['QR_CHILD_META'] || {},
    added: 0,
    addedCoin
  };
}

export async function getQrData() {
  return await qrPayload_(0);
}

export async function syncQr() {
  const { data, error } = await supabase.functions.invoke('sync-qr');
  if (error) throw new Error(error.message || 'ซิงก์ข้อมูลไม่สำเร็จ');
  if (data?.error) throw new Error(data.error);
  return await qrPayload_(data?.added || 0);
}

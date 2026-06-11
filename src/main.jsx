import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';
import { Upload, LayoutDashboard, FileText, Settings, Users, LogOut, Trash2, Plus, Download } from 'lucide-react';
import './styles.css';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

const api = async (path, opts = {}) => {
  const res = await fetch('/api' + path, {
    ...opts,
    headers: { 'content-type': 'application/json', ...(opts.headers || {}) },
    credentials: 'include'
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || 'Request failed');
  return data;
};

const money = (n) => Number(n || 0).toLocaleString(undefined, { style: 'currency', currency: 'USD' });
const cleanNum = (v) => {
  if (typeof v === 'number') return v;
  if (v == null) return 0;
  let s = String(v).trim().replace(/[$,]/g, '');
  if (/^\(.+\)$/.test(s)) s = '-' + s.slice(1, -1);
  if (s.startsWith('- ')) s = '-' + s.slice(2);
  return Number(s) || 0;
};
const titleCase = (s) => String(s || '').toLowerCase().replace(/\b[a-z]/g, c => c.toUpperCase()).replace(/\bNyc\b/g, 'NYC').replace(/\bNyct\b/g, 'NYCT').replace(/\bUsa\b/g, 'USA');

function applyClean(desc, rules = []) {
  const original = String(desc || '').trim().replace(/\s+/g, ' ');
  const upper = original.toUpperCase();
  const rule = rules.find(r => upper.includes(String(r.match_text).toUpperCase()));
  if (rule) return rule.clean_name;
  let s = original
    .replace(/\b[A-Z]{2}\s*$/i, '')
    .replace(/\b(NEW YORK|BROOKLYN|ASTORIA|MANHATTAN|WHITE PLAINS|PORT CHESTER|SCARSDALE|MUMBAI|DELHI|NY|NJ|CA|NC|WA|FL|TX)\b\s*$/i, '')
    .replace(/\b\d{3}-\d{3}-\d{4}\b/g, '')
    .replace(/\bHELP\.[A-Z0-9.*-]+\b/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
  return titleCase(s || original);
}

function normalizeDate(value) {
  if (!value) return '';
  if (typeof value === 'number') {
    const d = XLSX.SSF.parse_date_code(value);
    if (!d) return String(value);
    return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  const s = String(value).trim();
  const parsed = new Date(s);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0,10);
  return s;
}

async function parseAmex(file, rules) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  return rows.slice(7).filter(r => r && (r[0] || r[1] || r[4])).map((r, idx) => {
    const original = String(r[1] || '').trim();
    const person = String(r[2] || '').trim();
    const amt = cleanNum(r[4]);
    return {
      id: crypto.randomUUID(),
      date: normalizeDate(r[0]),
      person,
      merchant: applyClean(original, rules),
      original,
      amount: amt,
      originalAmount: amt,
      lineItems: [{ person, amount: amt }],
      confirmed: false,
      idx
    };
  });
}

async function pdfLines(file, startPageIndex = 2) {
  const data = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const lines = [];
  for (let pageNo = startPageIndex + 1; pageNo <= pdf.numPages; pageNo++) {
    const page = await pdf.getPage(pageNo);
    const content = await page.getTextContent();
    const buckets = [];
    for (const item of content.items) {
      const x = item.transform[4];
      const y = item.transform[5];
      let b = buckets.find(v => Math.abs(v.y - y) < 3);
      if (!b) { b = { y, items: [] }; buckets.push(b); }
      b.items.push({ x, text: item.str });
    }
    buckets.sort((a,b) => b.y - a.y);
    for (const b of buckets) {
      b.items.sort((a,c) => a.x - c.x);
      const text = b.items.map(i => i.text).join(' ').replace(/\s+/g, ' ').trim();
      if (text) lines.push({ page: pageNo, text, items: b.items });
    }
  }
  return lines;
}

async function parseChase(file, rules) {
  const lines = await pdfLines(file, 2);
  const out = [];
  let inPurchase = false;
  let current = null;
  const amountRe = /(-?\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})|-?\$?\d+\.\d{2})$/;
  for (const line of lines) {
    const t = line.text;
    if (/^PURCHASE$/i.test(t)) { inPurchase = true; continue; }
    if (/^(FEES|INTEREST CHARGED|TOTAL)/i.test(t)) inPurchase = false;
    if (!inPurchase) continue;
    const m = t.match(/^(\d{2}\/\d{2})\s+(.+?)\s+(-?\$?\d[\d,]*\.\d{2})$/);
    if (m) {
      if (current) out.push(current);
      const original = m[2].trim();
      const amt = cleanNum(m[3]);
      current = { id: crypto.randomUUID(), date: m[1], person: '', merchant: applyClean(original, rules), original, amount: amt, originalAmount: amt, lineItems: [], confirmed: false };
    } else if (current && !amountRe.test(t) && !/^(Date of|Merchant Name|\$ Amount|ACCOUNT ACTIVITY)/i.test(t)) {
      current.original = (current.original + ' ' + t).trim();
    }
  }
  if (current) out.push(current);
  return out.map(r => ({ ...r, merchant: applyClean(r.original, rules), lineItems: [{ person: '', amount: r.amount }] }));
}

async function parseCapitalOne(file, rules) {
  const lines = await pdfLines(file, 2);
  const out = [];
  let person = '';
  const monthDay = '(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\\s+\\d{1,2}';
  const re = new RegExp(`^(${monthDay})\\s+(${monthDay})\\s+(.+?)\\s+(-?\\$?[\\d,]+\\.\\d{2})$`, 'i');
  for (const line of lines) {
    const t = line.text;
    const header = t.match(/^(.+?)\s+#\d+:\s+(Payments, Credits and Adjustments|Transactions)/i);
    if (header) { person = titleCase(header[1].trim()); continue; }
    if (!person || /^(Trans Date|SAUNAK|SHANAYA|Total Transactions|Transactions|Visit capitalone)/i.test(t)) continue;
    const m = t.match(re);
    if (!m) continue;
    const original = m[5].trim();
    if (/CAPITAL ONE ONLINE PYMT/i.test(original)) continue;
    const amt = cleanNum(m[6]);
    out.push({ id: crypto.randomUUID(), date: m[1], person, merchant: applyClean(original, rules), original, amount: amt, originalAmount: amt, lineItems: [{ person, amount: amt }], confirmed: false });
  }
  return out;
}

function Login({ onLogin }) {
  const [needsSetup, setNeedsSetup] = useState(false);
  const [form, setForm] = useState({ name: '', email: '', password: '' });
  const [err, setErr] = useState('');
  useEffect(() => { api('/setup-status').then(d => setNeedsSetup(d.needsSetup)).catch(() => {}); }, []);
  async function submit(e) {
    e.preventDefault(); setErr('');
    try {
      const path = needsSetup ? '/setup' : '/login';
      const data = await api(path, { method: 'POST', body: JSON.stringify(form) });
      onLogin(data.user);
    } catch (e) { setErr(e.message); }
  }
  return <div className="auth-wrap"><div className="auth-card">
    <h1>Statement Splitter</h1><p>{needsSetup ? 'Create your first admin login.' : 'Log in to your private dashboard.'}</p>
    <form onSubmit={submit}>
      {needsSetup && <input placeholder="Name" value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>} 
      <input placeholder="Email" value={form.email} onChange={e=>setForm({...form,email:e.target.value})}/>
      <input placeholder="Password" type="password" value={form.password} onChange={e=>setForm({...form,password:e.target.value})}/>
      {err && <div className="error">{err}</div>}
      <button className="primary">{needsSetup ? 'Create admin' : 'Log in'}</button>
    </form>
  </div></div>;
}

function UploadPage({ refreshKey }) {
  const [rules, setRules] = useState([]);
  const [issuer, setIssuer] = useState('amex');
  const [title, setTitle] = useState('');
  const [rows, setRows] = useState([]);
  const [people, setPeople] = useState([]);
  const [newPerson, setNewPerson] = useState('');
  const [splitId, setSplitId] = useState(null);
  const [msg, setMsg] = useState('');
  const [busy, setBusy] = useState(false);
  useEffect(() => { loadBasics(); }, [refreshKey]);
  async function loadBasics(){ const [r,p] = await Promise.all([api('/merchant-rules'), api('/people')]); setRules(r.rules); setPeople(p.people.map(x=>x.name)); }
  const allPeople = useMemo(() => Array.from(new Set([...people, ...rows.map(r=>r.person).filter(Boolean), ...rows.flatMap(r=>(r.lineItems||[]).map(i=>i.person).filter(Boolean))])).sort(), [people, rows]);
  function addPerson() { if (!newPerson.trim()) return; setPeople(Array.from(new Set([...people, titleCase(newPerson.trim())])).sort()); setNewPerson(''); }
  async function parseFile(file) {
    setBusy(true); setMsg('Parsing statement...');
    try {
      let parsed = [];
      if (issuer === 'amex') parsed = await parseAmex(file, rules);
      if (issuer === 'chase') parsed = await parseChase(file, rules);
      if (issuer === 'capitalone') parsed = await parseCapitalOne(file, rules);
      setRows(parsed);
      setPeople(Array.from(new Set([...people, ...parsed.map(r=>r.person).filter(Boolean)])));
      setMsg(`Pulled ${parsed.length} transactions. Review, edit, split, then save.`);
    } catch (e) { setMsg('Could not parse file: ' + e.message); }
    setBusy(false);
  }
  function updateRow(id, patch) {
    setRows(rows.map(r => {
      if (r.id !== id) return r;
      const next = { ...r, ...patch };
      const hasExplicitLineItems = Object.prototype.hasOwnProperty.call(patch, 'lineItems');
      if (!hasExplicitLineItems && patch.person !== undefined && (!r.lineItems || r.lineItems.length <= 1)) next.lineItems = [{ person: patch.person, amount: next.amount }];
      if (!hasExplicitLineItems && patch.amount !== undefined && (!r.lineItems || r.lineItems.length <= 1)) {
        next.originalAmount = Number(patch.amount);
        next.lineItems = [{ person: next.person, amount: Number(patch.amount) }];
      }
      return next;
    }));
  }
  async function save() {
    if (!title.trim()) return setMsg('Give the statement a title first.');
    const missing = rows.filter(r => (r.lineItems || []).some(i => !i.person));
    if (missing.length) return setMsg('Some rows still have no person assigned. Assign or delete them before saving.');
    setBusy(true);
    try {
      const res = await api('/statements', { method: 'POST', body: JSON.stringify({ issuer, title, rows }) });
      setMsg(`Saved statement #${res.id}.`); setRows([]); setTitle('');
    } catch(e) { setMsg(e.message); }
    setBusy(false);
  }
  return <div className="page"><div className="page-head"><h2>Upload Statement</h2><p>Upload one statement at a time, then review every parsed line before saving.</p></div>
    <div className="panel grid-4">
      <label>Issuer<select value={issuer} onChange={e=>setIssuer(e.target.value)}><option value="amex">American Express Excel</option><option value="chase">Chase PDF</option><option value="capitalone">Capital One PDF</option></select></label>
      <label>Statement title<input value={title} onChange={e=>setTitle(e.target.value)} placeholder="Capital One May 2026"/></label>
      <label>Upload file<input type="file" accept=".xlsx,.xls,.csv,.pdf" onChange={e=>e.target.files?.[0] && parseFile(e.target.files[0])}/></label>
      <div className="people-box"><div className="muted">People for this statement</div><div className="inline"><input value={newPerson} onChange={e=>setNewPerson(e.target.value)} placeholder="Add name"/><button onClick={addPerson}><Plus size={16}/></button></div></div>
    </div>
    {msg && <div className="notice">{msg}</div>}
    {rows.length > 0 && <div className="panel">
      <div className="toolbar"><strong>{rows.length} parsed rows</strong><button disabled={busy} className="primary" onClick={save}>Save Statement</button></div>
      <div className="table-wrap"><table><thead><tr><th>Date</th><th>Person</th><th>Merchant / Original</th><th>Amount</th><th>Split</th><th></th></tr></thead><tbody>
        {rows.map(r => <React.Fragment key={r.id}><tr>
          <td><input className="small" value={r.date} onChange={e=>updateRow(r.id,{date:e.target.value})}/></td>
          <td>{r.lineItems?.length > 1 ? <div className="split-lines">{r.lineItems.map((item,idx)=><div key={`${item.person}-${idx}`}><span>{item.person}</span><strong>{money(item.amount)}</strong></div>)}</div> : <select value={r.person || ''} onChange={e=>updateRow(r.id,{person:e.target.value})}><option value="">Select</option>{allPeople.map(p=><option key={p}>{p}</option>)}</select>}</td>
          <td><input value={r.merchant} onChange={e=>updateRow(r.id,{merchant:e.target.value})}/><div className="original">{r.original}</div></td>
          <td>{r.lineItems?.length > 1 ? <strong>{money(r.originalAmount ?? r.amount)}</strong> : <input className="money-input" type="number" step="0.01" value={r.amount} onChange={e=>updateRow(r.id,{amount:Number(e.target.value)})}/>}</td>
          <td>{r.lineItems?.length > 1 ? <span className="pill">{r.lineItems.length} lines</span> : <span className="muted">Single</span>}</td>
          <td className="actions"><button onClick={()=>setSplitId(splitId===r.id?null:r.id)}>{splitId===r.id?'Close':'Split'}</button><button className="danger" onClick={()=>setRows(rows.filter(x=>x.id!==r.id))}><Trash2 size={15}/></button></td>
        </tr>{splitId===r.id && <tr><td colSpan="6"><SplitEditor row={r} people={allPeople} onApply={(lineItems)=>{updateRow(r.id,{lineItems, person: lineItems.length === 1 ? lineItems[0].person : ''}); setSplitId(null);}} /></td></tr>}</React.Fragment>)}
      </tbody></table></div></div>}
  </div>;
}

function SplitEditor({ row, people, onApply }) {
  const [selected, setSelected] = useState(row.lineItems?.length > 1 ? row.lineItems.map(i=>i.person) : [row.person].filter(Boolean));
  const [custom, setCustom] = useState(row.lineItems?.length ? row.lineItems : []);
  const [mode, setMode] = useState('even');
  const total = Number(row.originalAmount ?? row.amount ?? 0);
  const alloc = (mode === 'even' ? evenItems() : custom).reduce((s,i)=>s+Number(i.amount||0),0);
  function toggle(p){ setSelected(selected.includes(p) ? selected.filter(x=>x!==p) : [...selected,p]); }
  function evenItems(){
    const names = selected.filter(Boolean); if (!names.length) return [];
    const cents = Math.round(total * 100); const base = Math.trunc(cents / names.length); const remainder = cents - base * names.length;
    return names.map((p,i) => ({ person:p, amount: (base + (i < Math.abs(remainder) ? Math.sign(remainder) : 0)) / 100 }));
  }
  const itemsToApply = mode === 'even' ? evenItems() : custom.filter(i=>i.person);
  const hasDuplicatePeople = new Set(itemsToApply.map(i=>i.person.toLowerCase())).size !== itemsToApply.length;
  const canApply = itemsToApply.length > 0 && !hasDuplicatePeople && Math.abs(total-alloc) < 0.005;
  function apply(){ if (canApply) onApply(itemsToApply); }
  return <div className="split-editor">
    <div className="split-top"><strong>Split: {row.merchant}</strong><span>Charge total: {money(total)}</span><span>Allocated: {money(alloc)}</span><span className={Math.abs(total-alloc)<0.005?'ok':'bad'}>Difference: {money(total-alloc)}</span></div>
    <div className="tabs"><button className={mode==='even'?'active':''} onClick={()=>setMode('even')}>Even split</button><button className={mode==='custom'?'active':''} onClick={()=>{setMode('custom'); if (!custom.length) setCustom(evenItems());}}>Custom amounts</button></div>
    {mode === 'even' && <div><div className="chips">{people.map(p=><button key={p} className={selected.includes(p)?'chip selected':'chip'} onClick={()=>toggle(p)}>{p}</button>)}</div><div className="mini-table">{evenItems().map(i=><div key={i.person}><span>{i.person}</span><strong>{money(i.amount)}</strong></div>)}</div></div>}
    {mode === 'custom' && <div className="mini-table">{custom.map((i,idx)=><div key={idx}><select value={i.person} onChange={e=>setCustom(custom.map((x,j)=>j===idx?{...x,person:e.target.value}:x))}><option value="">Person</option>{people.map(p=><option key={p}>{p}</option>)}</select><input type="number" step="0.01" value={i.amount} onChange={e=>setCustom(custom.map((x,j)=>j===idx?{...x,amount:Number(e.target.value)}:x))}/><button onClick={()=>setCustom(custom.filter((_,j)=>j!==idx))}>Remove</button></div>)}<button onClick={()=>setCustom([...custom,{person:'',amount:0}])}>Add line</button></div>}
    {hasDuplicatePeople && <div className="bad split-warning">Each person can only appear once in a split.</div>}
    <div className="toolbar right"><button className="primary" disabled={!canApply} onClick={apply}>Confirm split into line items</button></div>
  </div>
}

function Dashboard() {
  const [data, setData] = useState({ summary: [], detail: [], payments: [] });
  const [openPerson, setOpenPerson] = useState(null);
  const [form, setForm] = useState({ person:'', amount:'', payment_date:new Date().toISOString().slice(0,10), type:'Paid', method:'Venmo', notes:'' });
  const [resetConfirm, setResetConfirm] = useState('');
  const [msg, setMsg] = useState('');
  async function load(){ setData(await api('/dashboard')); }
  useEffect(()=>{ load(); }, []);
  const visiblePeople = data.summary.filter(p=>!p.hidden);
  const hiddenPeople = data.summary.filter(p=>p.hidden);
  const hiddenWithBalance = hiddenPeople.filter(p=>Math.abs(p.open) >= 0.005);
  async function addPayment(e){ e.preventDefault(); await api('/payments',{method:'POST',body:JSON.stringify(form)}); setForm({...form, amount:'', notes:''}); load(); }
  async function setHidden(person, hidden){ await api('/people/'+person.id,{method:'PATCH',body:JSON.stringify({hidden})}); if (hidden && openPerson===person.id) setOpenPerson(null); load(); }
  async function resetData(){
    try { await api('/reset-data',{method:'POST',body:JSON.stringify({confirm:resetConfirm})}); setResetConfirm(''); setOpenPerson(null); setMsg('All saved statements, charges, people, payments, and merchant rules were deleted.'); load(); }
    catch(e) { setMsg(e.message); }
  }
  function exportCSV(){
    const lines = [['Person','Assigned','Paid','Splitwise','Adjustment','Open'], ...data.summary.map(r=>[r.name,r.assigned,r.paid,r.splitwise,r.adjustment,r.open])];
    downloadCSV('dashboard-summary.csv', lines);
  }
  return <div className="page"><div className="page-head"><h2>Dashboard</h2><p>Click a person to open their profile, review every charge, or add a one-off charge.</p></div>
    {msg && <div className="notice">{msg}</div>}
    {hiddenWithBalance.length > 0 && <div className="notice hidden-alert"><strong>Hidden people have an open balance:</strong>{hiddenWithBalance.map(p=><span key={p.id}>{p.name} ({money(p.open)}) <button onClick={()=>setHidden(p,false)}>Unhide</button></span>)}</div>}
    <div className="stats">{visiblePeople.map(p=><div className={openPerson===p.id?'stat active':'stat'} key={p.id} onClick={()=>setOpenPerson(openPerson===p.id?null:p.id)}><div className="stat-head"><strong>{p.name}</strong><button onClick={e=>{e.stopPropagation();setHidden(p,true);}}>Hide</button></div><b>{money(p.open)}</b><span>Assigned {money(p.assigned)} · Paid {money(p.paid)} · Splitwise {money(p.splitwise)}</span><em>{openPerson===p.id?'Close profile':'View profile and charges'}</em></div>)}</div>
    {!visiblePeople.length && <div className="panel muted">No visible people yet. Import a statement or unhide someone below.</div>}
    {data.summary.map(p => openPerson===p.id && <PersonDetail key={p.id} person={p} data={data} reload={load}/>) }
    <div className="panel"><div className="toolbar"><h3>Record payment / Splitwise transfer</h3><button onClick={exportCSV}><Download size={16}/> Export summary</button></div>
      <form className="payment-form" onSubmit={addPayment}>
        <select required value={form.person} onChange={e=>setForm({...form,person:e.target.value})}><option value="">Person</option>{data.summary.map(p=><option key={p.id}>{p.name}</option>)}</select>
        <input required type="number" step="0.01" placeholder="Amount" value={form.amount} onChange={e=>setForm({...form,amount:e.target.value})}/>
        <input required type="date" value={form.payment_date} onChange={e=>setForm({...form,payment_date:e.target.value})}/>
        <select value={form.type} onChange={e=>setForm({...form,type:e.target.value})}><option>Paid</option><option>Moved to Splitwise</option><option>Adjustment</option></select>
        <select value={form.method} onChange={e=>setForm({...form,method:e.target.value})}><option>Venmo</option><option>Zelle</option><option>Cash</option><option>Check</option><option>Splitwise</option><option>Other</option></select>
        <input placeholder="Notes" value={form.notes} onChange={e=>setForm({...form,notes:e.target.value})}/>
        <button className="primary">Add</button>
      </form>
    </div>
    {hiddenPeople.length > 0 && <div className="panel"><h3>Hidden people</h3><div className="hidden-list">{hiddenPeople.map(p=><div key={p.id}><span>{p.name} · Open {money(p.open)}</span><button onClick={()=>setHidden(p,false)}>Unhide</button></div>)}</div></div>}
    <div className="panel danger-zone"><h3>Start from scratch</h3><p>Delete all of your saved statements, charges, people, payments, and merchant rules. Your login will remain active. This cannot be undone.</p><div className="reset-row"><input value={resetConfirm} onChange={e=>setResetConfirm(e.target.value)} placeholder="Type RESET to confirm"/><button className="danger" disabled={resetConfirm!=='RESET'} onClick={resetData}>Wipe all saved data</button></div></div>
  </div>;
}

function PersonDetail({ person, data, reload }) {
  const rows = data.detail.filter(d=>d.person_id===person.id);
  const byStatement = rows.reduce((acc,r)=>{ (acc[r.statement_id] ||= { title:r.statement_title, issuer:r.issuer, rows:[], total:0 }); acc[r.statement_id].rows.push(r); acc[r.statement_id].total += Number(r.line_amount); return acc; },{});
  const payments = data.payments.filter(p=>p.person_id===person.id);
  const [charge, setCharge] = useState({ description:'', amount:'', charge_date:new Date().toISOString().slice(0,10), notes:'' });
  const [msg, setMsg] = useState('');
  async function addCharge(e){ e.preventDefault(); try { await api('/manual-charges',{method:'POST',body:JSON.stringify({...charge,person:person.name})}); setCharge({...charge,description:'',amount:'',notes:''}); setMsg('One-off charge added.'); reload(); } catch(e) { setMsg(e.message); } }
  return <div className="panel detail"><div className="profile-head"><div><h3>{person.name}</h3><p>Open balance: <strong>{money(person.open)}</strong></p></div><div><span>Assigned {money(person.assigned)}</span><span>Paid {money(person.paid)}</span><span>Splitwise {money(person.splitwise)}</span></div></div>
    <h4>Charges by statement</h4>
    {Object.entries(byStatement).length ? Object.entries(byStatement).map(([sid, group]) => <div className="accordion" key={sid}>
      <div className="acc-head static"><strong>{group.title}</strong><span>{group.issuer}</span><b>{money(group.total)}</b></div>
      <div className="acc-body"><table><thead><tr><th>Date</th><th>Charge</th><th>Amount owed</th><th>Full charge</th></tr></thead><tbody>{group.rows.map((r,i)=><tr key={i}><td>{r.transaction_date}</td><td>{r.merchant}<div className="original">{r.original_description}</div></td><td><strong>{money(r.line_amount)}</strong></td><td>{money(r.original_amount)}</td></tr>)}</tbody></table></div>
    </div>) : <p className="muted">No charges assigned to this person.</p>}
    <div className="manual-charge"><h4>Add one-off charge</h4><p className="muted">Use this for cash purchases or anything that did not come from an imported statement.</p>{msg&&<div className="notice">{msg}</div>}<form onSubmit={addCharge}><input required placeholder="Description" value={charge.description} onChange={e=>setCharge({...charge,description:e.target.value})}/><input required type="number" step="0.01" placeholder="Amount" value={charge.amount} onChange={e=>setCharge({...charge,amount:e.target.value})}/><input required type="date" value={charge.charge_date} onChange={e=>setCharge({...charge,charge_date:e.target.value})}/><input placeholder="Notes (optional)" value={charge.notes} onChange={e=>setCharge({...charge,notes:e.target.value})}/><button className="primary">Add charge</button></form></div>
    <h4>Payments / transfers</h4>{payments.length ? <table><thead><tr><th>Date</th><th>Type</th><th>Method</th><th>Amount</th><th>Notes</th><th></th></tr></thead><tbody>{payments.map(p=><tr key={p.id}><td>{p.payment_date}</td><td>{p.type}</td><td>{p.method}</td><td>{money(p.amount)}</td><td>{p.notes}</td><td><button className="danger" onClick={async()=>{await api('/payments/'+p.id,{method:'DELETE'}); reload();}}>Delete</button></td></tr>)}</tbody></table> : <p className="muted">No payments recorded.</p>}
  </div>
}

function Statements() {
  const [items, setItems] = useState([]); const [selected, setSelected] = useState(null);
  async function load(){ setItems((await api('/statements')).statements); }
  useEffect(()=>{ load(); }, []);
  async function del(id){ if(confirm('Delete this statement?')) { await api('/statements/'+id,{method:'DELETE'}); setSelected(null); load(); } }
  return <div className="page"><div className="page-head"><h2>Statements</h2><p>Previously imported statements and their parsed line items.</p></div>
    <div className="cards">{items.map(s=><div className="statement-card" key={s.id} onClick={async()=>setSelected(await api('/statements/'+s.id))}><strong>{s.title}</strong><span>{s.issuer}</span><b>{money(s.total)}</b><button className="danger" onClick={(e)=>{e.stopPropagation();del(s.id)}}>Delete</button></div>)}</div>
    {selected && <div className="panel"><h3>{selected.statement.title}</h3><table><thead><tr><th>Date</th><th>Person</th><th>Merchant</th><th>Original</th><th>Line Amount</th><th>Full Charge</th></tr></thead><tbody>{selected.lineItems.map(i=><tr key={i.id}><td>{i.transaction_date}</td><td>{i.person}</td><td>{i.merchant}</td><td className="original">{i.original_description}</td><td>{money(i.amount)}</td><td>{money(i.original_amount)}</td></tr>)}</tbody></table></div>}
  </div>
}

function Rules() {
  const [rules, setRules] = useState([]); const [form,setForm]=useState({match_text:'',clean_name:''});
  async function load(){ setRules((await api('/merchant-rules')).rules); }
  useEffect(()=>{ load(); }, []);
  async function add(e){ e.preventDefault(); await api('/merchant-rules',{method:'POST',body:JSON.stringify(form)}); setForm({match_text:'',clean_name:''}); load(); }
  return <div className="page"><div className="page-head"><h2>Merchant Rules</h2><p>When the original description contains your code, the clean merchant name is applied on import.</p></div>
    <div className="panel"><form className="rule-form" onSubmit={add}><input placeholder="Original contains, e.g. MTA*NYCT PAYGO" value={form.match_text} onChange={e=>setForm({...form,match_text:e.target.value})}/><input placeholder="Clean name, e.g. Subway" value={form.clean_name} onChange={e=>setForm({...form,clean_name:e.target.value})}/><button className="primary">Add Rule</button></form>
    <table><thead><tr><th>Original contains</th><th>Clean name</th><th></th></tr></thead><tbody>{rules.map(r=><tr key={r.id}><td>{r.match_text}</td><td>{r.clean_name}</td><td><button className="danger" onClick={async()=>{await api('/merchant-rules/'+r.id,{method:'DELETE'});load();}}>Delete</button></td></tr>)}</tbody></table></div>
  </div>
}

function UsersPage() {
  const [users,setUsers]=useState([]); const [form,setForm]=useState({name:'',email:'',password:'',role:'user'}); const [msg,setMsg]=useState('');
  async function load(){ setUsers((await api('/users')).users); }
  useEffect(()=>{ load().catch(e=>setMsg(e.message)); }, []);
  async function add(e){ e.preventDefault(); await api('/users',{method:'POST',body:JSON.stringify(form)}); setForm({name:'',email:'',password:'',role:'user'}); load(); }
  return <div className="page"><div className="page-head"><h2>Users</h2><p>Create separate logins. Every user only sees their own statements, rules, people, and payments.</p></div>{msg&&<div className="notice">{msg}</div>}
    <div className="panel"><form className="rule-form" onSubmit={add}><input placeholder="Name" value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/><input placeholder="Email" value={form.email} onChange={e=>setForm({...form,email:e.target.value})}/><input placeholder="Temporary password" value={form.password} onChange={e=>setForm({...form,password:e.target.value})}/><select value={form.role} onChange={e=>setForm({...form,role:e.target.value})}><option>user</option><option>admin</option></select><button className="primary">Create</button></form>
    <table><thead><tr><th>Name</th><th>Email</th><th>Role</th><th>Created</th></tr></thead><tbody>{users.map(u=><tr key={u.id}><td>{u.name}</td><td>{u.email}</td><td>{u.role}</td><td>{u.created_at}</td></tr>)}</tbody></table></div>
  </div>
}

function downloadCSV(filename, rows) {
  const csv = rows.map(r => r.map(v => '"' + String(v ?? '').replace(/"/g,'""') + '"').join(',')).join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' })); a.download = filename; a.click(); URL.revokeObjectURL(a.href);
}

function App() {
  const [user,setUser]=useState(null); const [loading,setLoading]=useState(true); const [tab,setTab]=useState('dashboard');
  useEffect(()=>{ api('/me').then(d=>setUser(d.user)).catch(()=>{}).finally(()=>setLoading(false)); }, []);
  if (loading) return <div className="loading">Loading...</div>;
  if (!user) return <Login onLogin={setUser}/>;
  const nav = [ ['dashboard',LayoutDashboard,'Dashboard'], ['upload',Upload,'Upload'], ['statements',FileText,'Statements'], ['rules',Settings,'Rules'] ];
  if (user.role === 'admin') nav.push(['users',Users,'Users']);
  return <div className="app"><aside><div className="brand"><div className="logo">$</div><div><strong>Statement Splitter</strong><span>{user.name}</span></div></div>{nav.map(([id,Icon,label])=><button key={id} className={tab===id?'active':''} onClick={()=>setTab(id)}><Icon size={18}/>{label}</button>)}<button className="logout" onClick={async()=>{await api('/logout',{method:'POST'});location.reload();}}><LogOut size={18}/>Log out</button></aside><main>{tab==='dashboard'&&<Dashboard/>}{tab==='upload'&&<UploadPage/>}{tab==='statements'&&<Statements/>}{tab==='rules'&&<Rules/>}{tab==='users'&&<UsersPage/>}</main></div>;
}

createRoot(document.getElementById('root')).render(<App />);

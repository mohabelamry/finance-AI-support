import * as XLSX from 'xlsx/xlsx.mjs';
export const today = () => new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Riyadh',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
export function dateValue(value) {
 if (typeof value==='number') {const d=XLSX.SSF.parse_date_code(value);if(!d)throw Error('Invalid Excel date');return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;}
 if(value instanceof Date) return value.toISOString().slice(0,10);
 const s=String(value??'').trim(); if(!/^\d{4}-\d{2}-\d{2}$/.test(s)||!Number.isFinite(Date.parse(s))||new Date(s).toISOString().slice(0,10)!==s)throw Error('Use valid YYYY-MM-DD dates or Excel date cells');return s;
}
export function status(invoice,now=today()) {if(invoice.paid)return {label:'Paid',days:0};const days=Math.round((Date.parse(now)-Date.parse(invoice.dueDate))/86400000);return {label:days>0?'Overdue':'Due',days:Math.max(0,days)};}
export function parseRows(rows) {
 if(!rows.length) throw Error('The sheet is empty.');const ids=new Set();return rows.map((row,index)=>{
 const r=Object.fromEntries(Object.entries(row).map(([k,v])=>[k.toLowerCase().replace(/[^a-z]/g,''),v]));
 const id=String(r.invoice??r.invoicenumber??'').trim(),customer=String(r.customer??'').trim(),amount=Number(r.amount),currency=String(r.currency??'USD').toUpperCase().trim(),email=String(r.email??'').trim(),state=String(r.status??'Due').toLowerCase().trim();
 if(!id||!customer||r.amount==null||r.amount===''||!Number.isFinite(amount)||amount<0)throw Error(`Row ${index+2}: invoice, customer, and a nonnegative amount are required.`);
 if(ids.has(id))throw Error(`Row ${index+2}: duplicate invoice ${id}.`);ids.add(id);
 if(!/^[A-Z]{3}$/.test(currency))throw Error(`Row ${index+2}: use a three-letter currency code.`);try{new Intl.NumberFormat('en',{style:'currency',currency}).format(0)}catch{throw Error(`Row ${index+2}: invalid currency.`)}
 if(email&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw Error(`Row ${index+2}: invalid email.`);
 if(!['paid','due','overdue','unpaid'].includes(state))throw Error(`Row ${index+2}: status must be Paid, Due, Overdue, or Unpaid.`);
 try{return {id,customer,amount,currency,email,dueDate:dateValue(r.duedate),paid:state==='paid'}}catch(e){throw Error(`Row ${index+2}: ${e.message}`)}
 });
}
export function readWorkbook(buffer){const book=XLSX.read(buffer,{type:'array'});return parseRows(XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]],{defval:''}));}
export function downloadTemplate(){const sheet=XLSX.utils.json_to_sheet([{Invoice:'INV-1001',Customer:'Example customer',Amount:4800,Currency:'USD','Due date':'2026-10-01',Status:'Due',Email:'accounts@example.com'}]);const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,sheet,'Invoices');XLSX.writeFile(book,'InvoiceFlow-template.xlsx');}
export const demo=[{id:'INV-1042',customer:'Acme Studio',amount:4800,dueDate:'2026-10-01',paid:false,email:'accounts@acme.example'},{id:'INV-1041',customer:'Northstar Labs',amount:7250,dueDate:'2026-10-08',paid:false,email:'finance@northstar.example'},{id:'INV-1040',customer:'Horizon Co',amount:3600,dueDate:'2026-10-15',paid:false,email:'billing@horizon.example'},{id:'INV-1039',customer:'Atlas Group',amount:9100,dueDate:'2026-10-18',paid:false,email:'accounts@atlas.example'},{id:'INV-1038',customer:'Cedar Design',amount:5200,dueDate:'2026-09-28',paid:true,email:'finance@cedar.example'}].map(i=>({...i,currency:'USD'}));

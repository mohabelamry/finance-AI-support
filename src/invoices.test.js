import {test} from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';
import {status,parseRows,readWorkbook,dateValue} from './invoices.js';
const row={Invoice:'INV-1',Customer:'Acme',Amount:100,Currency:'USD','Due date':'2026-10-01',Status:'Due',Email:'accounts@example.com'};
test('overdue boundary, future invoice and paid precedence',()=>{assert.deepEqual(status({dueDate:'2026-10-01',paid:false},'2026-10-11'),{label:'Overdue',days:10});assert.equal(status({dueDate:'2026-10-11'},'2026-10-11').label,'Due');assert.equal(status({dueDate:'2026-10-15'},'2026-10-11').days,0);assert.equal(status({dueDate:'2026-10-01',paid:true},'2026-10-11').label,'Paid')});
test('reads actual Excel workbook and preserves values',()=>{const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.json_to_sheet([row]),'Invoices');const data=readWorkbook(XLSX.write(book,{type:'array',bookType:'xlsx'}));assert.equal(data[0].id,'INV-1');assert.equal(data[0].amount,100);assert.equal(data[0].email,'accounts@example.com')});
test('rejects duplicate IDs, bad amounts, impossible dates and invalid status',()=>{assert.throws(()=>parseRows([row,row]),/duplicate/);assert.throws(()=>parseRows([{...row,Amount:'abc'}]),/amount/);assert.throws(()=>parseRows([{...row,'Due date':'2026-02-30'}]),/valid/);assert.throws(()=>parseRows([{...row,Status:'unknown'}]),/status/);assert.throws(()=>parseRows([]),/empty/)});
test('parses Excel date serials',()=>{assert.equal(dateValue(45931),'2025-10-01')});

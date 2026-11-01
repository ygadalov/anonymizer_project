// Newly written recovery verification; not extracted from HTML.
const test = require('node:test');
const assert = require('node:assert/strict');
const L = require('../src/limits.js');
const Zip = require('../vendor/jszip.min.js');
test('6.2 resource checks accept valid ZIP and reject invalid XML/grid/archive', async () => {
 const z = new Zip(); z.file('xl/worksheets/sheet1.xml','<worksheet><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>');
 const b = await z.generateAsync({type:'nodebuffer'});
 assert.equal((await L.validateZip(b,Zip)).entries,3);
 assert.throws(()=>L.checkGrid(100001,1));
 assert.throws(()=>L.checkFile(26*1024*1024));
 assert.throws(()=>L.checkZip(new Uint8Array(30)));
 assert.throws(()=>L.inspectXml('<a><b></a>'));
});

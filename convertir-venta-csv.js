// convertir-venta-csv.js
// -----------------------------------------------------------------------------
// Convierte venta.xlsx (export pesado del POS) a venta.csv (liviano) para que
// el dashboard lo pueda leer en modo streaming, sin importar cuántas filas tenga.
//
// IMPORTANTE: este script usa "exceljs" en modo streaming (lee el XML interno
// del Excel como un flujo, fila por fila) en vez de "xlsx"/SheetJS, porque con
// archivos muy grandes SheetJS intenta armar el contenido de la hoja como UN
// SOLO string de JavaScript, y el motor V8 (el mismo de Chrome y Node) tiene un
// límite duro de ~512 millones de caracteres por string. Tu venta.xlsx actual
// supera ese límite, así que SheetJS falla silenciosamente tanto en el
// navegador como en Node. exceljs en modo streaming no tiene ese problema
// porque nunca junta todo el archivo en un solo string.
//
// USO:
//   1) Instalá la dependencia una sola vez (en la carpeta de este script):
//        npm install exceljs
//   2) Corré:
//        node convertir-venta-csv.js "C:\ruta\a\venta.xlsx" "C:\ruta\a\venta.csv"
//
//      Si no pasás rutas, por defecto busca "venta.xlsx" en la carpeta actual
//      y genera "venta.csv" al lado.
//
//   3) Subí venta.csv (no venta.xlsx) al repo ampm-dashboard-data-2026 con el
//      mismo nombre de siempre para que el dashboard lo detecte automáticamente.
// -----------------------------------------------------------------------------

const fs = require('fs');
const path = require('path');

let ExcelJS;
try {
  ExcelJS = require('exceljs');
} catch (e) {
  console.error('Falta la librería "exceljs". Instalala con:  npm install exceljs');
  process.exit(1);
}

const inputPath = process.argv[2] || path.join(process.cwd(), 'venta.xlsx');
const outputPath = process.argv[3] || path.join(process.cwd(), 'venta.csv');

// Estas son EXACTAMENTE las columnas que usa el dashboard.
const NEEDED = [
  'Store Name', 'Department', 'Category', 'Supplier', 'Item', 'Description',
  'Marca', 'Qty Sold', 'Total Sales', 'Total Gross Margin', 'Transaction', 'Date Sold'
];

if (!fs.existsSync(inputPath)) {
  console.error('No se encontró el archivo de entrada:', inputPath);
  process.exit(1);
}

function fmtDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    const p = n => String(n).padStart(2, '0');
    return v.getFullYear() + '-' + p(v.getMonth() + 1) + '-' + p(v.getDate()) + ' ' +
           p(v.getHours()) + ':' + p(v.getMinutes()) + ':' + p(v.getSeconds());
  }
  return v == null ? '' : String(v);
}

function csvEscape(v) {
  if (v == null) return '';
  const s = String(v);
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function cellValue(cell) {
  // exceljs a veces entrega objetos ricos ({result:...}, {text:...}, fechas, etc).
  const v = cell.value;
  if (v == null) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if (v.result != null) return v.result;
    if (v.text != null) return v.text;
    if (v.richText) return v.richText.map(p => p.text).join('');
  }
  return v;
}

async function main() {
  console.log('Leyendo (modo streaming)', inputPath, '...');
  const t0 = Date.now();

  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(inputPath, {
    entries: 'emit',
    sharedStrings: 'cache',
    hyperlinks: 'ignore',
    styles: 'ignore',
    worksheets: 'emit'
  });

  const out = fs.createWriteStream(outputPath, { encoding: 'utf8' });
  out.write(NEEDED.join(',') + '\n');

  let header = null;
  let idx = null;
  let dateColPos = NEEDED.indexOf('Date Sold');
  let written = 0;
  let rowNum = 0;

  for await (const worksheetReader of workbookReader) {
    for await (const row of worksheetReader) {
      rowNum++;
      const values = row.values; // 1-indexed array (values[0] es undefined)
      if (!header) {
        header = values.map(v => (v == null ? '' : String(v).trim()));
        idx = NEEDED.map(col => {
          const i = header.indexOf(col);
          if (i === -1) {
            console.error('No se encontró la columna requerida "' + col + '" en el archivo de origen.');
            console.error('Encabezados encontrados:', header.filter(Boolean).join(' | '));
            process.exit(1);
          }
          return i;
        });
        continue;
      }
      const vals = idx.map(i => {
        const cell = row.getCell(i);
        return cellValue(cell);
      });
      vals[dateColPos] = fmtDate(vals[dateColPos]);
      out.write(vals.map(csvEscape).join(',') + '\n');
      written++;
      if (written % 100000 === 0) {
        console.log(written, 'filas escritas...', `(${((Date.now()-t0)/1000).toFixed(1)}s)`);
      }
    }
  }

  out.end(() => {
    const secs = ((Date.now() - t0) / 1000).toFixed(1);
    console.log('Listo. Filas escritas:', written, `(${secs}s total)`);
    console.log('Archivo generado:', outputPath);
    const stat = fs.statSync(outputPath);
    console.log('Peso:', (stat.size / 1024 / 1024).toFixed(1), 'MB');
  });
}

main().catch(err => {
  console.error('ERROR:', err && err.message ? err.message : err);
  process.exit(1);
});

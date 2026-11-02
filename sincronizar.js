// Copia la lógica de fechas y obligaciones de index.html a Calendario.gs,
// para que la página y el envío de correos calculen exactamente lo mismo.
// Uso (antes de cada "clasp push"):  node sincronizar.js
const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
const bloque = html.match(/\/\/ <calendario>[^\n]*\n([\s\S]*?)\/\/ <\/calendario>/);
if (!bloque) {
  console.error('No se encontró el bloque "// <calendario> ... // </calendario>" en index.html');
  process.exit(1);
}
fs.writeFileSync(
  path.join(__dirname, 'Calendario.gs'),
  '// ARCHIVO GENERADO por sincronizar.js a partir de index.html: no lo edites aquí.\n\n' + bloque[1]
);
console.log('Calendario.gs actualizado');

// ARCHIVO GENERADO por sincronizar.js a partir de index.html: no lo edites aquí.

// ================= Fechas =================
const MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
const DIAS = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
const DIAS_C = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];

// Periodos (años) cuyo cronograma mensual ya publicó la SUNAT; los demás son provisionales
const CRONOGRAMA_PUBLICADO = [2026];

// Día hábil del mes siguiente al periodo, según el último dígito del RUC (R.S. 281-2022/SUNAT)
const DIA_DECLARACION = { 0: 11, 1: 12, 2: 13, 3: 13, 4: 14, 5: 14, 6: 15, 7: 15, 8: 16, 9: 16 }; // Anexo I
const DIA_REGISTROS   = { 0: 10, 1: 11, 2: 12, 3: 12, 4: 13, 5: 13, 6: 14, 7: 14, 8: 15, 9: 15 }; // Anexo II

// Declaración anual del Impuesto a la Renta: fechas por último dígito (0..9), por ejercicio
const RENTA_ANUAL = {
  2025: { // R.S. 000386-2025/SUNAT
    general: ['2026-03-26','2026-03-27','2026-03-30','2026-03-31','2026-04-01','2026-04-06','2026-04-07','2026-04-08','2026-04-09','2026-04-10'],
    mype:    ['2026-05-27','2026-05-28','2026-05-29','2026-06-01','2026-06-02','2026-06-03','2026-06-04','2026-06-05','2026-06-08','2026-06-09']
  }
};

const pad = n => String(n).padStart(2, '0');
const iso = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const deIso = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
const corta = d => `${DIAS_C[d.getDay()]} ${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
const larga = d => `${DIAS[d.getDay()]} ${d.getDate()} de ${MESES[d.getMonth()]} de ${d.getFullYear()}`;
const diasEntre = (a, b) => Math.round((Date.UTC(b.getFullYear(), b.getMonth(), b.getDate()) - Date.UTC(a.getFullYear(), a.getMonth(), a.getDate())) / 86400000);
const periodoTexto = (y, m) => `${MESES[m]} ${y}`;

// Domingo de Pascua (algoritmo gregoriano anónimo)
function pascua(y) {
  const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4;
  const f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
  return new Date(y, Math.floor((h + l - 7 * m + 114) / 31) - 1, ((h + l - 7 * m + 114) % 31) + 1);
}

// Feriados nacionales del Perú
const cacheFeriados = {};
function feriados(y) {
  if (cacheFeriados[y]) return cacheFeriados[y];
  const fijos = ['01-01','05-01','06-07','06-29','07-23','07-28','07-29','08-06','08-30','10-08','11-01','12-08','12-09','12-25'];
  const s = new Set(fijos.map(md => `${y}-${md}`));
  const p = pascua(y);
  s.add(iso(new Date(y, p.getMonth(), p.getDate() - 3))); // Jueves Santo
  s.add(iso(new Date(y, p.getMonth(), p.getDate() - 2))); // Viernes Santo
  return (cacheFeriados[y] = s);
}
const esHabil = d => d.getDay() !== 0 && d.getDay() !== 6 && !feriados(d.getFullYear()).has(iso(d));

// n-ésimo día hábil del mes siguiente al periodo (y, m)
function diaHabil(y, m, n) {
  const d = new Date(y, m + 1, 1);
  for (let c = 0; ; d.setDate(d.getDate() + 1)) {
    if (esHabil(d) && ++c === n) return new Date(d);
  }
}

// ================= RUC =================
function errorRuc(ruc) {
  if (ruc.length !== 11) return `El RUC tiene 11 dígitos y escribiste ${ruc.length}. Revisa que no falte ni sobre ninguno.`;
  if (!/^(10|15|16|17|20)/.test(ruc)) return 'Un RUC empieza con 10, 15, 16, 17 o 20. Revisa los primeros dígitos.';
  const pesos = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2];
  let dv = 11 - pesos.reduce((s, p, i) => s + p * Number(ruc[i]), 0) % 11;
  if (dv === 10) dv = 0; else if (dv === 11) dv = 1;
  if (dv !== Number(ruc[10])) return 'Ese RUC no existe: su último dígito no coincide con el dígito de verificación. Revisa que esté bien escrito.';
  return '';
}

// ================= Obligaciones según el perfil =================
function obligaciones(p) {
  const L = [];
  if (p.regimen === 'NRUS') {
    L.push({ id: 'nrus', grupo: 'mensual', nombre: 'Cuota del Nuevo RUS', form: 'Formulario 1611 o Guía Pago Fácil' });
  } else {
    L.push({ id: 'f621', grupo: 'mensual', nombre: p.regimen === 'RER' ? 'IGV y Renta del RER' : 'IGV y pago a cuenta de Renta', form: 'Formulario 621 (Declara Fácil)' });
  }
  if (p.trabajadores) L.push({ id: 'plame', grupo: 'mensual', nombre: 'Planilla electrónica (PLAME)', form: 'Formulario 601: EsSalud, ONP y retenciones de renta' });
  if (p.regimen !== 'NRUS') {
    if (p.retencion) L.push({ id: 'ret', grupo: 'mensual', nombre: 'Retenciones del IGV', form: 'PDT 626' });
    if (p.percepcion) L.push({ id: 'per', grupo: 'mensual', nombre: 'Percepciones del IGV', form: 'PDT 697' });
    if (p.itan) L.push({ id: 'itan', grupo: 'mensual', nombre: 'Cuota del ITAN', form: 'Nueve cuotas, de abril a diciembre' });
    if (p.registros) {
      L.push({ id: 'rvie', grupo: 'registro', nombre: 'Registro de Ventas e Ingresos', form: 'SIRE o PLE' });
      L.push({ id: 'rce', grupo: 'registro', nombre: 'Registro de Compras', form: 'SIRE o PLE' });
    }
  }
  if (p.regimen === 'RMT' || p.regimen === 'RG') {
    L.push({ id: 'anual', grupo: 'anual', nombre: 'Declaración anual de Renta', form: p.regimen === 'RMT' ? 'Formulario 710, cronograma MYPE' : 'Formulario 710' });
  }
  return L;
}

const colorPorDias = d => d <= 2 ? 'r' : d <= 7 ? 'a' : 'v';

// Calcula el periodo que toca mostrar, su vencimiento y su estado
function evaluar(ob, hoy, digito, marcas, regimen) {
  if (ob.grupo === 'anual') return evaluarAnual(ob, hoy, digito, marcas, regimen);

  const tabla = ob.grupo === 'registro' ? DIA_REGISTROS : DIA_DECLARACION;
  // ITAN: solo periodos de marzo a noviembre (cuotas que vencen de abril a diciembre)
  const ajustar = (y, m) => ob.id !== 'itan' ? [y, m] : m < 2 ? [y, 2] : m > 10 ? [y + 1, 2] : [y, m];

  // Periodo que se declara este mes: el mes anterior (se muestra todo el mes, con su sello si ya se presentó)
  const [y, m] = ajustar(...(hoy.getMonth() === 0 ? [hoy.getFullYear() - 1, 11] : [hoy.getFullYear(), hoy.getMonth() - 1]));
  const vence = diaHabil(y, m, tabla[digito]);
  const clave = `${ob.id}:${y}-${pad(m + 1)}`;
  const dias = diasEntre(hoy, vence);
  const extra = ob.id === 'itan' ? `cuota ${m - 1} de 9, ` : '';
  return {
    ob, clave, vence, dias, hecho: marcas[clave] || null,
    periodo: `${extra}periodo ${periodoTexto(y, m)}`,
    provisional: !CRONOGRAMA_PUBLICADO.includes(y)
  };
}

function evaluarAnual(ob, hoy, digito, marcas, regimen) {
  const fecha = ej => RENTA_ANUAL[ej] ? deIso(RENTA_ANUAL[ej][regimen === 'RMT' ? 'mype' : 'general'][digito]) : null;
  const ej = hoy.getFullYear() - 1; // se presenta este año
  const vence = fecha(ej), clave = `anual:${ej}`;
  return {
    ob, clave, vence, dias: vence ? diasEntre(hoy, vence) : null, hecho: marcas[clave] || null,
    periodo: `ejercicio ${ej}`, sinFecha: !vence, anioPresenta: ej + 1
  };
}

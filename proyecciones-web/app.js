// ============================================================
// Tablas ISR (fuente: Configuracion sheet del xlsm original)
// ============================================================
const ISR_ANUAL = [
  [1,          6942.20,    0,          0.0192],
  [6942.21,    58922.16,   133.28,     0.064],
  [58922.17,   103550.44,  3460.01,    0.1088],
  [103550.45,  120372.83,  8315.57,    0.16],
  [120372.84,  144119.23,  11007.14,   0.1792],
  [144119.24,  290667.75,  15262.49,   0.2136],
  [290667.76,  458132.29,  46565.26,   0.2352],
  [458132.30,  874650.00,  85952.92,   0.30],
  [874650.01,  1166200.00, 210908.23,  0.32],
  [1166200.01, 3498600.00, 304204.21,  0.34],
  [3498600.01, 99999999,   1097220.21, 0.35],
];

const ISR_MENSUAL = [
  [1,          578.52,     0,         0.0192],
  [578.53,     4910.18,    11.11,     0.064],
  [4910.19,    8626.20,    288.33,    0.1088],
  [8626.21,    10031.07,   692.96,    0.16],
  [10031.08,   12009.94,   917.26,    0.1792],
  [12009.95,   24222.31,   1271.87,   0.2136],
  [24222.32,   38177.69,   3880.44,   0.2352],
  [38177.70,   72887.50,   7162.74,   0.30],
  [72887.51,   97183.33,   17575.69,  0.32],
  [97183.34,   291550.00,  25350.35,  0.34],
  [291550.01,  99999999,   91435.02,  0.35],
];

// ============================================================
// Tasas IMSS (trabajador) – tomadas de hoja Configuracion
// ============================================================
const IMSS_RATES = {
  especieExcedente:         0.004,   // H11
  prestacionesDinero:       0.0025,  // H13
  pensionadosBeneficiarios: 0.00375, // H12
  invalidezVida:            0.00625, // H14
  cesantiaVejez:            0.01125, // H16
  aporPatron:               0.05,    // E18
};

// ============================================================
// Motor de cálculo (traducción 1:1 del VBA addProjection.frm)
// ============================================================
function calcular(input) {
  const {
    nombre,
    sueldoBruto,   // Sueldo bruto total informado
    uma,
    diasVacaciones,
    primaVacacional, // en porcentaje, ej: 25
    diasAguinaldo,
    vales,           // mensuales
    fondoAhorro,     // boolean
    otros,           // mensuales SIN impuestos
    otrosTaxable,    // mensuales CON impuestos
    esquema,         // 'nominal' | 'mixto'
    mixSalary,       // monto o % de la parte ante el IMSS (solo si mixto)
    mixPerc,         // boolean: mixSalary es %?
  } = input;

  // ── 1. Salarios efectivos ──────────────────────────────────
  let WSueldo, sueldoMixto;
  if (esquema === 'mixto') {
    WSueldo = mixPerc
      ? sueldoBruto * mixSalary / 100
      : mixSalary;
    sueldoMixto = sueldoBruto - WSueldo;
  } else {
    WSueldo      = sueldoBruto;
    sueldoMixto  = 0;
  }

  const primaVac = primaVacacional / 100;

  // ── 2. SBC (Salario Base de Cotización) ────────────────────
  const topeVales     = uma * 0.4 * 30;
  const excedenteVales = Math.max(0, vales - topeVales);
  const factorVales   = excedenteVales > 0 ? excedenteVales / 30 : 0;
  const valesTaxable  = excedenteVales > 0 ? excedenteVales * 12 : 0;

  const SDI              = ((WSueldo + factorVales + otrosTaxable) * 12) / 365;
  const factorVacaciones = (diasVacaciones * primaVac) / 365;
  const factorAguinaldo  = diasAguinaldo / 365;

  let SBC = (1 + factorVacaciones + factorAguinaldo) * SDI;
  const topeSBC = 25 * uma;
  if (SBC > topeSBC) SBC = topeSBC;

  // ── 3. Fondo de Ahorro ─────────────────────────────────────
  let fondoMensual, fondoAnual;
  if (fondoAhorro) {
    fondoMensual = 1.3 * uma * 30.4;
    fondoAnual   = fondoMensual * 24; // 12 meses empleado + 12 patrón
  } else {
    fondoMensual = 0;
    fondoAnual   = 0;
  }

  // ── 4. Exentos y valores anuales ──────────────────────────
  const exentoPrima      = uma * 15.2;
  const exentoAguinaldo  = uma * 30.4;

  const aguinaldo        = ((WSueldo * 12) / 365) * diasAguinaldo;
  const primaVacMoney    = ((WSueldo * 12) / 365) * diasVacaciones * primaVac;

  const aguinaldoTaxable = aguinaldo       > exentoAguinaldo ? aguinaldo       - exentoAguinaldo : 0;
  const primaTaxable     = primaVacMoney   > exentoPrima     ? primaVacMoney   - exentoPrima     : 0;

  const aportacionesPatron = SBC * 365 * IMSS_RATES.aporPatron;

  const sueldoAnualTaxable  = ((WSueldo + otrosTaxable) * 12) + primaTaxable + aguinaldoTaxable + valesTaxable;
  const brutoBrutoAnual     = (WSueldo * 12) + primaVacMoney + aguinaldo + fondoAnual
                             + (vales * 12) + (sueldoMixto * 12)
                             + (otros * 12) + (otrosTaxable * 12);

  // ── 5. IMSS ────────────────────────────────────────────────
  function calcIMSS(periodo) {
    const ee  = Math.max(0, SBC - 3 * uma) * periodo * IMSS_RATES.especieExcedente;
    const pd  = SBC * periodo * IMSS_RATES.prestacionesDinero;
    const pb  = SBC * periodo * IMSS_RATES.pensionadosBeneficiarios;
    const iv  = SBC * periodo * IMSS_RATES.invalidezVida;
    const cv  = SBC * periodo * IMSS_RATES.cesantiaVejez;
    return ee + pd + pb + iv + cv;
  }

  const imssMonthly = calcIMSS(30.4);
  const imssAnnual  = calcIMSS(365);

  // ── 6. AFORE ───────────────────────────────────────────────
  const cesantiaVejezPatron  = 0.0315  * SBC * 365;
  const seguroRetiroPatron   = 0.02    * SBC * 365; // F15
  const gobiernoContrib      = 0.0023  * SBC * 365;
  const afore = (cesantiaVejezPatron + seguroRetiroPatron + gobiernoContrib
               + (SBC * 365 * IMSS_RATES.cesantiaVejez)) * 1.07;

  // ── 7. ISR ─────────────────────────────────────────────────
  function buscarBracket(ingreso, tabla) {
    for (const [li, ls, cuota, pct] of tabla) {
      if (ingreso >= li && ingreso <= ls) {
        return { li, cuota, pct };
      }
    }
    // Último bracket si excede
    const last = tabla[tabla.length - 1];
    return { li: last[0], cuota: last[2], pct: last[3] };
  }

  function calcISR(periodo) {
    let ingreso, tabla;
    if (periodo === 30) {
      ingreso = WSueldo + otrosTaxable;
      tabla   = ISR_MENSUAL;
    } else {
      ingreso = sueldoAnualTaxable;
      tabla   = ISR_ANUAL;
    }
    const { li, cuota, pct } = buscarBracket(ingreso, tabla);
    return (ingreso - li) * pct + cuota;
  }

  const isrMensual = calcISR(30);
  const isrAnual   = calcISR(365);

  // ── 8. Netos ───────────────────────────────────────────────
  const liquidoMensual = WSueldo + sueldoMixto + vales + otrosTaxable + otros
                        - fondoMensual - isrMensual - imssMonthly;
  const quincena       = (liquidoMensual - vales) / 2;
  const netoAnual      = brutoBrutoAnual - imssAnnual - isrAnual;

  return {
    // Inputs
    nombre, esquema: esquema === 'mixto' ? 'Mixto' : '100% Nominal',
    sueldoBruto, WSueldo, sueldoMixto,
    diasVacaciones, primaVacacional: primaVac,
    vales, fondoAhorro: fondoAhorro ? 'Sí' : 'No',
    diasAguinaldo, SBC, uma,
    // Mensuales
    monthly: {
      sueldoBruto: WSueldo,
      parteMixta: sueldoMixto,
      vales,
      fondoAhorro: fondoMensual,
      otros,
      otrosTaxable,
      isr: isrMensual,
      imss: imssMonthly,
      liquido: liquidoMensual,
      quincena,
    },
    // Anuales
    annual: {
      sueldoBruto: WSueldo * 12,
      parteMixta: sueldoMixto * 12,
      vales: vales * 12,
      fondoAhorro: fondoAnual,
      aguinaldo,
      primaVacacional: primaVacMoney,
      otros: (otros + otrosTaxable) * 12,
      bruto: brutoBrutoAnual,
      isr: isrAnual,
      imss: imssAnnual,
      neto: netoAnual,
    },
    extras: {
      aportacionesPatron,
      afore,
      exentoVales: topeVales,
      exentoAguinaldo,
      exentoPrima,
    },
  };
}

// ============================================================
// Utilidades UI
// ============================================================
const fmt = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 2 });
const fmtN = v => fmt.format(v);
const pct  = v => (v * 100).toFixed(0) + '%';

// ============================================================
// Estado de la aplicación
// ============================================================
let projections = JSON.parse(localStorage.getItem('proyecciones') || '[]');

function saveState() {
  localStorage.setItem('proyecciones', JSON.stringify(projections));
}

// ============================================================
// Render results panel
// ============================================================
function renderResults(r) {
  const panel = document.getElementById('resultsPanel');
  panel.innerHTML = `
    <h2 class="section-title">📊 ${r.nombre}</h2>

    <div class="grid-2col">

      <!-- DATOS GENERALES -->
      <div class="card card-info">
        <h3>Datos de la Proyección</h3>
        <table class="data-table">
          <tr><td>Esquema</td><td>${r.esquema}</td></tr>
          <tr><td>Sueldo Bruto (ante IMSS)</td><td>${fmtN(r.WSueldo)}</td></tr>
          ${r.sueldoMixto > 0 ? `<tr><td>Parte Mixta</td><td>${fmtN(r.sueldoMixto)}</td></tr>` : ''}
          <tr><td>Días de Vacaciones</td><td>${r.diasVacaciones}</td></tr>
          <tr><td>Prima Vacacional</td><td>${pct(r.primaVacacional)}</td></tr>
          <tr><td>Vales de Despensa</td><td>${fmtN(r.vales)}</td></tr>
          <tr><td>Fondo de Ahorro</td><td>${r.fondoAhorro}</td></tr>
          <tr><td>Días de Aguinaldo</td><td>${r.diasAguinaldo}</td></tr>
          <tr class="highlight"><td>SBC</td><td>${fmtN(r.SBC)}</td></tr>
          <tr><td>UMA</td><td>${fmtN(r.uma)}</td></tr>
        </table>
      </div>

      <!-- EXTRAS -->
      <div class="card card-extra">
        <h3>Datos Extra</h3>
        <table class="data-table">
          <tr><td>Aportaciones Patronales Anuales</td><td>${fmtN(r.extras.aportacionesPatron)}</td></tr>
          <tr><td>AFORE Anual (est. 7%)</td><td>${fmtN(r.extras.afore)}</td></tr>
          <tr><td>Exento Vales (mensual)</td><td>${fmtN(r.extras.exentoVales)}</td></tr>
          <tr><td>Exento Aguinaldo</td><td>${fmtN(r.extras.exentoAguinaldo)}</td></tr>
          <tr><td>Exento Prima Vacacional</td><td>${fmtN(r.extras.exentoPrima)}</td></tr>
        </table>
      </div>
    </div>

    <!-- PERCEPCIONES Y RETENCIONES -->
    <div class="card card-main">
      <table class="comparison-table">
        <thead>
          <tr>
            <th>Concepto</th>
            <th>Mensual</th>
            <th>Anual</th>
          </tr>
        </thead>
        <tbody>
          <tr class="section-header"><td colspan="3">Percepciones</td></tr>
          <tr><td>Sueldo Bruto ante el IMSS</td><td>${fmtN(r.monthly.sueldoBruto)}</td><td>${fmtN(r.annual.sueldoBruto)}</td></tr>
          ${r.monthly.parteMixta > 0 ? `<tr><td>Parte Mixta</td><td>${fmtN(r.monthly.parteMixta)}</td><td>${fmtN(r.annual.parteMixta)}</td></tr>` : ''}
          <tr><td>Vales de Despensa</td><td>${fmtN(r.monthly.vales)}</td><td>${fmtN(r.annual.vales)}</td></tr>
          <tr><td>Fondo de Ahorro</td><td>${fmtN(r.monthly.fondoAhorro)}</td><td>${fmtN(r.annual.fondoAhorro)}</td></tr>
          ${r.monthly.otros > 0 || r.monthly.otrosTaxable > 0 ? `
          <tr><td>Otros (SIN impuestos)</td><td>${fmtN(r.monthly.otros)}</td><td>${fmtN(r.annual.otros)}</td></tr>
          ` : ''}
          <tr><td>Aguinaldo</td><td>—</td><td>${fmtN(r.annual.aguinaldo)}</td></tr>
          <tr><td>Prima Vacacional</td><td>—</td><td>${fmtN(r.annual.primaVacacional)}</td></tr>
          <tr class="subtotal"><td><strong>Total Bruto</strong></td><td><strong>—</strong></td><td><strong>${fmtN(r.annual.bruto)}</strong></td></tr>

          <tr class="section-header"><td colspan="3">Retenciones Aproximadas</td></tr>
          <tr class="deduction"><td>ISR</td><td>−${fmtN(r.monthly.isr)}</td><td>−${fmtN(r.annual.isr)}</td></tr>
          <tr class="deduction"><td>IMSS</td><td>−${fmtN(r.monthly.imss)}</td><td>−${fmtN(r.annual.imss)}</td></tr>
          ${r.monthly.fondoAhorro > 0 ? `<tr class="deduction"><td>Fondo de Ahorro</td><td>−${fmtN(r.monthly.fondoAhorro)}</td><td>—</td></tr>` : ''}

          <tr class="section-header total-row"><td colspan="3">Totales</td></tr>
          <tr class="total-row">
            <td><strong>Líquido Mensual <em>(con vales)</em></strong></td>
            <td><strong>${fmtN(r.monthly.liquido)}</strong></td>
            <td></td>
          </tr>
          <tr class="total-row">
            <td><strong>La Quincena <em>(sin vales)</em></strong></td>
            <td><strong>${fmtN(r.monthly.quincena)}</strong></td>
            <td></td>
          </tr>
          <tr class="total-row">
            <td><strong>Neto Líquido Anual</strong></td>
            <td></td>
            <td><strong>${fmtN(r.annual.neto)}</strong></td>
          </tr>
        </tbody>
      </table>
    </div>
  `;
  panel.style.display = 'block';
}

// ============================================================
// Render projections table
// ============================================================
function renderTable() {
  const wrap = document.getElementById('projectionsWrap');
  if (projections.length === 0) {
    wrap.innerHTML = '<p class="empty-msg">Aún no hay proyecciones guardadas. Calcula una y haz clic en <strong>Guardar Proyección</strong>.</p>';
    return;
  }

  const cols = [
    { key: 'nombre',           label: 'Nombre',           fmt: v => v },
    { key: 'esquema',          label: 'Esquema',          fmt: v => v },
    { key: 'WSueldo',          label: 'Sueldo IMSS',      fmt: fmtN },
    { key: 'SBC',              label: 'SBC',              fmt: fmtN },
    { key: ['monthly','isr'],  label: 'ISR Mensual',      fmt: fmtN },
    { key: ['monthly','imss'], label: 'IMSS Mensual',     fmt: fmtN },
    { key: ['monthly','liquido'], label: 'Líquido Mensual', fmt: fmtN },
    { key: ['monthly','quincena'], label: 'Quincena',     fmt: fmtN },
    { key: ['annual','neto'],  label: 'Neto Anual',       fmt: fmtN },
  ];

  const getVal = (r, key) => Array.isArray(key) ? r[key[0]][key[1]] : r[key];

  wrap.innerHTML = `
    <div class="table-scroll">
      <table class="proj-table">
        <thead>
          <tr>
            ${cols.map(c => `<th>${c.label}</th>`).join('')}
            <th>Acciones</th>
          </tr>
        </thead>
        <tbody>
          ${projections.map((r, i) => `
            <tr>
              ${cols.map(c => `<td>${c.fmt(getVal(r, c.key))}</td>`).join('')}
              <td class="actions-cell">
                <button class="btn-sm btn-view"  onclick="viewProjection(${i})">Ver</button>
                <button class="btn-sm btn-delete" onclick="deleteProjection(${i})">✕</button>
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function viewProjection(i) {
  renderResults(projections[i]);
  document.getElementById('resultsPanel').scrollIntoView({ behavior: 'smooth' });
}

function deleteProjection(i) {
  if (!confirm(`¿Eliminar la proyección "${projections[i].nombre}"?`)) return;
  projections.splice(i, 1);
  saveState();
  renderTable();
}

// ============================================================
// Lógica del formulario
// ============================================================
let lastResult = null;

function onEsquemaChange() {
  const isMixto = document.getElementById('esquemaMixto').checked;
  document.getElementById('mixtoSection').style.display = isMixto ? 'flex' : 'none';
}

function onMixPercChange() {
  const isPerc = document.getElementById('mixPerc').checked;
  document.getElementById('mixSalaryLabel').textContent = isPerc ? '% ante IMSS' : 'Monto ante IMSS ($)';
}

function getFormValues() {
  const v = id => parseFloat(document.getElementById(id).value) || 0;
  const s = id => document.getElementById(id).value.trim();
  const b = id => document.getElementById(id).checked;

  return {
    nombre:          s('nombre') || 'Sin nombre',
    sueldoBruto:     v('sueldoBruto'),
    uma:             v('uma'),
    diasVacaciones:  v('diasVacaciones'),
    primaVacacional: v('primaVacacional'),
    diasAguinaldo:   v('diasAguinaldo'),
    vales:           v('vales'),
    fondoAhorro:     b('fondoAhorro'),
    otros:           v('otros'),
    otrosTaxable:    v('otrosTaxable'),
    esquema:         b('esquemaMixto') ? 'mixto' : 'nominal',
    mixSalary:       v('mixSalary'),
    mixPerc:         b('mixPerc'),
  };
}

function calcular_click() {
  const input = getFormValues();

  // Validaciones básicas
  if (input.sueldoBruto <= 0) {
    alert('Por favor ingresa un Sueldo Bruto válido.');
    return;
  }
  if (input.uma <= 0) {
    alert('Por favor ingresa el valor de la UMA.');
    return;
  }
  if (input.esquema === 'mixto' && input.mixSalary <= 0) {
    alert('Para esquema mixto, ingresa la parte ante el IMSS.');
    return;
  }

  lastResult = calcular(input);
  renderResults(lastResult);

  document.getElementById('btnSave').style.display = 'inline-block';
  document.getElementById('resultsPanel').scrollIntoView({ behavior: 'smooth' });
}

function guardar_click() {
  if (!lastResult) return;
  // Evitar duplicados por nombre
  const exists = projections.findIndex(p => p.nombre === lastResult.nombre);
  if (exists >= 0) {
    if (!confirm(`Ya existe una proyección con el nombre "${lastResult.nombre}". ¿Reemplazar?`)) return;
    projections[exists] = lastResult;
  } else {
    projections.push(lastResult);
  }
  saveState();
  renderTable();
  document.getElementById('projectionsSection').scrollIntoView({ behavior: 'smooth' });
}

function resetForm() {
  document.getElementById('nombre').value           = '';
  document.getElementById('sueldoBruto').value      = '';
  document.getElementById('uma').value              = '113.14';
  document.getElementById('diasVacaciones').value   = '12';
  document.getElementById('primaVacacional').value  = '25';
  document.getElementById('diasAguinaldo').value    = '15';
  document.getElementById('vales').value            = '0';
  document.getElementById('otros').value            = '0';
  document.getElementById('otrosTaxable').value     = '0';
  document.getElementById('mixSalary').value        = '0';
  document.getElementById('fondoAhorro').checked    = false;
  document.getElementById('esquemaNominal').checked = true;
  document.getElementById('mixPerc').checked        = false;
  onEsquemaChange();
  document.getElementById('btnSave').style.display  = 'none';
  document.getElementById('resultsPanel').style.display = 'none';
}

// ============================================================
// Init
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  onEsquemaChange();
  renderTable();

  document.getElementById('btnCalc').addEventListener('click', calcular_click);
  document.getElementById('btnSave').addEventListener('click', guardar_click);
  document.getElementById('btnReset').addEventListener('click', resetForm);
  document.getElementById('esquemaNominal').addEventListener('change', onEsquemaChange);
  document.getElementById('esquemaMixto').addEventListener('change', onEsquemaChange);
  document.getElementById('mixPerc').addEventListener('change', onMixPercChange);
});

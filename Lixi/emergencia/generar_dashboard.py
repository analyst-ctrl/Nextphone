import json, sys

sys.stdout.reconfigure(encoding='utf-8')

with open('datos.json', encoding='utf-8') as f:
    raw = json.load(f)

months = list(raw.keys())
data_json = json.dumps(raw, ensure_ascii=False)

html = r'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Resultados por Campa&ntilde;a &mdash; Reuni&oacute;n Puyol</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<style>
:root {
  --azul: #1F4E79; --azul2: #2E75B6; --celeste: #5B9BD5; --celeste2: #A8C7E7;
  --verde: #27ae60; --rojo: #c0392b; --naranja: #f39c12;
  --bg: #eef2f7; --card: #ffffff; --tx: #2c3e50; --tx2: #7f8c8d;
  --radius: 16px; --shadow: 0 6px 20px rgba(31,78,121,.08);
}
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Segoe UI',system-ui,-apple-system,Arial,sans-serif; background:var(--bg); color:var(--tx); min-height:100vh; }
.header {
  background: linear-gradient(135deg,#16385b,#1F4E79 45%,#2E75B6);
  color:#fff; padding:34px 40px; position:relative; overflow:hidden;
}
.header::after {
  content:''; position:absolute; right:-60px; top:-60px; width:280px; height:280px;
  background:radial-gradient(circle,rgba(255,255,255,.14),transparent 70%); border-radius:50%;
}
.header::before {
  content:''; position:absolute; right:90px; top:-90px; width:200px; height:200px;
  background:radial-gradient(circle,rgba(255,255,255,.10),transparent 70%); border-radius:50%;
}
.header .wrap { position:relative; z-index:2; display:flex; justify-content:space-between; align-items:flex-end; flex-wrap:wrap; gap:16px; }
.header h1 { font-size:30px; font-weight:700; letter-spacing:.4px; }
.header .sub { font-size:14px; opacity:.85; margin-top:6px; }
.btn-descarga {
  display:inline-block; background:rgba(255,255,255,.18); border:1px solid rgba(255,255,255,.35);
  color:#fff; padding:9px 18px; border-radius:10px; text-decoration:none; font-size:14px;
  backdrop-filter:blur(4px); transition:background .2s;
}
.btn-descarga:hover { background:rgba(255,255,255,.28); }
.wrap { max-width:1200px; margin:0 auto; padding:24px 28px 60px; }
.menu-meses {
  display:flex; gap:10px; background:#fff; padding:8px; border-radius:14px;
  box-shadow:var(--shadow); margin:-28px auto 24px; max-width:fit-content; position:relative; z-index:5;
}
.btn-mes {
  border:none; background:transparent; color:var(--tx2); font-size:14px; font-weight:600;
  padding:10px 22px; border-radius:10px; cursor:pointer; transition:all .2s;
}
.btn-mes.activo { background:linear-gradient(135deg,var(--azul),var(--azul2)); color:#fff; box-shadow:0 4px 12px rgba(46,117,182,.35); }
.kpi-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:16px; margin:26px 0 10px; }
.kpi-card {
  background:var(--card); border-radius:var(--radius); padding:20px 18px; text-align:center;
  box-shadow:var(--shadow); border-top:4px solid var(--azul2); transition:transform .2s;
}
.kpi-card:hover { transform:translateY(-3px); }
.kpi-card .label { font-size:12px; text-transform:uppercase; color:var(--tx2); letter-spacing:.6px; }
.kpi-card .value { font-size:32px; font-weight:800; color:var(--azul); margin:6px 0; }
.kpi-card .sub { font-size:12px; color:var(--tx2); }
.kpi-card.verde { border-top-color:var(--verde); } .kpi-card.verde .value { color:var(--verde); }
.kpi-card.naranja { border-top-color:var(--naranja); } .kpi-card.naranja .value { color:var(--naranja); }
.kpi-card.rojo { border-top-color:var(--rojo); } .kpi-card.rojo .value { color:var(--rojo); }
.kpi-card.celeste { border-top-color:var(--celeste); } .kpi-card.celeste .value { color:var(--celeste); }
.section { margin-top:34px; }
.section-title {
  font-size:19px; font-weight:800; color:var(--azul); margin-bottom:14px;
  display:flex; align-items:center; gap:10px;
}
.section-title::before { content:''; width:6px; height:22px; border-radius:3px; background:linear-gradient(180deg,var(--azul2),var(--celeste)); }
.chart-row { display:grid; grid-template-columns:1fr 1fr; gap:18px; }
.chart-box { background:var(--card); border-radius:var(--radius); padding:20px; box-shadow:var(--shadow); }
.chart-box h3 { font-size:15px; color:var(--azul); margin-bottom:12px; }
table { width:100%; border-collapse:collapse; font-size:13px; }
th { background:var(--azul); color:#fff; padding:10px 8px; text-align:center; font-weight:600; }
td { padding:8px; border-bottom:1px solid #eef2f7; text-align:center; }
tr:nth-child(even) td { background:#f7fafd; }
tr:hover td { background:#eef5fc; }
.pos { color:var(--verde); font-weight:700; }
.neg { color:var(--rojo); font-weight:700; }
.zero { color:var(--tx2); }
.campana-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(280px,1fr)); gap:16px; }
.campana-box { background:var(--card); border-radius:var(--radius); padding:18px; box-shadow:var(--shadow); border-top:4px solid var(--celeste); }
.campana-box h4 { font-size:15px; color:var(--azul); margin-bottom:10px; display:flex; align-items:center; justify-content:space-between; }
.badge { background:var(--azul2); color:#fff; border-radius:20px; padding:3px 12px; font-size:12px; }
.badge2 { background:#eef5fc; color:var(--azul); border-radius:20px; padding:3px 12px; font-size:12px; font-weight:700; }
.mini-table { margin-top:8px; }
.mini-title { font-size:11px; text-transform:uppercase; color:var(--tx2); letter-spacing:.5px; margin:12px 0 6px; }
.nota {
  background:#eef5fc; border-left:4px solid var(--celeste); border-radius:8px;
  padding:12px 16px; font-size:13px; color:var(--tx); margin-top:18px; line-height:1.5;
}
.nota b { color:var(--azul); }
.expl {
  grid-column:1 / -1; background:#fff; border:1px solid #dbe7f3; border-left:4px solid var(--azul2);
  border-radius:10px; padding:14px 18px; font-size:13.5px; color:#445; line-height:1.65; margin-top:14px;
}
.expl b { color:var(--azul); }
.expl .tag {
  display:inline-block; background:var(--azul); color:#fff; border-radius:5px; font-size:10.5px;
  letter-spacing:.8px; text-transform:uppercase; padding:3px 8px; margin-bottom:6px; font-weight:700;
}
.fade-in { animation:fadeIn .4s ease; }
@keyframes fadeIn { from{opacity:0; transform:translateY(6px);} to{opacity:1; transform:none;} }
@media (max-width:860px) { .chart-row { grid-template-columns:1fr; } .header { padding:24px; } }
</style>
</head>
<body>

<div class="header">
  <div class="wrap" style="padding:0">
    <div>
      <h1>RESULTADOS POR CAMPA&Ntilde;A</h1>
      <div class="sub">Preparaci&oacute;n reuni&oacute;n Jos&eacute; Puyol &bull; Base: HC Actualizado &bull; Corte al d&iacute;a 13 de cada mes</div>
    </div>
    <a class="btn-descarga" href="HC%20Actualizado%20Agosto.xlsx">Descargar HC</a>
  </div>
</div>

<div class="wrap">

  <div class="menu-meses" id="menuMeses"></div>

  <div id="contenido"></div>

  <div class="nota fade-in" id="nota">
    <b>Nota metodol&oacute;gica:</b> el HC es una base de planilla (headcount) y no contiene m&eacute;tricas de ventas. Los resultados mostrados corresponden a
    composici&oacute;n y tama&ntilde;o de plantilla por campa&ntilde;a; el comparativo usa el acumulado de personas con fecha de ingreso al d&iacute;a 13 de cada mes
    (mismo corte en ambos meses), y el an&aacute;lisis de cuartiles se calcula sobre la antig&uuml;edad de la plantilla (meses desde la fecha de ingreso al corte).
  </div>

</div>

<script>
const DATA = __DATA__;
const MESES = Object.keys(DATA);
const CORTE = 13;

const fmt = new Intl.NumberFormat('es-PA');
let charts = {};

function dieEnMes(d, mes) {
  return d.y >= mes.y && d.m === mes.m;
}

function corteMes(mes) { return { y: mes.y, m: mes.m, d: CORTE }; }

function antiguedadDias(fIngreso, corte) {
  if (!fIngreso) return null;
  const [ay, am, ad] = fIngreso.split('-').map(Number);
  const f = new Date(ay, am - 1, ad);
  const c = new Date(corte.y, corte.m - 1, corte.d);
  return Math.round((c - f) / 86400000);
}

function parseMes(nombre) {
  const [mes, anio] = nombre.split(' ');
  return { y: +anio, m: +({ Julio: 7, Agosto: 8 }[mes] || 0), label: nombre, corto: mes };
}

function esUltimo(nombre) { return nombre === MESES[MESES.length - 1]; }

function kpis(rows, mes) {
  const corte = corteMes(mes);
  const act = rows.filter(r => r.e === 'Activo').length;
  const capa = rows.filter(r => r.e !== 'Activo').length;
  const cut13 = rows.filter(r => r.f && antiguedadDias(r.f, corte) >= 0).length;
  const total = rows.length;
  return { total, activos: act, capa, cut13 };
}

function campañas(rows) {
  const map = new Map();
  rows.forEach(r => { if (!map.has(r.c)) map.set(r.c, []); map.get(r.c).push(r); });
  return map;
}

function cutByCampana(rows, mes) {
  const corte = corteMes(mes);
  const map = new Map();
  rows.forEach(r => {
    if (r.f && antiguedadDias(r.f, corte) < 0) return;
    if (!map.has(r.c)) map.set(r.c, 0);
    map.set(r.c, map.get(r.c) + 1);
  });
  return map;
}

function qsData(rows, mes) {
  const corte = corteMes(mes);
  const dias = rows.map(r => antiguedadDias(r.f, corte)).filter(d => d !== null).sort((a, b) => a - b);
  const n = dias.length;
  if (!n) return null;
  const bounds = [];
  for (let i = 1; i <= 4; i++) bounds.push(dias[Math.min(Math.floor(i * n / 4), n - 1)]);
  const buckets = [[], [], [], []];
  dias.forEach(d => { for (let i = 0; i < 4; i++) if (d <= bounds[i]) { buckets[i].push(d); break; } });
  const meses = d => (d / 30.44).toFixed(1);
  return {
    n,
    q: buckets.map((b, i) => ({
      nombre: ['Q1 (m\u00e1s nuevos)', 'Q2', 'Q3', 'Q4 (m\u00e1s antiguos)'][i],
      n: b.length,
      desde: meses(Math.min(...b)), hasta: meses(Math.max(...b)),
      diasMin: Math.min(...b), diasMax: Math.max(...b),
    }))
  };
}

function renderMenu() {
  document.getElementById('menuMeses').innerHTML = MESES.map(m =>
    `<button class="btn-mes ${esUltimo(m) ? 'activo' : ''}" onclick="selectMes('${m}')">${m}</button>`).join('');
}

function selectMes(nombre) {
  document.querySelectorAll('.btn-mes').forEach(b =>
    b.classList.toggle('activo', b.textContent === nombre));
  render(nombre);
}

function render(nombre) {
  const mes = parseMes(nombre);
  const rows = DATA[nombre];
  const prev = MESES[MESES.indexOf(nombre) - 1];
  const prevRows = prev ? DATA[prev] : null;
  const k = kpis(rows, mes);
  const kPrev = prevRows ? kpis(prevRows, parseMes(prev)) : null;
  const camps = campañas(rows);
  const cuartiles = qsData(rows, mes);

  const campOrden = [...camps.keys()].sort((a, b) => camps.get(b).length - camps.get(a).length);

  const varCut = kPrev ? k.cut13 - kPrev.cut13 : null;
  const varPct = kPrev && kPrev.cut13 ? ((k.cut13 - kPrev.cut13) / kPrev.cut13 * 100).toFixed(1) : null;
  const varClass = varCut > 0 ? 'pos' : (varCut < 0 ? 'neg' : 'zero');
  const varSigno = varCut > 0 ? '+' : '';

  let html = '';

  html += `<div class="kpi-grid fade-in">
    <div class="kpi-card">
      <div class="label">Plantilla ${mes.corto}</div>
      <div class="value">${fmt.format(k.total)}</div>
      <div class="sub">Registros en el HC</div>
    </div>
    <div class="kpi-card verde">
      <div class="label">Activos</div>
      <div class="value">${fmt.format(k.activos)}</div>
      <div class="sub">${(k.activos / k.total * 100).toFixed(1)}% de la plantilla</div>
    </div>
    <div class="kpi-card celeste">
      <div class="label">Acumulado al ${CORTE}/${mes.m}</div>
      <div class="value">${fmt.format(k.cut13)}</div>
      <div class="sub">Ingresos con corte al d\u00eda ${CORTE}</div>
    </div>
    ${kPrev ? `<div class="kpi-card ${varCut >= 0 ? 'verde' : 'rojo'}">
      <div class="label">vs ${prev}</div>
      <div class="value">${varSigno}${varCut}</div>
      <div class="sub">${varPct}% en acumulado al ${CORTE}</div>
    </div>` : ''}
  </div>`;

  const pctAct = (k.activos / k.total * 100).toFixed(1);
  const resumen = [];
  resumen.push(`La plantilla de <b>${mes.corto} ${mes.y}</b> se compone de <b>${fmt.format(k.total)} personas</b> (${pctAct}% activas). El acumulado de ingresos al d\u00eda ${CORTE} es de <b>${fmt.format(k.cut13)} personas</b>` +
    (kPrev ? `, lo que representa una variaci\u00f3n de <b>${varCut >= 0 ? '+' : ''}${varCut}</b> (${varPct}%) frente a los ${fmt.format(kPrev.cut13)} del ${prev}.` : '.'));

  html += `<div class="section">
    <div class="expl fade-in">
      <span class="tag">Resumen ejecutivo</span><br>${resumen.join(' ')}
    </div>
  </div>`;

  const campVentas = campOrden.filter(c => /ventas/i.test(c) || /reno|pre|s2s/i.test(c));
  const totalVentas = campVentas.reduce((a, c) => a + camps.get(c).length, 0);
  const pctVentas = totalVentas ? (totalVentas / k.total * 100).toFixed(1) : null;

  html += `<div class="section">
    <div class="section-title">Resultados por campa&ntilde;a</div>
    <div class="chart-row">
      <div class="chart-box"><h3>Plantilla por campa&ntilde;a</h3><canvas id="chCamp"></canvas></div>
      <div class="chart-box"><h3>${prev ? 'Comparativo corte al ' + CORTE + ': ' + prev.replace('2026','') + ' vs ' + nombre.replace('2026','') : 'Composici\u00f3n'}</h3><canvas id="chComp"></canvas></div>
      <div class="expl">
        <span class="tag">Lectura del gr\u00e1fico</span><br>
        La campa\u00f1a con mayor dotaci\u00f3n en ${mes.corto} es <b>${campOrden[0]}</b> con ${fmt.format(camps.get(campOrden[0]).length)} personas
        (${(camps.get(campOrden[0]).length / k.total * 100).toFixed(1)}% de la plantilla), seguida de <b>${campOrden[1]}</b> con ${fmt.format(camps.get(campOrden[1]).length)}.
        ${pctVentas !== null ? `El esfuerzo comercial (campa\u00f1as de ventas) concentra el <b>${pctVentas}%</b> de la dotaci\u00f3n.` : ''}
        ${prevRows ? `El comparativo usa el <b>acumulado al d\u00eda ${CORTE}</b> de cada mes (solo personas ingresadas a esa fecha), tal como se pidi\u00f3: color claro = ${prev} al 13, azul = ${nombre} al 13. El saldo neto del mes es <b>${varCut >= 0 ? '+' : ''}${varCut} persona(s)</b>.` : ''}
      </div>
    </div>
  </div>`;

  html += `<div class="section"><div class="section-title">Desglose por campa&ntilde;a</div><div class="campana-grid">`;
  campOrden.forEach(c => {
    const rowsC = camps.get(c);
    const cargo = new Map(), estado = new Map();
    rowsC.forEach(r => {
      cargo.set(r.r, (cargo.get(r.r) || 0) + 1);
      estado.set(r.e, (estado.get(r.e) || 0) + 1);
    });
    const diff = prevRows ? rowsC.length - campañas(prevRows).get(c).length : null;
    const dCls = diff > 0 ? 'pos' : (diff < 0 ? 'neg' : 'zero');
    const dSigno = diff > 0 ? '+' : '';
    html += `<div class="campana-box">
      <h4>${c} <span class="badge">${rowsC.length} pers.</span>
        ${diff !== null ? `<span class="badge2 ${dCls}">vs ${prev}: ${dSigno}${diff}</span>` : ''}</h4>
      <div class="mini-title">Cargos</div>
      <table class="mini-table"><thead><tr><th>Cargo</th><th>#</th></tr></thead><tbody>
        ${[...cargo.entries()].sort((a,b) => b[1] - a[1]).map(([k, v]) => `<tr><td>${k}</td><td>${v}</td></tr>`).join('')}
      </tbody></table>
      <div class="mini-title">Estado</div>
      <table class="mini-table"><thead><tr><th>Estado</th><th>#</th></tr></thead><tbody>
        ${[...estado.entries()].sort((a,b) => b[1] - a[1]).map(([k, v]) => `<tr><td>${k}</td><td>${v}</td></tr>`).join('')}
      </tbody></table>
    </div>`;
  });
  html += `</div>
  <div class="expl fade-in" style="margin-top:14px">
    <span class="tag">Lectura del detalle</span><br>
    Cada tarjeta desglosa la campa\u00f1a por <b>cargo</b> (distribuci\u00f3n de roles) y por <b>estado</b>
    (activos vs. capacitaci\u00f3n). La columna <i>vs ${prev}</i> muestra el movimiento neto de personal
    respecto al mes anterior: verde indica crecimiento de dotaci\u00f3n y rojo, reducci\u00f3n.
  </div>
  </div>`;

  if (cuartiles) {
    const qTop = cuartiles.q[0].n;
    const pctQ1 = (qTop / cuartiles.n * 100).toFixed(1);
    const qUltimo = cuartiles.q[3];
    html += `<div class="section">
      <div class="section-title">An&aacute;lisis por cuartiles (antig&uuml;edad al corte ${CORTE}/${mes.m})</div>
      <div class="chart-row">
        <div class="chart-box"><h3>Distribuci&oacute;n por cuartil</h3><canvas id="chCuart"></canvas></div>
        <div class="chart-box"><h3>Cuartiles &mdash; detalle</h3>
          <table><thead><tr><th>Cuartil</th><th>Personas</th><th>Rango</th><th>D\u00edas</th></tr></thead><tbody>
          ${cuartiles.q.map(q => `<tr><td>${q.nombre}</td><td>${q.n}</td><td>${q.desde} &ndash; ${q.hasta} meses</td><td>${q.diasMin} &ndash; ${q.diasMax}</td></tr>`).join('')}
          </tbody></table>
        </div>
        <div class="expl">
          <span class="tag">Lectura del gr\u00e1fico</span><br>
          La plantilla se dividi\u00f3 en 4 cuartiles por <b>antig\u00fcedad</b> (meses desde la fecha de ingreso al corte).
          El <b>Q1</b> agrupa al ${pctQ1}% del personal (${qTop} personas) con menos experiencia en campa\u00f1a,
          mientras el <b>Q4</b> concentra al ${fmt.format(qUltimo.n)}% m\u00e1s antiguo (${qUltimo.desde} &ndash; ${qUltimo.hasta} meses),
          que representa la base de experiencia operativa del equipo.
        </div>
      </div>
    </div>`;
  }

  const sup = new Map();
  rows.forEach(r => sup.set(r.s, (sup.get(r.s) || 0) + 1));
  const supMax = [...sup.entries()].sort((a, b) => b[1] - a[1])[0];
  const supTotal = sup.size;
  html += `<div class="section">
    <div class="section-title">Plantilla por supervisor</div>
    <div class="chart-row">
      <div class="chart-box"><h3>Personas a cargo</h3><canvas id="chSup"></canvas></div>
      <div class="expl">
        <span class="tag">Lectura del gr\u00e1fico</span><br>
        La distribuci\u00f3n de la carga operativa recae en <b>${supMax[0]}</b> con ${fmt.format(supMax[1])} personas a cargo
        (${(supMax[1] / k.total * 100).toFixed(1)}% de la plantilla), entre ${fmt.format(supTotal)} supervisores.
        Esto permite dimensionar el alcance de supervisi\u00f3n y el balance de carga entre l\u00edderes.
      </div>
    </div>
  </div>`;

  document.getElementById('contenido').innerHTML = html;

  // charts
  const labels = campOrden;
  const datos = labels.map(c => camps.get(c).length);
  chart('chCamp', 'bar', {
    labels, datos,
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, grid: { display: false } }, x: { grid: { display: false } } } }
  });

  if (prevRows) {
    const prevCut = cutByCampana(prevRows, parseMes(prev));
    const curCut = cutByCampana(rows, mes);
    chart('chComp', 'bar', {
      labels,
      datasets: [
        { label: prev + ' (al 13)', data: labels.map(c => prevCut.get(c) || 0), backgroundColor: '#A8C7E7', borderRadius: 4 },
        { label: nombre + ' (al 13)', data: labels.map(c => curCut.get(c) || 0), backgroundColor: '#2E75B6', borderRadius: 4 }
      ],
      options: { plugins: { legend: { position: 'bottom' } }, scales: { y: { beginAtZero: true, grid: { display: false } }, x: { grid: { display: false } } } }
    });
  }

  if (cuartiles) {
    chart('chCuart', 'doughnut', {
      labels: cuartiles.q.map(q => q.nombre),
      datos: cuartiles.q.map(q => q.n),
      options: { cutout: '58%', plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: ctx => ctx.label + ': ' + ctx.parsed + ' personas' } } } }
    });
  }

  chart('chSup', 'bar', {
    labels: [...sup.keys()].sort((a, b) => sup.get(b) - sup.get(a)),
    datasets: [{ label: 'Personas', data: [...sup.keys()].sort((a, b) => sup.get(b) - sup.get(a)).map(s => sup.get(s)), backgroundColor: '#2E75B6', borderRadius: 4 }],
    options: { indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true, grid: { display: false } }, y: { grid: { display: false } } } }
  });

  document.getElementById('contenido').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function chart(id, tipo, cfg) {
  const el = document.getElementById(id);
  if (!el) return;
  if (charts[id]) charts[id].destroy();
  const data = cfg.datos
    ? { labels: cfg.labels, datasets: [{ data: cfg.datos, backgroundColor: ['#1F4E79', '#2E75B6', '#5B9BD5', '#A8C7E7', '#8DB4E2', '#6B9BCF', '#35689E', '#163B5E'], borderWidth: 0 }] }
    : { labels: cfg.labels, datasets: cfg.datasets };
  charts[id] = new Chart(el, { type: tipo, data, options: cfg.options || {} });
}

renderMenu();
selectMes(MESES[MESES.length - 1]);
</script>
</body>
</html>'''

html = html.replace('__DATA__', data_json)

with open('dashboard_resultados.html', 'w', encoding='utf-8') as f:
    f.write(html)

print('Dashboard generado con meses:', months)

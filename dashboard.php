<?php
$db = __DIR__ . '/Lixi/Cobros/cobros.db';
if (!file_exists($db)) die("Base de datos no encontrada. Ejecuta primero el script de importacion.");

$dsn = 'sqlite:' . $db;
$pdo = new PDO($dsn);
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

$dateFrom = $_GET['desde'] ?? '';
$dateTo = $_GET['hasta'] ?? '';
$pCall = [];
$pSale = [];

if ($dateFrom !== '' && $dateTo !== '') {
    $pCall = [$dateFrom, $dateTo];
    $pSale = [$dateFrom, $dateTo];
}

// ---- QUERIES ----
function prep($pdo, $sql, $params) { $s=$pdo->prepare($sql); $s->execute($params); return $s; }
$p = $pSale;

$totalCuentas = prep($pdo, "SELECT COUNT(*) FROM list_51124" . ($dateFrom?" WHERE date(entry_date) BETWEEN ? AND ?":""), $p)->fetchColumn() ?: 0;
$alDia       = prep($pdo, "SELECT COUNT(*) FROM list_51124 WHERE estatus='Al dia'"    .($dateFrom?" AND date(entry_date) BETWEEN ? AND ?":""), $p)->fetchColumn() ?: 0;
$enMora      = prep($pdo, "SELECT COUNT(*) FROM list_51124 WHERE estatus='Mora'"      .($dateFrom?" AND date(entry_date) BETWEEN ? AND ?":""), $p)->fetchColumn() ?: 0;
$ceroPago    = prep($pdo, "SELECT COUNT(*) FROM list_51124 WHERE estatus='Cero Pago'" .($dateFrom?" AND date(entry_date) BETWEEN ? AND ?":""), $p)->fetchColumn() ?: 0;
$moraReal    = prep($pdo, "SELECT COUNT(*) FROM list_51124 WHERE estatus='Mora Real'" .($dateFrom?" AND date(entry_date) BETWEEN ? AND ?":""), $p)->fetchColumn() ?: 0;
$totalMora = $enMora + $ceroPago + $moraReal;

$ventas    = prep($pdo, "SELECT ROUND(SUM(CAST(REPLACE(precio,',','') AS REAL)),2) FROM list_51124 WHERE CAST(REPLACE(precio,',','') AS REAL) > 0" .($dateFrom?" AND date(entry_date) BETWEEN ? AND ?":""), $p)->fetchColumn() ?: 0;
$promDias  = prep($pdo, "SELECT ROUND(AVG(CAST(dias_atraso AS INTEGER)),0) FROM list_51124 WHERE CAST(dias_atraso AS INTEGER) > 0 AND estatus IN ('Mora','Mora Real','Cero Pago')" .($dateFrom?" AND date(entry_date) BETWEEN ? AND ?":""), $p)->fetchColumn() ?: 0;
$statusDist = prep($pdo, "SELECT estatus, COUNT(*) FROM list_51124 WHERE estatus != ''" .($dateFrom?" AND date(entry_date) BETWEEN ? AND ?":"")." GROUP BY estatus ORDER BY 2 DESC", $p)->fetchAll(PDO::FETCH_KEY_PAIR) ?: ['Sin datos' => 0];
$topDeud   = prep($pdo, "SELECT first_name||' '||last_name, CAST(dias_atraso AS INTEGER), CAST(precio AS REAL) FROM list_51124 WHERE CAST(dias_atraso AS INTEGER) > 0 AND estatus IN ('Mora','Mora Real','Cero Pago')" .($dateFrom?" AND date(entry_date) BETWEEN ? AND ?":"")." ORDER BY CAST(dias_atraso AS INTEGER) DESC LIMIT 10", $p)->fetchAll(PDO::FETCH_NUM) ?: [['Sin deudores', 0, 0]];
$marcas    = prep($pdo, "SELECT marca, COUNT(*) FROM list_51124 WHERE marca != ''" .($dateFrom?" AND date(entry_date) BETWEEN ? AND ?":"")." GROUP BY marca ORDER BY 2 DESC LIMIT 10", $p)->fetchAll(PDO::FETCH_KEY_PAIR) ?: ['Sin datos' => 0];

$totalLlamadas = prep($pdo, "SELECT COUNT(*) FROM call_report" .($dateFrom?" WHERE date(call_date) BETWEEN ? AND ?":""), $pCall)->fetchColumn() ?: 0;
$callStatus    = prep($pdo, "SELECT status_name, COUNT(*) FROM call_report" .($dateFrom?" WHERE date(call_date) BETWEEN ? AND ?":"")." GROUP BY status_name ORDER BY 2 DESC", $pCall)->fetchAll(PDO::FETCH_KEY_PAIR);
if (!$callStatus) $callStatus = ['Sin datos' => 0];

?>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard Cobros</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;padding:20px;color:#333}
.header{background:linear-gradient(135deg,#1F4E79,#2E75B6);color:#fff;padding:25px 30px;border-radius:12px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center}
.header h1{font-size:26px}
.filtros{background:#fff;padding:14px 20px;border-radius:10px;margin-bottom:20px;display:flex;gap:14px;align-items:end;flex-wrap:wrap;box-shadow:0 2px 6px rgba(0,0,0,0.05)}
.filtros label{font-size:12px;color:#666;display:block;margin-bottom:2px}
.filtros select,.filtros input{border:1px solid #ccc;border-radius:6px;padding:7px 12px;font-size:13px}
.filtros .btn{background:#1F4E79;color:#fff;border:none;padding:8px 20px;border-radius:6px;cursor:pointer;font-size:13px}
.filtros .btn:hover{background:#14365c}
.filtros .btn2{background:#6c757d}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(165px,1fr));gap:14px;margin-bottom:24px}
.kpi-card{background:#fff;border-radius:10px;padding:16px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.06);border-left:4px solid #2E75B6}
.kpi-card .label{font-size:11px;text-transform:uppercase;color:#888;letter-spacing:.5px}
.kpi-card .value{font-size:26px;font-weight:700;color:#1F4E79;margin:3px 0}
.kpi-card .sub{font-size:11px;color:#999}
.kpi-card.danger{border-left-color:#c0392b}.kpi-card.danger .value{color:#c0392b}
.kpi-card.success{border-left-color:#27ae60}.kpi-card.success .value{color:#27ae60}
.kpi-card.warning{border-left-color:#f39c12}
.chart-row{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:24px}
.chart-box{background:#fff;border-radius:10px;padding:18px;box-shadow:0 2px 8px rgba(0,0,0,0.06)}
.chart-box h3{font-size:15px;margin-bottom:12px;color:#1F4E79}
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:#1F4E79;color:#fff;padding:10px 8px;text-align:center;font-weight:600}
td{padding:8px;border-bottom:1px solid #e0e0e0;text-align:center}
tr:nth-child(even){background:#f7f9fc}
tr:hover{background:#e8f0fa}
.section-title{font-size:18px;font-weight:700;color:#1F4E79;margin:20px 0 12px}
.top-mora{color:#c0392b;font-weight:600}
@media(max-width:768px){.chart-row{grid-template-columns:1fr}}
</style>
</head>
<body>

<div class="header">
  <div><h1>REPORTE DE COBROS</h1><span>Dashboard ejecutivo &mdash; <?=date('d/m/Y')?></span></div>
  <div><a href="Reporte_Cobros.xlsx" style="color:#fff;background:rgba(255,255,255,0.2);padding:8px 16px;border-radius:6px;text-decoration:none">Descargar Excel</a></div>
</div>

<form class="filtros" method="get">
  <div>
    <label>Desde</label>
    <input type="date" name="desde" value="<?=htmlspecialchars($dateFrom)?>">
  </div>
  <div>
    <label>Hasta</label>
    <input type="date" name="hasta" value="<?=htmlspecialchars($dateTo)?>">
  </div>
  <div><button class="btn" type="submit">Filtrar</button></div>
  <div><a href="?" class="btn btn2" style="display:inline-block;text-decoration:none">Limpiar</a></div>
</form>

<div class="kpi-grid">
  <div class="kpi-card"><div class="label">Total Cuentas</div><div class="value"><?=number_format($totalCuentas)?></div><div class="sub">Clientes registrados</div></div>
  <div class="kpi-card success"><div class="label">Al Dia</div><div class="value"><?=number_format($alDia)?></div><div class="sub"><?=$totalCuentas?round($alDia/$totalCuentas*100,1):0?>% del total</div></div>
  <div class="kpi-card danger"><div class="label">En Mora</div><div class="value"><?=number_format($totalMora)?></div><div class="sub"><?=$totalCuentas?round($totalMora/$totalCuentas*100,1):0?>% del total</div></div>
  <div class="kpi-card"><div class="label">Ventas Totales</div><div class="value">$<?=number_format($ventas,0)?></div><div class="sub">Suma financiada</div></div>
  <div class="kpi-card warning"><div class="label">Llamadas</div><div class="value"><?=number_format($totalLlamadas)?></div><div class="sub"><?=$dateFrom?"Filtrado":"Total historico"?></div></div>
  <div class="kpi-card danger"><div class="label">Dias Prom Mora</div><div class="value"><?=number_format($promDias)?></div><div class="sub">Promedio atraso</div></div>
</div>

<div class="chart-row">
  <div class="chart-box">
    <h3>Estado de Cuentas</h3>
    <canvas id="pieStatus"></canvas>
  </div>
  <div class="chart-box">
    <h3>Resultado de Llamadas</h3>
    <canvas id="barLlamadas"></canvas>
  </div>
</div>

<div class="section-title">TOP 10 DEUDORES (mayor atraso)</div>
<table>
  <thead><tr><th>#</th><th>Nombre</th><th>Dias Atraso</th><th>Monto</th></tr></thead>
  <tbody>
    <?php $i=1; foreach($topDeud as $d): ?>
    <tr><td><?=$i++?></td><td class="top-mora"><?=htmlspecialchars($d[0])?></td><td><?=number_format($d[1])?></td><td>$<?=number_format($d[2],0)?></td></tr>
    <?php endforeach; ?>
  </tbody>
</table>

<div class="section-title" style="margin-top:24px">Top 10 Marcas</div>
<div class="chart-box" style="max-width:600px">
  <canvas id="barMarcas"></canvas>
</div>

<script>
const statusData = <?=json_encode(array_values($statusDist))?>;
const statusLabels = <?=json_encode(array_keys($statusDist))?>;
const callLabels = <?=json_encode(array_keys($callStatus))?>;
const callData = <?=json_encode(array_values($callStatus))?>;
const totalCuentas = <?=$totalCuentas ?: 1?>;
const marcasLabels = <?=json_encode(array_keys($marcas))?>;
const marcasData = <?=json_encode(array_values($marcas))?>;

new Chart(document.getElementById('pieStatus'), {
  type: 'doughnut',
  data: { labels: statusLabels, datasets: [{ data: statusData, backgroundColor: ['#27ae60','#e74c3c','#f39c12','#95a5a6','#3498db'], borderWidth: 0 }] },
  options: { plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: ctx => ctx.parsed + ' (' + (ctx.parsed/totalCuentas*100).toFixed(1) + '%)' } } }, cutout: '55%' }
});

new Chart(document.getElementById('barLlamadas'), {
  type: 'bar',
  data: { labels: callLabels, datasets: [{ label: 'Llamadas', data: callData, backgroundColor: '#2E75B6', borderRadius: 4 }] },
  options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } }, y: { grid: { display: false } } } }
});

new Chart(document.getElementById('barMarcas'), {
  type: 'bar',
  data: { labels: marcasLabels, datasets: [{ label: 'Cantidad', data: marcasData, backgroundColor: '#1F4E79', borderRadius: 4 }] },
  options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } }, y: { grid: { display: false } } } }
});
</script>
</body>
</html>

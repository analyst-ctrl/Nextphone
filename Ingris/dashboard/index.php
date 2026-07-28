<?php
/**
 * INGRIS DASHBOARD - Calidad Nextphone & OJT
 * Dashboard interactivo con PHP + Chart.js + SQLite
 */
require_once __DIR__ . '/config.php';

// Verificar que la DB existe
if (!file_exists(DB_PATH)) {
    $db_error = "⚠️ Base de datos no encontrada. Ejecuta primero: python scripts/extract_data.py";
} else {
    $db = getDB();
    $db_ok = true;
    $count = $db->query("SELECT COUNT(*) as c FROM evaluaciones")->fetch();
    $db_total = $count['c'];
}
?><!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Calidad - Ingris</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="assets/css/style.css">
</head>
<body>
    <div class="app">
        <!-- SIDEBAR -->
        <nav class="sidebar" id="sidebar">
            <div class="sidebar-logo">
                <span class="logo-icon">📊</span>
                <span class="logo-text">Calidad Ingris</span>
            </div>
            <ul class="nav-list">
                <li><a href="#" class="nav-link active" data-section="resumen"><span class="nav-icon">📈</span> Resumen</a></li>
                <li><a href="#" class="nav-link" data-section="supervisores"><span class="nav-icon">👔</span> Supervisores</a></li>
                <li><a href="#" class="nav-link" data-section="agentes"><span class="nav-icon">👤</span> Agentes</a></li>
                <li><a href="#" class="nav-link" data-section="campanas"><span class="nav-icon">🎯</span> Campañas</a></li>
                <li><a href="#" class="nav-link" data-section="atributos"><span class="nav-icon">📋</span> Atributos</a></li>
                <li><a href="#" class="nav-link" data-section="historico"><span class="nav-icon">📅</span> Histórico</a></li>
                <li><a href="#" class="nav-link" data-section="criticas"><span class="nav-icon">⚠️</span> Críticas</a></li>
            </ul>
            <div class="sidebar-footer">
                <span>Nextphone Analytics</span>
            </div>
        </nav>

        <!-- MAIN -->
        <main class="main-content">
            <!-- HEADER -->
            <header class="topbar">
                <button class="menu-toggle" id="menuToggle">☰</button>
                <h1 class="page-title">Dashboard de Calidad</h1>
                <div class="header-info">
                    <span class="badge-db" id="dbInfo"><?= isset($db_total) ? "{$db_total} eval." : 'Sin datos' ?></span>
                    <span class="badge-date" id="currentDate"></span>
                </div>
            </header>

            <!-- FILTROS -->
            <section class="filters-bar" id="filtersBar">
                <div class="filter-group">
                    <label>Supervisor</label>
                    <select id="filterSupervisor" onchange="applyFilters()">
                        <option value="todos">Todos</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>Proyecto</label>
                    <select id="filterProyecto" onchange="applyFilters()">
                        <option value="todos">Todos</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>Mes</label>
                    <select id="filterMes" onchange="applyFilters()">
                        <option value="todos">Todos</option>
                    </select>
                </div>
                <button class="btn-refresh" onclick="refreshData()" title="Actualizar datos">🔄</button>
            </section>

            <?php if (isset($db_error)): ?>
            <div class="error-banner"><?= $db_error ?></div>
            <?php endif; ?>

            <!-- CONTENIDO -->
            <div class="content-wrapper">

                <!-- SECCIÓN: Resumen -->
                <section id="section-resumen" class="dashboard-section active">
                    <div class="kpi-grid" id="kpiGrid">
                        <div class="kpi-card"><div class="kpi-value" id="kpiScore">--</div><div class="kpi-label">Score General</div><div class="kpi-trend" id="kpiScoreTrend"></div></div>
                        <div class="kpi-card"><div class="kpi-value" id="kpiEvaluaciones">--</div><div class="kpi-label">Evaluaciones</div></div>
                        <div class="kpi-card"><div class="kpi-value" id="kpiAgentes">--</div><div class="kpi-label">Agentes</div></div>
                        <div class="kpi-card"><div class="kpi-value" id="kpiSupervisores">--</div><div class="kpi-label">Supervisores</div></div>
                        <div class="kpi-card critica"><div class="kpi-value" id="kpiCriticas">--</div><div class="kpi-label">Críticas</div><div class="kpi-trend" id="kpiCriticasPct"></div></div>
                    </div>
                    <div class="chart-grid-2">
                        <div class="card"><div class="card-header"><h3>Score por Componente</h3></div><div class="card-body"><canvas id="chartComponentes"></canvas></div></div>
                        <div class="card"><div class="card-header"><h3>Distribución de Scores</h3></div><div class="card-body"><div style="text-align:center;padding:40px;color:var(--text-muted);font-size:14px;">📊 <br><strong>Próximamente</strong><br><small>Distribución por rangos con datos reales</small></div></div></div>
                    </div>
                </section>

                <!-- SECCIÓN: Supervisores -->
                <section id="section-supervisores" class="dashboard-section">
                    <div class="card full">
                        <div class="card-header"><h3>Ranking de Supervisores</h3></div>
                        <div class="card-body">
                            <div class="table-responsive">
                                <table class="data-table" id="tablaSupervisores">
                                    <thead>
                                        <tr>
                                            <th>#</th><th>Supervisor</th><th>Score</th><th>PENC</th><th>PECUF</th><th>PECNE</th><th>PECCUM</th>
                                            <th>Eval.</th><th>Agentes</th><th>Críticas</th><th>% Crít.</th>
                                        </tr>
                                    </thead>
                                    <tbody id="tbodySupervisores"></tbody>
                                </table>
                            </div>
                            <div class="chart-container"><canvas id="chartSupervisores"></canvas></div>
                        </div>
                    </div>
                </section>

                <!-- SECCIÓN: Agentes -->
                <section id="section-agentes" class="dashboard-section">
                    <div class="card full">
                        <div class="card-header">
                            <h3>Ranking de Agentes</h3>
                            <div class="card-actions">
                                <select id="ordenAgentes" onchange="cargarAgentes()">
                                    <option value="score_desc">Score ↓</option>
                                    <option value="score_asc">Score ↑</option>
                                    <option value="nombre">Nombre A-Z</option>
                                    <option value="evaluaciones">Más evaluados</option>
                                </select>
                                <input type="text" id="buscarAgente" placeholder="🔍 Buscar agente..." oninput="cargarAgentes()">
                            </div>
                        </div>
                        <div class="card-body">
                            <div class="table-responsive">
                                <table class="data-table" id="tablaAgentes">
                                    <thead>
                                        <tr>
                                            <th>#</th><th>Agente</th><th>Supervisor</th><th>Score</th><th>Cuartil</th><th>Cumple</th>
                                            <th>PENC</th><th>PECUF</th><th>PECNE</th><th>PECCUM</th><th>Eval.</th><th>Crít.</th>
                                        </tr>
                                    </thead>
                                    <tbody id="tbodyAgentes"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </section>

                <!-- SECCIÓN: Campañas -->
                <section id="section-campanas" class="dashboard-section">
                    <div class="chart-grid-2">
                        <div class="card"><div class="card-header"><h3>Score por Campaña</h3></div><div class="card-body"><canvas id="chartCampanas"></canvas></div></div>
                        <div class="card"><div class="card-header"><h3>Comparativa Componentes</h3></div><div class="card-body"><canvas id="chartCampanasRadar"></canvas></div></div>
                    </div>
                    <div class="card full mt-2">
                        <div class="card-header"><h3>Detalle por Campaña</h3></div>
                        <div class="card-body">
                            <div class="table-responsive">
                                <table class="data-table" id="tablaCampanas">
                                    <thead><tr><th>Campaña</th><th>Score</th><th>PENC</th><th>PECUF</th><th>PECNE</th><th>PECCUM</th><th>Eval.</th><th>Agentes</th><th>Crít.</th></tr></thead>
                                    <tbody id="tbodyCampanas"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </section>

                <!-- SECCIÓN: Atributos -->
                <section id="section-atributos" class="dashboard-section">
                    <div class="chart-grid-2">
                        <div class="card"><div class="card-header"><h3>Precisión por Tipo de Error</h3></div><div class="card-body"><canvas id="chartAtributos"></canvas></div></div>
                        <div class="card"><div class="card-header"><h3>Pesos vs Precisión</h3></div><div class="card-body"><canvas id="chartPesos"></canvas></div></div>
                    </div>
                    <div class="card full mt-2">
                        <div class="card-header"><h3>Factores de Calidad</h3></div>
                        <div class="card-body">
                            <div class="info-grid" id="factoresGrid">
                                <div class="factor-card penc"><div class="factor-sigla">PENC</div><div class="factor-nombre">Error No Crítico</div><div class="factor-valor" id="factorPENC">--</div></div>
                                <div class="factor-card pecuf"><div class="factor-sigla">PECUF</div><div class="factor-nombre">Error Crítico Usuario Final</div><div class="factor-valor" id="factorPECUF">--</div></div>
                                <div class="factor-card pecne"><div class="factor-sigla">PECNE</div><div class="factor-nombre">Error Crítico Negocio</div><div class="factor-valor" id="factorPECNE">--</div></div>
                                <div class="factor-card peccum"><div class="factor-sigla">PECCUM</div><div class="factor-nombre">Error Crítico Cumplimiento</div><div class="factor-valor" id="factorPECCUM">--</div></div>
                            </div>
                        </div>
                    </div>
                </section>

                <!-- SECCIÓN: Histórico -->
                <section id="section-historico" class="dashboard-section">
                    <div class="card full">
                        <div class="card-header"><h3>Tendencia Mensual de Scores</h3></div>
                        <div class="card-body"><canvas id="chartHistorico"></canvas></div>
                    </div>
                    <div class="card full mt-2">
                        <div class="card-header"><h3>Evolución por Componente</h3></div>
                        <div class="card-body"><canvas id="chartHistoricoComponentes"></canvas></div>
                    </div>
                </section>

                <!-- SECCIÓN: Críticas -->
                <section id="section-criticas" class="dashboard-section">
                    <div class="card full">
                        <div class="card-header"><h3>⚠️ Evaluaciones Críticas (Score &lt; 90%)</h3></div>
                        <div class="card-body">
                            <div class="table-responsive">
                                <table class="data-table" id="tablaCriticas">
                                    <thead><tr><th>Agente</th><th>Supervisor</th><th>Fecha</th><th>Proyecto</th><th>Score</th><th>PENC</th><th>PECUF</th><th>PECNE</th><th>PECCUM</th></tr></thead>
                                    <tbody id="tbodyCriticas"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </section>

            </div>

            <!-- FOOTER -->
            <footer class="footer">
                <span>Ingris Dashboard v1.0</span>
                <span>Datos corregidos: Fórmula F17*1/G17/100 → F17/G17 ✅</span>
                <span id="footerDate"></span>
            </footer>
        </main>
    </div>

    <script src="assets/js/app.js"></script>
    <script>
        // Inicializar dashboard al cargar
        document.addEventListener('DOMContentLoaded', () => {
            cargarFiltros();
            cargarResumen();
        });
    </script>
</body>
</html>

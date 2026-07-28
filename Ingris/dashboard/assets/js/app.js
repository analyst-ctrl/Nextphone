/**
 * INGRIS DASHBOARD - Frontend Application
 * Interactive charts and data visualizations for Quality Control
 */

// ═══ STATE ═══════════════════════════════════════════════
const STATE = {
    charts: {},
    filters: {},
    data: {}
};

// ═══ UTILITIES ═══════════════════════════════════════════
const API = 'api.php';
let chartInstances = {};

function $(id) { return document.getElementById(id); }

async function fetchAPI(params) {
    const qs = new URLSearchParams(params).toString();
    const resp = await fetch(`${API}?${qs}`);
    return resp.json();
}

function getQuery() {
    const sup = $('filterSupervisor')?.value || 'todos';
    const proy = $('filterProyecto')?.value || 'todos';
    const mes = $('filterMes')?.value || 'todos';
    const q = {};
    if (sup !== 'todos') q.supervisor = sup;
    if (proy !== 'todos') q.proyecto = proy;
    if (mes !== 'todos') q.mes = mes;
    return q;
}

function scoreClass(val) {
    if (val >= 85) return 'score-high';
    if (val >= 70) return 'score-mid';
    return 'score-low';
}

function getColor(val) {
    if (val >= 85) return '#34d399';
    if (val >= 70) return '#fbbf24';
    return '#f87171';
}

// ═══ MENU ═══════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
    // Menu toggle
    $('menuToggle').addEventListener('click', () => {
        document.getElementById('sidebar').classList.toggle('open');
    });

    // Navigation
    document.querySelectorAll('.nav-link').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const section = link.dataset.section;
            showSection(section);
            if (window.innerWidth <= 768) {
                document.getElementById('sidebar').classList.remove('open');
            }
        });
    });

    // Date
    const now = new Date();
    $('currentDate').textContent = now.toLocaleDateString('es-PA', { month: 'long', year: 'numeric' });
    $('footerDate').textContent = `Última actualización: ${now.toLocaleDateString('es-PA')} ${now.toLocaleTimeString('es-PA')}`;
});

function showSection(id) {
    document.querySelectorAll('.dashboard-section').forEach(s => s.classList.remove('active'));
    document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
    $(`section-${id}`)?.classList.add('active');
    document.querySelector(`[data-section="${id}"]`)?.classList.add('active');
    
    // Cargar datos de la sección
    switch(id) {
        case 'supervisores': cargarSupervisores(); break;
        case 'agentes': cargarAgentes(); break;
        case 'campanas': cargarCampanas(); break;
        case 'atributos': cargarAtributos(); break;
        case 'historico': cargarHistorico(); break;
        case 'criticas': cargarCriticas(); break;
    }
}

// ═══ FILTROS ════════════════════════════════════════════
async function cargarFiltros() {
    const data = await fetchAPI({ action: 'filtros' });
    
    const selSup = $('filterSupervisor');
    const selProy = $('filterProyecto');
    const selMes = $('filterMes');
    
    data.supervisores.forEach(s => {
        selSup.innerHTML += `<option value="${s}">${s}</option>`;
    });
    data.proyectos.forEach(p => {
        selProy.innerHTML += `<option value="${p}">${p}</option>`;
    });
    data.meses.forEach(m => {
        selMes.innerHTML += `<option value="${m.mes}">${m.mes_texto}</option>`;
    });
}

async function applyFilters() {
    cargarResumen();
    // Recargar secciones visibles
    document.querySelectorAll('.dashboard-section.active').forEach(s => {
        const id = s.id.replace('section-', '');
        if (id !== 'resumen') showSection(id);
    });
}

async function refreshData() {
    const btn = document.querySelector('.btn-refresh');
    btn.style.transform = 'rotate(360deg)';
    btn.style.transition = 'transform 0.5s';
    await Promise.all([cargarResumen(), cargarFiltros()]);
    setTimeout(() => btn.style.transform = '', 500);
}

// ═══ SECCIÓN: RESUMEN ═══════════════════════════════════
async function cargarResumen() {
    const q = getQuery();
    q.action = 'resumen';
    const data = await fetchAPI(q);
    
    // Animar números
    animateValue('kpiScore', data.score_promedio, '%');
    animateValue('kpiEvaluaciones', data.total_evaluaciones, '');
    animateValue('kpiAgentes', data.total_agentes, '');
    animateValue('kpiSupervisores', data.total_supervisores, '');
    animateValue('kpiCriticas', data.total_criticas, '');
    
    $('kpiCriticasPct').textContent = `${data.pct_criticas}% del total`;
    
    // Si hay score de otros periodos, mostrar tendencia
    if (data.score_promedio) {
        const trend = data.score_promedio >= 85 ? '📈 Bueno' : data.score_promedio >= 70 ? '📊 Regular' : '📉 Crítico';
        $('kpiScoreTrend').textContent = trend;
    }
    
    // Gráfico de componentes
    renderComponentesChart(data);
    // renderDistribucionChart(data);  // Placeholder - requiere datos reales de distribucion
}

function animateValue(id, value, suffix) {
    const el = $(id);
    if (!el) return;
    el.textContent = (value || 0) + suffix;
}

function renderComponentesChart(data) {
    const ctx = $('chartComponentes');
    if (!ctx) return;
    
    if (chartInstances.componentes) chartInstances.componentes.destroy();
    
    chartInstances.componentes = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['PENC', 'PECUF', 'PECNE', 'PECCUM'],
            datasets: [{
                label: 'Precisión (%)',
                data: [data.penc_promedio, data.pecuf_promedio, data.pecne_promedio, data.peccum_promedio],
                backgroundColor: ['rgba(79,140,255,0.7)', 'rgba(167,139,250,0.7)', 'rgba(251,191,36,0.7)', 'rgba(248,113,113,0.7)'],
                borderColor: ['#4f8cff', '#a78bfa', '#fbbf24', '#f87171'],
                borderWidth: 2,
                borderRadius: 6,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
            },
            scales: {
                y: { beginAtZero: true, max: 100, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#8b8fa3' } },
                x: { grid: { display: false }, ticks: { color: '#8b8fa3' } }
            }
        }
    });
}

function renderDistribucionChart(data) {
    const ctx = $('chartDistribucion');
    if (!ctx) return;
    if (chartInstances.distribucion) chartInstances.distribucion.destroy();
    
    // Crear texto informativo de que la distribución requiere datos reales
    const parent = ctx.parentElement;
    if (!parent.querySelector('.chart-placeholder')) {
        const msg = document.createElement('div');
        msg.className = 'chart-placeholder';
        msg.style.cssText = 'text-align:center;padding:40px;color:var(--text-muted);font-size:13px;';
        msg.innerHTML = '📊 Distribución por rango de scores<br><small>Disponible cuando se conecten datos reales</small>';
        parent.appendChild(msg);
    }
    
    chartInstances.distribucion = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Crítico (<70)', 'Regular (70-79)', 'Bueno (80-89)', 'Excelente (90+)'],
            datasets: [{
                data: [25, 25, 25, 25],
                backgroundColor: ['rgba(248,113,113,0.8)', 'rgba(251,191,36,0.8)', 'rgba(79,140,255,0.8)', 'rgba(52,211,153,0.8)'],
                borderColor: ['#f87171', '#fbbf24', '#4f8cff', '#34d399'],
                borderWidth: 2,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#8b8fa3', padding: 12, font: { size: 11 } } }
            }
        }
    });
}

// ═══ SECCIÓN: SUPERVISORES ══════════════════════════════
async function cargarSupervisores() {
    const q = getQuery();
    q.action = 'supervisores';
    const data = await fetchAPI(q);
    
    if (!data.length) return;
    
    // Tabla
    const tbody = $('tbodySupervisores');
    tbody.innerHTML = data.map((s, i) => `
        <tr>
            <td>${i + 1}</td>
            <td><strong>${s.supervisor}</strong></td>
            <td class="${scoreClass(s.score_promedio)}">${s.score_promedio}%</td>
            <td>${s.penc}%</td>
            <td>${s.pecuf}%</td>
            <td>${s.pecne}%</td>
            <td>${s.peccum}%</td>
            <td>${s.evaluaciones}</td>
            <td>${s.agentes}</td>
            <td>${s.criticas}</td>
            <td class="${s.pct_criticas > 20 ? 'score-low' : 'score-high'}">${s.pct_criticas}%</td>
        </tr>
    `).join('');
    
    // Gráfico
    const ctx = $('chartSupervisores');
    if (!ctx) return;
    if (chartInstances.supervisores) chartInstances.supervisores.destroy();
    
    chartInstances.supervisores = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: data.map(s => s.supervisor.split(' ').slice(0,2).join(' ')),
            datasets: [
                { label: 'Score', data: data.map(s => s.score_promedio), backgroundColor: 'rgba(79,140,255,0.7)', borderColor: '#4f8cff', borderWidth: 2, borderRadius: 4 },
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: {
                legend: { display: false },
            },
            scales: {
                x: { beginAtZero: true, max: 100, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#8b8fa3' } },
                y: { grid: { display: false }, ticks: { color: '#8b8fa3', font: { size: 11 } } }
            }
        }
    });
}

// ═══ SECCIÓN: AGENTES ═══════════════════════════════════
async function cargarAgentes() {
    const q = getQuery();
    q.action = 'agentes';
    q.orden = $('ordenAgentes')?.value || 'score_desc';
    
    const data = await fetchAPI(q);
    if (!data.agentes?.length) return;
    
    const busqueda = ($('buscarAgente')?.value || '').toLowerCase();
    const agentes = busqueda 
        ? data.agentes.filter(a => a.agente.toLowerCase().includes(busqueda))
        : data.agentes;
    
    const tbody = $('tbodyAgentes');
    tbody.innerHTML = agentes.map((a, i) => `
        <tr>
            <td>${i + 1}</td>
            <td><strong>${a.agente}</strong></td>
            <td>${a.supervisor || '—'}</td>
            <td class="${scoreClass(a.score_promedio)}">${a.score_promedio}%</td>
            <td><span class="badge-cuartil ${(a.cuartil||'Q4').toLowerCase()}">${a.cuartil || '—'}</span></td>
            <td class="${a.cumple === 'Cumple' ? 'badge-cumple' : 'badge-nocumple'}">${a.cumple}</td>
            <td>${a.penc}%</td>
            <td>${a.pecuf}%</td>
            <td>${a.pecne}%</td>
            <td>${a.peccum}%</td>
            <td>${a.evaluaciones}</td>
            <td>${a.criticas}</td>
        </tr>
    `).join('');
}

// ═══ SECCIÓN: CAMPAÑAS ═════════════════════════════════
async function cargarCampanas() {
    const q = getQuery();
    q.action = 'campanas';
    const data = await fetchAPI(q);
    
    if (!data.length) return;
    
    // Tabla
    const tbody = $('tbodyCampanas');
    tbody.innerHTML = data.map(c => `
        <tr>
            <td><strong>${c.campana}</strong></td>
            <td class="${scoreClass(c.score_promedio)}">${c.score_promedio}%</td>
            <td>${c.penc}%</td>
            <td>${c.pecuf}%</td>
            <td>${c.pecne}%</td>
            <td>${c.peccum}%</td>
            <td>${c.evaluaciones}</td>
            <td>${c.agentes}</td>
            <td>${c.criticas}</td>
        </tr>
    `).join('');
    
    // Gráfico barras
    const ctx1 = $('chartCampanas');
    if (ctx1) {
        if (chartInstances.campanas) chartInstances.campanas.destroy();
        chartInstances.campanas = new Chart(ctx1, {
            type: 'bar',
            data: {
                labels: data.map(c => c.campana),
                datasets: [{
                    label: 'Score', data: data.map(c => c.score_promedio),
                    backgroundColor: ['rgba(79,140,255,0.7)', 'rgba(167,139,250,0.7)'],
                    borderColor: ['#4f8cff', '#a78bfa'], borderWidth: 2, borderRadius: 6,
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, max: 100, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#8b8fa3' } },
                    x: { grid: { display: false }, ticks: { color: '#8b8fa3' } }
                }
            }
        });
    }
    
    // Radar comparativo
    const ctx2 = $('chartCampanasRadar');
    if (ctx2) {
        if (chartInstances.campanasRadar) chartInstances.campanasRadar.destroy();
        chartInstances.campanasRadar = new Chart(ctx2, {
            type: 'radar',
            data: {
                labels: ['PENC', 'PECUF', 'PECNE', 'PECCUM'],
                datasets: data.map((c, i) => ({
                    label: c.campana,
                    data: [c.penc, c.pecuf, c.pecne, c.peccum],
                    backgroundColor: i === 0 ? 'rgba(79,140,255,0.1)' : 'rgba(167,139,250,0.1)',
                    borderColor: i === 0 ? '#4f8cff' : '#a78bfa',
                    borderWidth: 2,
                    pointBackgroundColor: i === 0 ? '#4f8cff' : '#a78bfa',
                }))
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                scales: {
                    r: { 
                        beginAtZero: true, max: 100,
                        grid: { color: 'rgba(255,255,255,0.1)' },
                        angleLines: { color: 'rgba(255,255,255,0.1)' },
                        ticks: { color: '#8b8fa3', backdropColor: 'transparent' },
                        pointLabels: { color: '#8b8fa3' }
                    }
                },
                plugins: {
                    legend: { labels: { color: '#8b8fa3', font: { size: 11 } } }
                }
            }
        });
    }
}

// ═══ SECCIÓN: ATRIBUTOS ════════════════════════════════
async function cargarAtributos() {
    const q = getQuery();
    q.action = 'atributos';
    const data = await fetchAPI(q);
    
    if (!data.length) return;
    
    // Actualizar tarjetas
    data.forEach(f => {
        const id = `factor${f.sigla}`;
        if ($(id)) $(id).textContent = f.valor + '%';
    });
    
    // Gráfico
    const ctx = $('chartAtributos');
    if (!ctx) return;
    if (chartInstances.atributos) chartInstances.atributos.destroy();
    
    chartInstances.atributos = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: data.map(f => f.sigla),
            datasets: [{
                label: 'Precisión',
                data: data.map(f => f.valor),
                backgroundColor: ['rgba(79,140,255,0.7)', 'rgba(167,139,250,0.7)', 'rgba(251,191,36,0.7)', 'rgba(248,113,113,0.7)'],
                borderColor: ['#4f8cff', '#a78bfa', '#fbbf24', '#f87171'],
                borderWidth: 2,
                borderRadius: 6,
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
            },
            scales: {
                y: { beginAtZero: true, max: 100, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#8b8fa3' } },
                x: { grid: { display: false }, ticks: { color: '#8b8fa3' } }
            }
        }
    });
    
    // Pesos vs Precisión
    const ctx2 = $('chartPesos');
    if (!ctx2) return;
    if (chartInstances.pesos) chartInstances.pesos.destroy();
    
    chartInstances.pesos = new Chart(ctx2, {
        type: 'radar',
        data: {
            labels: data.map(f => f.sigla),
            datasets: [
                { label: 'Precisión Actual', data: data.map(f => f.valor), borderColor: '#4f8cff', backgroundColor: 'rgba(79,140,255,0.1)', borderWidth: 2, pointBackgroundColor: '#4f8cff' },
                { label: 'Peso del Factor', data: data.map(() => 100), borderColor: '#8b8fa3', backgroundColor: 'rgba(139,143,163,0.05)', borderWidth: 1, borderDash: [5,5], pointBackgroundColor: '#8b8fa3' },
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                r: { beginAtZero: true, max: 100, grid: { color: 'rgba(255,255,255,0.1)' }, angleLines: { color: 'rgba(255,255,255,0.1)' }, ticks: { color: '#8b8fa3', backdropColor: 'transparent' }, pointLabels: { color: '#8b8fa3' } }
            },
            plugins: { legend: { labels: { color: '#8b8fa3', font: { size: 11 } } } }
        }
    });
}

// ═══ SECCIÓN: HISTÓRICO ════════════════════════════════
async function cargarHistorico() {
    const data = await fetchAPI({ action: 'historico' });
    
    if (!data.length) return;
    
    const ctx1 = $('chartHistorico');
    if (ctx1) {
        if (chartInstances.historico) chartInstances.historico.destroy();
        chartInstances.historico = new Chart(ctx1, {
            type: 'line',
            data: {
                labels: data.map(d => d.mes_texto),
                datasets: [{
                    label: 'Score General',
                    data: data.map(d => d.score_promedio),
                    borderColor: '#4f8cff',
                    backgroundColor: 'rgba(79,140,255,0.1)',
                    fill: true,
                    tension: 0.4,
                    pointBackgroundColor: '#4f8cff',
                    pointRadius: 5,
                    pointHoverRadius: 8,
                    borderWidth: 3,
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { labels: { color: '#8b8fa3' } } },
                scales: {
                    y: { beginAtZero: true, max: 100, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#8b8fa3' } },
                    x: { grid: { display: false }, ticks: { color: '#8b8fa3' } }
                }
            }
        });
    }
    
    // Componentes histórico
    const ctx2 = $('chartHistoricoComponentes');
    if (ctx2) {
        if (chartInstances.historicoComp) chartInstances.historicoComp.destroy();
        chartInstances.historicoComp = new Chart(ctx2, {
            type: 'line',
            data: {
                labels: data.map(d => d.mes_texto),
                datasets: [
                    { label: 'PENC', data: data.map(d => d.penc), borderColor: '#4f8cff', tension: 0.4, pointRadius: 4, borderWidth: 2, fill: false },
                    { label: 'PECUF', data: data.map(d => d.pecuf), borderColor: '#a78bfa', tension: 0.4, pointRadius: 4, borderWidth: 2, fill: false },
                    { label: 'PECNE', data: data.map(d => d.pecne), borderColor: '#fbbf24', tension: 0.4, pointRadius: 4, borderWidth: 2, fill: false },
                    { label: 'PECCUM', data: data.map(d => d.peccum), borderColor: '#f87171', tension: 0.4, pointRadius: 4, borderWidth: 2, fill: false },
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { labels: { color: '#8b8fa3', font: { size: 11 } } } },
                scales: {
                    y: { beginAtZero: true, max: 100, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#8b8fa3' } },
                    x: { grid: { display: false }, ticks: { color: '#8b8fa3' } }
                }
            }
        });
    }
}

// ═══ SECCIÓN: CRÍTICAS ═════════════════════════════════
async function cargarCriticas() {
    const q = getQuery();
    q.action = 'criticas';
    const data = await fetchAPI(q);
    
    if (!data.length) {
        $('tbodyCriticas').innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--text-muted);padding:40px;">✅ No hay evaluaciones críticas</td></tr>';
        return;
    }
    
    const tbody = $('tbodyCriticas');
    tbody.innerHTML = data.map(c => `
        <tr>
            <td><strong>${c.agente}</strong></td>
            <td>${c.supervisor || '—'}</td>
            <td>${c.fecha}</td>
            <td>${c.proyecto}</td>
            <td class="score-low"><strong>${c.score}%</strong></td>
            <td>${c.penc}%</td>
            <td>${c.pecuf}%</td>
            <td>${c.pecne}%</td>
            <td>${c.peccum}%</td>
        </tr>
    `).join('');
}

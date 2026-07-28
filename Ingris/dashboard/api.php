<?php
/**
 * API - Ingris Dashboard
 * Endpoints para obtener datos de calidad desde SQLite
 */

require_once __DIR__ . '/config.php';

$action = $_GET['action'] ?? 'resumen';

switch ($action) {
    case 'resumen':
        apiResumen();
        break;
    case 'supervisores':
        apiSupervisores();
        break;
    case 'agentes':
        apiAgentes();
        break;
    case 'campanas':
        apiCampanas();
        break;
    case 'atributos':
        apiAtributos();
        break;
    case 'historico':
        apiHistorico();
        break;
    case 'criticas':
        apiCriticas();
        break;
    case 'filtros':
        apiFiltros();
        break;
    case 'detalle_agente':
        apiDetalleAgente();
        break;
    default:
        jsonResponse(['error' => 'Acción no válida']);
}

function apiResumen() {
    $db = getDB();
    $filtro = getFiltrosSQL($db);

    // Score general
    $stmt = $db->query("SELECT 
        COUNT(*) as total_evaluaciones,
        COUNT(DISTINCT agente) as total_agentes,
        COUNT(DISTINCT supervisor) as total_supervisores,
        ROUND(AVG(score), 1) as score_promedio,
        ROUND(MIN(score), 1) as score_minimo,
        ROUND(MAX(score), 1) as score_maximo,
        SUM(es_critica) as total_criticas,
        ROUND(SUM(es_critica) * 100.0 / COUNT(*), 1) as pct_criticas
        FROM evaluaciones {$filtro['where']}");
    
    // Score por componente
    $stmt2 = $db->query("SELECT 
        ROUND(AVG(penc), 4) as avg_penc,
        ROUND(AVG(pecuf), 4) as avg_pecuf,
        ROUND(AVG(pecne), 4) as avg_pecne,
        ROUND(AVG(peccum), 4) as avg_peccum
        FROM evaluaciones {$filtro['where']}");

    $resumen = $stmt->fetch();
    $componentes = $stmt2->fetch();

    // Agregar scores corregidos (multiplicados por 100 si están en decimal)
    $resumen['penc_promedio'] = $componentes['avg_penc'] < 1 ? round($componentes['avg_penc'] * 100, 1) : round($componentes['avg_penc'], 1);
    $resumen['pecuf_promedio'] = $componentes['avg_pecuf'] < 1 ? round($componentes['avg_pecuf'] * 100, 1) : round($componentes['avg_pecuf'], 1);
    $resumen['pecne_promedio'] = $componentes['avg_pecne'] < 1 ? round($componentes['avg_pecne'] * 100, 1) : round($componentes['avg_pecne'], 1);
    $resumen['peccum_promedio'] = $componentes['avg_peccum'] < 1 ? round($componentes['avg_peccum'] * 100, 1) : round($componentes['avg_peccum'], 1);

    jsonResponse($resumen);
}

function apiSupervisores() {
    $db = getDB();
    $filtro = getFiltrosSQL($db);

    $stmt = $db->query("SELECT 
        supervisor,
        COUNT(*) as evaluaciones,
        COUNT(DISTINCT agente) as agentes,
        ROUND(AVG(score), 1) as score_promedio,
        ROUND(AVG(penc) * CASE WHEN AVG(penc) < 1 THEN 100 ELSE 1 END, 1) as penc,
        ROUND(AVG(pecuf) * CASE WHEN AVG(pecuf) < 1 THEN 100 ELSE 1 END, 1) as pecuf,
        ROUND(AVG(pecne) * CASE WHEN AVG(pecne) < 1 THEN 100 ELSE 1 END, 1) as pecne,
        ROUND(AVG(peccum) * CASE WHEN AVG(peccum) < 1 THEN 100 ELSE 1 END, 1) as peccum,
        SUM(es_critica) as criticas,
        ROUND(SUM(es_critica) * 100.0 / COUNT(*), 1) as pct_criticas
        FROM evaluaciones {$filtro['where']}
        AND supervisor != ''
        GROUP BY supervisor
        ORDER BY score_promedio DESC");

    jsonResponse($stmt->fetchAll());
}

function apiAgentes() {
    $db = getDB();
    $filtro = getFiltrosSQL($db);
    $sup = $_GET['supervisor'] ?? '';
    $orden = $_GET['orden'] ?? 'score_desc';

    $whereExtra = '';
    $params = [];
    
    if ($sup && $sup !== 'todos') {
        $whereExtra = ' AND supervisor = ' . $db->quote($sup);
    }

    $orderSQL = match($orden) {
        'score_asc' => 'score_promedio ASC',
        'nombre' => 'agente ASC',
        'evaluaciones' => 'evaluaciones DESC',
        'criticas' => 'criticas DESC',
        default => 'score_promedio DESC',
    };

    $stmt = $db->query("SELECT 
        agente,
        supervisor,
        COUNT(*) as evaluaciones,
        ROUND(AVG(score), 1) as score_promedio,
        ROUND(AVG(penc) * CASE WHEN AVG(penc) < 1 THEN 100 ELSE 1 END, 1) as penc,
        ROUND(AVG(pecuf) * CASE WHEN AVG(pecuf) < 1 THEN 100 ELSE 1 END, 1) as pecuf,
        ROUND(AVG(pecne) * CASE WHEN AVG(pecne) < 1 THEN 100 ELSE 1 END, 1) as pecne,
        ROUND(AVG(peccum) * CASE WHEN AVG(peccum) < 1 THEN 100 ELSE 1 END, 1) as peccum,
        SUM(es_critica) as criticas,
        MIN(fecha) as primera_eval,
        MAX(fecha) as ultima_eval
        FROM evaluaciones {$filtro['where']} {$whereExtra}
        AND agente != ''
        GROUP BY agente
        ORDER BY {$orderSQL}
        LIMIT 100");

    $agentes = $stmt->fetchAll();

    // Calcular cuartiles
    $scores = array_column($agentes, 'score_promedio');
    sort($scores);
    $n = count($scores);
    $q1 = $n > 0 ? $scores[(int)($n * 0.25)] : 0;
    $q2 = $n > 0 ? $scores[(int)($n * 0.50)] : 0;
    $q3 = $n > 0 ? $scores[(int)($n * 0.75)] : 0;

    foreach ($agentes as &$a) {
        $s = $a['score_promedio'];
        if ($s >= $q3) $a['cuartil'] = 'Q1';
        elseif ($s >= $q2) $a['cuartil'] = 'Q2';
        elseif ($s >= $q1) $a['cuartil'] = 'Q3';
        else $a['cuartil'] = 'Q4';
        
        // Cumple/No Cumple (umbral 85% por defecto)
        $a['cumple'] = $s >= 85 ? 'Cumple' : 'No Cumple';
    }

    jsonResponse([
        'agentes' => $agentes,
        'cuartiles' => ['Q1' => $q3, 'Q2' => $q2, 'Q3' => $q1, 'Q4' => 0],
        'total' => count($agentes)
    ]);
}

function apiCampanas() {
    $db = getDB();
    $filtro = getFiltrosSQL($db);

    $stmt = $db->query("SELECT 
        proyecto as campana,
        COUNT(*) as evaluaciones,
        COUNT(DISTINCT agente) as agentes,
        ROUND(AVG(score), 1) as score_promedio,
        ROUND(AVG(penc) * CASE WHEN AVG(penc) < 1 THEN 100 ELSE 1 END, 1) as penc,
        ROUND(AVG(pecuf) * CASE WHEN AVG(pecuf) < 1 THEN 100 ELSE 1 END, 1) as pecuf,
        ROUND(AVG(pecne) * CASE WHEN AVG(pecne) < 1 THEN 100 ELSE 1 END, 1) as pecne,
        ROUND(AVG(peccum) * CASE WHEN AVG(peccum) < 1 THEN 100 ELSE 1 END, 1) as peccum,
        SUM(es_critica) as criticas
        FROM evaluaciones {$filtro['where']}
        GROUP BY proyecto
        ORDER BY score_promedio DESC");

    jsonResponse($stmt->fetchAll());
}

function apiAtributos() {
    $db = getDB();
    $filtro = getFiltrosSQL($db);

    // Los atributos son los 4 componentes del score
    $stmt = $db->query("SELECT 
        ROUND(AVG(penc) * CASE WHEN AVG(penc) < 1 THEN 100 ELSE 1 END, 1) as penc,
        ROUND(AVG(pecuf) * CASE WHEN AVG(pecuf) < 1 THEN 100 ELSE 1 END, 1) as pecuf,
        ROUND(AVG(pecne) * CASE WHEN AVG(pecne) < 1 THEN 100 ELSE 1 END, 1) as pecne,
        ROUND(AVG(peccum) * CASE WHEN AVG(peccum) < 1 THEN 100 ELSE 1 END, 1) as peccum
        FROM evaluaciones {$filtro['where']}");

    $row = $stmt->fetch();
    
    $atributos = [
        ['nombre' => 'Precisión Error No Crítico', 'sigla' => 'PENC', 'valor' => $row['penc']],
        ['nombre' => 'Precisión Error Crítico Usuario Final', 'sigla' => 'PECUF', 'valor' => $row['pecuf']],
        ['nombre' => 'Precisión Error Crítico Negocio', 'sigla' => 'PECNE', 'valor' => $row['pecne']],
        ['nombre' => 'Precisión Error Crítico Cumplimiento', 'sigla' => 'PECCUM', 'valor' => $row['peccum']],
    ];

    jsonResponse($atributos);
}

function apiHistorico() {
    $db = getDB();

    $stmt = $db->query("SELECT 
        mes,
        mes_texto,
        COUNT(*) as evaluaciones,
        COUNT(DISTINCT agente) as agentes,
        ROUND(AVG(score), 1) as score_promedio,
        ROUND(AVG(penc) * CASE WHEN AVG(penc) < 1 THEN 100 ELSE 1 END, 1) as penc,
        ROUND(AVG(pecuf) * CASE WHEN AVG(pecuf) < 1 THEN 100 ELSE 1 END, 1) as pecuf,
        ROUND(AVG(pecne) * CASE WHEN AVG(pecne) < 1 THEN 100 ELSE 1 END, 1) as pecne,
        ROUND(AVG(peccum) * CASE WHEN AVG(peccum) < 1 THEN 100 ELSE 1 END, 1) as peccum,
        SUM(es_critica) as criticas
        FROM evaluaciones
        WHERE mes != ''
        GROUP BY mes
        ORDER BY mes ASC");

    jsonResponse($stmt->fetchAll());
}

function apiCriticas() {
    $db = getDB();
    $filtro = getFiltrosSQL($db);

    $stmt = $db->query("SELECT 
        agente, supervisor, fecha, proyecto, score,
        ROUND(penc * CASE WHEN penc < 1 THEN 100 ELSE 1 END, 1) as penc,
        ROUND(pecuf * CASE WHEN pecuf < 1 THEN 100 ELSE 1 END, 1) as pecuf,
        ROUND(pecne * CASE WHEN pecne < 1 THEN 100 ELSE 1 END, 1) as pecne,
        ROUND(peccum * CASE WHEN peccum < 1 THEN 100 ELSE 1 END, 1) as peccum
        FROM evaluaciones {$filtro['where']}
        AND es_critica = 1
        ORDER BY score ASC
        LIMIT 50");

    jsonResponse($stmt->fetchAll());
}

function apiFiltros() {
    $db = getDB();

    $supervisores = $db->query("SELECT DISTINCT supervisor FROM evaluaciones WHERE supervisor != '' ORDER BY supervisor")->fetchAll(PDO::FETCH_COLUMN);
    $proyectos = $db->query("SELECT DISTINCT proyecto FROM evaluaciones WHERE proyecto != '' ORDER BY proyecto")->fetchAll(PDO::FETCH_COLUMN);
    $meses = $db->query("SELECT DISTINCT mes, mes_texto FROM evaluaciones WHERE mes != '' ORDER BY mes DESC")->fetchAll();

    jsonResponse([
        'supervisores' => $supervisores,
        'proyectos' => $proyectos,
        'meses' => $meses,
    ]);
}

function apiDetalleAgente() {
    $db = getDB();
    $agente = $_GET['agente'] ?? '';

    if (!$agente) {
        jsonResponse(['error' => 'Agente no especificado']);
    }

    $stmt = $db->prepare("SELECT 
        e.*,
        ROUND(e.penc * CASE WHEN e.penc < 1 THEN 100 ELSE 1 END, 1) as penc_pct,
        ROUND(e.pecuf * CASE WHEN e.pecuf < 1 THEN 100 ELSE 1 END, 1) as pecuf_pct,
        ROUND(e.pecne * CASE WHEN e.pecne < 1 THEN 100 ELSE 1 END, 1) as pecne_pct,
        ROUND(e.peccum * CASE WHEN e.peccum < 1 THEN 100 ELSE 1 END, 1) as peccum_pct
        FROM evaluaciones e 
        WHERE e.agente = ?
        ORDER BY e.fecha DESC
        LIMIT 50");
    $stmt->execute([$agente]);
    
    $evaluaciones = $stmt->fetchAll();
    
    // Estadísticas del agente
    $stats = $db->prepare("SELECT 
        COUNT(*) as total,
        ROUND(AVG(score), 1) as score_promedio,
        SUM(es_critica) as criticas,
        MIN(fecha) as desde,
        MAX(fecha) as hasta
        FROM evaluaciones WHERE agente = ?");
    $stats->execute([$agente]);
    
    jsonResponse([
        'agente' => $agente,
        'stats' => $stats->fetch(),
        'evaluaciones' => $evaluaciones,
    ]);
}

function getFiltrosSQL($db) {
    $where = [];
    $params = [];

    if (!empty($_GET['supervisor']) && $_GET['supervisor'] !== 'todos') {
        $where[] = 'supervisor = ' . $db->quote($_GET['supervisor']);
    }
    if (!empty($_GET['proyecto']) && $_GET['proyecto'] !== 'todos') {
        $where[] = 'proyecto = ' . $db->quote($_GET['proyecto']);
    }
    if (!empty($_GET['mes']) && $_GET['mes'] !== 'todos') {
        $where[] = 'mes = ' . $db->quote($_GET['mes']);
    }
    if (!empty($_GET['agente']) && $_GET['agente'] !== 'todos') {
        $where[] = 'agente = ' . $db->quote($_GET['agente']);
    }
    if (!empty($_GET['critica'])) {
        $where[] = 'es_critica = ' . intval($_GET['critica']);
    }

    $sqlWhere = count($where) > 0 ? 'WHERE ' . implode(' AND ', $where) : 'WHERE 1=1';

    return ['where' => $sqlWhere, 'params' => $params];
}

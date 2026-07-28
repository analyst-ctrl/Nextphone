<?php
/**
 * CONFIGURACIÓN - Ingris Dashboard
 * Conexión a la base de datos SQLite
 */

define('DB_PATH', __DIR__ . '/data/calidad.db');

function getDB() {
    static $db = null;
    if ($db === null) {
        try {
            $db = new PDO("sqlite:" . DB_PATH);
            $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
            $db->setAttribute(PDO::ATTR_DEFAULT_FETCH_MODE, PDO::FETCH_ASSOC);
            $db->exec("PRAGMA journal_mode=WAL");
            $db->exec("PRAGMA foreign_keys=ON");
        } catch (PDOException $e) {
            http_response_code(500);
            die(json_encode(['error' => 'Error de conexión: ' . $e->getMessage()]));
        }
    }
    return $db;
}

function jsonResponse($data) {
    header('Content-Type: application/json; charset=utf-8');
    header('Access-Control-Allow-Origin: *');
    echo json_encode($data, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
    exit;
}

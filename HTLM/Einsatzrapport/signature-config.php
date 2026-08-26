<?php
/**
 * signature-config.php
 *
 * Shared configuration + helpers for the Unterschriften-Verwaltung tool
 * (unterschriften.html + signature-save.php / signature-admin-list.php /
 * signature-delete.php). Deploy in the SAME folder as those files (next to,
 * but not linked from, einsatzrapport.html) so the relative fetch() calls
 * in unterschriften.html find them, and so the "signatures" folder below
 * sits alongside this config rather than somewhere else on the server.
 *
 * ---------------------------------------------------------------------
 * ADMIN PASSWORD - CHANGE THIS BEFORE RELYING ON THIS TOOL.
 * ---------------------------------------------------------------------
 * The hash below is for the placeholder password "changeme-please" - it
 * works out of the box so the tool is testable immediately, but it is
 * PUBLIC (it's sitting in this file) and MUST be replaced.
 *
 * To generate a new hash for your own password, run this once on any
 * machine with PHP installed (nothing is sent anywhere):
 *
 *   php -r "echo password_hash('YOUR-NEW-PASSWORD', PASSWORD_DEFAULT), PHP_EOL;"
 *
 * Then paste the output (starts with $2y$...) as the value below.
 */
const ADMIN_PASSWORD_HASH = '$2y$12$by3AokYyyN0pv3JReisF3OQH9Nq7CV4dpFSOPZZItjeyHPSVQmweO';

// PIN length required when adding/overwriting a stored signature.
const PIN_LENGTH = 6;

// After this many wrong PINs (per stored signature) or wrong admin
// passwords (globally), further attempts are locked out for LOCKOUT_SECONDS.
// There's no database here, so this is tracked in small JSON files.
const MAX_FAILED_ATTEMPTS = 5;
const LOCKOUT_SECONDS      = 300; // 5 minutes

// Where stored signatures live. Protected from direct web access by
// signatures/.htaccess (Apache) - always go through the PHP endpoints below,
// never fetch a .json file in there directly.
const SIGNATURES_DIR = __DIR__ . '/signatures';

// A signature's stroke data is a handful of KB at most; this is a generous
// ceiling to reject anything absurd (abuse / bugs) rather than a real limit.
const MAX_REQUEST_BYTES = 300000; // ~300 KB

const ADMIN_LOCK_FILE = SIGNATURES_DIR . '/.admin-lock.json';

header('Content-Type: application/json; charset=utf-8');

function respond($success, $error = null, $extra = [], $httpCode = 200) {
    http_response_code($httpCode);
    $payload = $success ? array_merge(['success' => true], $extra) : ['success' => false, 'error' => $error];
    echo json_encode($payload);
    exit;
}

function require_post() {
    if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
        respond(false, 'Method not allowed', [], 405);
    }
}

/** Reads and JSON-decodes the raw POST body, with a size guard. */
function read_json_body() {
    $raw = file_get_contents('php://input', false, null, 0, MAX_REQUEST_BYTES + 1);
    if ($raw === false || strlen($raw) > MAX_REQUEST_BYTES) {
        respond(false, 'Anfrage ist zu gross.', [], 400);
    }
    $data = json_decode($raw, true);
    if (!is_array($data)) {
        respond(false, 'Ungültige Anfrage.', [], 400);
    }
    return $data;
}

/** Turns a display name into a safe filename stem: lowercase ascii-ish, dashes only. */
function slugify($name) {
    $name = trim((string) $name);
    if (function_exists('iconv')) {
        $translit = @iconv('UTF-8', 'ASCII//TRANSLIT', $name);
        if ($translit !== false) $name = $translit;
    }
    $name = strtolower($name);
    $name = preg_replace('/[^a-z0-9]+/', '-', $name);
    $name = trim($name, '-');
    return $name;
}

/** Validates an id/slug coming back from the client (e.g. for delete) before touching the filesystem. */
function is_safe_slug($slug) {
    return is_string($slug) && $slug !== '' && preg_match('/^[a-z0-9-]+$/', $slug) === 1;
}

function ensure_signatures_dir() {
    if (!is_dir(SIGNATURES_DIR)) {
        mkdir(SIGNATURES_DIR, 0755, true);
    }
}

/** Atomic-ish read of a small JSON file; returns null if missing/corrupt. */
function read_json_file($path) {
    if (!is_file($path)) return null;
    $raw = file_get_contents($path);
    if ($raw === false) return null;
    $data = json_decode($raw, true);
    return is_array($data) ? $data : null;
}

function write_json_file($path, $data) {
    $tmp = $path . '.tmp-' . bin2hex(random_bytes(4));
    if (file_put_contents($tmp, json_encode($data)) === false) return false;
    return rename($tmp, $path);
}

/**
 * Shared lockout check/update, used both for a single stored signature's PIN
 * (state kept inside that signature's own record) and for the global admin
 * password (state kept in ADMIN_LOCK_FILE). $state is the ['failedAttempts'
 * => int, 'lockedUntil' => int] slice of whichever record is relevant.
 */
function lockout_remaining_seconds($state) {
    $lockedUntil = isset($state['lockedUntil']) ? (int) $state['lockedUntil'] : 0;
    $remaining = $lockedUntil - time();
    return $remaining > 0 ? $remaining : 0;
}

function lockout_register_failure(&$state) {
    $state['failedAttempts'] = (isset($state['failedAttempts']) ? (int) $state['failedAttempts'] : 0) + 1;
    if ($state['failedAttempts'] >= MAX_FAILED_ATTEMPTS) {
        $state['lockedUntil'] = time() + LOCKOUT_SECONDS;
        $state['failedAttempts'] = 0;
    }
}

function lockout_register_success(&$state) {
    $state['failedAttempts'] = 0;
    $state['lockedUntil'] = 0;
}

/**
 * Verifies the admin password against ADMIN_PASSWORD_HASH, with a global
 * failed-attempt lockout (there's only one admin password, so this can't be
 * scoped per-record like the per-signature PIN lockout is).
 */
function verify_admin_password($password) {
    ensure_signatures_dir();
    $lock = read_json_file(ADMIN_LOCK_FILE) ?? ['failedAttempts' => 0, 'lockedUntil' => 0];

    $remaining = lockout_remaining_seconds($lock);
    if ($remaining > 0) {
        respond(false, "Zu viele Fehlversuche. Bitte in " . ceil($remaining / 60) . " Minute(n) erneut versuchen.", [], 429);
    }

    if (!is_string($password) || $password === '' || !password_verify($password, ADMIN_PASSWORD_HASH)) {
        lockout_register_failure($lock);
        write_json_file(ADMIN_LOCK_FILE, $lock);
        respond(false, 'Falsches Passwort.', [], 403);
    }

    lockout_register_success($lock);
    write_json_file(ADMIN_LOCK_FILE, $lock);
}

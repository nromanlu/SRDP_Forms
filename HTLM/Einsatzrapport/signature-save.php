<?php
/**
 * signature-save.php
 *
 * Stores (or, with the correct PIN, overwrites) one named signature as a
 * JSON file in signatures/<slug>.json. Called from unterschriften.html's
 * "Unterschrift hinzufügen" button, after the user has entered a name,
 * drawn a signature, and confirmed a PIN in the on-page modal.
 *
 * Body (JSON): { name: string, pin: "1234", strokes: [[x0,y0,x1,y1,...], ...] }
 *
 * - New name -> creates the record, PIN is hashed and becomes that
 *   signature's write-protection going forward.
 * - Existing name -> the submitted PIN must match the stored one
 *   (password_verify) before the signature is overwritten. This is a
 *   self-service update path (re-sign under the same name), not a way to
 *   change the PIN itself.
 */

require __DIR__ . '/signature-config.php';

require_post();
$data = read_json_body();

$name    = isset($data['name']) ? trim((string) $data['name']) : '';
$pin     = isset($data['pin']) ? (string) $data['pin'] : '';
$strokes = isset($data['strokes']) ? $data['strokes'] : null;

if ($name === '') {
    respond(false, 'Name fehlt.', [], 400);
}
if (!preg_match('/^\d{' . PIN_LENGTH . '}$/', $pin)) {
    respond(false, 'PIN muss genau ' . PIN_LENGTH . ' Ziffern haben.', [], 400);
}
if (!is_array($strokes) || count($strokes) === 0) {
    respond(false, 'Keine Unterschrift erfasst.', [], 400);
}
// Sanity-check the stroke payload shape: array of arrays of finite numbers.
$pointCount = 0;
foreach ($strokes as $stroke) {
    if (!is_array($stroke)) respond(false, 'Ungültige Unterschrift-Daten.', [], 400);
    foreach ($stroke as $n) {
        if (!is_numeric($n)) respond(false, 'Ungültige Unterschrift-Daten.', [], 400);
        $pointCount++;
    }
}
if ($pointCount > 20000) {
    respond(false, 'Unterschrift-Daten sind zu umfangreich.', [], 400);
}

$slug = slugify($name);
if ($slug === '') {
    respond(false, 'Name enthält keine verwendbaren Zeichen.', [], 400);
}

ensure_signatures_dir();
$path = SIGNATURES_DIR . '/' . $slug . '.json';
$existing = read_json_file($path);

if ($existing !== null) {
    $remaining = lockout_remaining_seconds($existing);
    if ($remaining > 0) {
        respond(false, "Zu viele Fehlversuche für diesen Namen. Bitte in " . ceil($remaining / 60) . " Minute(n) erneut versuchen.", [], 429);
    }
    if (!isset($existing['pinHash']) || !password_verify($pin, $existing['pinHash'])) {
        lockout_register_failure($existing);
        write_json_file($path, $existing);
        respond(false, 'PIN falsch für "' . $existing['name'] . '".', [], 403);
    }
    lockout_register_success($existing);
    $record = $existing;
    $record['name']      = $name;
    $record['strokes']   = $strokes;
    $record['updatedAt'] = gmdate('Y-m-d H:i:s') . ' UTC';
} else {
    $record = [
        'name'           => $name,
        'slug'           => $slug,
        'pinHash'        => password_hash($pin, PASSWORD_DEFAULT),
        'strokes'        => $strokes,
        'updatedAt'      => gmdate('Y-m-d H:i:s') . ' UTC',
        'failedAttempts' => 0,
        'lockedUntil'    => 0,
    ];
}

if (!write_json_file($path, $record)) {
    respond(false, 'Speichern fehlgeschlagen.', [], 500);
}

respond(true, null, ['updatedAt' => $record['updatedAt']]);

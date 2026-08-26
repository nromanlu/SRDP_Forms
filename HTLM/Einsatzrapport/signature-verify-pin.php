<?php
/**
 * signature-verify-pin.php
 *
 * Verifies a PIN against one stored signature by id - used by the
 * "Bearbeitet von" dropdown in einsatzrapport.html to confirm a person's
 * identity before their selected name is accepted.
 *
 * This is a SEPARATE, stricter policy from signature-save.php's overwrite
 * check: wrong attempts here count against their own `selectFailedAttempts`
 * field (not the save flow's `failedAttempts` / `lockedUntil`, which is a
 * temporary 5-minute lockout instead) so the two flows can't interfere with
 * each other. After MAX_FAILED_ATTEMPTS (5) wrong PINs, the stored
 * signature is deleted outright rather than just locked out - the person
 * would need to re-add it from scratch (with a new PIN) in unterschriften.html.
 *
 * Body (JSON): { id: "max-muster", pin: "123456" }
 */

require __DIR__ . '/signature-config.php';

require_post();
$data = read_json_body();

$id  = isset($data['id']) ? (string) $data['id'] : '';
$pin = isset($data['pin']) ? (string) $data['pin'] : '';

if (!is_safe_slug($id)) {
    respond(false, 'Ungültige ID.', [], 400);
}
if (!preg_match('/^\d{' . PIN_LENGTH . '}$/', $pin)) {
    respond(false, 'PIN muss genau ' . PIN_LENGTH . ' Ziffern haben.', [], 400);
}

ensure_signatures_dir();
$path = SIGNATURES_DIR . '/' . $id . '.json';
$record = read_json_file($path);

if ($record === null) {
    respond(false, 'Signatur nicht gefunden. Die Seite evtl. neu laden.', [], 404);
}

if (!isset($record['pinHash']) || !password_verify($pin, $record['pinHash'])) {
    $attempts = (isset($record['selectFailedAttempts']) ? (int) $record['selectFailedAttempts'] : 0) + 1;

    if ($attempts >= MAX_FAILED_ATTEMPTS) {
        @unlink($path);
        respond(
            false,
            'PIN falsch. Diese Signatur wurde nach ' . MAX_FAILED_ATTEMPTS . ' Fehlversuchen automatisch gelöscht.',
            ['deleted' => true, 'remaining' => 0],
            403
        );
    }

    $record['selectFailedAttempts'] = $attempts;
    write_json_file($path, $record);
    $remaining = MAX_FAILED_ATTEMPTS - $attempts;
    respond(
        false,
        "PIN falsch. Noch {$remaining} Versuch(e), bevor die Signatur automatisch gelöscht wird.",
        ['deleted' => false, 'remaining' => $remaining],
        403
    );
}

// Correct PIN - reset the counter and confirm.
$record['selectFailedAttempts'] = 0;
write_json_file($path, $record);
respond(true, null, ['name' => $record['name']]);

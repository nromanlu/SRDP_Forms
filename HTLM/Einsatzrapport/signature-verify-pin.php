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
 * each other. After SELECT_MAX_FAILED_ATTEMPTS wrong PINs, the stored
 * signature is deleted outright rather than just locked out - the person
 * would need to re-add it from scratch (with a new PIN) in unterschriften.html.
 * The person at the form is never told a deletion happened, at any attempt -
 * every wrong attempt (including the last one) only reports the remaining
 * count, which simply reaches 0. The admin is notified separately by email
 * (see signature-notify-deletion.php) via a one-time token issued below.
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

    if ($attempts >= SELECT_MAX_FAILED_ATTEMPTS) {
        $deletedName = $record['name'];
        @unlink($path);

        // The person at the form is never told a signature was deleted -
        // just that no attempts are left, same wording style as the
        // countdown above. The admin is told separately, by email (see
        // signature-notify-deletion.php) using the one-time token below.
        $notifyToken = bin2hex(random_bytes(16));
        write_json_file(
            SIGNATURES_DIR . '/.notify-' . $notifyToken . '.json',
            ['name' => $deletedName, 'createdAt' => time()]
        );

        respond(
            false,
            'PIN falsch. Noch 0 Versuch(e).',
            ['deleted' => true, 'remaining' => 0, 'notifyToken' => $notifyToken],
            403
        );
    }

    $record['selectFailedAttempts'] = $attempts;
    write_json_file($path, $record);
    $remaining = SELECT_MAX_FAILED_ATTEMPTS - $attempts;
    respond(
        false,
        "PIN falsch. Noch {$remaining} Versuch(e).",
        ['deleted' => false, 'remaining' => $remaining],
        403
    );
}

// Correct PIN - reset the counter and confirm.
$record['selectFailedAttempts'] = 0;
write_json_file($path, $record);
respond(true, null, ['name' => $record['name']]);

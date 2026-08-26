<?php
/**
 * signature-delete.php
 *
 * Deletes one stored signature by id (slug), once the correct admin
 * password is supplied. Used by the "Löschen" button per row in
 * unterschriften.html's "Signaturen verwalten" section.
 *
 * Body (JSON): { adminPassword: string, id: "max-muster" }
 */

require __DIR__ . '/signature-config.php';

require_post();
$data = read_json_body();

$adminPassword = isset($data['adminPassword']) ? (string) $data['adminPassword'] : '';
verify_admin_password($adminPassword); // exits with an error response on failure

$id = isset($data['id']) ? (string) $data['id'] : '';
if (!is_safe_slug($id)) {
    respond(false, 'Ungültige ID.', [], 400);
}

$path = SIGNATURES_DIR . '/' . $id . '.json';
if (!is_file($path)) {
    respond(false, 'Signatur nicht gefunden.', [], 404);
}
if (!unlink($path)) {
    respond(false, 'Löschen fehlgeschlagen.', [], 500);
}

respond(true);

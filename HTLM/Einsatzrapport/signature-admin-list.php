<?php
/**
 * signature-admin-list.php
 *
 * Lists all stored signatures (name, id/slug, last-updated, stroke data for
 * a preview render) once the correct admin password is supplied. Never
 * returns PIN hashes. Used by unterschriften.html's "Signaturen verwalten"
 * section after the admin unlocks it.
 *
 * Body (JSON): { adminPassword: string }
 */

require __DIR__ . '/signature-config.php';

require_post();
$data = read_json_body();

$adminPassword = isset($data['adminPassword']) ? (string) $data['adminPassword'] : '';
verify_admin_password($adminPassword); // exits with an error response on failure

ensure_signatures_dir();

$signatures = [];
foreach (glob(SIGNATURES_DIR . '/*.json') as $file) {
    $slug = basename($file, '.json');
    if (!is_safe_slug($slug)) continue; // skips .admin-lock.json etc.
    $record = read_json_file($file);
    if ($record === null) continue;
    $signatures[] = [
        'id'        => $slug,
        'name'      => $record['name'] ?? $slug,
        'rank'      => $record['rank'] ?? '',
        'updatedAt' => $record['updatedAt'] ?? '',
        'strokes'   => $record['strokes'] ?? [],
    ];
}

usort($signatures, fn($a, $b) => strcasecmp($a['name'], $b['name']));

respond(true, null, ['signatures' => $signatures]);

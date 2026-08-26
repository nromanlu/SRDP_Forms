<?php
/**
 * signature-names-list.php
 *
 * Public (no password) list of stored-signature names - powers the
 * "Bearbeitet von" dropdown in einsatzrapport.html. Deliberately lighter
 * than signature-admin-list.php: only id + name, never strokes or PIN
 * hashes, so exposing it without the admin password is fine. Anyone filling
 * out the Einsatzrapport needs to see the list of names to pick their own;
 * the PIN (checked separately by signature-verify-pin.php once a name is
 * chosen) is what actually protects each signature, not this list.
 */

require __DIR__ . '/signature-config.php';

ensure_signatures_dir();

$names = [];
foreach (glob(SIGNATURES_DIR . '/*.json') as $file) {
    $slug = basename($file, '.json');
    if (!is_safe_slug($slug)) continue; // skips .admin-lock.json etc.
    $record = read_json_file($file);
    if ($record === null) continue;
    $names[] = [
        'id'   => $slug,
        'name' => $record['name'] ?? $slug,
    ];
}

usort($names, fn($a, $b) => strcasecmp($a['name'], $b['name']));

respond(true, null, ['signatures' => $names]);

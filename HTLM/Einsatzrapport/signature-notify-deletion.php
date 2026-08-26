<?php
/**
 * signature-notify-deletion.php
 *
 * Emails the admin when a stored signature was just auto-deleted (3 wrong
 * PIN attempts in einsatzrapport.html's "Bearbeitet von" dropdown - see
 * signature-verify-pin.php). The person filling out the form is NOT told a
 * signature was deleted (the on-screen message just counts down to "0
 * Versuch(e)") - this is purely an admin notification, sent server-side.
 *
 * Called from the client right after signature-verify-pin.php responds with
 * deleted:true - the client re-runs generatePDF() (the same function used
 * for "Absenden" / "PDF herunterladen") to build a PDF of the Einsatzrapport
 * exactly as far as it had been filled in, and posts it here together with
 * the one-time `notifyToken` signature-verify-pin.php issued. That token is
 * required and single-use (a small file in signatures/, deleted the moment
 * it's read here, valid for NOTIFY_TOKEN_TTL_SECONDS) - so this endpoint
 * can't be used to send arbitrary emails; it only works right after a real
 * deletion, and only once per deletion.
 *
 * Deploy in the SAME folder as einsatzrapport.html / signature-verify-pin.php.
 */

require __DIR__ . '/signature-config.php';

// ---------------------------------------------------------------------------
// Configuration - same recipient/sender pattern as send-report.php.
// ---------------------------------------------------------------------------
$RECIPIENT  = 'n.romanlu@gmail.com';
$FROM_ADDR  = 'noreply@nr-works.ch';
$FROM_NAME  = 'Einsatzrapport SRDP';
$MAX_PDF_MB = 10;
// ---------------------------------------------------------------------------

require_post();

$token = isset($_POST['token']) ? (string) $_POST['token'] : '';
if (!preg_match('/^[a-f0-9]{32}$/', $token)) {
    respond(false, 'Ungültiges Token.', [], 400);
}

ensure_signatures_dir();
$tokenPath = SIGNATURES_DIR . '/.notify-' . $token . '.json';
$tokenData = read_json_file($tokenPath);
// Single-use: consumed immediately, whether or not it turns out valid below.
@unlink($tokenPath);

if ($tokenData === null || !isset($tokenData['name'], $tokenData['createdAt'])) {
    respond(false, 'Unbekanntes oder bereits verwendetes Token.', [], 403);
}
if (time() - (int) $tokenData['createdAt'] > NOTIFY_TOKEN_TTL_SECONDS) {
    respond(false, 'Token ist abgelaufen.', [], 403);
}

$signatureName = (string) $tokenData['name'];

if (!isset($_FILES['pdf']) || $_FILES['pdf']['error'] !== UPLOAD_ERR_OK) {
    respond(false, 'PDF fehlt oder Upload ist fehlgeschlagen.', [], 400);
}
if ($_FILES['pdf']['size'] > $MAX_PDF_MB * 1024 * 1024) {
    respond(false, 'PDF ist zu gross.', [], 400);
}

// Verify the uploaded file is really a PDF (don't trust the client-supplied MIME type).
if (function_exists('finfo_open')) {
    $finfo = finfo_open(FILEINFO_MIME_TYPE);
    $realType = finfo_file($finfo, $_FILES['pdf']['tmp_name']);
    finfo_close($finfo);
    if ($realType !== 'application/pdf') {
        respond(false, 'Ungültiger Dateityp.', [], 400);
    }
}

$pdfData = file_get_contents($_FILES['pdf']['tmp_name']);
if ($pdfData === false) {
    respond(false, 'Datei konnte nicht gelesen werden.', [], 500);
}

$attachmentName = 'Einsatzrapport-bei-Loeschung.pdf';

$subject = 'Signatur automatisch gelöscht: ' . $signatureName;
$bodyText =
    "Die gespeicherte Signatur von \"{$signatureName}\" wurde soeben automatisch gelöscht, " .
    'nachdem im Einsatzrapport-Formular (Feld "Bearbeitet von") ' . SELECT_MAX_FAILED_ATTEMPTS . " Mal hintereinander " .
    "die falsche PIN eingegeben wurde.\n\n" .
    "Im Anhang der Einsatzrapport mit dem Stand, wie er zu diesem Zeitpunkt ausgefüllt war.";

// ---------------------------------------------------------------------------
// Build the email: a plain-text body + the PDF as an attachment.
// ---------------------------------------------------------------------------
$mixedBoundary = md5(uniqid((string) microtime(true), true));

$headers  = "From: {$FROM_NAME} <{$FROM_ADDR}>\r\n";
$headers .= "Reply-To: {$FROM_ADDR}\r\n";
$headers .= "MIME-Version: 1.0\r\n";
$headers .= "Content-Type: multipart/mixed; boundary=\"{$mixedBoundary}\"\r\n";

$textPart  = "--{$mixedBoundary}\r\n";
$textPart .= "Content-Type: text/plain; charset=UTF-8\r\n";
$textPart .= "Content-Transfer-Encoding: quoted-printable\r\n\r\n";
$textPart .= quoted_printable_encode($bodyText) . "\r\n\r\n";

$attachment  = "--{$mixedBoundary}\r\n";
$attachment .= "Content-Type: application/pdf; name=\"{$attachmentName}\"\r\n";
$attachment .= "Content-Transfer-Encoding: base64\r\n";
$attachment .= "Content-Disposition: attachment; filename=\"{$attachmentName}\"\r\n\r\n";
$attachment .= chunk_split(base64_encode($pdfData)) . "\r\n";

$message = $textPart . $attachment . "--{$mixedBoundary}--";

// RFC 2047 encode the subject so umlauts (ä/ö/ü) render correctly in the recipient's inbox.
$encodedSubject = '=?UTF-8?B?' . base64_encode($subject) . '?=';

$sent = mail($RECIPIENT, $encodedSubject, $message, $headers);

respond($sent, $sent ? null : 'mail() ist fehlgeschlagen.', [], $sent ? 200 : 500);

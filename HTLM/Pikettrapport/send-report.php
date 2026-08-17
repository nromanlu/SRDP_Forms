<?php
/**
 * send-report.php
 *
 * Receives the generated Pikettrapport PDF from pikettrapport.html and emails it
 * to the fixed recipient below, with the PDF as an attachment. No mail app opens
 * on the device — the message is sent directly from the server.
 *
 * Deploy this file in the SAME folder as pikettrapport.html so the relative
 * fetch('send-report.php') call in the form finds it. It ships through the
 * existing GitHub Actions FTP pipeline automatically, just like the HTML file.
 */

header('Content-Type: application/json; charset=utf-8');

// ---------------------------------------------------------------------------
// Configuration — adjust these two if needed.
// ---------------------------------------------------------------------------
$RECIPIENT   = 'n.romanlu@gmail.com';
// Must be an address on a domain this Infomaniak hosting account controls,
// ideally a real mailbox you've created (better deliverability / less likely
// to land in spam than a made-up noreply@ address). If you have a real mailbox
// on nr-works.dev, use it here instead.
$FROM_ADDR   = 'noreply@nr-works.ch';
$FROM_NAME   = 'Pikettrapport SRDP';
$MAX_PDF_MB  = 10;

// ---------------------------------------------------------------------------
function respond($success, $error = null, $httpCode = 200) {
    http_response_code($httpCode);
    echo json_encode($success ? ['success' => true] : ['success' => false, 'error' => $error]);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    respond(false, 'Method not allowed', 405);
}

if (!isset($_FILES['pdf']) || $_FILES['pdf']['error'] !== UPLOAD_ERR_OK) {
    respond(false, 'PDF fehlt oder Upload ist fehlgeschlagen.', 400);
}

if ($_FILES['pdf']['size'] > $MAX_PDF_MB * 1024 * 1024) {
    respond(false, 'PDF ist zu gross.', 400);
}

// Verify the uploaded file is really a PDF (don't trust the client-supplied MIME type).
if (function_exists('finfo_open')) {
    $finfo = finfo_open(FILEINFO_MIME_TYPE);
    $realType = finfo_file($finfo, $_FILES['pdf']['tmp_name']);
    finfo_close($finfo);
    if ($realType !== 'application/pdf') {
        respond(false, 'Ungültiger Dateityp.', 400);
    }
}

$pdfData = file_get_contents($_FILES['pdf']['tmp_name']);
if ($pdfData === false) {
    respond(false, 'Datei konnte nicht gelesen werden.', 500);
}

$subject        = isset($_POST['subject']) ? trim($_POST['subject']) : 'Pikettrapport';
$bodyText       = isset($_POST['body']) ? trim($_POST['body']) : '';
$attachmentName = isset($_POST['filename']) ? basename($_POST['filename']) : 'Pikettrapport.pdf';
if (substr(strtolower($attachmentName), -4) !== '.pdf') {
    $attachmentName .= '.pdf';
}

// ---------------------------------------------------------------------------
// Build a multipart/mixed MIME email with the PDF as a base64 attachment.
// ---------------------------------------------------------------------------
$boundary = md5(uniqid((string) microtime(true), true));

$headers  = "From: {$FROM_NAME} <{$FROM_ADDR}>\r\n";
$headers .= "Reply-To: {$FROM_ADDR}\r\n";
$headers .= "MIME-Version: 1.0\r\n";
$headers .= "Content-Type: multipart/mixed; boundary=\"{$boundary}\"\r\n";

$message  = "--{$boundary}\r\n";
$message .= "Content-Type: text/plain; charset=UTF-8\r\n";
$message .= "Content-Transfer-Encoding: 8bit\r\n\r\n";
$message .= $bodyText . "\r\n\r\n";

$message .= "--{$boundary}\r\n";
$message .= "Content-Type: application/pdf; name=\"{$attachmentName}\"\r\n";
$message .= "Content-Transfer-Encoding: base64\r\n";
$message .= "Content-Disposition: attachment; filename=\"{$attachmentName}\"\r\n\r\n";
$message .= chunk_split(base64_encode($pdfData)) . "\r\n";

$message .= "--{$boundary}--";

// RFC 2047 encode the subject so umlauts (ä/ö/ü) render correctly in the recipient's inbox.
$encodedSubject = '=?UTF-8?B?' . base64_encode($subject) . '?=';

$sent = mail($RECIPIENT, $encodedSubject, $message, $headers);

if ($sent) {
    respond(true);
} else {
    respond(false, 'mail() ist fehlgeschlagen.', 500);
}

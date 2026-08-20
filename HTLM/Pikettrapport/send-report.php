<?php
/**
 * send-report.php
 *
 * Receives the generated Pikettrapport PDF from pikettrapport.html and emails it
 * to the fixed recipient below, with the PDF as an attachment. No mail app opens
 * on the device — the message is sent directly from the server.
 *
 * The email body also contains a "Rapport bearbeiten" link. That link carries
 * the entire filled-in form (as a base64-encoded fragment in the URL, built by
 * buildEditUrl() in pikettrapport.html) so that opening it re-fills the form,
 * e.g. to fix a typo, without needing to type everything again. The link is
 * NOT included in the PDF itself — only in the email body.
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
// Edit links only make sense if they point back at nr-works.ch. Anything else
// (e.g. a manipulated value) is dropped rather than emailed out.
$ALLOWED_EDIT_URL_HOSTS = ['nr-works.ch', 'www.nr-works.ch'];

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
// Edit link ("Rapport bearbeiten") — validate before using it anywhere.
// ---------------------------------------------------------------------------
$editUrl = isset($_POST['editUrl']) ? trim($_POST['editUrl']) : '';
$editUrl = preg_replace('/[\r\n]+/', '', $editUrl); // strip any injected line breaks

if ($editUrl !== '') {
    $parts = parse_url($editUrl);
    $isHttps   = isset($parts['scheme']) && $parts['scheme'] === 'https';
    $hostOk    = isset($parts['host']) && in_array(strtolower($parts['host']), $ALLOWED_EDIT_URL_HOSTS, true);
    if (!$isHttps || !$hostOk) {
        // Don't fail the whole request over this — just omit the link.
        $editUrl = '';
    }
}

// ---------------------------------------------------------------------------
// Build the email:
//   multipart/mixed
//     ├─ multipart/alternative
//     │    ├─ text/plain  (body + edit link as a plain URL)
//     │    └─ text/html   (body + edit link as a clickable <a>)
//     └─ application/pdf  (the report, unchanged)
// ---------------------------------------------------------------------------
$mixedBoundary = md5(uniqid((string) microtime(true), true));
$altBoundary   = md5(uniqid((string) microtime(true) . 'alt', true));

$headers  = "From: {$FROM_NAME} <{$FROM_ADDR}>\r\n";
$headers .= "Reply-To: {$FROM_ADDR}\r\n";
$headers .= "MIME-Version: 1.0\r\n";
$headers .= "Content-Type: multipart/mixed; boundary=\"{$mixedBoundary}\"\r\n";

// --- plain text body ---
$plainBody = $bodyText;
if ($editUrl !== '') {
    $plainBody .= "\r\n\r\nRapport bearbeiten (z.B. bei einem Tippfehler):\r\n{$editUrl}";
}

// --- html body ---
$htmlBodyText = nl2br(htmlspecialchars($bodyText, ENT_QUOTES, 'UTF-8'));
$htmlBody  = "<div style=\"font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a;\">";
$htmlBody .= "<p>{$htmlBodyText}</p>";
if ($editUrl !== '') {
    $safeEditUrl = htmlspecialchars($editUrl, ENT_QUOTES, 'UTF-8');
    $htmlBody .= "<p style=\"margin-top:18px;\">";
    $htmlBody .= "<a href=\"{$safeEditUrl}\" style=\"display:inline-block;padding:8px 16px;background:#1a4fa0;color:#ffffff;text-decoration:none;border-radius:4px;\">Rapport bearbeiten</a>";
    $htmlBody .= "</p>";
    $htmlBody .= "<p style=\"font-size:11px;color:#777;margin-top:6px;\">Falls der Button nicht funktioniert, diesen Link öffnen:<br>{$safeEditUrl}</p>";
}
$htmlBody .= "</div>";

// --- assemble multipart/alternative ---
$alternative  = "--{$mixedBoundary}\r\n";
$alternative .= "Content-Type: multipart/alternative; boundary=\"{$altBoundary}\"\r\n\r\n";

// Quoted-printable (not 8bit) is important here: the edit link can push a
// single line (e.g. the href="...") well past ~1000 characters. Some MTAs
// hard-wrap long 8bit lines in transit by inserting a stray line break /
// space wherever the limit is hit, which silently corrupts the base64 data
// in the middle of the link. Quoted-printable encodes long lines with
// explicit "=\r\n" soft breaks that mail clients reassemble correctly, so
// the link (and the rest of the body) always arrives intact.
$alternative .= "--{$altBoundary}\r\n";
$alternative .= "Content-Type: text/plain; charset=UTF-8\r\n";
$alternative .= "Content-Transfer-Encoding: quoted-printable\r\n\r\n";
$alternative .= quoted_printable_encode($plainBody) . "\r\n\r\n";

$alternative .= "--{$altBoundary}\r\n";
$alternative .= "Content-Type: text/html; charset=UTF-8\r\n";
$alternative .= "Content-Transfer-Encoding: quoted-printable\r\n\r\n";
$alternative .= quoted_printable_encode($htmlBody) . "\r\n\r\n";

$alternative .= "--{$altBoundary}--\r\n";

// --- PDF attachment (unchanged, no edit link inside it) ---
$attachment  = "--{$mixedBoundary}\r\n";
$attachment .= "Content-Type: application/pdf; name=\"{$attachmentName}\"\r\n";
$attachment .= "Content-Transfer-Encoding: base64\r\n";
$attachment .= "Content-Disposition: attachment; filename=\"{$attachmentName}\"\r\n\r\n";
$attachment .= chunk_split(base64_encode($pdfData)) . "\r\n";

$message = $alternative . $attachment . "--{$mixedBoundary}--";

// RFC 2047 encode the subject so umlauts (ä/ö/ü) render correctly in the recipient's inbox.
$encodedSubject = '=?UTF-8?B?' . base64_encode($subject) . '?=';

$sent = mail($RECIPIENT, $encodedSubject, $message, $headers);

if ($sent) {
    respond(true);
} else {
    respond(false, 'mail() ist fehlgeschlagen.', 500);
}
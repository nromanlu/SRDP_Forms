<?php
/**
 * map-snapshot.php
 *
 * Server-side proxy for Geoapify's Static Maps API, so the API key never
 * reaches the browser - embedding it directly in einsatzrapport.html's
 * JavaScript would expose it to anyone who views page source or the
 * network tab. The client (fetchGpsMapImage() in einsatzrapport.html)
 * calls this endpoint with just lat/lon; this file attaches the key and
 * forwards the resulting PNG straight through.
 *
 * Called from generatePDF() when the "Position (GPS)" fields are filled
 * in, to embed a small OpenStreetMap snapshot in the "Gemeindegebiet &
 * Position" section of the PDF. On any failure here the client simply
 * skips the map and generates the rest of the report normally (see
 * fetchGpsMapImage()) - a missing map must never block sending an actual
 * report, e.g. out on the water with no signal.
 *
 * Deploy in the SAME folder as einsatzrapport.html.
 *
 * ---------------------------------------------------------------------
 * API KEY
 * ---------------------------------------------------------------------
 * Free Geoapify tier - no credit card, ~3000 credits/day as of writing.
 * If this key ever needs replacing (stops working, or you want to
 * restrict it to your own domain in the Geoapify dashboard for extra
 * safety), get a new one at https://www.geoapify.com/ and swap it in
 * below - nothing else in this file needs to change.
 */
const GEOAPIFY_API_KEY = '5d0494d021a84a36b381c240d7b37466';

// ---------------------------------------------------------------------
// Map appearance - change here, not in einsatzrapport.html. The client
// reads back the resulting image's real pixel size to keep the aspect
// ratio correct in the PDF, so these can be changed freely.
// ---------------------------------------------------------------------
const MAP_WIDTH  = 640;
const MAP_HEIGHT = 420;
const MAP_ZOOM   = 12;
// "osm-carto" is Geoapify's raster style based on the same OSM Carto
// renderer that powers openstreetmap.org itself - the closest match to
// the classic OpenStreetMap look, unlike their various vector styles.
const MAP_STYLE  = 'maptiler-3d';
const MARKER_COLOR = '%23e63329'; // URL-encoded #e63329 (red)

function respond_error($message, $httpCode = 400) {
    http_response_code($httpCode);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(['success' => false, 'error' => $message]);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'GET') {
    respond_error('Method not allowed', 405);
}

$lat = isset($_GET['lat']) ? filter_var($_GET['lat'], FILTER_VALIDATE_FLOAT) : false;
$lon = isset($_GET['lon']) ? filter_var($_GET['lon'], FILTER_VALIDATE_FLOAT) : false;

if ($lat === false || $lon === false) {
    respond_error('lat/lon fehlen oder sind ungueltig.');
}
if ($lat < -90 || $lat > 90 || $lon < -180 || $lon > 180) {
    respond_error('lat/lon ausserhalb des gueltigen Bereichs.');
}

$url = 'https://maps.geoapify.com/v1/staticmap'
    . '?style=' . MAP_STYLE
    . '&width=' . MAP_WIDTH
    . '&height=' . MAP_HEIGHT
    . '&center=lonlat:' . rawurlencode($lon) . ',' . rawurlencode($lat)
    . '&zoom=' . MAP_ZOOM
    . '&marker=lonlat:' . rawurlencode($lon) . ',' . rawurlencode($lat) . ';color:' . MARKER_COLOR . ';size:large'
    . '&apiKey=' . GEOAPIFY_API_KEY;

if (!function_exists('curl_init')) {
    respond_error('Server-Konfiguration: PHP curl-Erweiterung fehlt.', 500);
}

$ch = curl_init($url);
curl_setopt_array($ch, [
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_TIMEOUT        => 10,
    CURLOPT_FOLLOWLOCATION => true,
]);
$imageData = curl_exec($ch);
$httpCode  = curl_getinfo($ch, CURLINFO_HTTP_CODE);
$curlError = curl_error($ch);
curl_close($ch);

if ($imageData === false || $curlError) {
    respond_error('Geoapify konnte nicht erreicht werden: ' . $curlError, 502);
}
if ($httpCode !== 200) {
    respond_error('Geoapify-Fehler (HTTP ' . $httpCode . ').', 502);
}

header('Content-Type: image/png');
header('Cache-Control: private, max-age=3600');
echo $imageData;

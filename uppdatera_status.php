<?php
session_start();
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

// ── Enbart POST tillåtet ──
if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    header('Content-Type: application/json');
    echo json_encode(['success' => false, 'message' => 'Endast POST tillåtet']);
    exit;
}

// ── CSRF-validering ──
$token = $_POST['csrf_token'] ?? '';
if (empty($token) || !hash_equals($_SESSION['csrf_token'] ?? '', $token)) {
    http_response_code(403);
    header('Content-Type: application/json');
    echo json_encode(['success' => false, 'message' => 'Ogiltig CSRF-token']);
    exit;
}

// ── Fillåsning ──
$lockFile = __DIR__ . '/source.xlsx.lock';
$lockFp = fopen($lockFile, 'w');
if (!flock($lockFp, LOCK_EX)) {
    header('Content-Type: application/json');
    echo json_encode(['success' => false, 'message' => 'Kunde inte låsa filen, försök igen']);
    fclose($lockFp);
    exit;
}

$spreadsheet = IOFactory::load('source.xlsx');
$sheet = $spreadsheet->getActiveSheet();

// ── Tillåtna statusvärden ──
$allowedStatuses = ['Testad OK', 'Inte testad', ''];

function normalizeSystemName($name) {
    $name = urldecode($name);
    $name = trim($name);
    $name = preg_replace('/\s+/', ' ', $name);
    return $name;
}

function getSystemStats($sheet) {
    $testedOkCount = 0;
    $totalSystems = 0;
    $highestRow = $sheet->getHighestDataRow();
    $systemNamesInSheet = [];

    for ($row = 1; $row <= $highestRow; $row++) {
        $systemName = $sheet->getCell('A' . $row)->getValue();
        if (empty($systemName)) {
            continue;
        }
        $normalizedSystemName = normalizeSystemName($systemName);

        if (!in_array($normalizedSystemName, $systemNamesInSheet) && !empty($normalizedSystemName)) {
            $systemNamesInSheet[] = $normalizedSystemName;
            $status = trim($sheet->getCell('B' . $row)->getValue());
            $totalSystems++;
            if ($status === 'Testad OK') {
                $testedOkCount++;
            }
        }
    }
    $percentageTestedOk = ($totalSystems > 0) ? round(($testedOkCount / $totalSystems) * 100, 2) : 0;
    return [
        'testedOkCount' => $testedOkCount,
        'totalSystems' => $totalSystems,
        'percentageTestedOk' => $percentageTestedOk
    ];
}

$response = ['success' => false, 'message' => ''];
$updateNeeded = false;

$requestedSystem = isset($_POST['system']) ? normalizeSystemName($_POST['system']) : null;
$requestedStatus = isset($_POST['status']) ? trim($_POST['status']) : null;

if (isset($_POST['action'])) {
    $action = $_POST['action'];

    if ($action === 'clear_status') {
        $highestRow = $sheet->getHighestDataRow();
        for ($row = 2; $row <= $highestRow; $row++) {
            $sheet->setCellValue('B' . $row, '');
        }
        $response['message'] = "All status har nollställts.";
        $updateNeeded = true;

    } elseif ($action === 'add_system') {
        if (empty($requestedSystem)) {
            $response['message'] = "Systemnamn kan inte vara tomt.";
        } elseif (mb_strlen($requestedSystem) > 200) {
            $response['message'] = "Systemnamn får max vara 200 tecken.";
        } else {
            $systemExists = false;
            $highestRow = $sheet->getHighestDataRow();
            for ($row = 1; $row <= $highestRow; $row++) {
                if (normalizeSystemName($sheet->getCell('A' . $row)->getValue()) === $requestedSystem) {
                    $systemExists = true;
                    break;
                }
            }

            if ($systemExists) {
                $response['message'] = "System '$requestedSystem' finns redan.";
            } else {
                $stats = getSystemStats($sheet);
                if ($stats['totalSystems'] >= 500) {
                    $response['message'] = "Max 500 system tillåtet.";
                } else {
                    $nextRow = $sheet->getHighestRow() + 1;
                    $sheet->setCellValue('A' . $nextRow, $requestedSystem);
                    $sheet->setCellValue('B' . $nextRow, '');
                    $response['message'] = "System '$requestedSystem' har lagts till.";
                    $updateNeeded = true;
                }
            }
        }

    } elseif ($action === 'remove_system') {
        if (empty($requestedSystem)) {
            $response['message'] = "Systemnamn kan inte vara tomt.";
        } else {
            $rowToRemove = null;
            $highestRow = $sheet->getHighestDataRow();
            for ($row = 1; $row <= $highestRow; $row++) {
                if (normalizeSystemName($sheet->getCell('A' . $row)->getValue()) === $requestedSystem) {
                    $rowToRemove = $row;
                    break;
                }
            }

            if ($rowToRemove) {
                if ($rowToRemove === 1) {
                    $response['message'] = "Kan inte ta bort rubrikraden.";
                } else {
                    $sheet->removeRow($rowToRemove);
                    $response['message'] = "System '$requestedSystem' har tagits bort.";
                    $updateNeeded = true;
                }
            } else {
                $response['message'] = "Systemet hittades inte.";
            }
        }
    }
} else {
    if (empty($requestedSystem) || !isset($requestedStatus)) {
        $response['message'] = "System eller status saknas.";
    } elseif (!in_array($requestedStatus, $allowedStatuses)) {
        $response['message'] = "Otillåtet statusvärde.";
    } else {
        $rowToUpdate = null;
        $highestRow = $sheet->getHighestDataRow();
        for ($row = 1; $row <= $highestRow; $row++) {
            if (normalizeSystemName($sheet->getCell('A' . $row)->getValue()) === $requestedSystem) {
                $rowToUpdate = $row;
                break;
            }
        }

        if ($rowToUpdate) {
            $sheet->setCellValue('B' . $rowToUpdate, $requestedStatus);
            $response['message'] = "Status uppdaterad för $requestedSystem.";
            $updateNeeded = true;
        } else {
            $response['message'] = "Systemet hittades inte.";
        }
    }
}

if ($updateNeeded) {
    try {
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('source.xlsx');
        $response['success'] = true;
    } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
        $response['success'] = false;
        $response['message'] = "Fel vid sparning: " . $e->getMessage();
    }
}

flock($lockFp, LOCK_UN);
fclose($lockFp);

$stats = getSystemStats($sheet);
$response = array_merge($response, $stats);

header('Content-Type: application/json');
echo json_encode($response);
?>

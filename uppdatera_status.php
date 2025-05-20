<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

$spreadsheet = IOFactory::load('source.xlsx');
$sheet = $spreadsheet->getActiveSheet();

/**
 * Normalizes a system name string for consistent comparison.
 * - Trims whitespace.
 * - Replaces multiple spaces with a single space.
 * - Decodes URL-encoded characters (useful if names were mistakenly saved encoded).
 */
function normalizeSystemName($name) {
    // 1. URL decode in case it was saved encoded in the Excel file
    $name = urldecode($name);
    // 2. Trim leading/trailing whitespace
    $name = trim($name);
    // 3. Replace multiple spaces with a single space
    $name = preg_replace('/\s+/', ' ', $name);
    // 4. Optionally, ensure consistent case for comparison (e.g., lowercase all)
    //    Commented out for now, but useful if "System A" should match "system a"
    // $name = mb_strtolower($name, 'UTF-8');
    return $name;
}

// Function to calculate and return statistics
function getSystemStats($sheet) {
    $testedOkCount = 0;
    $totalSystems = 0;
    $highestRow = $sheet->getHighestDataRow(); // Get the highest row with data
    $systemNamesInSheet = []; // To store normalized system names from sheet

    for ($row = 1; $row <= $highestRow; $row++) {
        $systemName = $sheet->getCell('A' . $row)->getValue();
        if (empty($systemName)) {
            continue; // Skip entirely empty rows in column A
        }
        $normalizedSystemName = normalizeSystemName($systemName);

        // Check if this normalized system name is already counted (e.g., header row)
        // This is important to avoid double counting if header is included in iteration
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
$updateNeeded = false; // Flag to indicate if we need to save the spreadsheet

// Normalize the incoming system name from GET request
$requestedSystem = isset($_GET['system']) ? normalizeSystemName($_GET['system']) : null;
$requestedStatus = isset($_GET['status']) ? normalizeSystemName($_GET['status']) : null; // Also normalize status

if (isset($_GET['action'])) {
    $action = $_GET['action'];

    if ($action === 'clear_status') {
        // Clear all status cells except potentially header row (row 1)
        $highestRow = $sheet->getHighestDataRow();
        for ($row = 2; $row <= $highestRow; $row++) { // Start from row 2 if row 1 is header
            $sheet->setCellValue('B' . $row, '');
        }
        $response['message'] = "All status cells have been cleared.";
        $updateNeeded = true;
    } elseif ($action === 'add_system') {
        if (empty($requestedSystem)) {
            $response['message'] = "System name cannot be empty.";
        } else {
            // Check if system already exists using normalized names
            $systemExists = false;
            $highestRow = $sheet->getHighestDataRow();
            for ($row = 1; $row <= $highestRow; $row++) {
                if (normalizeSystemName($sheet->getCell('A' . $row)->getValue()) === $requestedSystem) {
                    $systemExists = true;
                    break;
                }
            }

            if ($systemExists) {
                $response['message'] = "System '$requestedSystem' already exists.";
            } else {
                $nextRow = $sheet->getHighestRow() + 1; // getHighestRow gives the last populated row number
                $sheet->setCellValue('A' . $nextRow, $requestedSystem); // Write the normalized name
                $sheet->setCellValue('B' . $nextRow, ''); // Initial empty status
                $response['message'] = "System '$requestedSystem' has been added.";
                $updateNeeded = true;
            }
        }
    } elseif ($action === 'remove_system') {
        if (empty($requestedSystem)) {
            $response['message'] = "System name cannot be empty.";
        } else {
            $rowToRemove = null;
            $highestRow = $sheet->getHighestDataRow();
            for ($row = 1; $row <= $highestRow; $row++) {
                // Compare normalized names
                if (normalizeSystemName($sheet->getCell('A' . $row)->getValue()) === $requestedSystem) {
                    $rowToRemove = $row;
                    break;
                }
            }

            if ($rowToRemove) {
                if ($rowToRemove === 1) { // Assuming row 1 is your header and should not be removed
                    $response['message'] = "Cannot remove header row.";
                } else {
                    $sheet->removeRow($rowToRemove);
                    $response['message'] = "System '$requestedSystem' has been removed.";
                    $updateNeeded = true;
                }
            } else {
                $response['message'] = "System not found.";
            }
        }
    }
} else {
    // This is the "update status" path
    if (empty($requestedSystem) || !isset($requestedStatus)) {
        $response['message'] = "System or status not provided.";
    } else {
        $rowToUpdate = null;
        $highestRow = $sheet->getHighestDataRow();
        for ($row = 1; $row <= $highestRow; $row++) {
            // Compare normalized names
            if (normalizeSystemName($sheet->getCell('A' . $row)->getValue()) === $requestedSystem) {
                $rowToUpdate = $row;
                break;
            }
        }

        if ($rowToUpdate) {
            $sheet->setCellValue('B' . $rowToUpdate, $requestedStatus); // Write the normalized status
            $response['message'] = "Updating status for $requestedSystem to $requestedStatus.";
            $updateNeeded = true;
        } else {
            $response['message'] = "System not found.";
        }
    }
}

// Only save if an update was actually made
if ($updateNeeded) {
    try {
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('source.xlsx');
        $response['success'] = true;
    } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
        $response['success'] = false;
        $response['message'] = "Error saving file: " . $e->getMessage();
    }
} else {
    // If no update was made, but it's a valid request (e.g., trying to add an existing system)
    // still return success true if no specific error message implies failure.
    if (!isset($response['success'])) {
         $response['success'] = true;
    }
}

// After any action, re-calculate and include the latest stats
$stats = getSystemStats($sheet);
$response = array_merge($response, $stats);

header('Content-Type: application/json');
echo json_encode($response);
?>

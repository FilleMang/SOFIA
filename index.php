<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

$spreadsheet = IOFactory::load('source.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$data = array();
$testedOkCount = 0;
$totalSystems = 0;

foreach ($sheet->getRowIterator() as $row) {
    $system = $sheet->getCell('A' . $row->getRowIndex())->getValue();
    if (empty($system)) {
        continue;
    }

    $status = $sheet->getCell('B' . $row->getRowIndex())->getValue();
    $data[] = [
        'system' => $system,
        'status' => $status,
    ];

    $totalSystems++;
    if ($status == 'Testad OK') {
        $testedOkCount++;
    }
}

$percentageTestedOk = ($totalSystems > 0) ? round(($testedOkCount / $totalSystems) * 100, 2) : 0;
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SOFIA - System för Operativ Förvaltning, Incidentöversikt och Analys</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <style>
        .section { margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        button { margin-right: 5px; }
        .stats-container { text-align: right; margin-top: 20px; }
    </style>
    <script>
    $(document).ready(function() {
        function updateStatsDisplay(stats) {
            $('#tested-ok-count').text(stats.testedOkCount);
            $('#total-systems').text(stats.totalSystems);
            $('#percentage-tested-ok').text(stats.percentageTestedOk + '%');
        }

        $('.update-status-btn').on('click', function() {
            var systemIdSafe = $(this).data('system-id-safe'); // Get the *safe* ID from the button
            var newStatus = $(this).data('status');
            var systemToUpdateInPHP = $(this).data('system'); // Get the *decoded* name for PHP (e.g., "My System")

            console.log("--- Update Button Clicked ---");
            console.log("systemIdSafe (for DOM update):", systemIdSafe);
            console.log("systemToUpdateInPHP (for AJAX):", systemToUpdateInPHP);
            console.log("newStatus:", newStatus);

            $.ajax({
                url: 'uppdatera_status.php',
                type: 'GET',
                data: { system: systemToUpdateInPHP, status: newStatus }, // Send the decoded name to PHP
                dataType: 'json',
                success: function(response) {
                    console.log("AJAX Success Response:", response);
                    if (response.success) {
                        // Select the <td> using the generated safe ID
                        $('#status-cell-' + systemIdSafe).text(newStatus); 
                        
                        updateStatsDisplay(response);
                    } else {
                        console.error('Server error: ' + response.message);
                    }
                },
                error: function(xhr, status, error) {
                    console.error('AJAX Error:', error, xhr, status);
                }
            });
        });

        $('#clear-status-btn').on('click', function() {
            $.ajax({
                url: 'uppdatera_status.php',
                type: 'GET',
                data: { action: 'clear_status' },
                dataType: 'json',
                success: function(response) {
                    console.log("AJAX Success Response (Clear):", response);
                    if (response.success) {
                        $('.status-cell').text('Okänd');
                        updateStatsDisplay(response);
                    } else {
                        console.error('Server error: ' + response.message);
                    }
                },
                error: function(xhr, status, error) {
                    console.error('AJAX Error (Clear):', error, xhr, status);
                }
            });
        });

        $('#add-system-form').on('submit', function(e) {
            e.preventDefault();
            var system = $('#new-system').val();

            console.log("--- Add System Clicked ---");
            console.log("New System Name:", system);

            $.ajax({
                url: 'uppdatera_status.php',
                type: 'GET',
                data: { action: 'add_system', system: system },
                dataType: 'json',
                success: function(response) {
                    console.log("AJAX Success Response (Add):", response);
                    alert(response.message);
                    if (response.success) {
                        location.reload(); // Reload the page to reflect changes in the table
                    }
                },
                error: function(xhr, status, error) {
                    console.error('AJAX Error (Add):', error, xhr, status);
                }
            });
        });

        $('.remove-system-btn').on('click', function() {
            var systemToRemoveInPHP = $(this).data('system'); // Decoded system name for PHP
            var systemIdSafe = $(this).data('system-id-safe'); // Safe ID for DOM removal

            console.log("--- Remove Button Clicked ---");
            console.log("systemToRemoveInPHP (for AJAX):", systemToRemoveInPHP);
            console.log("systemIdSafe (for DOM removal):", systemIdSafe);

            $.ajax({
                url: 'uppdatera_status.php',
                type: 'GET',
                data: { action: 'remove_system', system: systemToRemoveInPHP },
                dataType: 'json',
                success: function(response) {
                    console.log("AJAX Success Response (Remove):", response);
                    alert(response.message);
                    if (response.success) {
                        // Remove the row using the safe ID
                        $('#status-cell-' + systemIdSafe).closest('tr').remove();
                        updateStatsDisplay(response);
                    }
                },
                error: function(xhr, status, error) {
                    console.error('AJAX Error (Remove):', error, xhr, status);
                }
            });
        });
    });
    </script>
</head>
<body>
    <div class="section">
        <table>
            <tr>
                <th>System</th>
                <th>Status</th>
                <th>Action</th>
            </tr>
            <?php foreach ($data as $row) {
                // Generate a safe ID for HTML: replace non-alphanumeric (except underscore and hyphen) characters with underscores.
                // This creates IDs like "My_System" which are perfectly valid for HTML IDs and jQuery selectors.
                $safeSystemId = preg_replace('/[^a-zA-Z0-9_\-]/', '_', $row['system']);
                // Trim any leading/trailing underscores that might result from the replacement
                $safeSystemId = trim($safeSystemId, '_');
                // Fallback for very strange system names that might result in empty safe ID
                if (empty($safeSystemId)) {
                    $safeSystemId = 'system_' . md5($row['system']); // Use a hash as a unique fallback
                }
            ?>
            <tr>
                <td><?php echo htmlspecialchars($row['system']); ?></td>
                <td class="status-cell" id="status-cell-<?php echo $safeSystemId; ?>" data-system="<?php echo htmlspecialchars($row['system']); ?>">
                    <?php echo empty($row['status']) ? 'Okänd' : htmlspecialchars($row['status']); ?>
                </td>
                <td>
                    <button class="update-status-btn" data-system="<?php echo htmlspecialchars($row['system']); ?>" data-status="Testad OK" data-system-id-safe="<?php echo $safeSystemId; ?>">Testad OK</button>
                    <button class="update-status-btn" data-system="<?php echo htmlspecialchars($row['system']); ?>" data-status="Inte testad" data-system-id-safe="<?php echo $safeSystemId; ?>">Inte testad</button>
                    <button class="remove-system-btn" data-system="<?php echo htmlspecialchars($row['system']); ?>" data-system-id-safe="<?php echo $safeSystemId; ?>">X</button>
                </td>
            </tr>
            <?php } ?>
        </table>
    </div>

    <div class="section">
        <form id="add-system-form" style="display: inline-block;">
            <input type="text" id="new-system" placeholder="Lägg till system" required>
            <button type="submit">Lägg till system</button>
        </form>
        <button id="clear-status-btn" style="display: inline-block; margin-left: 10px;">Nollställ status</button>
    </div>

    <div class="stats-container">
        Antal "Testad OK" system: <span id="tested-ok-count"><?php echo $testedOkCount; ?></span> / <span id="total-systems"><?php echo $totalSystems; ?></span> (<span id="percentage-tested-ok"><?php echo $percentageTestedOk; ?>%</span>)
    </div>
</body>
</html>

<?php
session_start();
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

// ── CSRF-token ──
if (empty($_SESSION['csrf_token'])) {
    $_SESSION['csrf_token'] = bin2hex(random_bytes(32));
}
$csrfToken = $_SESSION['csrf_token'];

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
    <script src="https://code.jquery.com/jquery-3.7.1.min.js" integrity="sha256-/JqT3SQfawRcv/BIHPThkBvs0OEvtFFmqPF/lYI/Cxo=" crossorigin="anonymous"></script>
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
        var csrfToken = <?php echo json_encode($csrfToken); ?>;

        function updateStatsDisplay(stats) {
            $('#tested-ok-count').text(stats.testedOkCount);
            $('#total-systems').text(stats.totalSystems);
            $('#percentage-tested-ok').text(stats.percentageTestedOk + '%');
        }

        $('.update-status-btn').on('click', function() {
            var systemIdSafe = $(this).data('system-id-safe');
            var newStatus = $(this).data('status');
            var systemToUpdateInPHP = $(this).data('system');

            $.ajax({
                url: 'uppdatera_status.php',
                type: 'POST',
                data: { system: systemToUpdateInPHP, status: newStatus, csrf_token: csrfToken },
                dataType: 'json',
                success: function(response) {
                    if (response.success) {
                        $('#status-cell-' + systemIdSafe).text(newStatus);
                        updateStatsDisplay(response);
                    } else {
                        alert('Fel: ' + response.message);
                    }
                },
                error: function(xhr, status, error) {
                    console.error('AJAX Error:', error);
                }
            });
        });

        $('#clear-status-btn').on('click', function() {
            if (!confirm('Nollställ all status?')) return;
            $.ajax({
                url: 'uppdatera_status.php',
                type: 'POST',
                data: { action: 'clear_status', csrf_token: csrfToken },
                dataType: 'json',
                success: function(response) {
                    if (response.success) {
                        $('.status-cell').text('Okänd');
                        updateStatsDisplay(response);
                    } else {
                        alert('Fel: ' + response.message);
                    }
                },
                error: function(xhr, status, error) {
                    console.error('AJAX Error:', error);
                }
            });
        });

        $('#add-system-form').on('submit', function(e) {
            e.preventDefault();
            var system = $('#new-system').val();

            $.ajax({
                url: 'uppdatera_status.php',
                type: 'POST',
                data: { action: 'add_system', system: system, csrf_token: csrfToken },
                dataType: 'json',
                success: function(response) {
                    alert(response.message);
                    if (response.success) {
                        location.reload();
                    }
                },
                error: function(xhr, status, error) {
                    console.error('AJAX Error:', error);
                }
            });
        });

        $('.remove-system-btn').on('click', function() {
            var systemToRemoveInPHP = $(this).data('system');
            var systemIdSafe = $(this).data('system-id-safe');

            if (!confirm('Ta bort "' + systemToRemoveInPHP + '"?')) return;

            $.ajax({
                url: 'uppdatera_status.php',
                type: 'POST',
                data: { action: 'remove_system', system: systemToRemoveInPHP, csrf_token: csrfToken },
                dataType: 'json',
                success: function(response) {
                    if (response.success) {
                        $('#status-cell-' + systemIdSafe).closest('tr').remove();
                        updateStatsDisplay(response);
                    } else {
                        alert('Fel: ' + response.message);
                    }
                },
                error: function(xhr, status, error) {
                    console.error('AJAX Error:', error);
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
                $safeSystemId = preg_replace('/[^a-zA-Z0-9_\-]/', '_', $row['system']);
                $safeSystemId = trim($safeSystemId, '_');
                if (empty($safeSystemId)) {
                    $safeSystemId = 'system_' . md5($row['system']);
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
            <input type="text" id="new-system" placeholder="Lägg till system" required maxlength="200">
            <button type="submit">Lägg till system</button>
        </form>
        <button id="clear-status-btn" style="display: inline-block; margin-left: 10px;">Nollställ status</button>
    </div>

    <div class="stats-container">
        Antal "Testad OK" system: <span id="tested-ok-count"><?php echo $testedOkCount; ?></span> / <span id="total-systems"><?php echo $totalSystems; ?></span> (<span id="percentage-tested-ok"><?php echo $percentageTestedOk; ?>%</span>)
    </div>
</body>
</html>

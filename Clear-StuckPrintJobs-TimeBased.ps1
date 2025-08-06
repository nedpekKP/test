# Clear-StuckPrintJobs-TimeBased.ps1
# Automatically clears the print queue if 4 or more jobs are stuck (older than X minutes)

# Set threshold in minutes
$minutesThreshold = 5
$minStuckJobs = 2
$now = Get-Date
$stuckJobs = @()

# Get all jobs older than threshold
$printers = Get-Printer -ErrorAction SilentlyContinue
foreach ($printer in $printers) {
    $jobs = Get-PrintJob -PrinterName $printer.Name -ErrorAction SilentlyContinue
    foreach ($job in $jobs) {
        if ($job.TimeSubmitted -lt $now.AddMinutes(-$minutesThreshold)) {
            $stuckJobs += $job
        }
    }
}

if ($stuckJobs.Count -ge $minStuckJobs) {
    Write-Host "Detected $($stuckJobs.Count) stuck print jobs older than $minutesThreshold minutes. Clearing spooler..."

    try {
        Stop-Service -Name Spooler -Force
        Start-Sleep -Seconds 3

        Remove-Item -Path "$env:SystemRoot\System32\spool\PRINTERS\*.*" -Force -ErrorAction SilentlyContinue

        Start-Sleep -Seconds 2
        Start-Service -Name Spooler

        Write-Host "Spooler restarted and jobs cleared."
    } catch {
        Write-Host "Failed to restart spooler or clear jobs: $_"
    }
} else {
    Write-Host "Only $($stuckJobs.Count) old jobs found. No action taken."
}
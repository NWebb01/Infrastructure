# Define the Hyper-V host (use 'localhost' for local host)
$HyperVHost = "localhost"

# Define output path
$OutputCsvPath = "C:\Temp\HyperV_Host_Metrics.csv"

# Create an object to hold the metrics
$hostMetrics = [PSCustomObject]@{
    Timestamp           = (Get-Date).ToString("s")  # ISO 8601 format
    HostName            = $null
    OSVersion           = $null
    CPUCount            = $null
    ProcessorType       = $null
    AvgCPUUsagePercent  = $null
    TotalMemoryGB       = $null
    FreeMemoryGB        = $null
    UsedMemoryGB        = $null
}

# Try to get metrics
try {
    $os   = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $HyperVHost
    $comp = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $HyperVHost
    $cpu  = Get-CimInstance -ClassName Win32_Processor -ComputerName $HyperVHost

    $hostMetrics.HostName      = $comp.Name
    $hostMetrics.OSVersion     = $os.Caption
    $hostMetrics.ProcessorType = $cpu[0].Name.Trim()
    $hostMetrics.CPUCount      = $cpu.Count

    # Memory
    $totalMemoryGB = [math]::Round($comp.TotalPhysicalMemory / 1GB, 2)
    $freeMemoryGB  = [math]::Round($os.FreePhysicalMemory * 1KB / 1GB, 2)
    $usedMemoryGB  = [math]::Round($totalMemoryGB - $freeMemoryGB, 2)

    $hostMetrics.TotalMemoryGB = $totalMemoryGB
    $hostMetrics.FreeMemoryGB  = $freeMemoryGB
    $hostMetrics.UsedMemoryGB  = $usedMemoryGB

    # CPU Load
    $cpuLoads = $cpu | Select-Object -ExpandProperty LoadPercentage
    if ($cpuLoads.Count -gt 0) {
        $hostMetrics.AvgCPUUsagePercent = [math]::Round(($cpuLoads | Measure-Object -Average).Average, 2)
    }

    # Export to CSV
    if (-not (Test-Path $OutputCsvPath)) {
        # If the file doesn't exist, create it with headers
        $hostMetrics | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
    } else {
        # Append without headers
        $hostMetrics | Export-Csv -Path $OutputCsvPath -Append -NoTypeInformation -Encoding UTF8
    }

    Write-Host "Metrics saved to: $OutputCsvPath"
}
catch {
    Write-Error "Failed to retrieve metrics from host '$HyperVHost': $_"
}

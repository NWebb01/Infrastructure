# Define the Hyper-V host (use 'localhost' for local host)
$HyperVHost = "localhost"

# Create an object to hold results
$hostMetrics = [PSCustomObject]@{
    HostName           = $null
    OSVersion          = $null
    CPUCount           = $null
    AvgCPUUsagePercent = $null
    TotalMemoryGB      = $null
    FreeMemoryGB       = $null
    UsedMemoryGB       = $null
}

# Connect to the host using CIM (works locally and remotely)
try {
    # Get basic OS info
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $HyperVHost
    $comp = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $HyperVHost

    $hostMetrics.HostName = $comp.Name
    $hostMetrics.OSVersion = $os.Caption

    # Get memory info
    $totalMemoryGB = [math]::Round($comp.TotalPhysicalMemory / 1GB, 2)
    $freeMemoryGB = [math]::Round($os.FreePhysicalMemory * 1KB / 1GB, 2)
    $usedMemoryGB = [math]::Round($totalMemoryGB - $freeMemoryGB, 2)

    $hostMetrics.TotalMemoryGB = $totalMemoryGB
    $hostMetrics.FreeMemoryGB = $freeMemoryGB
    $hostMetrics.UsedMemoryGB = $usedMemoryGB

    # Get CPU load (average across all CPUs)
    $cpuLoadSamples = Get-CimInstance -ClassName Win32_Processor -ComputerName $HyperVHost |
        Select-Object -ExpandProperty LoadPercentage

    if ($cpuLoadSamples.Count -gt 0) {
        $avgLoad = [math]::Round(($cpuLoadSamples | Measure-Object -Average).Average, 2)
        $hostMetrics.AvgCPUUsagePercent = $avgLoad
        $hostMetrics.CPUCount = $cpuLoadSamples.Count
    }

    # Output the results
    $hostMetrics | Format-List

} catch {
    Write-Error "Failed to connect to Hyper-V host '$HyperVHost': $_"
}

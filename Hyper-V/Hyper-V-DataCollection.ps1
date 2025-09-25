# Requires ImportExcel module
# Install if needed:
# Install-Module ImportExcel -Scope CurrentUser

$HyperVHost = "localhost"
$OutputExcelPath = "C:\Temp\HyperV_Metrics.xlsx"

# --- Collect Host Metrics ---
$hostMetrics = [PSCustomObject]@{
    Timestamp           = (Get-Date).ToString("s")
    HostName            = $null
    OSVersion           = $null
    CPUCount            = $null
    ProcessorType       = $null
    AvgCPUUsagePercent  = $null
    TotalMemoryGB       = $null
    FreeMemoryGB        = $null
    UsedMemoryGB        = $null
    Manufacturer        = $null
    Model               = $null
}

try {
    $os   = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $HyperVHost
    $comp = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $HyperVHost
    $cpu  = Get-CimInstance -ClassName Win32_Processor -ComputerName $HyperVHost

    $hostMetrics.HostName      = $comp.Name
    $hostMetrics.OSVersion     = $os.Caption
    $hostMetrics.ProcessorType = $cpu[0].Name.Trim()
    $hostMetrics.CPUCount      = $cpu.Count

    $hostMetrics.Manufacturer  = $comp.Manufacturer
    $hostMetrics.Model         = $comp.Model

    $totalMemoryGB = [math]::Round($comp.TotalPhysicalMemory / 1GB, 2)
    $freeMemoryGB  = [math]::Round($os.FreePhysicalMemory * 1KB / 1GB, 2)
    $usedMemoryGB  = [math]::Round($totalMemoryGB - $freeMemoryGB, 2)

    $hostMetrics.TotalMemoryGB = $totalMemoryGB
    $hostMetrics.FreeMemoryGB  = $freeMemoryGB
    $hostMetrics.UsedMemoryGB  = $usedMemoryGB

    $cpuLoads = $cpu | Select-Object -ExpandProperty LoadPercentage
    if ($cpuLoads.Count -gt 0) {
        $hostMetrics.AvgCPUUsagePercent = [math]::Round(($cpuLoads | Measure-Object -Average).Average, 2)
    }
}
catch {
    Write-Error "Failed to retrieve Host metrics: $_"
}

# --- Collect VM Metrics ---
$vmMetrics = @()

try {
    $vms = Get-VM -ComputerName $HyperVHost
    foreach ($vm in $vms) {
        $metrics = [PSCustomObject]@{
            Timestamp          = (Get-Date).ToString("s")
            VMName             = $vm.Name
            State              = $vm.State
            CPUUsagePercent    = $vm.CPUUsage
            MemoryAssignedGB   = [math]::Round($vm.MemoryAssigned / 1GB, 2)
            Uptime             = $vm.Uptime
        }
        $vmMetrics += $metrics
    }
}
catch {
    Write-Error "Failed to retrieve VM metrics: $_"
}

# --- Export both to Excel ---
# Host metrics in one sheet
$hostMetrics | Export-Excel -Path $OutputExcelPath -WorksheetName "HostMetrics" -AutoSize -AutoFilter -ClearSheet

# VM metrics in another sheet
if ($vmMetrics.Count -gt 0) {
    $vmMetrics | Export-Excel -Path $OutputExcelPath -WorksheetName "VMMetrics" -AutoSize -AutoFilter
}

Write-Host "Metrics saved to Excel file: $OutputExcelPath"

<#
.SYNOPSIS
  Collects Windows Server inventory across many servers, and writes ONE Excel workbook (.xlsx)
  with four worksheets (Summary, NICs, Disks, Volumes) using Excel COM automation—no external modules.

.DESCRIPTION
  - OS: Caption, Version, Build
  - Hardware: Manufacturer, Model, Serial Number
  - CPU: Model(s), Physical sockets, Total cores
  - Memory: Total physical memory in GB
  - Performance (30d averages if available, else short live sample):
      * CPU_AvgPct_30d  -> \Processor(_Total)\% Processor Time
      * Mem_AvgPct_30d  -> \Memory\% Committed Bytes In Use
      * Perf_Method     -> 'PerfLogs (UNC)' | 'ShortSample Ns' | 'Unavailable'
  - NICs (physical only): Vendor, Model, Interface, MAC, Link speed (Gbps), Status
  - Disks (physical): Index, Model, Interface, Serial, SizeGB
  - Volumes (fixed): Label, DriveLetter, Path, FileSystem, CapacityGB, FreeGB, UsedGB, IsSystemVolume

.PARAMETER ComputerName
  One or more target servers (NetBIOS, FQDN, or IP). You can pipe input or pass an array.

.PARAMETER Credential
  Optional credential for CIM/WMI, SMB admin shares (for PerfLogs), and remoting (for short live sample).

.PARAMETER OutputFolder
  Folder where the Excel file will be saved. Defaults to current directory.

.PARAMETER ExcelPath
  Full path to the output .xlsx. If omitted, a timestamped name is generated in OutputFolder.

.PARAMETER PerfLookbackDays
  Days to look back for perf logs (default 30).

.PARAMETER PerfSearchPaths
  Subpaths under C:\ to search for perf logs (default 'PerfLogs','PerfLogs\Admin').

.PARAMETER UseShortSampleFallback
  If set, performs a short live sample (default 15s) when perf logs are not found.

.PARAMETER ShortSampleSeconds
  Seconds for short live sample (min 3, max 300). Default 15.

.PARAMETER ExcelVisible
  If set, Excel UI is shown during export (useful for debugging). Hidden by default.

.NOTES
  - Requires Excel to be installed on the machine running this script (COM automation).
  - No external modules required.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('CN','Server')]
    [string[]]$ComputerName,

    [Parameter()]
    [pscredential]$Credential,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder = ".",

    [Parameter()]
    [string]$ExcelPath,

    [Parameter()]
    [ValidateRange(1,365)]
    [int]$PerfLookbackDays = 30,

    [Parameter()]
    [string[]]$PerfSearchPaths = @('PerfLogs','PerfLogs\Admin'),

    [Parameter()]
    [switch]$UseShortSampleFallback = $true,

    [Parameter()]
    [ValidateRange(3,300)]
    [int]$ShortSampleSeconds = 15,

    [Parameter()]
    [switch]$ExcelVisible
)

begin {
    function New-InventoryCimSession {
        param([string]$Computer, [pscredential]$Cred)
        try {
            $wsmanOpt = New-CimSessionOption -Protocol Wsman
            return New-CimSession -ComputerName $Computer -Credential $Cred -SessionOption $wsmanOpt -ErrorAction Stop
        } catch {
            Write-Verbose "WSMan failed for $Computer ($($_.Exception.Message)). Trying DCOM..."
            try {
                $dcomOpt = New-CimSessionOption -Protocol Dcom
                return New-CimSession -ComputerName $Computer -Credential $Cred -SessionOption $dcomOpt -ErrorAction Stop
            } catch {
                Write-Warning "Failed to create CIM session to $Computer via WSMan and DCOM: $($_.Exception.Message)"
                return $null
            }
        }
    }

    function Convert-BitsPerSecondToGbps { param([nullable[uint64]]$bps)
        if (-not $bps) { return $null }
        [math]::Round(($bps / 1e9), 3)
    }

    # Get 30-day averages from PerfLogs if available; else do short live sample
    function Get-PerfAverages {
        param(
            [string]$Computer,
            [int]$LookbackDays = 30,
            [string[]]$SearchPaths = @('PerfLogs','PerfLogs\Admin'),
            [pscredential]$Cred,
            [switch]$UseShortSampleFallback,
            [int]$ShortSampleSeconds = 15
        )

        $start = (Get-Date).AddDays(-$LookbackDays)
        $rootUNC = "\\$Computer\c$"
        $driveName = "INV_$([System.Guid]::NewGuid().ToString('N').Substring(0,6))"
        $mapped = $false

        # --- Try reading PerfLogs via admin share ---
        try {
            if ($Cred) {
                New-PSDrive -Name $driveName -PSProvider FileSystem -Root $rootUNC -Credential $Cred -ErrorAction Stop | Out-Null
            } else {
                New-PSDrive -Name $driveName -PSProvider FileSystem -Root $rootUNC -ErrorAction Stop | Out-Null
            }
            $mapped = $true

            $files = @()
            foreach ($sp in $SearchPaths) {
                # FIX: brace variable to avoid "$driveName:" parsing error
                $p = "${driveName}:\$sp"
                if (Test-Path -LiteralPath $p) {
                    # Use -Recurse and filter by extension/time
                    $files += Get-ChildItem -Path $p -File -Recurse -ErrorAction SilentlyContinue |
                              Where-Object { $_.LastWriteTime -ge $start -and $_.Extension -in '.blg','.csv' }
                }
            }
            if ($files -and $files.Count -gt 0) {
                $paths = $files.FullName
                $data = Import-Counter -Path $paths -StartTime $start -ErrorAction Stop
                $cpuVals = $data.CounterSamples | Where-Object { $_.Path -match '\\Processor\(_Total\)\\% Processor Time' } | Select-Object -ExpandProperty CookedValue
                $memVals = $data.CounterSamples | Where-Object { $_.Path -match '\\Memory\\% Committed Bytes In Use' }          | Select-Object -ExpandProperty CookedValue
                $cpuAvg = if ($cpuVals) { [math]::Round( ($cpuVals | Measure-Object -Average).Average , 2) } else { $null }
                $memAvg = if ($memVals) { [math]::Round( ($memVals | Measure-Object -Average).Average , 2) } else { $null }
                if ($cpuAvg -ne $null -or $memAvg -ne $null) {
                    return [pscustomobject]@{ CpuAvg=$cpuAvg; MemAvg=$memAvg; Method='PerfLogs (UNC)' }
                }
            }
        } catch {
            # ignore; fallbacks below
        } finally {
            if ($mapped) { Remove-PSDrive -Name $driveName -Force -ErrorAction SilentlyContinue }
        }

        # --- Fallback: short live sample via Get-Counter / Invoke-Command ---
        if ($UseShortSampleFallback) {
            try {
                $samples = [math]::Max(3, [math]::Min(60, $ShortSampleSeconds))
                if ($Cred) {
                    $sb = {
                        param($samples)
                        Get-Counter -Counter '\Processor(_Total)\% Processor Time','\Memory\% Committed Bytes In Use' `
                                    -SampleInterval 1 -MaxSamples $samples
                    }
                    $c = Invoke-Command -ComputerName $Computer -Credential $Cred -ScriptBlock $sb -ArgumentList $samples -ErrorAction Stop
                } else {
                    $c = Get-Counter -ComputerName $Computer -Counter '\Processor(_Total)\% Processor Time','\Memory\% Committed Bytes In Use' `
                                     -SampleInterval 1 -MaxSamples $samples -ErrorAction Stop
                }
                $cpuVals = $c.CounterSamples | Where-Object { $_.Path -match '\\Processor\(_Total\)\\% Processor Time' } | Select-Object -ExpandProperty CookedValue
                $memVals = $c.CounterSamples | Where-Object { $_.Path -match '\\Memory\\% Committed Bytes In Use' }          | Select-Object -ExpandProperty CookedValue
                $cpuAvg = if ($cpuVals) { [math]::Round( ($cpuVals | Measure-Object -Average).Average , 2) } else { $null }
                $memAvg = if ($memVals) { [math]::Round( ($memVals | Measure-Object -Average).Average , 2) } else { $null }
                return [pscustomobject]@{ CpuAvg=$cpuAvg; MemAvg=$memAvg; Method="ShortSample ${samples}s" }
            } catch {
                # ignore
            }
        }

        return [pscustomobject]@{ CpuAvg=$null; MemAvg=$null; Method='Unavailable' }
    }

    # Efficiently write a collection to a worksheet using a 2D array (fast; avoids cell-by-cell)
    function Write-Worksheet {
        param(
            [Parameter(Mandatory=$true)] $Worksheet,
            [Parameter(Mandatory=$true)] [System.Collections.IEnumerable] $Rows,
            [Parameter(Mandatory=$true)] [string[]] $Columns,
            [string]$TableName  # optional
        )

        # Convert to array for counting
        $arr = @($Rows)
        $rowCount = $arr.Count
        $colCount = $Columns.Count

        # Header row
        for ($c=0; $c -lt $colCount; $c++) {
            $Worksheet.Cells.Item(1, $c+1) = $Columns[$c]
        }

        if ($rowCount -gt 0) {
            # Build a 2D array [rows, cols]
            $data = New-Object 'object[,]' $rowCount, $colCount
            for ($r=0; $r -lt $rowCount; $r++) {
                $obj = $arr[$r]
                for ($c=0; $c -lt $colCount; $c++) {
                    $val = $null
                    try { $val = $obj.($Columns[$c]) } catch { $val = $null }
                    if ($null -eq $val) { $val = '' }
                    $data[$r, $c] = $val
                }
            }
            $start = $Worksheet.Cells.Item(2,1)
            $end   = $Worksheet.Cells.Item($rowCount+1, $colCount)
            $Worksheet.Range($start, $end).Value2 = $data

            # Optional: turn into a Table (ListObject) for filtering & header styling
            try {
                $used = $Worksheet.UsedRange
                # xlSrcRange=1; xlYes=1
                $listObj = $Worksheet.ListObjects.Add(1, $used, $null, 1)
                if ($TableName) { $listObj.Name = $TableName }
                # Apply a built-in style if available
                $listObj.TableStyle = 'TableStyleMedium9'
            } catch {
                # If table creation fails (older Excel, protected, etc.), continue
            }
        }

        # Header formatting & freeze
        $Worksheet.Rows.Item(1).Font.Bold = $true
        try { $Worksheet.Rows.Item(1).Interior.Color = 0xEEEEEE } catch { }

        # Freeze top row: select A2 then freeze panes
        $Worksheet.Activate() | Out-Null
        $Worksheet.Cells.Item(2,1).Select() | Out-Null
        $Worksheet.Application.ActiveWindow.FreezePanes = $true

        # Autofit columns
        $Worksheet.UsedRange.Columns.AutoFit() | Out-Null
    }

    function Remove-ComObjectSafely {
        [CmdletBinding(SupportsShouldProcess=$false)]
        param(
            [Parameter(ValueFromPipeline=$true)]
            $ComObj
        )
        try {
            if ($null -ne $ComObj -and ($ComObj -isnot [int])) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObj)
            }
        } catch {
            # Swallow any release errors; GC will clean up
        }
    }

    # Data buckets
    $serverSummaries = New-Object System.Collections.Generic.List[object]
    $nicDetails      = New-Object System.Collections.Generic.List[object]
    $diskDetails     = New-Object System.Collections.Generic.List[object]
    $volumeDetails   = New-Object System.Collections.Generic.List[object]

    # Output path prep
    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    if (-not $ExcelPath -or [string]::IsNullOrWhiteSpace($ExcelPath)) {
        $ExcelPath = Join-Path $OutputFolder ("ServerInventory_{0}.xlsx" -f $timestamp)
    } elseif (-not (Split-Path $ExcelPath -IsAbsolute)) {
        $ExcelPath = Join-Path $OutputFolder $ExcelPath
    }
}

process {
    foreach ($computer in $ComputerName) {
        Write-Host ">>> Collecting inventory from $computer ..." -ForegroundColor Cyan
        $cs = New-InventoryCimSession -Computer $computer -Cred $Credential
        if (-not $cs) {
            $serverSummaries.Add([pscustomobject]@{
                ComputerName            = $computer
                Reachable               = $false
                OS_Caption              = $null
                OS_Version              = $null
                OS_Build                = $null
                HW_Manufacturer         = $null
                HW_Model                = $null
                HW_SerialNumber         = $null
                CPU_Model               = $null
                CPU_PhysicalSockets     = $null
                CPU_TotalCores          = $null
                Mem_TotalGB             = $null
                CPU_AvgPct_30d          = $null
                Mem_AvgPct_30d          = $null
                Perf_Method             = 'Unavailable'
                NIC_Count               = $null
                Disk_PhysicalCount      = $null
                Disk_RawPhysicalGB      = $null
                Vol_FixedCount          = $null
                Vol_TotalCapacityGB     = $null
                Vol_UsedGB              = $null
                Vol_FreeGB              = $null
            })
            continue
        }

        try {
            # OS & HW
            $os   = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $cs -ErrorAction Stop
            $csys = Get-CimInstance -ClassName Win32_ComputerSystem   -CimSession $cs -ErrorAction Stop

            $osCaption = $os.Caption
            $osVersion = $os.Version
            $osBuild   = $os.BuildNumber
            $manu      = $csys.Manufacturer
            $model     = $csys.Model

            # Serial Number: preferred from BIOS; fallback to ComputerSystemProduct
            $serial = $null
            try {
                $bios = Get-CimInstance -ClassName Win32_BIOS -CimSession $cs -ErrorAction Stop
                $serial = $bios.SerialNumber
            } catch { }
            if (-not $serial) {
                try {
                    $csp = Get-CimInstance -ClassName Win32_ComputerSystemProduct -CimSession $cs -ErrorAction Stop
                    $serial = $csp.IdentifyingNumber
                } catch { }
            }

            # Memory GB
            $memTotalGB = if ($csys.TotalPhysicalMemory) { [math]::Round(($csys.TotalPhysicalMemory / 1GB), 2) } else { $null }

            # CPU
            $procs         = Get-CimInstance -ClassName Win32_Processor -CimSession $cs -ErrorAction Stop
            $socketCount   = ($procs | Measure-Object).Count
            $totalCores    = ($procs | Measure-Object -Property NumberOfCores -Sum).Sum
            $cpuModels     = ($procs | Select-Object -ExpandProperty Name -Unique) -join '; '

            # NICs (physical)
            $msftNics = @()
            $wmiNics  = @()
            try {
                $msftNics = Get-CimInstance -Namespace root/StandardCimv2 -ClassName MSFT_NetAdapter -CimSession $cs -ErrorAction Stop |
                            Where-Object { $_.HardwareInterface -and -not $_.Virtual -and -not $_.Hidden }
            } catch { }
            $wmiNics = Get-CimInstance -ClassName Win32_NetworkAdapter -CimSession $cs -ErrorAction SilentlyContinue |
                       Where-Object { $_.PhysicalAdapter -eq $true }

            $nicCount = 0
            if ($msftNics -and $msftNics.Count -gt 0) {
                $wmiByPNP = @{}
                foreach ($w in $wmiNics) { if ($w.PNPDeviceID) { $wmiByPNP[$w.PNPDeviceID] = $w } }

                foreach ($n in $msftNics) {
                    $pnp = $n.PnPDeviceID
                    $man = $null
                    if ($pnp -and $wmiByPNP.ContainsKey($pnp)) {
                        $man = $wmiByPNP[$pnp].Manufacturer
                    } elseif ($wmiNics) {
                        $match = $wmiNics | Where-Object {
                            ($_.MACAddress -and $_.MACAddress -eq $n.MacAddress) -or
                            ($_.Name -eq $n.InterfaceDescription)
                        } | Select-Object -First 1
                        $man = $match.Manufacturer
                    }

                    $nicDetails.Add([pscustomobject]@{
                        ComputerName       = $computer
                        Vendor             = $man
                        Model              = $n.DriverDescription
                        Interface          = $n.InterfaceDescription
                        MACAddress         = $n.MacAddress
                        LinkSpeedGbps      = Convert-BitsPerSecondToGbps -bps $n.LinkSpeed
                        Status             = $n.Status
                    })
                }
                $nicCount = $msftNics.Count
            } elseif ($wmiNics) {
                foreach ($n in $wmiNics) {
                    $nicDetails.Add([pscustomobject]@{
                        ComputerName       = $computer
                        Vendor             = $n.Manufacturer
                        Model              = $n.ProductName
                        Interface          = $n.Name
                        MACAddress         = $n.MACAddress
                        LinkSpeedGbps      = Convert-BitsPerSecondToGbps -bps $n.Speed
                        Status             = $n.NetConnectionStatus
                    })
                }
                $nicCount = ($wmiNics | Measure-Object).Count
            }

            # Disks (physical)
            $diskDrives = Get-CimInstance -ClassName Win32_DiskDrive -CimSession $cs -ErrorAction SilentlyContinue
            $physicalDisks = @()
            if ($diskDrives) {
                $physicalDisks = $diskDrives | Where-Object { -not $_.MediaType -or $_.MediaType -notmatch 'Removable' }
                foreach ($d in $physicalDisks) {
                    $diskDetails.Add([pscustomobject]@{
                        ComputerName   = $computer
                        Index          = $d.Index
                        Model          = $d.Model
                        InterfaceType  = $d.InterfaceType
                        SerialNumber   = $d.SerialNumber
                        SizeGB         = if ($d.Size) { [math]::Round(($d.Size / 1GB),2) } else { $null }
                    })
                }
            }
            $diskCount        = ($physicalDisks | Measure-Object).Count
            $rawPhysicalBytes = ($physicalDisks | Measure-Object -Property Size -Sum).Sum
            $rawPhysicalGB    = if ($rawPhysicalBytes) { [math]::Round(($rawPhysicalBytes / 1GB),2) } else { 0 }

            # Volumes (fixed)
            $volumes = Get-CimInstance -ClassName Win32_Volume -CimSession $cs -ErrorAction SilentlyContinue |
                       Where-Object { $_.DriveType -eq 3 -and $_.FileSystem }
            foreach ($v in $volumes) {
                $volumeDetails.Add([pscustomobject]@{
                    ComputerName   = $computer
                    Label          = $v.Label
                    DriveLetter    = $v.DriveLetter
                    Path           = $v.DeviceID
                    FileSystem     = $v.FileSystem
                    CapacityGB     = if ($v.Capacity)   { [math]::Round(($v.Capacity / 1GB),2) } else { $null }
                    FreeGB         = if ($v.FreeSpace)  { [math]::Round(($v.FreeSpace / 1GB),2) } else { $null }
                    UsedGB         = if ($v.Capacity -and $v.FreeSpace -ne $null) { [math]::Round((($v.Capacity - $v.FreeSpace) / 1GB),2) } else { $null }
                    IsSystemVolume = $v.SystemVolume
                })
            }
            $volCount    = ($volumes | Measure-Object).Count
            $volCapBytes = ($volumes | Measure-Object -Property Capacity -Sum).Sum
            $volFreeBytes= ($volumes | Measure-Object -Property FreeSpace -Sum).Sum
            $volUsedBytes= if ($volCapBytes -ne $null -and $volFreeBytes -ne $null) { $volCapBytes - $volFreeBytes } else { $null }

            # Performance averages (30d preferred; else short sample)
            $perf = Get-PerfAverages -Computer $computer -LookbackDays $PerfLookbackDays -SearchPaths $PerfSearchPaths `
                                     -Cred $Credential -UseShortSampleFallback:$UseShortSampleFallback -ShortSampleSeconds $ShortSampleSeconds

            # Summary row
            $serverSummaries.Add([pscustomobject]@{
                ComputerName            = $computer
                Reachable               = $true
                OS_Caption              = $osCaption
                OS_Version              = $osVersion
                OS_Build                = $osBuild
                HW_Manufacturer         = $manu
                HW_Model                = $model
                HW_SerialNumber         = $serial
                CPU_Model               = $cpuModels
                CPU_PhysicalSockets     = $socketCount
                CPU_TotalCores          = $totalCores
                Mem_TotalGB             = $memTotalGB
                CPU_AvgPct_30d          = $perf.CpuAvg
                Mem_AvgPct_30d          = $perf.MemAvg
                Perf_Method             = $perf.Method
                NIC_Count               = $nicCount
                Disk_PhysicalCount      = $diskCount
                Disk_RawPhysicalGB      = $rawPhysicalGB
                Vol_FixedCount          = $volCount
                Vol_TotalCapacityGB     = if ($volCapBytes)  { [math]::Round(($volCapBytes  / 1GB),2) } else { 0 }
                Vol_UsedGB              = if ($volUsedBytes) { [math]::Round(($volUsedBytes / 1GB),2) } else { 0 }
                Vol_FreeGB              = if ($volFreeBytes) { [math]::Round(($volFreeBytes / 1GB),2) } else { 0 }
            })
        }
        catch {
            # FIX: brace $computer to avoid "$computer:" parsing error
            Write-Warning "Inventory failed on ${computer}: $($_.Exception.Message)"
            $serverSummaries.Add([pscustomobject]@{
                ComputerName            = $computer
                Reachable               = $true
                OS_Caption              = $null
                OS_Version              = $null
                OS_Build                = $null
                HW_Manufacturer         = $null
                HW_Model                = $null
                HW_SerialNumber         = $null
                CPU_Model               = $null
                CPU_PhysicalSockets     = $null
                CPU_TotalCores          = $null
                Mem_TotalGB             = $null
                CPU_AvgPct_30d          = $null
                Mem_AvgPct_30d          = $null
                Perf_Method             = 'Unavailable'
                NIC_Count               = $null
                Disk_PhysicalCount      = $null
                Disk_RawPhysicalGB      = $null
                Vol_FixedCount          = $null
                Vol_TotalCapacityGB     = $null
                Vol_UsedGB              = $null
                Vol_FreeGB              = $null
            })
        }
        finally {
            if ($cs) { $cs | Remove-CimSession }
        }
    }
}

end {
    # Sort data like the previous outputs for consistency
    $serverSummaries = $serverSummaries | Sort-Object ComputerName
    $nicDetails      = $nicDetails      | Sort-Object ComputerName, Interface
    $diskDetails     = $diskDetails     | Sort-Object ComputerName, Index
    $volumeDetails   = $volumeDetails   | Sort-Object ComputerName, DriveLetter, Path

    # Create Excel via COM (requires Excel installed)
    $excel = $null
    $workbook = $null
    $wsSummary = $null; $wsNICs = $null; $wsDisks = $null; $wsVolumes = $null

    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $excel.Visible = [bool]$ExcelVisible
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()

        # Create/rename worksheets
        $wsSummary = $workbook.Worksheets.Item(1)
        $wsSummary.Name = 'Summary'

        # Add sheets AFTER the last sheet to keep order Summary, NICs, Disks, Volumes
        $last = $workbook.Worksheets.Item($workbook.Worksheets.Count)
        $wsNICs    = $workbook.Worksheets.Add($null, $last);   $wsNICs.Name = 'NICs'
        $last = $workbook.Worksheets.Item($workbook.Worksheets.Count)
        $wsDisks   = $workbook.Worksheets.Add($null, $last);   $wsDisks.Name = 'Disks'
        $last = $workbook.Worksheets.Item($workbook.Worksheets.Count)
        $wsVolumes = $workbook.Worksheets.Add($null, $last);   $wsVolumes.Name = 'Volumes'

        # Remove any extra default sheets Excel may have created that aren't ours
        $desired = @('Summary','NICs','Disks','Volumes')
        for ($i = $workbook.Worksheets.Count; $i -ge 1; $i--) {
            $ws = $workbook.Worksheets.Item($i)
            if ($desired -notcontains $ws.Name) {
                $ws.Delete() | Out-Null
            }
        }

        # Column orders (Summary includes HW_SerialNumber)
        $summaryCols = @(
            'ComputerName','Reachable','OS_Caption','OS_Version','OS_Build',
            'HW_Manufacturer','HW_Model','HW_SerialNumber',
            'CPU_Model','CPU_PhysicalSockets','CPU_TotalCores',
            'Mem_TotalGB','CPU_AvgPct_30d','Mem_AvgPct_30d','Perf_Method',
            'NIC_Count','Disk_PhysicalCount','Disk_RawPhysicalGB',
            'Vol_FixedCount','Vol_TotalCapacityGB','Vol_UsedGB','Vol_FreeGB'
        )
        $nicsCols  = @('ComputerName','Vendor','Model','Interface','MACAddress','LinkSpeedGbps','Status')
        $disksCols = @('ComputerName','Index','Model','InterfaceType','SerialNumber','SizeGB')
        $volCols   = @('ComputerName','Label','DriveLetter','Path','FileSystem','CapacityGB','FreeGB','UsedGB','IsSystemVolume')

        # Write each sheet (fast array method + basic formatting)
        Write-Worksheet -Worksheet $wsSummary -Rows $serverSummaries -Columns $summaryCols -TableName 'Summary'
        Write-Worksheet -Worksheet $wsNICs    -Rows $nicDetails      -Columns $nicsCols    -TableName 'NICs'
        Write-Worksheet -Worksheet $wsDisks   -Rows $diskDetails     -Columns $disksCols   -TableName 'Disks'
        Write-Worksheet -Worksheet $wsVolumes -Rows $volumeDetails   -Columns $volCols     -TableName 'Volumes'

        # Save as .xlsx (51 = xlOpenXMLWorkbook)
        if (-not (Split-Path $ExcelPath -IsAbsolute)) {
            $ExcelPath = Join-Path (Resolve-Path .) $ExcelPath
        }
        $workbook.SaveAs($ExcelPath, 51)
        Write-Host ""
        Write-Host "Inventory complete. Excel written to: $ExcelPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create or write Excel workbook: $($_.Exception.Message)"
    }
    finally {
        if ($workbook) { $workbook.Close($true) | Out-Null }
        if ($excel)    { $excel.Quit() | Out-Null }

        # Release COM objects (approved verb function)
        $wsSummary | Remove-ComObjectSafely
        $wsNICs    | Remove-ComObjectSafely
        $wsDisks   | Remove-ComObjectSafely
        $wsVolumes | Remove-ComObjectSafely
        $workbook  | Remove-ComObjectSafely
        $excel     | Remove-ComObjectSafely

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
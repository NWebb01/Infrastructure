<#
.SYNOPSIS
  One-shot Hyper-V capacity planning collector + RVTools-style Excel report.

.DESCRIPTION
  - Remotely inventories Hyper-V hosts (VMs, hosts, cluster, networking, storage, optional perf).
  - Writes normalized CSV outputs into a timestamped folder.
  - Immediately merges those CSVs into a single Excel workbook with RVTools-like tabs:
      Summary, vInfo, vCPU, vMemory, vNIC, vDisk, vHost, vCluster, vCSV,
      vClusterNet, vSwitch, vPNIC, vSANiSCSI, vMPIO, vLocalDisk, vLocalVolume, vS2D, vPerf.
  - Uses ImportExcel if available; otherwise Excel COM; otherwise CSV-only.

.PARAMETER ComputerName
  One or more Hyper-V hosts (FQDN/NetBIOS).

.PARAMETER HostListPath
  Text file with one hostname per line.

.PARAMETER Credential
  Credentials for remote collection.

.PARAMETER CollectPerf
  Collect average perf counter samples during run.

.PARAMETER PerfSampleSeconds
  Duration for perf sampling (default 60s, sampled at 5s interval).

.PARAMETER OutputRoot
  Root folder for results (folder 'HyperV_Sizing_yyyyMMdd_HHmmss' is created inside).

.PARAMETER ExcelPath
  Optional custom output path for Excel workbook (.xlsx). Defaults under OutputRoot.

.PARAMETER Recurse
  (For the consolidation step) Not typically needed since we write to a single folder, but supported.

.PARAMETER ForceCOM
  Force Excel COM automation even if ImportExcel is present.

.PARAMETER AutoOpen
  Open the Excel file after creation.

.EXAMPLE
  .\Get-HyperVInventoryAndReport.ps1 -ComputerName HV01,HV02,HV03 -CollectPerf -PerfSampleSeconds 120 -AutoOpen

.EXAMPLE
  .\Get-HyperVInventoryAndReport.ps1 -HostListPath .\hosts.txt -Credential (Get-Credential)

.NOTES
  - Run from an elevated PowerShell session with network access to target hosts.
  - Remote hosts must have Hyper-V role (for Hyper-V cmdlets) and, if clustered, FailoverClusters module.
  - ImportExcel path (if you want nicer output): https://www.powershellgallery.com/packages/ImportExcel
#>

[CmdletBinding()]
param(
  [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
  [string[]]$ComputerName,

  [string]$HostListPath,

  [System.Management.Automation.PSCredential]$Credential,

  [switch]$CollectPerf,

  [int]$PerfSampleSeconds = 60,

  [string]$OutputRoot = (Get-Location).Path,

  [string]$ExcelPath,

  [switch]$Recurse,

  [switch]$ForceCOM,

  [switch]$AutoOpen
)

###############################################################################
### Phase 0: Prepare inputs and output folder
###############################################################################
$ErrorActionPreference = 'Stop'

# Build host list
$hosts = @()
if ($ComputerName) { $hosts += $ComputerName }
if ($HostListPath) {
  if (!(Test-Path $HostListPath)) { throw "HostListPath not found: $HostListPath" }
  $hosts += Get-Content -Path $HostListPath | Where-Object { $_ -and $_.Trim() -ne "" }
}
$hosts = $hosts | Select-Object -Unique
if (-not $hosts) { throw "No hosts provided. Use -ComputerName or -HostListPath." }

# Create the output folder
$runStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$OutputPath = Join-Path -Path $OutputRoot -ChildPath ("HyperV_Sizing_{0}" -f $runStamp)
New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null

if (-not $ExcelPath) {
  $ExcelPath = Join-Path -Path $OutputPath -ChildPath ("HyperV_RVTools_Style_Report_{0}.xlsx" -f $runStamp)
}

Write-Host "Collecting from $($hosts.Count) host(s)..." -ForegroundColor Cyan
Write-Host "Output folder: $OutputPath" -ForegroundColor Cyan
Write-Host "Excel output:  $ExcelPath" -ForegroundColor Cyan

###############################################################################
### Phase 1: Collection (per-host via PS Remoting)
###############################################################################
# Collections (local accumulators)
$hostsInfo = New-Object System.Collections.Generic.List[object]
$vmInfo    = New-Object System.Collections.Generic.List[object]
$nicInfo   = New-Object System.Collections.Generic.List[object]
$vSwitches = New-Object System.Collections.Generic.List[object]
$storageLocal = New-Object System.Collections.Generic.List[object]
$storageSAN   = New-Object System.Collections.Generic.List[object]
$mpioInfo  = New-Object System.Collections.Generic.List[object]
$clusterInfo = New-Object System.Collections.Generic.List[object]
$csvInfo     = New-Object System.Collections.Generic.List[object]
$clusterNets = New-Object System.Collections.Generic.List[object]
$s2dInfo     = New-Object System.Collections.Generic.List[object]
$perfInfo    = New-Object System.Collections.Generic.List[object]
$seenClusters = New-Object System.Collections.Generic.HashSet[string]

foreach ($h in $hosts) {
  Write-Host ">>> $h" -ForegroundColor Green
  $sessionParams = @{ ComputerName = $h; ErrorAction = 'Stop' }
  if ($Credential) { $sessionParams.Credential = $Credential }

  try {
    $result = Invoke-Command @sessionParams -ScriptBlock {
      $ErrorActionPreference = 'Stop'

      # -- Host basics
      $os  = Get-CimInstance -ClassName Win32_OperatingSystem
      $cs  = Get-CimInstance -ClassName Win32_ComputerSystem
      $procs = Get-CimInstance -ClassName Win32_Processor
      $memMods = Get-CimInstance -ClassName Win32_PhysicalMemory
      $bios = Get-CimInstance -ClassName Win32_BIOS

      # Hyper-V host config
      $hv = $null
      try { $hv = Get-VMHost -ErrorAction Stop } catch {}

      # NUMA
      $numa = $null
      try { $numa = Get-VMHostNumaNode -ErrorAction Stop } catch {}

      # Live Migration (subset)
      $lm = $null
      try {
        $lm = Get-VMHost -ErrorAction Stop | Select-Object `
          VirtualMachineMigrationEnabled,VirtualMachineMigrationAuthenticationType,VirtualMachineMigrationPerformanceOption
      } catch {}

      # -- Networking
      $nics = @()
      try {
        $nics = Get-NetAdapter | Select-Object `
          @{n='Host';e={$env:COMPUTERNAME}},
          Name, InterfaceDescription, Status, LinkSpeed, MacAddress, VlanID, DriverInformation, DriverFileName
      } catch {}

      $vSwitch = @()
      try {
        $vSwitch = Get-VMSwitch -ErrorAction SilentlyContinue | ForEach-Object {
          $uplinks = $null
          try {
            # SET switching: adapter descs tied to the switch team (if cmdlet exists)
            if (Get-Command Get-VMSwitchTeam -ErrorAction SilentlyContinue) {
              $uplinks = (Get-VMSwitchTeam -SwitchName $_.Name -ErrorAction SilentlyContinue |
                         Select-Object -ExpandProperty NetAdapterInterfaceDescription) -join ';'
            }
          } catch {}
          [pscustomobject]@{
            Host = $env:COMPUTERNAME
            Name = $_.Name
            SwitchType = $_.SwitchType
            AllowManagementOS = $_.AllowManagementOS
            BandwidthReservationMode = $_.BandwidthReservationMode
            NetAdapters = $uplinks
          }
        }
      } catch {}

      # -- Local storage
      $disks = @()
      $vols  = @()
      try {
        $disks = Get-Disk | Select-Object `
          @{n='Host';e={$env:COMPUTERNAME}},
          Number, FriendlyName, Model, SerialNumber, BusType, PartitionStyle, OperationalStatus, Size, AllocationUnitSize
        $vols = Get-Volume | Select-Object `
          @{n='Host';e={$env:COMPUTERNAME}},
          DriveLetter, FileSystemLabel, FileSystem, HealthStatus, Size, SizeRemaining, Path
      } catch {}

      # -- SAN / iSCSI
      $iscsi = @()
      try {
        $initiators = Get-InitiatorPort -ErrorAction SilentlyContinue
        $sessions = Get-IscsiSession -ErrorAction SilentlyContinue
        foreach ($s in $sessions) {
          $iscsi += [pscustomobject]@{
            Host        = $env:COMPUTERNAME
            TargetNode  = $s.TargetNodeAddress
            TargetPortal= $s.TargetPortalAddress
            Initiator   = ($initiators | Select-Object -First 1).NodeAddress
            SessionId   = $s.SessionIdentifier
            Connection  = $s.ConnectionIdentifier
          }
        }
      } catch {}

      # -- MPIO
      $mpio = @()
      try {
        if (Get-WindowsFeature -Name Multipath-IO -ErrorAction SilentlyContinue | Where-Object {$_.InstallState -eq 'Installed'}) {
          $devices = Get-MSDSMSupportedHW -ErrorAction SilentlyContinue
          $claims  = Get-MSDSMAutomaticClaim -ErrorAction SilentlyContinue
          $mpio += [pscustomobject]@{
            Host      = $env:COMPUTERNAME
            SupportedVendors = ($devices | ForEach-Object {$_.VendorId + ':' + $_.ProductId}) -join ';'
            AutoClaims = ($claims | ForEach-Object {$_.BusType}) -join ';'
          }
        }
      } catch {}

      # -- Cluster / CSV / S2D
      $cluster = $null
      $clusterNodes = @()
      $clusterNetworks = @()
      $csvs = @()
      $s2d = $null
      $s2dVDs = @()

      try {
        $clusSvc = Get-Service -Name clussvc -ErrorAction SilentlyContinue
        if ($clusSvc -and $clusSvc.Status -eq 'Running') {
          Import-Module FailoverClusters -ErrorAction Stop | Out-Null
          $cluster = Get-Cluster
          $clusterNodes = Get-ClusterNode | Select-Object Name, State, NodeWeight
          $clusterNetworks = Get-ClusterNetwork | Select-Object Name, Role, Metric, AutoMetric, Address, AddressMask
          $csvs = Get-ClusterSharedVolume | ForEach-Object {
            $o = $_.SharedVolumeInfo
            [pscustomobject]@{
              Cluster   = $cluster.Name
              Name      = $_.Name
              Path      = $o.FriendlyVolumeName
              CSVFSPath = $o.FriendlyVolumeName + "\"
              Size      = $o.TotalSize
              Used      = $o.UsedSize
              Free      = $o.TotalSize - $o.UsedSize
            }
          }

          try {
            $s2d = Get-ClusterS2D -ErrorAction Stop
            $subsys = Get-StorageSubSystem | Where-Object {$_.FriendlyName -like "*$($cluster.Name)*"}
            if ($subsys) {
              $s2dVDs = Get-VirtualDisk -StorageSubSystemFriendlyName $subsys.FriendlyName |
                        Select-Object FriendlyName, ResiliencySettingName, FaultDomainAwareness, HealthStatus, OperationalStatus, Size
            }
          } catch {}
        }
      } catch {}

      # -- VMs
      $vmsOut = @()
      try {
        $vmsOut = Get-VM | ForEach-Object {
          $vm = $_
          $cpu = $null; $mem = $null; $nics = $null; $disks = $null
          try { $cpu = $vm | Get-VMProcessor } catch {}
          try { $mem = $vm | Get-VMMemory } catch {}
          try { $nics = $vm | Get-VMNetworkAdapter } catch {}
          try { $disks = $vm | Get-VMHardDiskDrive } catch {}

          $diskInfo = @()
          foreach ($d in ($disks | Sort-Object -Property ControllerType, ControllerNumber, ControllerLocation)) {
            try {
              $vhd = Get-VHD -Path $d.Path -ErrorAction Stop
              $diskInfo += [pscustomobject]@{
                Path         = $d.Path
                DiskType     = $vhd.VHDType
                Format       = $vhd.VHDFormat
                SizeBytes    = $vhd.Size
                FileSizeBytes= $vhd.FileSize
                Controller   = "$($d.ControllerType) $($d.ControllerNumber):$($d.ControllerLocation)"
              }
            } catch {
              $diskInfo += [pscustomobject]@{
                Path         = $d.Path
                DiskType     = $null
                Format       = $null
                SizeBytes    = $null
                FileSizeBytes= $null
                Controller   = "$($d.ControllerType) $($d.ControllerNumber):$($d.ControllerLocation)"
              }
            }
          }

          [pscustomobject]@{
            Host                = $env:COMPUTERNAME
            VMName              = $vm.Name
            State               = $vm.State
            Generation          = $vm.Generation
            Uptime              = $vm.Uptime
            CPUCount            = $cpu.Count
            CPUReservePct       = $cpu.Reserve
            CPURelativeWeight   = $cpu.RelativeWeight
            ExposeVirtualNUMA   = $cpu.ExposeVirtualizationExtensions
            MemoryAssignedMB    = [int]($vm.MemoryAssigned / 1MB)
            MemoryStartupMB     = $mem.Startup
            DynamicMemoryEnabled= $mem.DynamicMemoryEnabled
            MinMemoryMB         = $mem.Minimum
            MaxMemoryMB         = $mem.Maximum
            Switches            = (($nics | Select-Object -ExpandProperty SwitchName) -join ';')
            NICCount            = ($nics | Measure-Object).Count
            DiskCount           = $diskInfo.Count
            Disks               = ($diskInfo | ConvertTo-Json -Depth 3 -Compress)
          }
        }
      } catch {}

      # -- Perf (optional)
      $perf = @()
      if ($using:CollectPerf) {
        $sampleInterval = 5
        $samples = [Math]::Max([Math]::Round($using:PerfSampleSeconds / $sampleInterval, 0), 1)

        $counters = @(
          '\Hyper-V Hypervisor Logical Processor(_Total)\% Total Run Time',
          '\Memory\Available MBytes',
          '\PhysicalDisk(_Total)\Avg. Disk sec/Read',
          '\PhysicalDisk(_Total)\Avg. Disk sec/Write'
        )
        try {
          $clusSvc = Get-Service -Name clussvc -ErrorAction SilentlyContinue
          if ($clusSvc -and $clusSvc.Status -eq 'Running') {
            $counters += '\Cluster CSVFS(_Total)\IO Read Latency',
                         '\Cluster CSVFS(_Total)\IO Write Latency'
          }
        } catch {}

        $data = Get-Counter -Counter $counters -SampleInterval $sampleInterval -MaxSamples $samples
        $perf = $data.CounterSamples |
          Group-Object -Property Path |
          ForEach-Object {
            $avg = ($_.Group | Measure-Object -Property CookedValue -Average).Average
            [pscustomobject]@{
              Host     = $env:COMPUTERNAME
              Counter  = $_.Name
              AvgValue = [double]::Round($avg, 6)
              Samples  = $samples
              IntervalSeconds = $sampleInterval
            }
          }
      }

      # Return all
      [pscustomobject]@{
        HostInfo = [pscustomobject]@{
          HostName      = $env:COMPUTERNAME
          Domain        = $env:USERDNSDOMAIN
          Manufacturer  = $bios.Manufacturer
          Model         = $cs.Model
          OSMajorMinor  = "$($os.Caption) ($($os.Version))"
          InstallDate   = $os.InstallDate
          TotalRAMGB    = [math]::Round(($cs.TotalPhysicalMemory/1GB),2)
          CPUModel      = ($procs | Select-Object -First 1 -ExpandProperty Name)
          CPUSockets    = ($procs | Measure-Object).Count
          CoresPerCPU   = ($procs | Select-Object -First 1 -ExpandProperty NumberOfCores)
          LogicalPerCPU = ($procs | Select-Object -First 1 -ExpandProperty NumberOfLogicalProcessors)
          NUMANodes     = ($numa | Measure-Object).Count
          VMMigration   = ($lm | ConvertTo-Json -Compress)
        }

        NICs    = $nics
        VSwitch = $vSwitch
        LocalStorage_Disks  = $disks
        LocalStorage_Volumes= $vols
        SAN_iSCSI           = $iscsi
        MPIO                = $mpio
        Cluster             = $cluster
        ClusterNodes        = $clusterNodes
        ClusterNetworks     = $clusterNetworks
        CSVs                = $csvs
        S2D                 = $s2d
        S2D_VirtualDisks    = $s2dVDs
        VMs                 = $vmsOut
        PerfCounters        = $perf
      }
    }

    # Accumulate
    if ($result.HostInfo) { $hostsInfo.Add($result.HostInfo) }
    if ($result.NICs) { $nicInfo.AddRange($result.NICs) }
    if ($result.VSwitch) { $vSwitches.AddRange($result.VSwitch) }
    if ($result.LocalStorage_Disks) { $storageLocal.AddRange($result.LocalStorage_Disks) }
    if ($result.LocalStorage_Volumes) { $storageLocal.AddRange($result.LocalStorage_Volumes) }
    if ($result.SAN_iSCSI) { $storageSAN.AddRange($result.SAN_iSCSI) }
    if ($result.MPIO) { $mpioInfo.AddRange($result.MPIO) }
    if ($result.VMs) { $vmInfo.AddRange($result.VMs) }
    if ($result.PerfCounters) { $perfInfo.AddRange($result.PerfCounters) }

    if ($result.Cluster) {
      $clusName = $result.Cluster.Name
      if ($clusName -and -not $seenClusters.Contains($clusName)) {
        $seenClusters.Add($clusName) | Out-Null
        $clusterInfo.Add([pscustomobject]@{
          ClusterName  = $result.Cluster.Name
          CLVersion    = $result.Cluster.ClusterFunctionalLevel
          Quorum       = $result.Cluster.QuorumResource
          DynamicQuorum= $result.Cluster.DynamicQuorum
          Nodes        = ($result.ClusterNodes | Select-Object -ExpandProperty Name) -join ';'
        })
        if ($result.CSVs) { $csvInfo.AddRange($result.CSVs) }
        if ($result.ClusterNetworks) { $clusterNets.AddRange($result.ClusterNetworks) }
        if ($result.S2D) {
          $s2dInfo.Add([pscustomobject]@{
            ClusterName = $result.Cluster.Name
            S2DEnabled  = $result.S2D.S2DEnabled
            Health      = $result.S2D.OperationalStatus
          })
          if ($result.S2D_VirtualDisks) {
            foreach ($vd in $result.S2D_VirtualDisks) {
              $s2dInfo.Add([pscustomobject]@{
                ClusterName = $result.Cluster.Name
                S2D_VDisk   = $vd.FriendlyName
                Resiliency  = $vd.ResiliencySettingName
                Health      = $vd.HealthStatus
                Size        = $vd.Size
              })
            }
          }
        }
      }
    }
  }
  catch {
    Write-Warning "Failed to collect from $h : $($_.Exception.Message)"
    continue
  }
}

###############################################################################
### Phase 2: Write CSVs to the output folder
###############################################################################
Write-Host "Exporting CSVs..." -ForegroundColor Cyan
$hostsInfo   | Export-Csv -Path (Join-Path $OutputPath 'hosts.csv') -NoTypeInformation -Encoding UTF8
$vmInfo      | Export-Csv -Path (Join-Path $OutputPath 'vms.csv') -NoTypeInformation -Encoding UTF8
$nicInfo     | Export-Csv -Path (Join-Path $OutputPath 'nics.csv') -NoTypeInformation -Encoding UTF8
$vSwitches   | Export-Csv -Path (Join-Path $OutputPath 'vswitches.csv') -NoTypeInformation -Encoding UTF8
$storageLocal| Export-Csv -Path (Join-Path $OutputPath 'storage_local.csv') -NoTypeInformation -Encoding UTF8
$storageSAN  | Export-Csv -Path (Join-Path $OutputPath 'storage_san_iscsi.csv') -NoTypeInformation -Encoding UTF8
$mpioInfo    | Export-Csv -Path (Join-Path $OutputPath 'mpio.csv') -NoTypeInformation -Encoding UTF8
$clusterInfo | Export-Csv -Path (Join-Path $OutputPath 'clusters.csv') -NoTypeInformation -Encoding UTF8
$csvInfo     | Export-Csv -Path (Join-Path $OutputPath 'csvs.csv') -NoTypeInformation -Encoding UTF8
$clusterNets | Export-Csv -Path (Join-Path $OutputPath 'cluster_networks.csv') -NoTypeInformation -Encoding UTF8
$s2dInfo     | Export-Csv -Path (Join-Path $OutputPath 's2d.csv') -NoTypeInformation -Encoding UTF8
$perfInfo    | Export-Csv -Path (Join-Path $OutputPath 'perf_counters.csv') -NoTypeInformation -Encoding UTF8

# Summary JSON (optional)
$summary = [pscustomobject]@{
  GeneratedOn = (Get-Date)
  HostCount   = ($hostsInfo | Measure-Object).Count
  VMCount     = ($vmInfo   | Measure-Object).Count
  TotalHostRAM_GB = [math]::Round((($hostsInfo | Select-Object -ExpandProperty TotalRAMGB) | Measure-Object -Sum).Sum, 2)
}
$summary | ConvertTo-Json -Depth 4 | Out-File (Join-Path $OutputPath 'summary.json') -Encoding UTF8

###############################################################################
### Phase 3: Consolidate into one Excel workbook (RVTools-style)
###############################################################################
Write-Host "Building Excel workbook..." -ForegroundColor Cyan

# Helper: import CSV if present
function Import-CsvIfExists {
  param([string]$FileName)
  $file = Join-Path $OutputPath $FileName
  if (Test-Path $file) { Import-Csv -Path $file } else { $null }
}

# Helper: safe JSON parse
function ConvertFrom-JsonSafe {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
  try { return $Text | ConvertFrom-Json -Depth 6 } catch { return $null }
}

# Load
$hostsCsv   = Import-CsvIfExists 'hosts.csv'
$vmsCsv     = Import-CsvIfExists 'vms.csv'
$nicsCsv    = Import-CsvIfExists 'nics.csv'
$vswCsv     = Import-CsvIfExists 'vswitches.csv'
$storLocal  = Import-CsvIfExists 'storage_local.csv'
$storSAN    = Import-CsvIfExists 'storage_san_iscsi.csv'
$mpioCsv    = Import-CsvIfExists 'mpio.csv'
$clusters   = Import-CsvIfExists 'clusters.csv'
$csvs       = Import-CsvIfExists 'csvs.csv'
$clusNets   = Import-CsvIfExists 'cluster_networks.csv'
$s2dCsv     = Import-CsvIfExists 's2d.csv'
$perfCsv    = Import-CsvIfExists 'perf_counters.csv'

# Build tables
$tables = @{}

# vHost
if ($hostsCsv) {
  $tables['vHost'] = $hostsCsv | Select-Object `
    HostName, Domain, Manufacturer, Model, OSMajorMinor, InstallDate, TotalRAMGB,
    CPUModel, CPUSockets, CoresPerCPU, LogicalPerCPU, NUMANodes
}

# vInfo / vCPU / vMemory / vNIC / vDisk
if ($vmsCsv) {
  $tables['vInfo'] = $vmsCsv | Select-Object Host, VMName, State, Generation, Uptime
  $tables['vCPU']  = $vmsCsv | Select-Object `
      Host, VMName,
      @{n='vCPU';e={$_.CPUCount}},
      @{n='CPUReservePct';e={$_.CPUReservePct}},
      @{n='CPURelativeWeight';e={$_.CPURelativeWeight}},
      @{n='ExposeVirtualNUMA';e={$_.ExposeVirtualNUMA}}
  $tables['vMemory'] = $vmsCsv | Select-Object `
      Host, VMName,
      @{n='StartupMB';e={$_.MemoryStartupMB}},
      @{n='AssignedMB';e={$_.MemoryAssignedMB}},
      @{n='DynamicMemoryEnabled';e={$_.DynamicMemoryEnabled}},
      @{n='MinMB';e={$_.MinMemoryMB}},
      @{n='MaxMB';e={$_.MaxMemoryMB}}
  # vNIC
  $vNIC = foreach ($row in $vmsCsv) {
    $switches = @()
    if ($row.Switches) { $switches = ($row.Switches -split ';') | Where-Object { $_ -and $_.Trim() -ne '' } }
    if ($switches.Count -gt 0) {
      foreach ($sw in $switches) {
        [pscustomobject]@{ Host=$row.Host; VMName=$row.VMName; Switch=$sw.Trim(); NICCount=$row.NICCount }
      }
    } else {
      [pscustomobject]@{ Host=$row.Host; VMName=$row.VMName; Switch=$null; NICCount=$row.NICCount }
    }
  }
  $tables['vNIC'] = $vNIC | Sort-Object Host, VMName, Switch

  # vDisk
  $vDisk = foreach ($row in $vmsCsv) {
    $disks = ConvertFrom-JsonSafe $row.Disks
    if ($disks -isnot [System.Collections.IEnumerable]) { $disks = @($disks) }
    foreach ($d in ($disks | Where-Object { $_ })) {
      [pscustomobject]@{
        Host          = $row.Host
        VMName        = $row.VMName
        Path          = $d.Path
        DiskType      = $d.DiskType
        Format        = $d.Format
        SizeBytes     = $d.SizeBytes
        FileSizeBytes = $d.FileSizeBytes
        Controller    = $d.Controller
      }
    }
  }
  $tables['vDisk'] = $vDisk | Sort-Object Host, VMName, Path
}

# vSwitch
if ($vswCsv) { $tables['vSwitch'] = $vswCsv | Select-Object Host, Name, SwitchType, AllowManagementOS, BandwidthReservationMode, NetAdapters }

# vPNIC
if ($nicsCsv) { $tables['vPNIC'] = $nicsCsv | Select-Object Host, Name, InterfaceDescription, Status, LinkSpeed, MacAddress, VlanID, DriverInformation, DriverFileName }

# vCluster / vCSV / vClusterNet
if ($clusters)   { $tables['vCluster']   = $clusters | Select-Object ClusterName, CLVersion, Quorum, DynamicQuorum, Nodes }
if ($csvs)       { $tables['vCSV']       = $csvs     | Select-Object Cluster, Name, Path, CSVFSPath, Size, Used, Free }
if ($clusNets)   { $tables['vClusterNet']= $clusNets | Select-Object Name, Role, Metric, AutoMetric, Address, AddressMask }

# vSANiSCSI / vMPIO
if ($storSAN)    { $tables['vSANiSCSI']  = $storSAN  | Select-Object Host, TargetNode, TargetPortal, Initiator, SessionId, Connection }
if ($mpioCsv)    { $tables['vMPIO']      = $mpioCsv  | Select-Object Host, SupportedVendors, AutoClaims }

# vLocalDisk / vLocalVolume
if ($storLocal) {
  $vLocalDisk = $storLocal | Where-Object { $_.Number -ne $null } | Select-Object `
    Host, Number, FriendlyName, Model, SerialNumber, BusType, PartitionStyle, OperationalStatus, Size, AllocationUnitSize
  $vLocalVolume = $storLocal | Where-Object { $_.DriveLetter -ne $null -or $_.Path } | Select-Object `
    Host, DriveLetter, FileSystemLabel, FileSystem, HealthStatus, Size, SizeRemaining, Path
  if ($vLocalDisk)   { $tables['vLocalDisk']   = $vLocalDisk }
  if ($vLocalVolume) { $tables['vLocalVolume'] = $vLocalVolume }
}

# vS2D
if ($s2dCsv) { $tables['vS2D'] = $s2dCsv }

# vPerf
if ($perfCsv) { $tables['vPerf'] = $perfCsv }

# Summary
$summaryRows = @()
if ($hostsCsv) {
  $summaryRows += [pscustomobject]@{ Metric='Host count'; Value=($hostsCsv | Measure-Object).Count }
  $summaryRows += [pscustomobject]@{ Metric='Total host RAM (GB)'; Value=[math]::Round(($hostsCsv.TotalRAMGB | Measure-Object -Sum).Sum,2) }
  $summaryRows += [pscustomobject]@{ Metric='Distinct CPU models'; Value=(($hostsCsv.CPUModel | Sort-Object -Unique) -join '; ') }
}
if ($vmsCsv) {
  $summaryRows += [pscustomobject]@{ Metric='VM count'; Value=($vmsCsv | Measure-Object).Count }
  $dyn = ($vmsCsv | Where-Object { $_.DynamicMemoryEnabled -match 'True' }).Count
  $summaryRows += [pscustomobject]@{ Metric='VMs with Dynamic Memory'; Value=$dyn }
}
if ($csvs) {
  $summaryRows += [pscustomobject]@{ Metric='CSV volumes'; Value=($csvs | Measure-Object).Count }
  $cap = (($csvs.Size | ForEach-Object { [double]$_ }) | Measure-Object -Sum).Sum
  $summaryRows += [pscustomobject]@{ Metric='CSV total capacity (bytes)'; Value=$cap }
}
$tables['Summary'] = if ($summaryRows) { $summaryRows } else { [pscustomobject]@{ Metric='Info'; Value='No data available' } }

# Try ImportExcel (unless forced COM)
$hasImportExcel = $false
if (-not $ForceCOM) { $hasImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel) }

if ($hasImportExcel) {
  Write-Host "Using ImportExcel module for output..." -ForegroundColor Green
  $first = $true
  foreach ($name in ($tables.Keys | Sort-Object)) {
    $data = $tables[$name]
    if (-not $data) { continue }
    if ($first) {
      $data | Export-Excel -Path $ExcelPath -WorksheetName $name -AutoSize -FreezeTopRow -TableName ($name + "_tbl") -BoldTopRow -ClearSheet
      $first = $false
    } else {
      $data | Export-Excel -Path $ExcelPath -WorksheetName $name -AutoSize -FreezeTopRow -TableName ($name + "_tbl") -BoldTopRow -Append
    }
  }
}
else {
  # Excel COM fallback
  try {
    Write-Host "ImportExcel not found or -ForceCOM set. Using Excel COM..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { ($wb.Worksheets.Item(1)).Delete() }

    function Write-ComSheet {
      param([Parameter(Mandatory)]$Worksheet, [Parameter(Mandatory)][System.Collections.IEnumerable]$Data)
      $arr = @()
      $props = @()
      foreach ($item in $Data) {
        if (-not $props) { $props = $item.PSObject.Properties.Name }
        $row = @()
        foreach ($p in $props) { $row += $item.$p }
        $arr += ,$row
      }
      # Headers
      for ($c=0; $c -lt $props.Count; $c++) { $Worksheet.Cells.Item(1, $c+1) = $props[$c] }
      # Data
      if ($arr.Count -gt 0) {
        $rows = $arr.Count; $cols = $props.Count
        $range = $Worksheet.Range($Worksheet.Cells.Item(2,1), $Worksheet.Cells.Item(1+$rows, $cols))
        $data2D = New-Object 'object[,]' $rows, $cols
        for ($r=0; $r -lt $rows; $r++) {
          for ($c=0; $c -lt $cols; $c++) { $data2D[$r,$c] = $arr[$r][$c] }
        }
        $range.Value2 = $data2D
      }
      $Worksheet.Rows.Item(1).Font.Bold = $true
      $Worksheet.Columns.AutoFit() | Out-Null
      $Worksheet.Rows.Item(1).FreezePanes = $true
    }

    foreach ($name in ($tables.Keys | Sort-Object)) {
      $data = $tables[$name]
      if (-not $data) { continue }
      $ws = $wb.Worksheets.Add()
      $ws.Name = ($name.Length -gt 31) ? $name.Substring(0,31) : $name
      Write-ComSheet -Worksheet $ws -Data $data
    }

    # Save as xlsx (51 = xlOpenXMLWorkbook)
    $null = New-Item -ItemType Directory -Force -Path (Split-Path -Path $ExcelPath -Parent) -ErrorAction SilentlyContinue
    $wb.SaveAs((Resolve-Path (New-Item -ItemType File -Path $ExcelPath -Force)).Path, 51)
    $wb.Close($true)
    $excel.Quit()
  }
  catch {
    Write-Warning "Excel COM failed or Excel not installed. CSVs are available in $OutputPath. Error: $($_.Exception.Message)"
  }
}

Write-Host "All done." -ForegroundColor Green
if (Test-Path $ExcelPath) { Write-Host "Excel created: $ExcelPath" -ForegroundColor Green; if ($AutoOpen) { Invoke-Item -Path $ExcelPath } }
else { Write-Host "Excel file not created; see CSVs in: $OutputPath" -ForegroundColor Yellow }
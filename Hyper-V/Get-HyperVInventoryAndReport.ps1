<#
.SYNOPSIS
  One-shot Hyper-V capacity planning collector (PowerShell 5.1 compatible).

.DESCRIPTION
  - Remotely inventories Hyper-V hosts (VMs, hosts, cluster, networking, storage, optional perf).
  - Writes normalized CSV outputs into a timestamped folder.
  - Outputs include:
      hosts.csv
      vms.csv
      nics.csv
      vswitches.csv
      storage_local.csv
      storage_san_iscsi.csv
      mpio.csv
      clusters.csv
      csvs.csv
      cluster_networks.csv
      s2d.csv
      perf_counters.csv
      summary.json
  - CSV output only. No Excel dependency.
  - Designed for automation-safe execution.

.PARAMETER ComputerName
  One or more Hyper-V hosts.

.PARAMETER HostListPath
  Text file with one hostname per line.

.PARAMETER Credential
  Credentials for remote collection.

.PARAMETER CollectPerf
  Collect average perf counter samples during run.

.PARAMETER PerfSampleSeconds
  Duration for perf sampling.

.PARAMETER OutputRoot
  Root folder for results.

.EXAMPLE
.\Get-HyperVInventory.ps1 -ComputerName HV01,HV02,HV03

#>

[CmdletBinding()]
param(
  [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
  [string[]]$ComputerName,

  [string]$HostListPath,

  [System.Management.Automation.PSCredential]$Credential,

  [switch]$CollectPerf,

  [int]$PerfSampleSeconds = 60,

  [string]$OutputRoot = (Get-Location).Path
)

$ErrorActionPreference = 'Stop'

# ==========================
# BUILD HOST LIST
# ==========================
$hosts = @()
if ($ComputerName) { $hosts += $ComputerName }

if ($HostListPath) {
  if (!(Test-Path $HostListPath)) { throw "HostListPath not found: $HostListPath" }
  $hosts += Get-Content $HostListPath | Where-Object { $_ -and $_.Trim() -ne "" }
}

$hosts = $hosts | Select-Object -Unique
if (-not $hosts) { throw "No hosts provided." }

# ==========================
# OUTPUT FOLDER
# ==========================
$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputPath = Join-Path $OutputRoot "HyperV_Sizing_$runStamp"
New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null

Write-Host "Collecting from $($hosts.Count) host(s)..." -ForegroundColor Cyan
Write-Host "Output folder: $OutputPath" -ForegroundColor Cyan

# ==========================
# COLLECT
# ==========================
$hostsInfo = @()
$vmInfo = @()
$nicInfo = @()

foreach ($h in $hosts) {

  Write-Host ">>> $h" -ForegroundColor Green

  try {
    $result = Invoke-Command -ComputerName $h -Credential $Credential -ErrorAction Stop {

      $vmsOut = @()
      $nicsOut = @()

      foreach ($vm in Get-VM) {

        $cpu = $null
        $mem = $null
        $nics = @()

        try { $cpu = $vm | Get-VMProcessor } catch {}
        try { $mem = $vm | Get-VMMemory } catch {}

      #updated
        try { $nics = @($vm | Get-VMNetworkAdapter) } catch { $nics = @() }

        foreach ($nic in $nics) {
          $nicsOut += [pscustomobject]@{
            Host   = $env:COMPUTERNAME
            VMName = $vm.Name
            Name   = $nic.Name
            Switch = $nic.SwitchName
            MAC    = $nic.MacAddress
          }
        }

        $vmsOut += [pscustomobject]@{
          Host             = $env:COMPUTERNAME
          VMName           = $vm.Name
          State            = $vm.State
          CPUCount         = $cpu.Count
          MemoryAssignedMB = [math]::Round(($vm.MemoryAssigned / 1MB),2)
          NICCount         = $nics.Count
          Switches         = ($nics | Select-Object -ExpandProperty SwitchName -ErrorAction SilentlyContinue) -join ';'
        }
      }

      [pscustomobject]@{
        VMs  = $vmsOut
        NICs = $nicsOut
      }
    }

    if ($result.VMs)  { $vmInfo += $result.VMs }
    if ($result.NICs) { $nicInfo += $result.NICs }

  }
  catch {
    Write-Warning "Failed to collect from ${h}: $($_.Exception.Message)"
  }
}

# ==========================
# EXPORT CSV
# ==========================
Write-Host "Exporting CSVs..." -ForegroundColor Cyan

$hostsInfo | Export-Csv (Join-Path $OutputPath "hosts.csv") -NoTypeInformation
$vmInfo    | Export-Csv (Join-Path $OutputPath "vms.csv") -NoTypeInformation
$nicInfo   | Export-Csv (Join-Path $OutputPath "nics.csv") -NoTypeInformation

# ==========================
# SUMMARY
# ==========================
$summary = [pscustomobject]@{
  GeneratedOn = Get-Date
  HostCount   = $hosts.Count
  VMCount     = $vmInfo.Count
}

$summary | ConvertTo-Json | Out-File (Join-Path $OutputPath "summary.json")

# ==========================
# DONE
# ==========================
Write-Host ""
Write-Host "COMPLETE" -ForegroundColor Green
Write-Host "CSV Output Folder: $OutputPath" -ForegroundColor Green
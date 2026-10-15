# system-info.ps1
# Lightweight Windows System Information Utility

Write-Host ""
Write-Host "========================================"
Write-Host "          SYSTEM INFORMATION"
Write-Host "========================================"


# --------------------------------------------------
# Operating System
# --------------------------------------------------

$os = Get-CimInstance Win32_OperatingSystem

Write-Host ""
Write-Host "[Operating System]"
Write-Host "Computer Name : $env:COMPUTERNAME"
Write-Host "OS            : $($os.Caption)"
Write-Host "Version       : $($os.Version)"
Write-Host "Build Number  : $($os.BuildNumber)"
Write-Host "Architecture  : $($os.OSArchitecture)"


# --------------------------------------------------
# Hardware / RAM
# --------------------------------------------------

$computer = Get-CimInstance Win32_ComputerSystem

$totalRAMGB = [math]::Round($computer.TotalPhysicalMemory / 1GB, 2)
$usableRAMGB = [math]::Round(([double]$os.TotalVisibleMemorySize * 1KB) / 1GB, 2)
$freeRAMGB = [math]::Round(([double]$os.FreePhysicalMemory * 1KB) / 1GB, 2)
$usedRAMGB = [math]::Round(($usableRAMGB - $freeRAMGB), 2)
$ramUsagePercent = if ($usableRAMGB -gt 0) {
    [math]::Round(($usedRAMGB / $usableRAMGB) * 100, 1)
} else {
    0
}
$uptime = New-TimeSpan -Start $os.LastBootUpTime -End (Get-Date)

Write-Host ""
Write-Host "[Hardware]"
Write-Host "Manufacturer  : $($computer.Manufacturer)"
Write-Host "Model         : $($computer.Model)"
Write-Host "Installed RAM : $totalRAMGB GB"
Write-Host "Usable RAM    : $usableRAMGB GB"
Write-Host "RAM In Use    : $usedRAMGB GB ($ramUsagePercent%)"
Write-Host "RAM Available : $freeRAMGB GB"
Write-Host ("System Uptime : {0}d {1}h {2}m" -f [int]$uptime.TotalDays, $uptime.Hours, $uptime.Minutes)

$pageFiles = Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue
$pageFileTotalKB = ($pageFiles | Measure-Object -Property AllocatedBaseSize -Sum).Sum
$pageFileUsedKB = ($pageFiles | Measure-Object -Property CurrentUsage -Sum).Sum
if ($null -ne $pageFileTotalKB -and $pageFileTotalKB -gt 0) {
    $pageFileFreeKB = $pageFileTotalKB - $pageFileUsedKB
    $pageFileUsagePercent = [math]::Round(($pageFileUsedKB / $pageFileTotalKB) * 100, 1)
    Write-Host "Page File     : $([math]::Round($pageFileTotalKB / 1MB, 2)) GB total, $([math]::Round($pageFileUsedKB / 1MB, 2)) GB used, $([math]::Round($pageFileFreeKB / 1MB, 2)) GB free ($pageFileUsagePercent%)"
} else {
    Write-Host "Page File     : Unavailable"
}


# --------------------------------------------------
# Processor
# --------------------------------------------------

$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1

Write-Host ""
Write-Host "[Processor]"
Write-Host "CPU           : $($cpu.Name)"
Write-Host "Cores         : $($cpu.NumberOfCores)"
Write-Host "Logical CPUs  : $($cpu.NumberOfLogicalProcessors)"


# --------------------------------------------------
# Disks
# --------------------------------------------------

$disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"

Write-Host ""
Write-Host "[Disks]"

foreach ($disk in $disks) {

    $sizeGB = [math]::Round($disk.Size / 1GB, 2)
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    $usedGB = [math]::Round(($disk.Size - $disk.FreeSpace) / 1GB, 2)
    $usagePercent = if ($disk.Size -gt 0) {
        [math]::Round((($disk.Size - $disk.FreeSpace) / $disk.Size) * 100, 1)
    } else {
        0
    }

    Write-Host ""
    Write-Host "Drive         : $($disk.DeviceID)"
    Write-Host "Volume Label  : $(if ($disk.VolumeName) { $disk.VolumeName } else { 'No label' })"
    Write-Host "File System   : $($disk.FileSystem)"
    Write-Host "Total Space   : $sizeGB GB"
    Write-Host "Used Space    : $usedGB GB ($usagePercent%)"
    Write-Host "Free Space    : $freeGB GB"
}


# --------------------------------------------------
# Installed Software
# --------------------------------------------------

Write-Host ""
Write-Host "[Installed Software]"
Write-Host ""

$registryPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$software = foreach ($path in $registryPaths) {

    Get-ItemProperty $path -ErrorAction SilentlyContinue |
        Where-Object {
            $null -ne $_.DisplayName
        } |
        Select-Object DisplayName, DisplayVersion, Publisher
}

$software |
    Sort-Object DisplayName, DisplayVersion, Publisher -Unique |
    Format-Table DisplayName, DisplayVersion, Publisher -AutoSize


# --------------------------------------------------
# Completed
# --------------------------------------------------

Write-Host ""
Write-Host "========================================"
Write-Host "       SYSTEM SCAN COMPLETED"
Write-Host "========================================"
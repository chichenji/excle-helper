# Register EXCLE助手 as a per-user Excel & WPS Spreadsheets (ET) COM add-in (HKCU only).
# Registers BOTH 64-bit and 32-bit (WOW6432Node) COM CLSID views so the AnyCPU DLL works in
# 64-bit Excel, 32-bit Excel, 64-bit WPS, and 32-bit WPS.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$dll  = Join-Path $here 'bin\ExcleHelper.dll'
if (-not (Test-Path $dll)) { throw "not built: $dll (run build.ps1 first)" }
try { Unblock-File -Path $dll -ErrorAction SilentlyContinue } catch { }

$guid       = '{5E1C7A93-2B48-4D16-9F77-8C3A5B1E6D24}'
$netCatGuid = '{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}'
$progId     = 'ExcleHelper.Connect'
$asmName    = 'ExcleHelper, Version=1.1.0.0, Culture=neutral, PublicKeyToken=null'
$codeBase   = ([Uri]$dll).AbsoluteUri

Write-Host "DLL      : $dll"
Write-Host "CodeBase : $codeBase"

function Register-ComClsid([string]$clsidRoot) {
  $clsid = "$clsidRoot\$guid"
  New-Item -Path $clsid -Force | Out-Null
  Set-Item -Path $clsid -Value $progId | Out-Null

  New-Item -Path "$clsid\InprocServer32" -Force | Out-Null
  Set-Item -Path "$clsid\InprocServer32" -Value 'mscoree.dll' | Out-Null
  New-ItemProperty -Path "$clsid\InprocServer32" -Name 'ThreadingModel' -Value 'Both'       -PropertyType String -Force | Out-Null
  New-ItemProperty -Path "$clsid\InprocServer32" -Name 'Assembly'       -Value $asmName     -PropertyType String -Force | Out-Null
  New-ItemProperty -Path "$clsid\InprocServer32" -Name 'Class'          -Value $progId      -PropertyType String -Force | Out-Null
  New-ItemProperty -Path "$clsid\InprocServer32" -Name 'RuntimeVersion' -Value 'v4.0.30319' -PropertyType String -Force | Out-Null
  New-ItemProperty -Path "$clsid\InprocServer32" -Name 'CodeBase'       -Value $codeBase    -PropertyType String -Force | Out-Null

  # Version subkey (matches regasm /codebase layout for both 32-bit and 64-bit hosts)
  New-Item -Path "$clsid\InprocServer32\1.1.0.0" -Force | Out-Null
  New-ItemProperty -Path "$clsid\InprocServer32\1.1.0.0" -Name 'Assembly'       -Value $asmName     -PropertyType String -Force | Out-Null
  New-ItemProperty -Path "$clsid\InprocServer32\1.1.0.0" -Name 'Class'          -Value $progId      -PropertyType String -Force | Out-Null
  New-ItemProperty -Path "$clsid\InprocServer32\1.1.0.0" -Name 'RuntimeVersion' -Value 'v4.0.30319' -PropertyType String -Force | Out-Null
  New-ItemProperty -Path "$clsid\InprocServer32\1.1.0.0" -Name 'CodeBase'       -Value $codeBase    -PropertyType String -Force | Out-Null

  New-Item -Path "$clsid\ProgId" -Force | Out-Null
  Set-Item -Path "$clsid\ProgId" -Value $progId | Out-Null
  New-Item -Path "$clsid\Implemented Categories\$netCatGuid" -Force | Out-Null
}

function Register-AddinEntry([string]$addinPath) {
  New-Item -Path $addinPath -Force | Out-Null
  New-ItemProperty -Path $addinPath -Name 'FriendlyName'    -Value 'EXCLE助手'                                      -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $addinPath -Name 'Description'     -Value 'Excel / WPS 表格效率工具箱（全功能免费，纯本地运算）' -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $addinPath -Name 'LoadBehavior'    -Value 3                                                -PropertyType DWord  -Force | Out-Null
  New-ItemProperty -Path $addinPath -Name 'CommandLineSafe' -Value 0                                                -PropertyType DWord  -Force | Out-Null
}

# ---------- per-user COM registration (64-bit + 32-bit WOW6432Node) ----------
Register-ComClsid 'HKCU:\Software\Classes\CLSID'
Register-ComClsid 'HKCU:\Software\Classes\WOW6432Node\CLSID'

$progKey = "HKCU:\Software\Classes\$progId"
New-Item -Path $progKey -Force | Out-Null
Set-Item -Path $progKey -Value $progId | Out-Null
New-Item -Path "$progKey\CLSID" -Force | Out-Null
Set-Item -Path "$progKey\CLSID" -Value $guid | Out-Null

# ---------- Microsoft Excel add-in entry ----------
$excelAddin = "HKCU:\Software\Microsoft\Office\Excel\Addins\$progId"
Register-AddinEntry $excelAddin

# ---------- WPS Spreadsheets (ET) add-in entry & whitelist ----------
$wpsAddin = "HKCU:\Software\Kingsoft\Office\6.0\et\Addins\$progId"
Register-AddinEntry $wpsAddin

$wpsWl = 'HKCU:\Software\Kingsoft\Office\ET\AddinsWL'
New-Item -Path $wpsWl -Force | Out-Null
New-ItemProperty -Path $wpsWl -Name $progId -Value '' -PropertyType String -Force | Out-Null

# Remove from WPS blacklist if previously disabled
$wpsBl = 'HKCU:\Software\Kingsoft\Office\ET\AddinsBL'
if (Test-Path $wpsBl) {
  try { Remove-ItemProperty -Path $wpsBl -Name $progId -ErrorAction SilentlyContinue } catch { }
}

# ---------- retire the earlier prototype so two identical ribbons do not stack ----------
$legacy = @(
  'HKCU:\Software\Microsoft\Office\Excel\Addins\GridKit.Connect',
  'HKCU:\Software\Kingsoft\Office\6.0\et\Addins\GridKit.Connect',
  'HKCU:\Software\Classes\CLSID\{A7C4E1B2-9F3D-4E8A-B6C1-2D5F7A9E3B41}',
  'HKCU:\Software\Classes\WOW6432Node\CLSID\{A7C4E1B2-9F3D-4E8A-B6C1-2D5F7A9E3B41}',
  'HKCU:\Software\Classes\GridKit.Connect'
)
$backupDir = Join-Path $here 'docs'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
foreach ($k in $legacy) {
  if (Test-Path $k) {
    $safe = ($k -replace '[\\:*?"<>|]', '_')
    $bak = Join-Path $backupDir "legacy-$safe.reg"
    $regPath = $k -replace '^HKCU:', 'HKEY_CURRENT_USER' -replace '^HKLM:', 'HKEY_LOCAL_MACHINE'
    & reg.exe export $regPath $bak /y | Out-Null
    Remove-Item -Path $k -Recurse -Force
    Write-Host "retired legacy key: $k   (backup: $bak)"
  }
}

Write-Host ''
Write-Host 'Registered (HKCU only, 64-bit + 32-bit):'
Write-Host "  CLSID (x64) : HKCU:\Software\Classes\CLSID\$guid"
Write-Host "  CLSID (x86) : HKCU:\Software\Classes\WOW6432Node\CLSID\$guid"
Write-Host "  ProgID      : $progId"
Write-Host "  Excel Addin : $excelAddin  (LoadBehavior=3)"
Write-Host "  WPS Addin   : $wpsAddin  (LoadBehavior=3)"
Write-Host "  WPS WL      : $wpsWl -> $progId"
Write-Host ''
Write-Host 'Restart Excel or WPS Spreadsheets: tabs 「EXCLE助手」「格式与数据」「工具与输出」.'
Write-Host 'Note for WPS: ensure "启用第三方 COM 加载项" is checked in WPS Options -> Trust Center.'

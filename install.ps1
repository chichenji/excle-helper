# Register EXCLE助手 as a per-user Excel COM add-in (HKCU only).
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$dll  = Join-Path $here 'bin\ExcleHelper.dll'
if (-not (Test-Path $dll)) { throw "not built: $dll (run build.ps1 first)" }

$guid     = '{5E1C7A93-2B48-4D16-9F77-8C3A5B1E6D24}'
$progId   = 'ExcleHelper.Connect'
$asmName  = 'ExcleHelper, Version=1.1.0.0, Culture=neutral, PublicKeyToken=null'
$codeBase = ([Uri]$dll).AbsoluteUri

Write-Host "DLL      : $dll"
Write-Host "CodeBase : $codeBase"

# ---------- per-user COM registration ----------
$clsid = "HKCU:\Software\Classes\CLSID\$guid"
New-Item -Path $clsid -Force | Out-Null
New-Item -Path "$clsid\InprocServer32" -Force | Out-Null
Set-Item -Path "$clsid\InprocServer32" -Value 'mscoree.dll' | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'ThreadingModel' -Value 'Both'      -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'Assembly'       -Value $asmName    -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'Class'          -Value $progId     -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'RuntimeVersion' -Value 'v4.0.30319' -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'CodeBase'       -Value $codeBase   -PropertyType String -Force | Out-Null
New-Item -Path "$clsid\ProgId" -Force | Out-Null
Set-Item -Path "$clsid\ProgId" -Value $progId | Out-Null

$progKey = "HKCU:\Software\Classes\$progId"
New-Item -Path $progKey -Force | Out-Null
Set-Item -Path $progKey -Value $progId | Out-Null
New-Item -Path "$progKey\CLSID" -Force | Out-Null
Set-Item -Path "$progKey\CLSID" -Value $guid | Out-Null

# ---------- Excel add-in entry ----------
$addin = "HKCU:\Software\Microsoft\Office\Excel\Addins\$progId"
New-Item -Path $addin -Force | Out-Null
New-ItemProperty -Path $addin -Name 'FriendlyName'    -Value 'EXCLE助手'                              -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addin -Name 'Description'     -Value '表格效率工具箱（全功能免费，纯本地运算）' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addin -Name 'LoadBehavior'    -Value 3                                        -PropertyType DWord  -Force | Out-Null
New-ItemProperty -Path $addin -Name 'CommandLineSafe' -Value 0                                        -PropertyType DWord  -Force | Out-Null

# ---------- retire the earlier prototype so two identical ribbons do not stack ----------
$legacy = @(
  'HKCU:\Software\Microsoft\Office\Excel\Addins\GridKit.Connect',
  'HKCU:\Software\Classes\CLSID\{A7C4E1B2-9F3D-4E8A-B6C1-2D5F7A9E3B41}',
  'HKCU:\Software\Classes\GridKit.Connect'
)
$backupDir = Join-Path $here 'docs'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
foreach ($k in $legacy) {
  if (Test-Path $k) {
    $safe = ($k -replace '[\\:*?"<>|]', '_')
    $bak = Join-Path $backupDir "legacy-$safe.reg"
    $regPath = $k -replace "^HKCU:", "HKEY_CURRENT_USER" -replace "^HKLM:", "HKEY_LOCAL_MACHINE"`r`n    & reg.exe export $regPath $bak /y | Out-Null
    Remove-Item -Path $k -Recurse -Force
    Write-Host "retired legacy key: $k   (backup: $bak)"
  }
}

Write-Host ''
Write-Host 'Registered (HKCU only):'
Write-Host "  CLSID  : $guid"
Write-Host "  ProgID : $progId"
Write-Host "  Add-in : $addin  (LoadBehavior=3)"
Write-Host ''
Write-Host 'Restart Excel: tabs 「EXCLE助手」「数据工具」「格式与对象」「汇总拆分」.'

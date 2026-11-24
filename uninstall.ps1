# Unregister EXCLE助手 from both Excel and WPS Spreadsheets (HKCU only).
# 反注册前会先把现有注册表项导出成 .reg 备份到 docs\，再移除。
# 用法：
#   powershell -ExecutionPolicy Bypass -File .\uninstall.ps1            # 导出备份并反注册
#   powershell -ExecutionPolicy Bypass -File .\uninstall.ps1 -DryRun    # 只导出备份，不改注册表

param([switch]$DryRun)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$docs = Join-Path $here 'docs'
if (-not (Test-Path $docs)) { New-Item -ItemType Directory -Force -Path $docs | Out-Null }

$guid   = '{5E1C7A93-2B48-4D16-9F77-8C3A5B1E6D24}'
$progId = 'ExcleHelper.Connect'
$stamp  = Get-Date -Format 'yyyyMMdd-HHmmss'

# 需要完整移除的注册表键（64位 + 32位 + Excel + WPS）
$targets = @(
  @{ Ps = "HKCU:\Software\Microsoft\Office\Excel\Addins\$progId"; Reg = "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\$progId" },
  @{ Ps = "HKCU:\Software\Kingsoft\Office\6.0\et\Addins\$progId"; Reg = "HKEY_CURRENT_USER\Software\Kingsoft\Office\6.0\et\Addins\$progId" },
  @{ Ps = "HKCU:\Software\Classes\$progId";                       Reg = "HKEY_CURRENT_USER\Software\Classes\$progId" },
  @{ Ps = "HKCU:\Software\Classes\CLSID\$guid";                   Reg = "HKEY_CURRENT_USER\Software\Classes\CLSID\$guid" },
  @{ Ps = "HKCU:\Software\Classes\WOW6432Node\CLSID\$guid";       Reg = "HKEY_CURRENT_USER\Software\Classes\WOW6432Node\CLSID\$guid" }
)
$wpsWlPs  = 'HKCU:\Software\Kingsoft\Office\ET\AddinsWL'
$wpsWlReg = 'HKEY_CURRENT_USER\Software\Kingsoft\Office\ET\AddinsWL'

Write-Host '== 导出备份 =='
foreach ($t in $targets) {
  if (-not (Test-Path $t.Ps)) { Write-Host ("  跳过（不存在）：" + $t.Ps); continue }
  $slug = ($t.Reg -replace '^HKEY_CURRENT_USER\\', '') -replace '[\\{}]', '_'
  $out  = Join-Path $docs ("uninstall-" + $stamp + "-" + $slug + ".reg")
  & reg.exe export $t.Reg $out /y | Out-Null
  Write-Host ("  已导出： " + $out)
}
if (Test-Path $wpsWlPs) {
  $wlOut = Join-Path $docs ("uninstall-" + $stamp + "-Software_Kingsoft_Office_ET_AddinsWL.reg")
  & reg.exe export $wpsWlReg $wlOut /y | Out-Null
  Write-Host ("  已导出： " + $wlOut)
}

if ($DryRun) {
  Write-Host ''
  Write-Host 'DryRun：只做了备份，未修改注册表。'
  exit 0
}

Write-Host ''
Write-Host '== 移除注册项 =='
foreach ($t in $targets) {
  if (-not (Test-Path $t.Ps)) { Write-Host ("  跳过（不存在）：" + $t.Ps); continue }
  Remove-Item -Path $t.Ps -Recurse -Force
  Write-Host ("  已移除： " + $t.Ps)
}
if (Test-Path $wpsWlPs) {
  $props = Get-ItemProperty -Path $wpsWlPs -ErrorAction SilentlyContinue
  if ($null -ne $props -and $null -ne $props.PSObject.Properties[$progId]) {
    Remove-ItemProperty -Path $wpsWlPs -Name $progId -Force
    Write-Host ("  已移除 WPS 白名单值： $wpsWlPs -> $progId")
  }
}

Write-Host ''
Write-Host '已反注册。重启 Excel / WPS 表格后功能区不再出现 EXCLE助手。'
Write-Host '如需恢复：重新运行 install.ps1；或用上面的 .reg 备份手动导入。'

# Unregister EXCLE助手 (HKCU only).
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

# 需要处理的注册表位置（HKCU 路径 -> reg.exe 认识的 HK 路径）
$targets = @(
  @{ Ps  = "HKCU:\Software\Microsoft\Office\Excel\Addins\$progId"; Reg = "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\$progId" },
  @{ Ps  = "HKCU:\Software\Classes\$progId";                       Reg = "HKEY_CURRENT_USER\Software\Classes\$progId" },
  @{ Ps  = "HKCU:\Software\Classes\CLSID\$guid";                   Reg = "HKEY_CURRENT_USER\Software\Classes\CLSID\$guid" }
)

Write-Host '== 导出备份 =='
foreach ($t in $targets) {
  if (-not (Test-Path $t.Ps)) { Write-Host ("  跳过（不存在）：" + $t.Ps); continue }
  $slug = ($t.Reg -replace '^HKEY_CURRENT_USER\\', '') -replace '[\\{}]', '_'
  $out  = Join-Path $docs ("uninstall-" + $stamp + "-" + $slug + ".reg")
  & reg.exe export $t.Reg $out /y | Out-Null
  Write-Host ("  已导出： " + $out)
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

Write-Host ''
Write-Host '已反注册。重启 Excel 后功能区不再出现 EXCLE助手。'
Write-Host '如需恢复：重新运行 install.ps1；或用上面的 .reg 备份手动导入。'

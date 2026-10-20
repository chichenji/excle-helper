# 只做一件事：渲染 UI 设计稿（面板 / 关于 / AI 三个弹窗 / 图标 / 二维码）
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$env:EXCLE_SILENT = '1'
$env:EXCLE_SHOT = '1'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$temp = $env:TEMP
$targets = @(
  'exclehelper-panel.png',
  'exclehelper-about.png',
  'exclehelper-icons.png',
  'exclehelper-qr.png',
  'exclehelper-ai-progress.png',
  'exclehelper-ai-preview.png',
  'exclehelper-ai-settings.png',
  'exclehelper-ai-chat.png'
)
foreach ($f in $targets) { Remove-Item (Join-Path $temp $f) -ErrorAction SilentlyContinue }
Remove-Item (Join-Path $temp 'exclehelper-ui.log') -ErrorAction SilentlyContinue

Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Start-Sleep -Seconds 3

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
Start-Sleep -Seconds 2
$null = $xl.Workbooks.Add()
$xl.WindowState = -4137
Start-Sleep -Seconds 3

$win = $AE::RootElement.FindFirst($TS::Children,
  (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'XLMAIN')))
function FindAny([string]$n) {
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $n)
  $e = $win.FindFirst($TS::Descendants, $c)
  if ($e) { return $e }
  return $AE::RootElement.FindFirst($TS::Descendants, $c)
}
function Click([string]$name, [int]$waitMs) {
  for ($i = 0; $i -lt 4; $i++) {
    $b = FindAny $name
    if ($b) {
      try { $b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); Start-Sleep -Milliseconds $waitMs; return $true } catch { }
    }
    Start-Sleep -Milliseconds 600
  }
  return $false
}

# 切到 EXCLE助手 页签，点「功能面板」与「关于」
$tab = FindAny 'EXCLE助手'
if ($tab) { try { $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() } catch { }; Start-Sleep -Milliseconds 1000 }
[void](Click '功能面板' 1500)
[void](Click '关于' 2500)

$xl.Quit()
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }

Write-Host ''
Write-Host '== 产物 =='
foreach ($f in $targets) {
  $p = Join-Path $temp $f
  if (Test-Path $p) { Write-Host ("  OK   {0,-32} {1} bytes" -f $f, (Get-Item $p).Length) }
  else { Write-Host ("  缺失 {0}" -f $f) }
}
$log = Join-Path $temp 'exclehelper-ui.log'
if (Test-Path $log) {
  Write-Host ''
  Write-Host '== UI 渲染诊断 =='
  Get-Content $log | ForEach-Object { Write-Host ("  " + $_) }
}

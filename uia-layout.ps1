$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]
$env:EXCLE_TRACE = '1'
$env:EXCLE_SILENT = '1'

$log = Join-Path $env:TEMP 'exclehelper-load.log'
Remove-Item $log -ErrorAction SilentlyContinue
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Start-Sleep -Milliseconds 700

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
Start-Sleep -Seconds 2
$wb = $xl.Workbooks.Add()
$xl.WindowState = -4137
Start-Sleep -Seconds 3

$win = $AE::RootElement.FindFirst($TS::Children,
  (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'XLMAIN')))
$tabCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::TabItem)
$btnCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Button)

function CountVisibleButtons {
  $w = $AE::RootElement.FindFirst($TS::Children,
    (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'XLMAIN')))
  return ($w.FindAll($TS::Descendants, $btnCond) | Where-Object { $_.Current.Name }).Count
}

Write-Host '=== 每个选项卡可见按钮数（检测功能组是否被折叠）==='
foreach ($name in @('EXCLE助手', '编辑与格式', '汇总拆分')) {
  $t = $win.FindFirst($TS::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $name)))
  if ($null -eq $t) { Write-Host ("  {0}: TAB MISSING" -f $name); continue }
  $t.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
  Start-Sleep -Milliseconds 1400
  $n = CountVisibleButtons
  Write-Host ("  {0,-8} 可见控件={1}" -f $name, $n)
}

Write-Host ''
Write-Host '=== 图标回调 / Ribbon 接受情况 ==='
$lines = Get-Content $log
($lines | Where-Object { $_ -match 'GetCustomUI|OnRibbonLoad|OnGetImage #' }) | ForEach-Object { "  " + $_ }
$imgCalls = ($lines | Where-Object { $_ -match 'OnGetImage #' }).Count
Write-Host ("  getImage 采样行数=" + $imgCalls)

$xl.Quit()
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }

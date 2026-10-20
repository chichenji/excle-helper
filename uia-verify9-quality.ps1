$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]
$env:EXCLE_SILENT = '1'
$env:EXCLE_TRACE = '1'
$qlInput = Join-Path $env:TEMP 'excle-ql-inputs.txt'

$pass = 0; $fail = 0; $failed = @()
$script:xlWin = $null
$script:xl = $null

function Set-QlInput([hashtable]$kv) {
  $lines = @()
  foreach ($k in $kv.Keys) { $lines += ("{0}={1}" -f $k, $kv[$k]) }
  Set-Content -Path $qlInput -Value $lines -Encoding UTF8
}
function Clear-QlInput { if (Test-Path $qlInput) { Remove-Item $qlInput -Force } }
function FindAny([string]$name) {
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $name)
  if ($null -ne $script:xlWin) { $e = $script:xlWin.FindFirst($TS::Descendants, $c); if ($null -ne $e) { return $e } }
  return $AE::RootElement.FindFirst($TS::Descendants, $c)
}
function WaitFor([string]$name, [int]$sec) {
  $end = (Get-Date).AddSeconds($sec)
  while ((Get-Date) -lt $end) { $e = FindAny $name; if ($null -ne $e) { return $e }; Start-Sleep -Milliseconds 300 }
  return $null
}
function Activate([object]$e) {
  try { $p = $e.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern); if ($p) { $p.Invoke(); return $true } } catch { }
  try { $p = $e.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern); if ($p) { if ($p.Current.ExpandCollapseState -eq 'Expanded') { return $true }; $p.Expand(); return $true } } catch { }
  try { $p = $e.GetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern); if ($p) { $p.DoDefaultAction(); return $true } } catch { }
  return $false
}
$script:curTab = $null
function Select-Tab([string]$name) {
  $script:curTab = $name
  $t = WaitFor $name 15
  if ($null -eq $t) { throw "tab '$name' not found" }
  try { $t.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() } catch { Activate $t | Out-Null }
  Start-Sleep -Milliseconds 900
}
function Wait-ExcelIdle([int]$sec) {
  $end = (Get-Date).AddSeconds($sec)
  while ((Get-Date) -lt $end) {
    try { $n = $script:xl.ActiveSheet.Name; if ($null -ne $n) { Start-Sleep -Milliseconds 250; return $true } } catch { Start-Sleep -Milliseconds 400 }
  }
  return $false
}
function Click-Ribbon([string]$item, [string]$menu) {
  for ($attempt = 1; $attempt -le 4; $attempt++) {
    Start-Sleep -Milliseconds 400
    if ($attempt -gt 1 -and $script:curTab) { try { Select-Tab $script:curTab } catch { } }
    $menuOk = $true
    if ($menu) {
      $m = WaitFor $menu 8
      if ($null -eq $m) { $menuOk = $false }
      elseif (-not (Activate $m)) { $menuOk = $false }
      else { Start-Sleep -Milliseconds 1200 }
    }
    if ($menuOk) {
      $b = WaitFor $item 8
      if ($null -ne $b -and (Activate $b)) { Start-Sleep -Milliseconds 1500; try { Wait-ExcelIdle 25 | Out-Null } catch { }; return }
    }
    Start-Sleep -Milliseconds 700
  }
  throw "item '$item' not found after 3 attempts"
}
function Check([string]$name, $actual, $expected) {
  $a = if ($null -eq $actual) { '' } else { ([string]$actual).Trim() }
  $e = if ($null -eq $expected) { '' } else { ([string]$expected).Trim() }
  if ($a -eq $e) { $script:pass++; Write-Host ("  PASS  {0,-26} = {1}" -f $name, $a) }
  else { $script:fail++; $script:failed += $name; Write-Host ("  FAIL  {0,-26} got [{1}] want [{2}]" -f $name, $a, $e) }
}
function SheetExists($wb, [string]$n) { foreach ($s in $wb.Worksheets) { if ($s.Name -eq $n) { return $true } }; return $false }

Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Clear-QlInput
Start-Sleep -Milliseconds 800
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
Start-Sleep -Seconds 2
$wb = $xl.Workbooks.Add(); $xl.WindowState = -4137
Start-Sleep -Seconds 2
$script:xl = $xl
$script:xlWin = $AE::RootElement.FindFirst($TS::Children,
  (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'XLMAIN')))

$ws = $wb.Worksheets[1]

Write-Host ''
Write-Host '=== A2-1 异常值检测（IQR）==='
$ws.Range('A1').Value2 = '金额'
$vals = @(10, 12, 11, 13, 12, 11, 12, 1000, 11, 12)
for ($i = 0; $i -lt $vals.Count; $i++) { $ws.Cells.Item($i + 2, 1).Value2 = $vals[$i] }
$ws.Range('A1:A11').Select() | Out-Null
Set-QlInput @{ mode='2' }
Select-Tab '格式与数据'
Click-Ribbon '异常值检测…' '质量工具'
Check 'A2-1 生成异常清单' (SheetExists $wb '异常值清单') $true
Check 'A2-1 命中 1 个异常' $wb.Worksheets['异常值清单'].UsedRange.Rows.Count 2

Write-Host ''
Write-Host '=== A2-2 数据字典 ==='
$ws.Activate(); $ws.Range('A1:A11').Select() | Out-Null
Click-Ribbon '数据字典' '质量工具'
Check 'A2-2 生成字典' (SheetExists $wb '数据字典') $true
$dd = $wb.Worksheets['数据字典']
Check 'A2-2 字典行数=表头+1' $dd.UsedRange.Rows.Count 2
Check 'A2-2 类型识别为数值' $dd.Range('B2').Value2 '数值'

Write-Host ''
Write-Host '=== A2-3 单位统一 ==='
$ws.Activate()
$ws.Range('C1').Value2 = '1kg'; $ws.Range('C2').Value2 = '500g'; $ws.Range('C3').Value2 = '2吨'
$ws.Range('C1:C3').Select() | Out-Null
Set-QlInput @{ target='kg' }
Click-Ribbon '单位统一…' '质量工具'
Check 'A2-3 1kg→1' $ws.Range('D1').Value2 1
Check 'A2-3 500g→0.5' $ws.Range('D2').Value2 0.5
Check 'A2-3 2吨→2000' $ws.Range('D3').Value2 2000

Write-Host ''
Write-Host '=== A2-4 电话规范化 ==='
$env:EXCLE_SILENT = '1'
$env:EXCLE_TRACE = '1'
$phone = '1' + '39' + '1234' + '5678'
$ws.Range('E1').NumberFormat = '@'
$ws.Range('E1').Formula = "'" + $phone
$ws.Range('E1').Select() | Out-Null
Set-QlInput @{ mode='2' }
Click-Ribbon '电话规范化…' '质量工具'
Check 'A2-4 加 +86 前缀' $ws.Range('E1').Value2 ('+86 ' + $phone)

Write-Host ''
Write-Host '=== A2-5 缺失值处理（上方值）==='
$ws.Range('F1').Value2 = 5; $ws.Range('F2').ClearContents(); $ws.Range('F3').Value2 = 9
$ws.Range('F1:F3').Select() | Out-Null
Set-QlInput @{ mode='4'; header='0' }
Click-Ribbon '缺失值处理…' '质量工具'
Check 'A2-5 空值被上方值填充' $ws.Range('F2').Value2 5

$xl.Quit()
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Clear-QlInput

Write-Host ''
Write-Host ("RESULT: pass={0} fail={1}" -f $pass, $fail)
if ($fail -gt 0) { Write-Host ("failed: " + ($failed -join ', ')); exit 1 }

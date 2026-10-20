$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]
$env:EXCLE_SILENT = '1'
$env:EXCLE_TRACE = '1'

$inputFile = Join-Path $env:TEMP 'excle-jt-inputs.txt'
$pass = 0; $fail = 0; $failed = @()
$script:xlWin = $null

function Set-Input([hashtable]$kv) {
  $lines = @()
  foreach ($k in $kv.Keys) { $lines += ("{0}={1}" -f $k, $kv[$k]) }
  Set-Content -Path $inputFile -Value $lines -Encoding UTF8
}
function Clear-Input { if (Test-Path $inputFile) { Remove-Item $inputFile -Force } }
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
function Click-Ribbon([string]$item, [string]$menu) {
  for ($attempt = 1; $attempt -le 4; $attempt++) {
    Start-Sleep -Milliseconds 400
    $menuOk = $true
    if ($menu) {
      $m = WaitFor $menu 8
      if ($null -eq $m) { $menuOk = $false }
      elseif (-not (Activate $m)) { $menuOk = $false }
      else { Start-Sleep -Milliseconds 1200 }
    }
    if ($menuOk) {
      $b = WaitFor $item 8
      if ($null -ne $b -and (Activate $b)) { Start-Sleep -Milliseconds 1500; return }
    }
    Start-Sleep -Milliseconds 700
  }
  throw "item '$item' not found after 3 attempts"
}
function Check([string]$name, $actual, $expected) {
  $a = if ($null -eq $actual) { '' } else { ([string]$actual).Trim() }
  $e = if ($null -eq $expected) { '' } else { ([string]$expected).Trim() }
  if ($a -eq $e) { $script:pass++; Write-Host ("  PASS  {0,-30} = {1}" -f $name, $a) }
  else { $script:fail++; $script:failed += $name; Write-Host ("  FAIL  {0,-30} got [{1}] want [{2}]" -f $name, $a, $e) }
}
function SheetExists($wb, [string]$n) { foreach ($s in $wb.Worksheets) { if ($s.Name -eq $n) { return $true } }; return $false }

Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Clear-Input
Start-Sleep -Milliseconds 700

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
Start-Sleep -Seconds 2
$wb = $xl.Workbooks.Add()
$xl.WindowState = -4137
Start-Sleep -Seconds 2
$script:xlWin = $AE::RootElement.FindFirst($TS::Children,
  (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'XLMAIN')))

# ---------- 构造数据 ----------
$ws1 = $wb.Worksheets[1]; $ws1.Name = '左表'
$ws1.Range('A1:C1').Value2 = @('编号','名称','部门')
$ws1.Range('A2').Value2 = 1; $ws1.Range('B2').Value2 = '张三'; $ws1.Range('C2').Value2 = '销售'
$ws1.Range('A3').Value2 = 2; $ws1.Range('B3').Value2 = '李四'; $ws1.Range('C3').Value2 = '技术'
$ws1.Range('A4').Value2 = 3; $ws1.Range('B4').Value2 = '王五'; $ws1.Range('C4').Value2 = '销售'
$ws2 = $wb.Worksheets.Add($ws1); $ws2.Name = '右表'
$ws2.Range('A1:C1').Value2 = @('编号','邮箱','工资')
$at = [char]64
$mail1 = 'a' + $at + 'x' + '.' + 'com'
$mail2 = 'b' + $at + 'x' + '.' + 'com'
$mail3 = 'c' + $at + 'x' + '.' + 'com'
$ws2.Range('A2').Value2 = 1; $ws2.Range('B2').Value2 = $mail1; $ws2.Range('C2').Value2 = 8000
$ws2.Range('A3').Value2 = 2; $ws2.Range('B3').Value2 = $mail2; $ws2.Range('C3').Value2 = 9000
$ws2.Range('A4').Value2 = 4; $ws2.Range('B4').Value2 = $mail3; $ws2.Range('C4').Value2 = 7000
$ws3 = $wb.Worksheets.Add($ws1); $ws3.Name = '销售'
$ws3.Range('A1:C1').Value2 = @('地区','产品','金额')
$ws3.Range('A2').Value2 = '华东'; $ws3.Range('B2').Value2 = 'A'; $ws3.Range('C2').Value2 = 100
$ws3.Range('A3').Value2 = '华东'; $ws3.Range('B3').Value2 = 'B'; $ws3.Range('C3').Value2 = 200
$ws3.Range('A4').Value2 = '华南'; $ws3.Range('B4').Value2 = 'A'; $ws3.Range('C4').Value2 = 50
$ws3.Range('A5').Value2 = '华南'; $ws3.Range('B5').Value2 = 'B'; $ws3.Range('C5').Value2 = 80
$ws1.Activate()

Write-Host ''
Write-Host '=== A1-1 左连接 JOIN ==='
Set-Input @{ left='左表!A1:C4'; right='右表!A1:C4'; keyLeft='1'; keyRight='1'; type='1' }
Select-Tab '格式与数据'
Click-Ribbon '多表关联合并（JOIN）…' '多表关联'
Check '生成关联结果表' (SheetExists $wb '关联结果') $true
$jr = $wb.Worksheets['关联结果']
Check '行数=表头+3' $jr.UsedRange.Rows.Count 4
Check '列数=5' $jr.UsedRange.Columns.Count 5
Check 'A2=编号1' $jr.Range('A2').Value2 1
Check 'D2=右表邮箱' $jr.Range('D2').Value2 $mail1
Check '第3行未匹配邮箱为空' ([string]$jr.Range('D4').Value2) ''
for ($r = 1; $r -le 4; $r++) {
  $line = @()
  for ($c = 1; $c -le 5; $c++) { $line += ('[' + [string]$jr.Cells.Item($r, $c).Value2 + ']') }
  Write-Host ("   行{0}: {1}" -f $r, ($line -join ' '))
}
Write-Host '   源表右表 B1:B4 (Formula/Text):'
for ($r = 1; $r -le 4; $r++) {
  $cell = $ws2.Cells.Item($r, 2)
  Write-Host ("     B{0}: formula=[{1}] text=[{2}] value2=[{3}]" -f $r, $cell.Formula, $cell.Text, $cell.Value2)
}
Write-Host '   输出 D1:D4 (Formula/Text):'
for ($r = 1; $r -le 4; $r++) {
  $cell = $jr.Cells.Item($r, 4)
  Write-Host ("     D{0}: formula=[{1}] text=[{2}] value2=[{3}]" -f $r, $cell.Formula, $cell.Text, $cell.Value2)
}

Write-Host ''
Write-Host '=== A1-2 内连接 JOIN ==='
Set-Input @{ left='左表!A1:C4'; right='右表!A1:C4'; keyLeft='1'; keyRight='1'; type='2' }
$jr.Delete()
Select-Tab '格式与数据'
Click-Ribbon '多表关联合并（JOIN）…' '多表关联'
$jr2 = $wb.Worksheets['关联结果']
Check '内连接行数=表头+2' $jr2.UsedRange.Rows.Count 3
Check '内连接不含编号3' ([string]$jr2.Range('A4').Value2) ''

Write-Host ''
Write-Host '=== A1-3 右连接 JOIN ==='
Set-Input @{ left='左表!A1:C4'; right='右表!A1:C4'; keyLeft='1'; keyRight='1'; type='3' }
$jr2.Delete()
Select-Tab '格式与数据'
Click-Ribbon '多表关联合并（JOIN）…' '多表关联'
$jr3 = $wb.Worksheets['关联结果']
Check '右连接行数=表头+3' $jr3.UsedRange.Rows.Count 4
Check '含右表独有编号4' ([string]$jr3.Range('A4').Value2) '4'

Write-Host ''
Write-Host '=== A1-4 交叉表 ==='
Set-Input @{ left='销售!A1:C5'; rowCol='1'; colCol='2'; valCol='3'; fn='1' }
$jr3.Delete()
Select-Tab '格式与数据'
Click-Ribbon '交叉表（行转列）…' '多表关联'
Check '生成交叉表' (SheetExists $wb '交叉表') $true
$ct = $wb.Worksheets['交叉表']
Check '交叉表 华东=A 合计' $ct.Range('B2').Value2 100
Check '交叉表 行合计' $ct.Range('D2').Value2 300
Check '交叉表 总计' $ct.Range('D4').Value2 430

Write-Host ''
Write-Host '=== A1-5 记录级合并 ==='
Set-Input @{ left='销售!A1:C5'; keyLeft='1'; sep='/' }
$ct.Delete()
Select-Tab '格式与数据'
Click-Ribbon '记录级合并…' '多表关联'
Check '生成记录合并表' (SheetExists $wb '记录合并') $true
$gm = $wb.Worksheets['记录合并']
Check '合并后行数=表头+2' $gm.UsedRange.Rows.Count 3
Check '多值用 / 连接' $gm.Range('B2').Value2 'A/B'

Write-Host ''
Write-Host '=== A1-6 按主键刷新 ==='
Set-Input @{ left='左表!A1:D4'; right='右表!A1:C4'; keyLeft='1'; keyRight='1' }
$gm.Delete()
# 给左表加一个同名可更新列
$ws1.Range('D1').Value2 = '邮箱'
$ws1.Activate(); $ws1.Range('A1:D4').Select() | Out-Null
Select-Tab '格式与数据'
Click-Ribbon '按主键刷新字段…' '多表关联'
Check '刷新后 D2=右表邮箱' $ws1.Range('D2').Value2 $mail1
Check '未匹配行 D4 仍空' ([string]$ws1.Range('D4').Value2) ''

$xl.Quit()
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Clear-Input

Write-Host ''
Write-Host ("RESULT: pass={0} fail={1}" -f $pass, $fail)
if ($fail -gt 0) { Write-Host ("failed: " + ($failed -join ', ')); exit 1 }

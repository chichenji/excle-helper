$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]
$env:EXCLE_SILENT = '1'
$env:EXCLE_AI_KEY = 'test-key-123456'
$env:EXCLE_TRACE = '1'

$aiInput = Join-Path $env:TEMP 'excle-ai-inputs.txt'
$mockDir = Join-Path $env:TEMP 'excle-ai2-mock'
New-Item -ItemType Directory -Force -Path $mockDir | Out-Null
$modeFile = Join-Path $mockDir 'mode.txt'
Set-Content -Path $modeFile -Value 'fill' -Encoding ASCII

function Set-AiInput([hashtable]$kv) {
  $lines = @()
  foreach ($k in $kv.Keys) { $lines += ("{0}={1}" -f $k, $kv[$k]) }
  Set-Content -Path $aiInput -Value $lines -Encoding UTF8
}

# ---------- mock 供应商（按模式返回不同结构） ----------
$mockScript = {
  param($modeFile)
  $payloads = @{
    'fill'     = '{"values":["阿里","腾讯","百度"]}'
    'classify' = '{"labels":["咨询","投诉","咨询","其他"]}'
    'extract'  = '{"rows":[{"姓名":"张三","金额":"100"},{"姓名":"李四","金额":"200"},{"姓名":"王五","金额":"300"},{"姓名":"赵六","金额":"400"}]}'
    'filter'   = '{"logic":"and","conditions":[{"column":"地区","op":"eq","value":"华东"},{"column":"金额","op":"gt","value":"100"}]}'
    'dict'     = '{"columns":[{"name":"地区","type":"文本","meaning":"销售区域","sample":"华东"},{"name":"金额","type":"数值","meaning":"销售额","sample":"100"},{"name":"产品","type":"文本","meaning":"产品代号","sample":"A"}]}'
    'insight'  = '1) 数据概况：共 4 行，覆盖 2 个地区。2) 关键发现：华东金额明显高于华南。3) 风险：样本量过小。4) 建议：用柱形图对比地区。'
  }
  $listener = New-Object System.Net.HttpListener
  $listener.Prefixes.Add('http://127.0.0.1:8766/')
  $listener.Start()
  while ($true) {
    try {
      $ctx = $listener.GetContext()
      $mode = 'fill'
      if (Test-Path $modeFile) { $mode = (Get-Content $modeFile -Raw).Trim() }
      $reader = New-Object System.IO.StreamReader($ctx.Request.InputStream)
      $null = $reader.ReadToEnd()
      $content = '{}'
      if ($payloads.ContainsKey($mode)) { $content = $payloads[$mode] }
      $body = '{"choices":[{"message":{"content":' + (ConvertTo-Json $content -Compress) + '}}]}'
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
      $ctx.Response.ContentType = 'application/json'
      $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
      $ctx.Response.Close()
    } catch { }
  }
}
$job = Start-Job -ScriptBlock $mockScript -ArgumentList $modeFile
Start-Sleep -Seconds 2
"mock 已启动"

$cfgDir = Join-Path $env:APPDATA 'ExcleHelper'
New-Item -ItemType Directory -Force -Path $cfgDir | Out-Null
@'
provider=custom
baseurl=http://127.0.0.1:8766/v1
model=mock-model
temperature=0
maxtokens=1024
timeout=20
streaming=0
jsonmode=0
redact=1
contextrows=3
'@ | Set-Content -Path (Join-Path $cfgDir 'ai.ini') -Encoding UTF8
$env:EXCLE_AI_AUTOAPPLY = '1'

# ---------- 工具 ----------
$pass = 0; $fail = 0; $failed = @()
$script:xlWin = $null
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
      if ($null -ne $b -and (Activate $b)) { Start-Sleep -Milliseconds 600; Wait-ExcelIdle 25 | Out-Null; return }
    }
    Start-Sleep -Milliseconds 700
  }
  throw "item '$item' not found after 3 attempts"
}
function Wait-ExcelIdle([int]$sec) {
  $end = (Get-Date).AddSeconds($sec)
  while ((Get-Date) -lt $end) {
    try { $n = $script:xl.ActiveSheet.Name; if ($null -ne $n) { Start-Sleep -Milliseconds 250; return $true } }
    catch { Start-Sleep -Milliseconds 400 }
  }
  return $false
}
function Check([string]$name, $actual, $expected) {
  $a = if ($null -eq $actual) { '' } else { ([string]$actual).Trim() }
  $e = if ($null -eq $expected) { '' } else { ([string]$expected).Trim() }
  if ($a -eq $e) { $script:pass++; Write-Host ("  PASS  {0,-28} = {1}" -f $name, $a) }
  else { $script:fail++; $script:failed += $name; Write-Host ("  FAIL  {0,-28} got [{1}] want [{2}]" -f $name, $a, $e) }
}
function SheetExists($wb, [string]$n) { foreach ($s in $wb.Worksheets) { if ($s.Name -eq $n) { return $true } }; return $false }

Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
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
Write-Host '=== B-1 AI 智能填充 ==='
$ws.Range('A1').Value2 = '阿里巴巴集团'
$ws.Range('A2').Value2 = '腾讯控股'
$ws.Range('A3').Value2 = '百度在线'
$ws.Range('A4').ClearContents(); $ws.Range('A5').ClearContents(); $ws.Range('A6').ClearContents()
$ws.Range('A1:A6').Select() | Out-Null
Set-Content -Path $modeFile -Value 'fill' -Encoding ASCII
Set-AiInput @{ ask='把公司全称转成简称'; examples='3' }
Select-Tab '工具与输出'
Click-Ribbon '智能填充…' '数据加工'
Check 'B-1 填充 A4' $ws.Range('A4').Value2 '阿里'
Check 'B-1 填充 A6' $ws.Range('A6').Value2 '百度'

Write-Host ''
Write-Host '=== B-2 AI 分类打标 ==='
$ws.Range('C1').Value2 = '我要退款'
$ws.Range('C2').Value2 = '服务态度差'
$ws.Range('C3').Value2 = '怎么开通'
$ws.Range('C4').Value2 = '随便'
$ws.Range('C1:C4').Select() | Out-Null
Set-Content -Path $modeFile -Value 'classify' -Encoding ASCII
Set-AiInput @{ ask='投诉,咨询,建议' }
Click-Ribbon '分类打标…' '数据加工'
Check 'B-2 第1条标签' $ws.Range('D1').Value2 '咨询'
Check 'B-2 第2条标签' $ws.Range('D2').Value2 '投诉'

Write-Host ''
Write-Host '=== B-3 AI 自然语言筛选 ==='
$ws.Range('A10').Value2 = '地区'; $ws.Range('B10').Value2 = '金额'; $ws.Range('C10').Value2 = '产品'
$ws.Range('A11').Value2 = '华东'; $ws.Range('B11').Value2 = 500; $ws.Range('C11').Value2 = 'A'
$ws.Range('A12').Value2 = '华南'; $ws.Range('B12').Value2 = 800; $ws.Range('C12').Value2 = 'B'
$ws.Range('A13').Value2 = '华东'; $ws.Range('B13').Value2 = 80; $ws.Range('C13').Value2 = 'C'
$ws.Range('A14').Value2 = '华东'; $ws.Range('B14').Value2 = 300; $ws.Range('C14').Value2 = 'A'
$ws.Range('A10:C14').Select() | Out-Null
Set-Content -Path $modeFile -Value 'filter' -Encoding ASCII
Set-AiInput @{ ask='华东且金额大于100' }
Click-Ribbon '自然语言筛选…' '洞察与建议'
Check 'B-3 生成筛选结果表' (SheetExists $wb 'AI筛选结果') $true
$fs = $wb.Worksheets['AI筛选结果']
Check 'B-3 命中行数=表头+2' $fs.UsedRange.Rows.Count 3

Write-Host ''
Write-Host '=== B-4 AI 数据字典 ==='
$ws.Activate(); $ws.Range('A10:C14').Select() | Out-Null
Set-Content -Path $modeFile -Value 'dict' -Encoding ASCII
Set-AiInput @{ }
Click-Ribbon '数据字典' '洞察与建议'
Check 'B-4 生成数据字典表' (SheetExists $wb '数据字典') $true
$dd = $wb.Worksheets['数据字典']
Check 'B-4 字典列数=4' $dd.UsedRange.Columns.Count 4
Check 'B-4 首列含义' $dd.Range('C2').Value2 '销售区域'

Write-Host ''
Write-Host '=== B-5 AI 数据洞察（文本）==='
$ws.Activate(); $ws.Range('A10:C14').Select() | Out-Null
Set-Content -Path $modeFile -Value 'insight' -Encoding ASCII
Set-AiInput @{ }
Click-Ribbon '数据洞察' '洞察与建议'
Check 'B-5 洞察不写单元格' ([string]$ws.Range('Z1').Value2) ''

$xl.Quit()
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Stop-Job $job -ErrorAction SilentlyContinue | Out-Null
Remove-Job $job -Force -ErrorAction SilentlyContinue | Out-Null
Remove-Item $aiInput -Force -ErrorAction SilentlyContinue

Write-Host ''
Write-Host ("RESULT: pass={0} fail={1}" -f $pass, $fail)
if ($fail -gt 0) { Write-Host ("failed: " + ($failed -join ', ')); exit 1 }

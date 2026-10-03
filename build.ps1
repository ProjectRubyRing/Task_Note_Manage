<#
.SYNOPSIS
  src フォルダの VBA ソースから「付箋タスクボード.xlsm」を作ります。

.DESCRIPTION
  Excel を画面に出さずに起動し、新しいブックへモジュールと入力画面を組み込み、
  modSetup の SetupWorkbook でシート・ボタン・サンプルを作ってから .xlsm で保存します。

  ・VBA のコードを書き込むため、実行中だけ Excel のセキュリティ設定
    「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」(AccessVBOM) を有効にし、
    終わったら元の状態に戻します。
  ・保存先に同じ名前のファイルがある場合は、-Force を付けないと上書きしません
    （上書きすると、そのファイルに貼ってある付箋は消えます）。
  ・src の .bas / .cls は UTF-8 で保存しています。VBE の「ファイルのインポート」では
    文字化けするため、このスクリプトから組み込んでください。

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\build.ps1
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\build.ps1 -OutFile .\test.xlsm -Force
#>
param(
    [string]$OutFile = (Join-Path $PSScriptRoot '付箋タスクボード.xlsm'),
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$src = Join-Path $PSScriptRoot 'src'
$OutFile = [IO.Path]::GetFullPath($OutFile)

if ((Test-Path -LiteralPath $OutFile) -and -not $Force) {
    throw "$OutFile は既にあります。上書きする場合は -Force を付けてください（貼ってある付箋は消えます）。"
}

# VBA は Shift_JIS (CP932) で保存されるので、表せない文字がないかを先に確かめる
$sjis = [Text.Encoding]::GetEncoding(932, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
function Read-Source([string]$name) {
    $text = [IO.File]::ReadAllText((Join-Path $src $name), [Text.Encoding]::UTF8) -replace "`r?`n", "`r`n"
    try { [void]$sjis.GetBytes($text) }
    catch { throw "$name に Shift_JIS で表せない文字があります（ChrW で書いてください）: $($_.Exception.Message)" }
    return $text
}

function Set-ModuleCode($component, [string]$code) {
    $cm = $component.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.AddFromString($code)
}

$sources = @{}
foreach ($f in 'modConfig.bas', 'modNote.bas', 'modMain.bas', 'modSetup.bas', 'frmNote.cls', 'ThisWorkbook.cls') {
    $sources[$f] = Read-Source $f
}

# Excel のバージョン（例: Excel.Application.16 → 16.0）
$curVer = (Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Excel.Application\CurVer').'(default)'
$ver = ($curVer -replace '^Excel\.Application\.', '') + '.0'
$secKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
if (-not (Test-Path $secKey)) { New-Item -Path $secKey -Force | Out-Null }
$prevVbom = (Get-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM

Add-Type -Namespace StickyBuild -Name Win32 -MemberDefinition @'
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(System.IntPtr hWnd, out int pid);
'@

function Restore-Vbom {
    if ($null -eq $prevVbom) { Remove-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { Set-ItemProperty -Path $secKey -Name AccessVBOM -Value $prevVbom -Type DWord }
}

$xl = $null
$wb = $null
$xlPid = 0
try {
    Set-ItemProperty -Path $secKey -Name AccessVBOM -Value 1 -Type DWord

    $xl = New-Object -ComObject Excel.Application
    [void][StickyBuild.Win32]::GetWindowThreadProcessId([IntPtr][long]$xl.Hwnd, [ref]$xlPid)
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    # シート 4 枚の新しいブックを作る
    # （後からシートを追加すると VBA 側のモジュールがすぐには作られないため、最初から 4 枚で作る）
    $origSheets = $xl.SheetsInNewWorkbook
    $xl.SheetsInNewWorkbook = 4
    try { $wb = $xl.Workbooks.Add() } finally { $xl.SheetsInNewWorkbook = $origSheets }
    $vbp = $wb.VBProject

    # 各シートに VBA から使う名前（CodeName）を付ける
    $codeNames = 'shBoard', 'shDone', 'shConfig', 'shHelp'
    for ($i = 1; $i -le 4; $i++) {
        $cn = $wb.Worksheets.Item($i).CodeName
        if (-not $cn) { throw "シート $i の CodeName を取得できません。" }
        $vbp.VBComponents.Item($cn).Name = $codeNames[$i - 1]
    }

    # 標準モジュール・入力画面・ThisWorkbook
    foreach ($m in 'modConfig', 'modNote', 'modMain', 'modSetup') {
        $c = $vbp.VBComponents.Add(1)      # vbext_ct_StdModule
        $c.Name = $m
        Set-ModuleCode $c $sources["$m.bas"]
    }
    $frm = $vbp.VBComponents.Add(3)        # vbext_ct_MSForm
    $frm.Name = 'frmNote'
    Set-ModuleCode $frm $sources['frmNote.cls']
    Set-ModuleCode ($vbp.VBComponents.Item($wb.CodeName)) $sources['ThisWorkbook.cls']

    # シート・ボタン・サンプルを作る
    $result = $xl.Run('SetupWorkbook')
    if ($result) { throw "SetupWorkbook でエラーが発生しました: $result" }

    $wb.SaveAs($OutFile, 52)               # xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false)
    $wb = $null
    Write-Host "作成しました: $OutFile"
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $xl) {
        try { $xl.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
    }
    $wb = $null
    $vbp = $null
    $c = $null
    $frm = $null
    $xl = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    # Excel は終了するときに設定を書き戻すので、終了を待ってからセキュリティ設定を元に戻す
    if ($xlPid -gt 0) {
        Wait-Process -Id $xlPid -Timeout 60 -ErrorAction SilentlyContinue
        # このスクリプトが起動した（画面に出ていない）Excel が残っていたら終了させる
        if (Get-Process -Id $xlPid -ErrorAction SilentlyContinue) {
            Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 500
        }
    }
    Restore-Vbom
    Start-Sleep -Milliseconds 500
    Restore-Vbom
    Write-Host "Excel のセキュリティ設定（AccessVBOM）を元に戻しました。"
}

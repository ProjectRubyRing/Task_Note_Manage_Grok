# Builds タスク付箋ノート.xlsm and runs the in-Excel self-test twice
# (once in the build session, once after reopening the saved file).
$ErrorActionPreference = 'Stop'

$src = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $src
$outPath = Join-Path $root 'タスク付箋ノート.xlsm'
$logPath = Join-Path $env:TEMP 'sticky-selftest-result.txt'
$tracePath = Join-Path $env:TEMP 'sticky-build-trace.txt'
if (Test-Path -LiteralPath $tracePath) {
    Remove-Item -LiteralPath $tracePath -Force
}

$secPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$accessExisted = $false
$accessPrev = 0
if (-not (Test-Path $secPath)) {
    New-Item -Path $secPath -Force | Out-Null
}
$existingAccess = Get-ItemProperty -Path $secPath -Name 'AccessVBOM' -ErrorAction SilentlyContinue
if ($null -ne $existingAccess) {
    $accessExisted = $true
    $accessPrev = [int]$existingAccess.AccessVBOM
}
New-ItemProperty -Path $secPath -Name 'AccessVBOM' -Value 1 -PropertyType DWord -Force | Out-Null

if (-not ('NativePid' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class NativePid {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(int hWnd, out uint processId);
}
"@
}

function Get-ExcelProcessId($app) {
    $processId = [uint32]0
    try {
        $hwnd = [int]$app.Hwnd
        if ($hwnd -ne 0) {
            [NativePid]::GetWindowThreadProcessId($hwnd, [ref]$processId) | Out-Null
        }
    } catch {
        $processId = [uint32]0
    }
    return [int]$processId
}

function Read-Source([string]$path) {
    $text = [System.IO.File]::ReadAllText($path, [System.Text.UTF8Encoding]::new($false))
    if ($text.Length -gt 0 -and [int][char]$text[0] -eq 0xFEFF) {
        $text = $text.Substring(1)
    }
    return ($text -replace "`r`n", "`n" -replace "`n", "`r`n")
}

function Set-VbaCode($component, [string]$code) {
    $cm = $component.CodeModule
    if ($cm.CountOfLines -gt 0) {
        $cm.DeleteLines(1, $cm.CountOfLines)
    }
    $payload = $code.TrimEnd()
    if ($payload.Length -le 60000) {
        $cm.AddFromString($payload)
        return
    }
    $lines = $payload -split "`r`n", -1
    $index = 0
    foreach ($line in $lines) {
        $index++
        $cm.InsertLines($index, $line)
    }
}

function Release-Com($object) {
    if ($null -ne $object) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object) } catch { }
    }
}

function Stop-ExcelInstance($app, [int]$processId) {
    if ($null -ne $app) {
        try { $app.DisplayAlerts = $false } catch { }
        try { $app.Quit() } catch { }
    }
    Release-Com $app
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($processId -gt 0) {
        $still = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($null -ne $still) {
            Start-Sleep -Milliseconds 500
            $still = Get-Process -Id $processId -ErrorAction SilentlyContinue
            if ($null -ne $still) {
                Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

function New-ExcelApp {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.AskToUpdateLinks = $false
    $app.AutomationSecurity = 1
    $app.EnableEvents = $true
    return $app
}

function Add-StdModule($project, [string]$name, [string]$code) {
    $comp = $project.VBComponents.Add(1)
    $comp.Name = $name
    Set-VbaCode $comp $code
    return $comp
}

function Get-VBComponent($project, [string]$name) {
    foreach ($comp in @($project.VBComponents)) {
        if ([string]$comp.Name -eq $name) {
            return $comp
        }
    }
    throw "VBA component not found: $name"
}

function Remove-ModuleIfPresent($project, [string]$name) {
    foreach ($comp in @($project.VBComponents)) {
        if ([string]$comp.Name -eq $name) {
            $project.VBComponents.Remove($comp)
            return
        }
    }
}

function Get-ComponentNames($project) {
    $names = @()
    foreach ($comp in @($project.VBComponents)) {
        $names += $comp.Name
    }
    return ($names -join ', ')
}

function Find-StickyTitle($worksheet, [string]$title, [int]$lanes) {
    for ($lane = 0; $lane -lt $lanes; $lane++) {
        for ($slot = 0; $slot -lt 25; $slot++) {
            $row = 6 + ($slot * 9) + 2
            $col = 3 + ($lane * 3)
            $value = $worksheet.Cells.Item($row, $col).Value2
            if ([string]$value -eq $title) {
                return @{ Lane = $lane; Slot = $slot; Row = $row; Col = $col }
            }
        }
    }
    return $null
}

function Write-Log([string]$text) {
    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($logPath, $text, $utf8)
    Write-Output $text
}

function Restore-AccessVbom {
    if ($script:accessExisted) {
        New-ItemProperty -Path $script:secPath -Name 'AccessVBOM' -Value $script:accessPrev -PropertyType DWord -Force | Out-Null
    } else {
        Remove-ItemProperty -Path $script:secPath -Name 'AccessVBOM' -ErrorAction SilentlyContinue
    }
}

$excel = $null
$excelPid = 0
$wb = $null
$passedBuild = $false

try {
try {
    $excel = New-ExcelApp
    $excelPid = Get-ExcelProcessId $excel
    Write-Output "EXCEL_PID=$excelPid"

    $wb = $excel.Workbooks.Add()
    $project = $wb.VBProject
    Add-StdModule $project 'modCore' (Read-Source (Join-Path $src 'modCore.bas')) | Out-Null
    Add-StdModule $project 'modActions' (Read-Source (Join-Path $src 'modActions.bas')) | Out-Null
    $thisName = [string]$wb.CodeName
    Write-Output "CODENAME=$thisName"
    Set-VbaCode (Get-VBComponent $project $thisName) (Read-Source (Join-Path $src 'ThisWorkbook.cls'))
    Add-StdModule $project 'modBuild' "Option Explicit`r`nPublic Sub GoBuild()`r`n    modCore.BuildAll`r`nEnd Sub`r`n" | Out-Null

    $coreLines = (Get-VBComponent $project 'modCore').CodeModule.CountOfLines
    $buildLines = (Get-VBComponent $project 'modBuild').CodeModule.CountOfLines
    Write-Output "CORE_LINES=$coreLines BUILD_LINES=$buildLines"
    Write-Output 'BUILD_START'
    $excel.Run('GoBuild')
    Write-Output 'BUILD_DONE'
    Remove-ModuleIfPresent $project 'modBuild'

    Add-StdModule $project 'modSelfTest' (Read-Source (Join-Path $src 'modSelfTest.bas')) | Out-Null
    Write-Output 'TEST1_START'
    $result1 = [string]$excel.Run('RunSelfTest')
    Write-Log $result1
    Remove-ModuleIfPresent $project 'modSelfTest'
    $components = Get-ComponentNames $project
    Write-Output "COMPONENTS=$components"
    if ($components -match 'modSelfTest' -or $components -match 'modBuild') {
        throw "Temporary modules remained: $components"
    }
    if ($result1 -notmatch '(?m)^ALL PASS\s*$' -or $result1 -match '(?m)^FAIL ') {
        throw 'Self-test failed in the build session.'
    }

    $excel.DisplayAlerts = $false
    if (Test-Path $outPath) {
        Remove-Item -LiteralPath $outPath -Force
    }
    $wb.Worksheets.Item('ノート').Activate() | Out-Null
    # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.SaveAs($outPath, 52)
    $wb.Close($true)
    $passedBuild = $true
    Write-Output "SAVED=$outPath"
}
finally {
    if ($null -ne $wb -and -not $passedBuild) {
        try { $wb.Close($false) } catch { }
    }
    Release-Com $wb
    $wb = $null
    Stop-ExcelInstance $excel $excelPid
    $excel = $null
}

if (-not $passedBuild) {
    throw 'Build session did not finish.'
}

$excel2 = $null
$excel2Pid = 0
$wb2 = $null
try {
    $excel2 = New-ExcelApp
    $excel2Pid = Get-ExcelProcessId $excel2
    Write-Output "EXCEL2_PID=$excel2Pid"
    $wb2 = $excel2.Workbooks.Open($outPath)
    $board = $wb2.Worksheets.Item('ノート')
    $done = $wb2.Worksheets.Item('完了')

    $locked = [bool]$board.Range('A1').Locked
    $protected = [bool]$board.ProtectContents
    $titleText = [string]$board.Range('A1').Value2
    if (-not $protected -or -not $locked) {
        throw "Sheet protection is not active. ProtectContents=$protected Locked=$locked Title='$titleText'"
    }
    Write-Output "PROTECT_OK contents=$protected locked=$locked"

    $hit = Find-StickyTitle $board '週次レポートの自動化' 4
    if ($null -eq $hit) { throw 'Sample note was not on the board after reopen.' }
    $titleCell = $board.Cells.Item($hit.Row, $hit.Col)
    $fontName = [string]$titleCell.Font.Name
    $bodyColor = [int]$board.Cells.Item($hit.Row + 2, $hit.Col).Interior.Color
    $expectBody = 255 + (236 * 256) + (130 * 65536)
    Write-Output "REOPEN_SAMPLE lane=$($hit.Lane) slot=$($hit.Slot) font=$fontName bodyColor=$bodyColor expect=$expectBody"
    if ($hit.Lane -ne 0 -or $hit.Slot -ne 0) { throw 'Sample note moved after reopen.' }
    if ($fontName -ne 'Meiryo UI') { throw "Font after reopen was '$fontName'." }
    if ($bodyColor -ne $expectBody) { throw "Body color after reopen was $bodyColor." }

    $doneHit = Find-StickyTitle $done '開発環境の初期構築' 3
    if ($null -eq $doneHit) { throw 'Completed sample was not on the completed sheet after reopen.' }
    $doneFontColor = [int]$done.Cells.Item($doneHit.Row, $doneHit.Col).Font.Color
    $expectGray = 80 + (80 * 256) + (80 * 65536)
    Write-Output "REOPEN_DONE lane=$($doneHit.Lane) slot=$($doneHit.Slot) fontColor=$doneFontColor"
    if ($doneFontColor -ne $expectGray) { throw "Completed note font color was $doneFontColor." }

    $project2 = $wb2.VBProject
    Add-StdModule $project2 'modSelfTest' (Read-Source (Join-Path $src 'modSelfTest.bas')) | Out-Null
    Write-Output 'TEST2_START'
    $result2 = [string]$excel2.Run('RunSelfTest')
    Write-Log ($result1 + "`r`n----- REOPEN -----`r`n" + $result2)
    if ($result2 -notmatch '(?m)^ALL PASS\s*$' -or $result2 -match '(?m)^FAIL ') {
        throw 'Self-test failed after reopening the saved workbook.'
    }
    Write-Output 'TEST2_PASS'
    $wb2.Close($false)
    $wb2 = $null
}
finally {
    if ($null -ne $wb2) {
        try { $wb2.Close($false) } catch { }
    }
    Release-Com $wb2
    Stop-ExcelInstance $excel2 $excel2Pid
}

if (-not (Test-Path -LiteralPath $outPath)) {
    throw 'Workbook was not saved.'
}
$item = Get-Item -LiteralPath $outPath
Write-Output "FILE_BYTES=$($item.Length)"
Write-Output 'ALL_CHECKS_PASSED'

}
finally {
    Restore-AccessVbom
}

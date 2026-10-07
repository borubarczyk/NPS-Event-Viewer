#Requires -Version 5.1
<#
.SYNOPSIS
    Test dymny NPS Event Viewer (Windows PowerShell 5.1, WPF) - uruchamiany w GitHub Actions.
.DESCRIPTION
    Wczytuje Logi_NPS.ps1 bez pokazywania okna ($global:NpsViewerNoShow), sprawdza logikę osi
    czasu (tests\HistoryLogic.Tests.ps1), otwiera okno główne i okna historii (render do PNG w
    smoke-output), wykonuje prawdziwe zapytanie Get-WinEvent z filtrem MAC, wczytuje testowe pliki .log,
    sprawdza zamykanie okna w trakcie wczytywania i mierzy czas wątku okna przy 20000 zdarzeń.
    Uruchom: powershell -NoProfile -STA -ExecutionPolicy Bypass -File tests\Smoke-NpsViewer.ps1
#>
param([string]$OutDir = (Join-Path $PSScriptRoot '..\smoke-output'))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$global:NpsViewerNoShow = $true
. (Join-Path $PSScriptRoot '..\Logi_NPS.ps1')
$ErrorActionPreference = 'Stop'

function Invoke-Pump([int]$Ms = 300) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $Ms) {
        $frame = New-Object System.Windows.Threading.DispatcherFrame
        [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{ $frame.Continue = $false })
        [System.Windows.Threading.Dispatcher]::PushFrame($frame)
        Start-Sleep -Milliseconds 15
    }
}

function Save-Png($Window, [string]$Name) {
    $Window.UpdateLayout()
    $root = [System.Windows.Media.VisualTreeHelper]::GetChild($Window, 0)
    $w = [int][Math]::Ceiling($root.ActualWidth); $h = [int][Math]::Ceiling($root.ActualHeight)
    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($w, $h, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($root)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $path = Join-Path $OutDir "$Name.png"
    $fs = [IO.File]::Create($path); try { $enc.Save($fs) } finally { $fs.Close() }
    "PNG  $Name ($w x $h)"
}

# Tekst, który nie mieści się w swoim polu (ucięty bez wielokropka) - jak lint w SmartDeployTool.
function Find-Clipping($Window, [string]$Name) {
    $out = New-Object System.Collections.Generic.List[string]
    $stack = New-Object System.Collections.Stack
    $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($Window, 0))
    while ($stack.Count) {
        $v = $stack.Pop()
        if ($v -is [System.Windows.UIElement] -and -not $v.IsVisible) { continue }
        if ($v -is [System.Windows.Controls.TextBlock] -and $v.Text -and $v.TextTrimming -eq 'None' -and $v.TextWrapping -eq 'NoWrap') {
            $ft = New-Object System.Windows.Media.FormattedText($v.Text, [Globalization.CultureInfo]::CurrentCulture, 'LeftToRight',
                (New-Object System.Windows.Media.Typeface($v.FontFamily, $v.FontStyle, $v.FontWeight, $v.FontStretch)), $v.FontSize, [System.Windows.Media.Brushes]::Black, 1.0)
            if ($ft.Width -gt $v.ActualWidth + 4 -and $v.ActualWidth -gt 0) { $out.Add("CLIP $Name | '$($v.Text)' $([int]$ft.Width)px > $([int]$v.ActualWidth)px") }
        }
        for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v); $i++) { $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i)) }
    }
    return $out
}

$fail = 0
function Check($name, $cond) { if ($cond) { "ok   $name" } else { "FAIL $name"; $script:fail++ } }

try {
# --- Logika osi czasu (wspólne testy, także dla pwsh bez WPF) ---
. (Join-Path $PSScriptRoot 'HistoryLogic.Tests.ps1') -LogDir (Join-Path $OutDir 'logs')
$evs = $script:TestEvs; $lgs = $script:TestLgs

function Wait-History($H, [int]$Sec = 120) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($H.Busy -and $sw.Elapsed.TotalSeconds -lt $Sec) { Invoke-Pump 200 }
    Invoke-Pump 200
}

# --- Okno główne z danymi testowymi ---
$script:Ctx.Ev.All = New-Object System.Collections.Generic.List[object]
foreach ($e in $evs) { $script:Ctx.Ev.All.Add($e) }
$script:Ctx.Log.All = New-Object System.Collections.Generic.List[object]
foreach ($e in $lgs) { $script:Ctx.Log.All.Add($e) }
foreach ($n in 'Ev', 'Log') { Update-FilterCombos $script:Ctx[$n]; Invoke-Filter $script:Ctx[$n] }
$win.Show(); Invoke-Pump 500
Save-Png $win '01-okno-glowne'
Check 'main grid rows' ($script:Ctx.Ev.Grid.Items.Count -eq $evs.Count)
Check 'details text' ((Get-DetailsText 'Ev' $evs[4]) -like '*Złe poświadczenia*')

# --- Okno historii z danych zakładek (dopasowanie w runspace w tle) ---
$H = Show-HistoryWindow -Mode 'Mac' -Value 'aa:bb:cc:dd:ee:ff'
Wait-History $H
"STATUS: $($H.Ui.hStatus.Text)"
Check 'history window rows' ($H.Rows.Count -eq 13)
Check 'history list bound' ($H.Ui.hList.Items.Count -eq 13)
Check 'history chips' ($H.Ui.hChips.Children.Count -ge 6)
Check 'history related pivots' ($H.Ui.hRelated.Children.Count -ge 2)
Check 'history strip drawn' ($H.Ui.hStrip.Children.Count -gt 0)
foreach ($c in (Find-Clipping $H.Win 'historia')) { $c; $script:fail++ }
Save-Png $H.Win '02-historia-mac'
$H.Ui.hList.SelectedIndex = 1; Invoke-Pump 200
Check 'history details group' ($H.Ui.hDetails.Text -like 'GRUPA: 3*')
$H.Ui.hNewest.IsChecked = $true; Invoke-Pump 300
Check 'history newest' ($H.Rows[0].Kind -eq 'Day' -and $H.Rows[0].Time.Day -eq 7)
Select-HistoryBucket $H ($H.StripLeft + 1); Invoke-Pump 200
Check 'history zoom from strip' ($null -ne $H.ZoomFrom -and $H.Ui.hZoomClear.Visibility -eq 'Visible')
Select-HistoryBucket $H ($H.StripLeft + $H.StripBw * ($H.Buckets.Count - 0.5)); Invoke-Pump 200
Check 'last strip bar includes newest event' (@($H.View | Where-Object { $_.Time -eq [datetime]'2026-10-07 07:00:00' }).Count -eq 1)
Set-HistoryZoom $H $null $null; Invoke-Pump 200
Save-Png $H.Win '03-historia-najnowsze'
$text = Get-HistoryText $H
Check 'history text' ($text -like '*Odmowa - Złe poświadczenia (kod 16)*')
# Nieudany start (zły własny zakres) nie zmienia tytułu ani zapytania pokazanej osi czasu
$H.Ui.hMode.SelectedIndex = 1; $H.Ui.hValue.Text = 'ola'; $H.Ui.hRange.SelectedIndex = 6; $H.Ui.hFrom.Text = 'zła data'
Start-HistoryLoad $H; Invoke-Pump 200
Check 'failed start keeps query and title' ($H.Q.Mode -eq 'Mac' -and $H.Ui.hTitle.Text -like '*MAC AA-BB-CC-DD-EE-FF*' -and -not $H.Busy -and $H.Ui.hStatus.Text -like 'Nieprawidłowy zakres*')
$H.Ui.hMode.SelectedIndex = 0; $H.Ui.hValue.Text = 'aa:bb:cc:dd:ee:ff'; $H.Ui.hRange.SelectedIndex = 0; $H.Ui.hFrom.Text = ''
Open-HistoryPivot @{ Id = $H.Id; Mode = 'User'; Value = 'CONTOSO\jan' }; Invoke-Pump 300
$H2 = $script:HistWins[$H.Id + 1]
if ($H2) { Wait-History $H2 }
Check 'pivot window user' ($H2 -and $H2.Q.Mode -eq 'User' -and $H2.Items.Count -ge 6)
if ($H2) { Save-Png $H2.Win '04-historia-uzytkownik' }

# --- Prawdziwe zapytanie do dziennika Security (zapytanie XML z wariantami MAC) i pliki .log ---
$ui.txtLogPath.Text = $script:TestLogDir
$H3 = Show-HistoryWindow -Mode 'Mac' -Value 'AA-BB-CC-DD-EE-FF' -RangeIndex 2
$H3.Ui.hSrcLog.IsChecked = $true
Start-HistoryLoad $H3
Wait-History $H3
"STATUS (Security + .log): $($H3.Ui.hStatus.Text)"
Check 'load finished' (-not $H3.Busy)
Check 'load without errors' ($H3.Ui.hStatus.Text -notlike '*BŁĘDY*')
Check 'load log entries (no Access-Request)' (@($H3.Items | Where-Object Src -eq 'Log').Count -eq 2 -and -not $H3.ReqLoaded)
Save-Png $H3.Win '05-historia-wczytana'
$H3.Ui.hReq.IsChecked = $true
Wait-History $H3
Check 'Access-Request loaded on demand' (@($H3.Items | Where-Object Src -eq 'Log').Count -eq 3 -and $H3.ReqLoaded)

$H3.Ui.hPartial.IsChecked = $true
$H3.Ui.hSrcLog.IsChecked = $false
Start-HistoryLoad $H3
Wait-History $H3
"STATUS (Security, pełny skan): $($H3.Ui.hStatus.Text)"
Check 'full scan without errors' ($H3.Ui.hStatus.Text -notlike '*BŁĘDY*')

# --- Zamknięcie okna w trakcie wczytywania nie blokuje aplikacji ---
$bigDir = Join-Path $OutDir 'logs\big'; New-Item -ItemType Directory -Path $bigDir -Force | Out-Null
$tb = (Get-Date).AddMinutes(-50)
$fmt = 'MM/dd/yyyy HH:mm:ss.fff'; $inv = [Globalization.CultureInfo]::InvariantCulture
$sbig = New-Object System.Text.StringBuilder
for ($i = 0; $i -lt 40000; $i++) {
    [void]$sbig.AppendLine("<Event><Timestamp data_type=`"4`">$($tb.AddMilliseconds($i * 50).ToString($fmt, $inv))</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><User-Name data_type=`"1`">CONTOSO\u$($i % 50)</User-Name><Calling-Station-Id data_type=`"1`">AA-BB-CC-DD-EE-$('{0:X2}' -f ($i % 50))</Calling-Station-Id><Client-IP-Address data_type=`"3`">10.0.0.1</Client-IP-Address><Packet-Type data_type=`"0`">4</Packet-Type><Acct-Status-Type data_type=`"0`">3</Acct-Status-Type></Event>")
}
[IO.File]::WriteAllText((Join-Path $bigDir 'IN2610.log'), $sbig.ToString())
$ui.txtLogPath.Text = $bigDir
$H4 = Show-HistoryWindow -Mode 'User' -Value 'u7' -RangeIndex 1
$H4.Ui.hSrcEv.IsChecked = $false; $H4.Ui.hSrcLog.IsChecked = $true
Start-HistoryLoad $H4
Invoke-Pump 500
$busy = $H4.Busy
$sw = [Diagnostics.Stopwatch]::StartNew(); $H4.Win.Close(); $closeMs = $sw.ElapsedMilliseconds
"CLOSE while loading: busy=$busy, Close() took $closeMs ms"
Check 'close during load returns at once' ($closeMs -lt 1500)
$sw.Restart(); while ($script:HistStopping.Count -and $sw.Elapsed.TotalSeconds -lt 60) { Invoke-Pump 300 }
Check 'stopped loader disposed' ($script:HistStopping.Count -eq 0)

# --- Wydajność w oknie: 20000 zdarzeń z zakładki (czas wątku okna) ---
$ui.txtLogPath.Text = $script:TestLogDir
$script:Ctx.Ev.All = New-Object System.Collections.Generic.List[object]
$script:Ctx.Log.All = New-Object System.Collections.Generic.List[object]
$t0 = [datetime]'2026-09-01 00:00:00'
for ($i = 0; $i -lt 20000; $i++) { $script:Ctx.Log.All.Add((New-Lg $t0.AddMinutes($i * 3) 'Info' 'Acct Interim' 'Acct' 'CONTOSO\jan' 'AA-BB-CC-DD-EE-FF' $(if ($i % 500 -lt 250) { '19' } else { '15' }) $(if ($i % 500 -lt 250) { 'WLC-1' } else { 'SW-01' }) '1' '10.0.5.17')) }
$sw.Restart()
$H5 = Show-HistoryWindow -Mode 'Mac' -Value 'AA-BB-CC-DD-EE-FF' -RangeIndex 0
$showMs = $sw.ElapsedMilliseconds
$H5.Timer.Stop()
while (@($H5.Handles.Values | Where-Object { -not $_.IsCompleted }).Count -and $sw.Elapsed.TotalSeconds -lt 300) { Start-Sleep -Milliseconds 100 }
$bgMs = $sw.ElapsedMilliseconds
$sw.Restart(); Complete-HistoryLoad $H5; $completeMs = $sw.ElapsedMilliseconds
$sw.Restart(); $H5.Ui.hGroup.IsChecked = $false; $toggleMs = $sw.ElapsedMilliseconds
$sw.Restart(); $H5.Ui.hStripHost.Width = 500; Invoke-Pump 100; $H5.Ui.hStripHost.Width = [double]::NaN; Invoke-Pump 100; $resizeMs = $sw.ElapsedMilliseconds
"PERF okno, 20000 zdarzeń: otwarcie $showMs ms, tło do $bgMs ms; wątek okna: scalenie + widok $completeMs ms, przełączenie grupowania $toggleMs ms, 2 zmiany rozmiaru paska $resizeMs ms"
Check 'perf window: all items' ($H5.Items.Count -eq 20000 -and $H5.Rows.Count -ge 20000)
Check 'perf window: changes counted' ($H5.Ui.hCounts.Text -like '*Zmian miejsca w sieci: 79*')

foreach ($hh in @($script:HistWins.Values)) { $hh.Win.Close() }
Invoke-Pump 300
Check 'windows cleaned up' ($script:HistWins.Count -eq 0)
$win.Close()
}
catch {
    # Pełna informacja o wyjątku (także .NET) - przy błędzie w Windows PowerShell 5.1 sam komunikat nie wystarcza
    "EXCEPTION: $($_.Exception.ToString())"
    "SCRIPT STACK: $($_.ScriptStackTrace)"
    $fail++
}
"FAILED: $fail"
if ($fail) { exit 1 }

#Requires -Version 5.1
<#
.SYNOPSIS
    Test dymny NPS Event Viewer (Windows PowerShell 5.1, WPF) - uruchamiany w GitHub Actions.
.DESCRIPTION
    Wczytuje Logi_NPS.ps1 bez pokazywania okna ($global:NpsViewerNoShow), sprawdza logikę osi
    czasu na danych testowych, otwiera okno główne i okno historii (render do PNG w smoke-output),
    wykonuje prawdziwe zapytanie Get-WinEvent z filtrem MAC i wczytuje testowy plik .log.
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

Check 'NormUser domain' ((Get-NormUser 'CONTOSO\Jan.Kowalski') -eq 'jan.kowalski')
Check 'NormUser upn' ((Get-NormUser 'jan@contoso.com') -eq 'jan')
Check 'NormUser host kept' ((Get-NormUser 'host/PC01.contoso.com') -eq 'host/pc01.contoso.com')
Check 'NormComputer host' ((Get-NormComputer 'host/PC01.contoso.local') -eq 'pc01')
Check 'NormComputer dollar' ((Get-NormComputer 'CONTOSO\PC01$') -eq 'pc01')
Check 'Format-Mac' ((Format-Mac 'AABBCCDDEEFF') -eq 'AA-BB-CC-DD-EE-FF')
$dv = Get-HistoryDataValues (New-HistoryQuery 'Mac' 'aa:bb:cc:dd:ee:ff' $false)
Check 'DataValues count' ($dv.Count -eq 10)
Check 'DataValues forms' (($dv -contains 'AA-BB-CC-DD-EE-FF') -and ($dv -contains 'aabb.ccdd.eeff') -and ($dv -contains 'aabbcc-ddeeff') -and ($dv -contains 'AA:BB:CC:DD:EE:FF') -and ($dv -contains 'aabbccddeeff'))
Check 'DataValues none for user' ($null -eq (Get-HistoryDataValues (New-HistoryQuery 'User' 'jan' $false)))
Check 'Partial mac auto' ((New-HistoryQuery 'Mac' 'aa:bb:cc' $false).Partial)

function New-Ev($t, $lvl, $res, $user, $mac, $type, $client, $port, $called, $rc = '0', $reason = '', $machine = '-') {
    [pscustomobject]@{ Time = [datetime]$t; TimeStr = ''; Id = 6272; Result = $res; Level = $lvl; User = $user; Domain = 'CONTOSO'; UserFQ = "CONTOSO\$user"
        Machine = $machine; CallingStation = $mac; CalledStation = $called; NasIp = '10.0.0.1'; NasId = '-'; NasPortType = $type; NasPort = $port
        Client = $client; ClientIp = '10.0.0.1'; CRP = '-'; Policy = 'Firma-802.1X'; AuthProvider = 'Windows'; AuthServer = 'NPS01'; AuthType = 'PEAP'; EapType = 'MSCHAPv2'
        ReasonCode = $rc; Reason = $reason; Server = 'NPS01'; RecordId = 1; MacHex = ($mac -replace '[^0-9A-Fa-f]', '').ToUpper(); UserHex = ($user -replace '[^0-9A-Fa-f]', '').ToUpper(); SearchText = '' }
}
function New-Lg($t, $lvl, $res, $kind, $user, $mac, $type, $client, $port, $ip = '') {
    [pscustomobject]@{ Time = [datetime]$t; TimeStr = ''; Id = '2'; Result = $res; Level = $lvl; Kind = $kind; User = $user; Domain = ''; Sam = ''; UserFQ = ''
        Machine = 'NPS01'; CallingStation = $mac; CalledStation = ''; NasIp = '10.0.0.1'; NasId = ''; NasPortType = $type; NasPort = $port; Client = $client; ClientIp = '10.0.0.1'
        CRP = ''; Policy = 'Firma-802.1X'; AuthProvider = ''; AuthServer = ''; AuthType = 'PEAP'; EapType = ''; ReasonCode = '0'; Reason = ''; Server = 'NPS01'; RecordId = 1
        File = 'IN2610.log'; FilePath = 'C:\x\IN2610.log'; Merged = $false; FramedIp = $ip; SessionTime = ''; MacHex = ($mac -replace '[^0-9A-Fa-f]', '').ToUpper(); UserHex = ''; Attrs = @{}; SearchText = '' }
}
$mac = 'AA-BB-CC-DD-EE-FF'
$evs = @(
    New-Ev '2026-10-06 08:00:00' 'OK' 'Udzielono' 'jan' $mac 'Ethernet' 'SW-01' '12' ''
    New-Ev '2026-10-06 08:30:00' 'OK' 'Udzielono' 'jan' $mac 'Ethernet' 'SW-01' '12' ''
    New-Ev '2026-10-06 09:00:00' 'OK' 'Udzielono' 'jan' $mac 'Ethernet' 'SW-01' '12' ''
    New-Ev '2026-10-06 09:10:00' 'OK' 'Udzielono' 'jan' $mac 'Wireless - IEEE 802.11' 'WLC-1' '1' '00-11-22-33-44-55:CORP'
    New-Ev '2026-10-06 09:12:00' 'Error' 'Odmowa' 'jan' $mac 'Wireless - IEEE 802.11' 'WLC-1' '1' '00-11-22-33-44-55:CORP' '16' 'Złe poświadczenia'
    New-Ev '2026-10-06 12:00:00' 'OK' 'Udzielono' 'jan' $mac 'Ethernet' 'SW-02' '3' ''
    New-Ev '2026-10-07 07:00:00' 'OK' 'Udzielono' 'aabbccddeeff' $mac 'Ethernet' 'SW-02' '3' ''
    New-Ev '2026-10-06 10:00:00' 'OK' 'Udzielono' 'ola' '11-22-33-44-55-66' 'Ethernet' 'SW-09' '1' ''
)
$lgs = @(
    New-Lg '2026-10-06 12:00:01' 'OK' 'Udzielono' 'Resp' 'CONTOSO\jan' $mac '15' 'SW-02' '3'
    New-Lg '2026-10-06 12:05:00' 'Info' 'Acct Start' 'Acct' 'CONTOSO\jan' $mac '15' 'SW-02' '3' '10.0.5.17'
)
$q = New-HistoryQuery 'Mac' 'aabb.ccdd.eeff' $false
$ev = @($evs | Where-Object { Test-HistoryMatch $_ $q }); $lg = @($lgs | Where-Object { Test-HistoryMatch $_ $q })
Check 'match mac ev' ($ev.Count -eq 7)
Check 'match mac log' ($lg.Count -eq 2)
$items = Merge-HistoryItems $ev $lg $q
Check 'merge dedupe' ($items.Count -eq 8)
Check 'merge both flag' (@($items | Where-Object Both).Count -eq 1)
Check 'sorted' ($items[0].Time -eq [datetime]'2026-10-06 08:00:00' -and $items[$items.Count-1].Time.Day -eq 7)
$b = Build-HistoryRows $items $q @{ Group = $true; Markers = $true; Newest = $false }
$b.Rows | ForEach-Object { '{0,-6} {1,-9} {2,-16} {3}  [{4} {5}]  {6}' -f $_.Kind, $_.TimeStr, $_.SubTime, $_.Title, $_.Medium, $_.Location, $_.Detail }
$kinds = ($b.Rows | ForEach-Object Kind) -join ','
"KINDS: $kinds"
Check 'rows sequence' ($kinds -eq 'Day,Event,Change,Event,Event,Gap,Change,Event,Event,Day,Gap,Who,Event')
Check 'group x3' ($b.Rows[1].SubTime -like '×3*')
Check 'changes' ($b.Changes -eq 2)
Check 'acct ip detail' (@($b.Rows | Where-Object { $_.Detail -like '*IP 10.0.5.17*' }).Count -eq 1)
Check 'both source label' (@($b.Rows | Where-Object { $_.Detail -like '*Security + log*' }).Count -eq 1)
Check 'ssid location' (@($b.Rows | Where-Object { $_.Location -like 'SSID CORP*' }).Count -ge 1)
$bn = Build-HistoryRows $items $q @{ Group = $true; Markers = $true; Newest = $true }
$kn = ($bn.Rows | ForEach-Object Kind) -join ','
"NEWEST: $kn"
Check 'newest first day header' ($bn.Rows[0].Kind -eq 'Day' -and $bn.Rows[0].Time.Day -eq 7)
Check 'newest group time = last' ((@($bn.Rows | Where-Object { $_.SubTime -like '×3*' })[0]).TimeStr -eq '09:00:00')
$bg = Build-HistoryRows $items $q @{ Group = $false; Markers = $false; Newest = $false }
Check 'no group no markers' (@($bg.Rows | Where-Object Kind -eq 'Event').Count -eq 8 -and @($bg.Rows | Where-Object { $_.Kind -in 'Gap','Change','Who' }).Count -eq 0)

# User / computer
$qu = New-HistoryQuery 'User' 'CONTOSO\JAN' $false
Check 'match user' (@($evs | Where-Object { Test-HistoryMatch $_ $qu }).Count -eq 6)
$pc = New-Ev '2026-10-06 07:59:00' 'OK' 'Udzielono' 'host/PC01.contoso.local' $mac 'Ethernet' 'SW-01' '12' '' '0' '' 'CONTOSO\PC01$'
$qc = New-HistoryQuery 'Computer' 'pc01' $false
Check 'match computer machine' (Test-HistoryMatch $pc $qc)
Check 'computer of event' ((Get-EventComputer $pc) -eq 'pc01')
Check 'medium' ((Get-EventMedium $evs[3]) -eq 'Wi-Fi' -and (Get-EventMedium $evs[0]) -eq 'LAN' -and (Get-EventMedium $lgs[0]) -eq 'LAN')


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

# --- Okno historii z danych zakładek ---
$H = Show-HistoryWindow -Mode 'Mac' -Value 'aa:bb:cc:dd:ee:ff'
Invoke-Pump 800
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
Set-HistoryZoom $H $null $null; Invoke-Pump 200
Save-Png $H.Win '03-historia-najnowsze'
$text = Get-HistoryText $H
Check 'history text' ($text -like '*Odmowa - Złe poświadczenia (kod 16)*')
Open-HistoryPivot @{ Id = $H.Id; Mode = 'User'; Value = 'CONTOSO\jan' }; Invoke-Pump 600
$H2 = $script:HistWins[$H.Id + 1]
Check 'pivot window user' ($H2 -and $H2.Q.Mode -eq 'User' -and $H2.Items.Count -ge 6)
Save-Png $H2.Win '04-historia-uzytkownik'

# --- Prawdziwe zapytanie do dziennika Security (filtr MAC przez FilterHashtable Data) i plik .log ---
# LogLoader z filtrem MAC na prawdziwym pliku DTS
$dirBase = Join-Path $OutDir 'logs'; $dir = Join-Path $dirBase 'mac'; New-Item -ItemType Directory -Path $dir -Force | Out-Null
$t = (Get-Date).AddMinutes(-30)
$f = { param($dt) $dt.ToString('MM/dd/yyyy HH:mm:ss.fff', [Globalization.CultureInfo]::InvariantCulture) }
$lines = @(
  "<Event><Timestamp data_type=`"4`">$(& $f $t)</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><User-Name data_type=`"1`">CONTOSO\jan</User-Name><Calling-Station-Id data_type=`"1`">AA-BB-CC-DD-EE-FF</Calling-Station-Id><NAS-Port-Type data_type=`"0`">15</NAS-Port-Type><Client-IP-Address data_type=`"3`">10.0.0.1</Client-IP-Address><Client-Friendly-Name data_type=`"1`">SW-01</Client-Friendly-Name><Packet-Type data_type=`"0`">1</Packet-Type></Event>"
  "<Event><Timestamp data_type=`"4`">$(& $f $t.AddSeconds(1))</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><Client-IP-Address data_type=`"3`">10.0.0.1</Client-IP-Address><Client-Friendly-Name data_type=`"1`">SW-01</Client-Friendly-Name><Packet-Type data_type=`"0`">2</Packet-Type><Reason-Code data_type=`"0`">0</Reason-Code></Event>"
  "<Event><Timestamp data_type=`"4`">$(& $f $t.AddSeconds(5))</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><User-Name data_type=`"1`">ola</User-Name><Calling-Station-Id data_type=`"1`">11-22-33-44-55-66</Calling-Station-Id><Client-IP-Address data_type=`"3`">10.0.0.9</Client-IP-Address><Packet-Type data_type=`"0`">1</Packet-Type></Event>"
  "<Event><Timestamp data_type=`"4`">$(& $f $t.AddSeconds(6))</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><Client-IP-Address data_type=`"3`">10.0.0.9</Client-IP-Address><Packet-Type data_type=`"0`">3</Packet-Type><Reason-Code data_type=`"0`">16</Reason-Code></Event>"
  "<Event><Timestamp data_type=`"4`">$(& $f $t.AddMinutes(2))</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><User-Name data_type=`"1`">CONTOSO\jan</User-Name><Calling-Station-Id data_type=`"1`">aa-bb-cc-dd-ee-ff</Calling-Station-Id><Client-IP-Address data_type=`"3`">10.0.0.1</Client-IP-Address><Packet-Type data_type=`"0`">4</Packet-Type><Acct-Status-Type data_type=`"0`">1</Acct-Status-Type><Framed-IP-Address data_type=`"3`">10.0.5.17</Framed-IP-Address></Event>"
)
Set-Content -Path (Join-Path $dir 'IN2610.log') -Value $lines -Encoding UTF8
$out = @(& $script:LogLoader @($dir) 'IN*.log' ((Get-Date).AddHours(-1)) (Get-Date) 0 'Mac' 'AABBCCDDEEFF' $false)
$entries = @($out | Where-Object { -not $_.PSObject.Properties['IsSummary'] })
"LOG entries: " + (($entries | ForEach-Object { "$($_.Result)/$($_.User)/$($_.FramedIp)" }) -join ' ; ')
Check 'loader mac filter keeps req+accept+acct' ($entries.Count -eq 3)
Check 'loader merged accept has mac' (@($entries | Where-Object { $_.Result -eq 'Udzielono' -and $_.MacHex -eq 'AABBCCDDEEFF' }).Count -eq 1)
Check 'loader framed ip' (@($entries | Where-Object { $_.FramedIp -eq '10.0.5.17' }).Count -eq 1)
$outU = @(& $script:LogLoader @($dir) 'IN*.log' ((Get-Date).AddHours(-1)) (Get-Date) 0 'User' 'ola' $false)
Check 'loader user filter' (@($outU | Where-Object { -not $_.PSObject.Properties['IsSummary'] }).Count -eq 2)
$outAll = @(& $script:LogLoader @($dir) 'IN*.log' ((Get-Date).AddHours(-1)) (Get-Date) 0)
Check 'loader no filter = all' (@($outAll | Where-Object { -not $_.PSObject.Properties['IsSummary'] }).Count -eq 5)

# Parowanie odpowiedzi z żądaniem po Class (dwóch klientów tego samego WLC naraz)
$dir2 = Join-Path $dirBase 'class'; New-Item -ItemType Directory -Path $dir2 -Force | Out-Null
$t2 = (Get-Date).AddMinutes(-20)
$ev2 = { param($sec, $body) "<Event><Timestamp data_type=`"4`">$(& $f $t2.AddSeconds($sec))</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><Client-IP-Address data_type=`"3`">10.0.0.50</Client-IP-Address><Client-Friendly-Name data_type=`"1`">WLC-1</Client-Friendly-Name>$body</Event>" }
$lines2 = @(
  (& $ev2 0 '<User-Name data_type="1">CONTOSO\jan</User-Name><Calling-Station-Id data_type="1">AA-BB-CC-DD-EE-FF</Calling-Station-Id><NAS-Port-Type data_type="0">19</NAS-Port-Type><Class data_type="1">311 1 10.0.0.2 10/07/2026 08:00:00 1</Class><Packet-Type data_type="0">1</Packet-Type>')
  (& $ev2 0 '<User-Name data_type="1">CONTOSO\ola</User-Name><Calling-Station-Id data_type="1">11-22-33-44-55-66</Calling-Station-Id><NAS-Port-Type data_type="0">19</NAS-Port-Type><Class data_type="1">311 1 10.0.0.2 10/07/2026 08:00:00 2</Class><Packet-Type data_type="0">1</Packet-Type>')
  (& $ev2 1 '<Class data_type="1">311 1 10.0.0.2 10/07/2026 08:00:00 1</Class><Packet-Type data_type="0">3</Packet-Type><Reason-Code data_type="0">16</Reason-Code>')
  (& $ev2 1 '<Class data_type="1">311 1 10.0.0.2 10/07/2026 08:00:00 2</Class><Packet-Type data_type="0">2</Packet-Type><Reason-Code data_type="0">0</Reason-Code>')
  (& $ev2 9 '<User-Name data_type="1">CONTOSO\jan</User-Name><Calling-Station-Id data_type="1">AA-BB-CC-DD-EE-FF</Calling-Station-Id><Packet-Type data_type="0">4</Packet-Type><Acct-Status-Type data_type="0">2</Acct-Status-Type><Acct-Session-Time data_type="0">3725</Acct-Session-Time><Acct-Terminate-Cause data_type="0">2</Acct-Terminate-Cause><NAS-Port-Id data_type="1">GigabitEthernet1/0/12</NAS-Port-Id>')
)
Set-Content -Path (Join-Path $dir2 'IN2610.log') -Value $lines2 -Encoding UTF8
$c2 = @(& $script:LogLoader @($dir2) 'IN*.log' ((Get-Date).AddHours(-1)) (Get-Date) 0 'Mac' 'AABBCCDDEEFF' $false | Where-Object { -not $_.PSObject.Properties['IsSummary'] })
"CLASS entries: " + (($c2 | ForEach-Object { "$($_.Result)/$($_.User)/$($_.ReasonCode)/$($_.AcctTerminate)/$($_.NasPortId)" }) -join ' ; ')
Check 'class pairing: reject belongs to jan' (@($c2 | Where-Object { $_.Result -eq 'Odmowa' }).Count -eq 1 -and @($c2 | Where-Object { $_.Result -eq 'Udzielono' }).Count -eq 0)
Check 'acct terminate cause' (@($c2 | Where-Object { $_.AcctTerminate -like 'Lost-Carrier*' }).Count -eq 1)
Check 'nas-port-id' (@($c2 | Where-Object { $_.NasPortId -eq 'GigabitEthernet1/0/12' }).Count -eq 1)
$it2 = New-HistoryItem ($c2 | Where-Object { $_.AcctTerminate } | Select-Object -First 1) 'Log' $q
Check 'history detail: session + cause' (($it2.Parts -join ' ') -like '*czas sesji 1 godz. 2 min*koniec sesji: Lost-Carrier*')
Check 'vpn medium from ip' ((Get-EventMedium ([pscustomobject]@{ NasPortType = '-'; CalledStation = '198.51.100.10'; CallingStation = '203.0.113.45' })) -eq 'VPN')
Check 'german wifi' ((Get-EventMedium ([pscustomobject]@{ NasPortType = 'Drahtlos - IEEE 802.11'; CalledStation = ''; CallingStation = '' })) -eq 'Wi-Fi')
Check 'user abc.def is not a mac' (-not (Test-HistoryMatch ([pscustomobject]@{ User = 'abc.def'; MacHex = ''; UserFQ = ''; Sam = '' }) (New-HistoryQuery 'Mac' 'abcd' $false)))
$ui.txtLogPath.Text = $dir
$H3 = Show-HistoryWindow -Mode 'Mac' -Value 'AA-BB-CC-DD-EE-FF' -RangeIndex 2
$H3.Ui.hSrcLog.IsChecked = $true
Start-HistoryLoad $H3
$sw = [Diagnostics.Stopwatch]::StartNew()
while ($H3.Busy -and $sw.Elapsed.TotalSeconds -lt 120) { Invoke-Pump 300 }
"STATUS (Security + .log): $($H3.Ui.hStatus.Text)"
Check 'load finished' (-not $H3.Busy)
Check 'load without errors' ($H3.Ui.hStatus.Text -notlike '*BŁĘDY*')
Check 'load log entries' (@($H3.Items | Where-Object Src -eq 'Log').Count -eq 3)
Save-Png $H3.Win '05-historia-wczytana'

$H3.Ui.hPartial.IsChecked = $true
$H3.Ui.hSrcLog.IsChecked = $false
Start-HistoryLoad $H3
$sw = [Diagnostics.Stopwatch]::StartNew()
while ($H3.Busy -and $sw.Elapsed.TotalSeconds -lt 120) { Invoke-Pump 300 }
"STATUS (Security, pełny skan): $($H3.Ui.hStatus.Text)"
Check 'full scan without errors' ($H3.Ui.hStatus.Text -notlike '*BŁĘDY*')

foreach ($hh in @($script:HistWins.Values)) { $hh.Win.Close() }
Invoke-Pump 300
Check 'windows cleaned up' ($script:HistWins.Count -eq 0)
$win.Close()
"FAILED: $fail"
if ($fail) { exit 1 }

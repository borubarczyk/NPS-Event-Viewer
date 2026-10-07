#Requires -Version 5.1
<#
.SYNOPSIS
    Testy logiki osi czasu (historia urządzenia / użytkownika / komputera) bez okien.
.DESCRIPTION
    Wywoływane przez tests\Smoke-NpsViewer.ps1 po wczytaniu Logi_NPS.ps1 (funkcje, $script:LogLoader,
    $script:HistWorker). Wymaga funkcji Check i katalogu $LogDir na testowe pliki .log.
#>
param([Parameter(Mandatory)][string]$LogDir)

New-Item -ItemType Directory -Path $LogDir -Force | Out-Null

# --- Normalizacja i zapytanie ---
Check 'NormUser domain' ((Get-NormUser 'CONTOSO\Jan.Kowalski') -eq 'jan.kowalski')
Check 'NormUser upn' ((Get-NormUser 'jan@contoso.com') -eq 'jan')
Check 'NormUser host kept' ((Get-NormUser 'host/PC01.contoso.com') -eq 'host/pc01.contoso.com')
Check 'NormComputer host' ((Get-NormComputer 'host/PC01.contoso.local') -eq 'pc01')
Check 'NormComputer dollar' ((Get-NormComputer 'CONTOSO\PC01$') -eq 'pc01')
Check 'Format-Mac' ((Format-Mac 'AABBCCDDEEFF') -eq 'AA-BB-CC-DD-EE-FF')
$dv = Get-HistoryDataValues (New-HistoryQuery 'Mac' 'aa:bb:cc:dd:ee:ff' $false)
Check 'DataValues count' ($dv.Count -eq 14)
Check 'DataValues forms' (($dv -ccontains 'AA-BB-CC-DD-EE-FF') -and ($dv -ccontains 'aabb.ccdd.eeff') -and ($dv -ccontains 'aabbcc-ddeeff') -and ($dv -ccontains 'AA:BB:CC:DD:EE:FF') -and ($dv -ccontains 'aabbccddeeff'))
Check 'DataValues Huawei / dotted pairs' (($dv -ccontains 'aabb-ccdd-eeff') -and ($dv -ccontains 'AABB-CCDD-EEFF') -and ($dv -ccontains 'aa.bb.cc.dd.ee.ff'))
Check 'DataValues none for user' ($null -eq (Get-HistoryDataValues (New-HistoryQuery 'User' 'jan' $false)))
Check 'Partial mac auto' ((New-HistoryQuery 'Mac' 'aa:bb:cc' $false).Partial)

function New-Ev($t, $lvl, $res, $user, $mac, $type, $client, $port, $called, $rc = '0', $reason = '', $machine = '-') {
    $mh = $(if ($mac -match '^\d{1,3}(\.\d{1,3}){3}$') { '' } else { ($mac -replace '[^0-9A-Fa-f]', '').ToUpper() })
    [pscustomobject]@{ Time = [datetime]$t; TimeStr = ''; Id = 6272; Result = $res; Level = $lvl; User = $user; Domain = 'CONTOSO'; UserFQ = "CONTOSO\$user"
        Machine = $machine; CallingStation = $mac; CalledStation = $called; NasIp = '10.0.0.1'; NasId = '-'; NasPortType = $type; NasPort = $port
        Client = $client; ClientIp = '10.0.0.1'; CRP = '-'; Policy = 'Firma-802.1X'; AuthProvider = 'Windows'; AuthServer = 'NPS01'; AuthType = 'PEAP'; EapType = 'MSCHAPv2'
        ReasonCode = $rc; Reason = $reason; Server = 'NPS01'; RecordId = 1; MacHex = $mh; UserHex = ($user -replace '[^0-9A-Fa-f]', '').ToUpper(); SearchText = '' }
}
function New-Lg($t, $lvl, $res, $kind, $user, $mac, $type, $client, $port, $ip = '', $portId = '') {
    [pscustomobject]@{ Time = [datetime]$t; TimeStr = ''; Id = '2'; Result = $res; Level = $lvl; Kind = $kind; User = $user; Domain = ''; Sam = ''; UserFQ = ''
        Machine = 'NPS01'; CallingStation = $mac; CalledStation = ''; NasIp = '10.0.0.1'; NasId = ''; NasPortType = $type; NasPort = $port; Client = $client; ClientIp = '10.0.0.1'
        CRP = ''; Policy = 'Firma-802.1X'; AuthProvider = ''; AuthServer = ''; AuthType = 'PEAP'; EapType = ''; ReasonCode = '0'; Reason = ''; Server = 'NPS01'; RecordId = 1
        File = 'IN2610.log'; FilePath = 'C:\x\IN2610.log'; Merged = $false; FramedIp = $ip; SessionTime = ''; NasPortId = $portId; AcctTerminate = ''
        MacHex = ($mac -replace '[^0-9A-Fa-f]', '').ToUpper(); UserHex = ''; Attrs = @{}; SearchText = '' }
}
# Jak w oknie: dopasowanie, elementy, scalenie (w aplikacji dwa pierwsze kroki robi runspace w tle)
function Get-TestItems($Ev, $Lg, $Q) {
    $ei = New-Object System.Collections.Generic.List[object]; $li = New-Object System.Collections.Generic.List[object]
    foreach ($e in $Ev) { if (Test-HistoryMatch $e $Q) { $ei.Add((New-HistoryItem $e 'Ev' $Q)) } }
    foreach ($e in $Lg) { if (Test-HistoryMatch $e $Q) { $li.Add((New-HistoryItem $e 'Log' $Q)) } }
    return , (Merge-HistoryItems $ei $li)
}

$mac = 'AA-BB-CC-DD-EE-FF'
$script:TestEvs = @(
    New-Ev '2026-10-06 08:00:00' 'OK' 'Udzielono' 'jan' $mac 'Ethernet' 'SW-01' '12' ''
    New-Ev '2026-10-06 08:30:00' 'OK' 'Udzielono' 'jan' $mac 'Ethernet' 'SW-01' '12' ''
    New-Ev '2026-10-06 09:00:00' 'OK' 'Udzielono' 'jan' $mac 'Ethernet' 'SW-01' '12' ''
    New-Ev '2026-10-06 09:10:00' 'OK' 'Udzielono' 'jan' $mac 'Wireless - IEEE 802.11' 'WLC-1' '1' '00-11-22-33-44-55:CORP'
    New-Ev '2026-10-06 09:12:00' 'Error' 'Odmowa' 'jan' $mac 'Wireless - IEEE 802.11' 'WLC-1' '1' '00-11-22-33-44-55:CORP' '16' 'Złe poświadczenia'
    New-Ev '2026-10-06 12:00:00' 'OK' 'Udzielono' 'jan' $mac 'Ethernet' 'SW-02' '3' ''
    New-Ev '2026-10-07 07:00:00' 'OK' 'Udzielono' 'aabbccddeeff' $mac 'Ethernet' 'SW-02' '3' ''
    New-Ev '2026-10-06 10:00:00' 'OK' 'Udzielono' 'ola' '11-22-33-44-55-66' 'Ethernet' 'SW-09' '1' ''
)
$script:TestLgs = @(
    New-Lg '2026-10-06 12:00:01' 'OK' 'Udzielono' 'Resp' 'CONTOSO\jan' $mac '15' 'SW-02' '3'
    New-Lg '2026-10-06 12:05:00' 'Info' 'Acct Start' 'Acct' 'CONTOSO\jan' $mac '15' 'SW-02' '3' '10.0.5.17'
)
$evs = $script:TestEvs; $lgs = $script:TestLgs
$q = New-HistoryQuery 'Mac' 'aabb.ccdd.eeff' $false
Check 'match mac ev' (@($evs | Where-Object { Test-HistoryMatch $_ $q }).Count -eq 7)
Check 'match mac log' (@($lgs | Where-Object { Test-HistoryMatch $_ $q }).Count -eq 2)
$items = Get-TestItems $evs $lgs $q
Check 'merge dedupe' ($items.Count -eq 8)
Check 'merge both flag' (@($items | Where-Object Both).Count -eq 1)
Check 'sorted' ($items[0].Time -eq [datetime]'2026-10-06 08:00:00' -and $items[$items.Count - 1].Time.Day -eq 7)
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
"NEWEST: " + (($bn.Rows | ForEach-Object Kind) -join ',')
Check 'newest first day header' ($bn.Rows[0].Kind -eq 'Day' -and $bn.Rows[0].Time.Day -eq 7)
Check 'newest group time = last' ((@($bn.Rows | Where-Object { $_.SubTime -like '×3*' })[0]).TimeStr -eq '09:00:00')
$bg = Build-HistoryRows $items $q @{ Group = $false; Markers = $false; Newest = $false }
Check 'no group no markers' (@($bg.Rows | Where-Object Kind -eq 'Event').Count -eq 8 -and @($bg.Rows | Where-Object { $_.Kind -in 'Gap', 'Change', 'Who' }).Count -eq 0)

# --- Użytkownik / komputer ---
$qu = New-HistoryQuery 'User' 'CONTOSO\JAN' $false
Check 'match user' (@($evs | Where-Object { Test-HistoryMatch $_ $qu }).Count -eq 6)
Check 'match user partial' (@($evs | Where-Object { Test-HistoryMatch $_ (New-HistoryQuery 'User' 'ja' $true) }).Count -eq 6)
$pc = New-Ev '2026-10-06 07:59:00' 'OK' 'Udzielono' 'host/PC01.contoso.local' $mac 'Ethernet' 'SW-01' '12' '' '0' '' 'CONTOSO\PC01$'
$qc = New-HistoryQuery 'Computer' 'pc01' $false
Check 'match computer machine' (Test-HistoryMatch $pc $qc)
Check 'match computer host/' (Test-HistoryMatch (New-Ev '2026-10-06 07:59:00' 'OK' 'Udzielono' 'host/PC01.contoso.local' $mac 'Ethernet' 'SW-01' '12' '') $qc)
Check 'no computer match on .log server name' (-not (Test-HistoryMatch (New-Lg '2026-10-06 07:59:00' 'OK' 'Udzielono' 'Resp' 'jan' $mac '15' 'SW-01' '1') (New-HistoryQuery 'Computer' 'NPS01' $false)))
Check 'computer of event' ((Get-EventComputer $pc) -eq 'pc01')
Check 'medium' ((Get-EventMedium $evs[3]) -eq 'Wi-Fi' -and (Get-EventMedium $evs[0]) -eq 'LAN' -and (Get-EventMedium $lgs[0]) -eq 'LAN')
Check 'vpn medium from ip' ((Get-EventMedium ([pscustomobject]@{ NasPortType = '-'; CalledStation = '198.51.100.10'; CallingStation = '203.0.113.45' })) -eq 'VPN')
Check 'german wifi' ((Get-EventMedium ([pscustomobject]@{ NasPortType = 'Drahtlos - IEEE 802.11'; CalledStation = ''; CallingStation = '' })) -eq 'Wi-Fi')
Check 'user abc.def is not a mac' (-not (Test-HistoryMatch ([pscustomobject]@{ User = 'abc.def'; MacHex = ''; UserFQ = ''; Sam = '' }) (New-HistoryQuery 'Mac' 'abcd' $false)))

# --- Miejsce w sieci: bez fałszywych zmian ---
# Security ma NAS-Port (numer), .log także NAS-Port-Id (nazwa portu) - to ten sam port
$e1 = New-Ev '2026-10-06 08:00:00' 'OK' 'Udzielono' 'jan' $mac 'Ethernet' 'SW-01' '50112' ''
$l1 = New-Lg '2026-10-06 08:00:05' 'Info' 'Acct Start' 'Acct' 'CONTOSO\jan' $mac '15' 'SW-01' '50112' '10.0.5.17' 'GigabitEthernet1/0/12'
$l2 = New-Lg '2026-10-06 08:30:05' 'Info' 'Acct Interim' 'Acct' 'CONTOSO\jan' $mac '' 'SW-01' '' '10.0.5.17' ''
$e2 = New-Ev '2026-10-06 09:00:00' 'OK' 'Udzielono' 'jan' $mac 'Ethernet' 'SW-01' '50112' ''
$bp = Build-HistoryRows (Get-TestItems @($e1, $e2) @($l1, $l2) $q) $q @{ Group = $true; Markers = $true; Newest = $false }
Check 'NAS-Port vs NAS-Port-Id: no false change' ($bp.Changes -eq 0)
Check 'NAS-Port-Id shown in location' (@($bp.Rows | Where-Object { $_.Location -eq 'SW-01 · port GigabitEthernet1/0/12' }).Count -eq 1)
$e3 = New-Ev '2026-10-06 09:30:00' 'OK' 'Udzielono' 'jan' $mac 'Ethernet' 'SW-01' '50113' ''
Check 'other port is a change' ((Build-HistoryRows (Get-TestItems @($e1, $e2, $e3) @($l1, $l2) $q) $q @{ Group = $true; Markers = $true; Newest = $false }).Changes -eq 1)
# VPN: NAS-Port to numer tunelu - inny przy każdym połączeniu
$vp = @(12, 37, 5, 129 | ForEach-Object -Begin { $i = 0 } -Process { New-Ev ([datetime]'2026-10-06 08:00:00').AddHours($i++) 'OK' 'Udzielono' 'jan' "203.0.113.$_" 'Virtual (VPN)' 'VPN-GW' "$_" '' })
$bv = Build-HistoryRows (Get-TestItems $vp @() $qu) $qu @{ Group = $true; Markers = $true; Newest = $false }
Check 'vpn reconnects: no change' ($bv.Changes -eq 0)
Check 'vpn: no device markers from IP' (@($bv.Rows | Where-Object Kind -eq 'Who').Count -eq 0)
Check 'vpn grouped' (@($bv.Rows | Where-Object Kind -eq 'Event').Count -eq 1)

# --- Adres IP w Calling-Station-Id (VPN) to nie MAC ---
$vip = New-Ev '2026-10-06 08:00:00' 'OK' 'Udzielono' 'jan' '192.168.100.200' 'Virtual (VPN)' 'VPN-GW' '7' ''
Check 'ip is not mac (event)' ($vip.MacHex -eq '' -and -not (Test-HistoryMatch $vip (New-HistoryQuery 'Mac' (Format-Mac '192168100200') $false)))
$vit = New-HistoryItem $vip 'Ev' $qu
Check 'ip is not mac (item)' ($vit.MacHex -eq '' -and ($vit.Parts -join ' ') -like '*z adresu 192.168.100.200*')

# --- Pliki .log: filtr MAC / użytkownika już przy czytaniu ---
$f = { param($dt) $dt.ToString('MM/dd/yyyy HH:mm:ss.fff', [Globalization.CultureInfo]::InvariantCulture) }
$dir = Join-Path $LogDir 'mac'; New-Item -ItemType Directory -Path $dir -Force | Out-Null
$t = (Get-Date).AddMinutes(-30)
$lines = @(
    "<Event><Timestamp data_type=`"4`">$(& $f $t)</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><User-Name data_type=`"1`">CONTOSO\jan</User-Name><Calling-Station-Id data_type=`"1`">AA-BB-CC-DD-EE-FF</Calling-Station-Id><NAS-Port-Type data_type=`"0`">15</NAS-Port-Type><Client-IP-Address data_type=`"3`">10.0.0.1</Client-IP-Address><Client-Friendly-Name data_type=`"1`">SW-01</Client-Friendly-Name><Packet-Type data_type=`"0`">1</Packet-Type></Event>"
    "<Event><Timestamp data_type=`"4`">$(& $f $t.AddSeconds(1))</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><Client-IP-Address data_type=`"3`">10.0.0.1</Client-IP-Address><Client-Friendly-Name data_type=`"1`">SW-01</Client-Friendly-Name><Packet-Type data_type=`"0`">2</Packet-Type><Reason-Code data_type=`"0`">0</Reason-Code></Event>"
    "<Event><Timestamp data_type=`"4`">$(& $f $t.AddSeconds(5))</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><User-Name data_type=`"1`">ola</User-Name><Calling-Station-Id data_type=`"1`">11-22-33-44-55-66</Calling-Station-Id><Client-IP-Address data_type=`"3`">10.0.0.9</Client-IP-Address><Packet-Type data_type=`"0`">1</Packet-Type></Event>"
    "<Event><Timestamp data_type=`"4`">$(& $f $t.AddSeconds(6))</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><Client-IP-Address data_type=`"3`">10.0.0.9</Client-IP-Address><Packet-Type data_type=`"0`">3</Packet-Type><Reason-Code data_type=`"0`">16</Reason-Code></Event>"
    "<Event><Timestamp data_type=`"4`">$(& $f $t.AddMinutes(2))</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><User-Name data_type=`"1`">CONTOSO\jan</User-Name><Calling-Station-Id data_type=`"1`">aa-bb-cc-dd-ee-ff</Calling-Station-Id><Client-IP-Address data_type=`"3`">10.0.0.1</Client-IP-Address><Packet-Type data_type=`"0`">4</Packet-Type><Acct-Status-Type data_type=`"0`">1</Acct-Status-Type><Framed-IP-Address data_type=`"3`">10.0.5.17</Framed-IP-Address></Event>"
    "<Event><Timestamp data_type=`"4`">$(& $f $t.AddMinutes(3))</Timestamp><Computer-Name data_type=`"1`">NPS01</Computer-Name><User-Name data_type=`"1`">CONTOSO\vpnuser</User-Name><Calling-Station-Id data_type=`"1`">192.168.100.200</Calling-Station-Id><NAS-Port-Type data_type=`"0`">5</NAS-Port-Type><Client-IP-Address data_type=`"3`">10.0.0.7</Client-IP-Address><Packet-Type data_type=`"0`">1</Packet-Type></Event>"
)
Set-Content -Path (Join-Path $dir 'IN2610.log') -Value $lines -Encoding UTF8
$script:TestLogDir = $dir
function Get-LogEntries { param([object[]]$A) @(& $script:LogLoader @A | Where-Object { -not $_.PSObject.Properties['IsSummary'] }) }
$from = (Get-Date).AddHours(-1); $to = (Get-Date)
$entries = Get-LogEntries @(@($dir), 'IN*.log', $from, $to, 0, 'Mac', 'AABBCCDDEEFF', $false)
"LOG entries: " + (($entries | ForEach-Object { "$($_.Result)/$($_.User)/$($_.FramedIp)" }) -join ' ; ')
Check 'loader mac filter keeps req+accept+acct' ($entries.Count -eq 3)
Check 'loader merged accept has mac' (@($entries | Where-Object { $_.Result -eq 'Udzielono' -and $_.MacHex -eq 'AABBCCDDEEFF' }).Count -eq 1)
Check 'loader framed ip' (@($entries | Where-Object { $_.FramedIp -eq '10.0.5.17' }).Count -eq 1)
Check 'loader user filter' ((Get-LogEntries @(@($dir), 'IN*.log', $from, $to, 0, 'User', 'ola', $false)).Count -eq 2)
$all = Get-LogEntries @(@($dir), 'IN*.log', $from, $to, 0)
Check 'loader no filter = all' ($all.Count -eq 6)
Check 'loader: vpn ip is not a mac' (@($all | Where-Object { $_.CallingStation -eq '192.168.100.200' -and $_.MacHex -eq '' }).Count -eq 1)
Check 'loader: mac filter ignores ip' ((Get-LogEntries @(@($dir), 'IN*.log', $from, $to, 0, 'Mac', '192168100200', $false)).Count -eq 0)
$noReq = Get-LogEntries @(@($dir), 'IN*.log', $from, $to, 0, 'Mac', 'AABBCCDDEEFF', $false, $true)
Check 'loader SkipReq drops requests only' ($noReq.Count -eq 2 -and @($noReq | Where-Object Kind -eq 'Req').Count -eq 0)
$sum = @(& $script:LogLoader @($dir) 'IN*.log' $from $to 1 | Where-Object { $_.PSObject.Properties['IsSummary'] })[0]
Check 'loader reports dropped (limit)' ($sum.Dropped -eq 5)

# --- Parowanie odpowiedzi z żądaniem po Class (dwóch klientów tego samego WLC naraz) ---
$dir2 = Join-Path $LogDir 'class'; New-Item -ItemType Directory -Path $dir2 -Force | Out-Null
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
$c2 = Get-LogEntries @(@($dir2), 'IN*.log', $from, $to, 0, 'Mac', 'AABBCCDDEEFF', $false)
"CLASS entries: " + (($c2 | ForEach-Object { "$($_.Result)/$($_.User)/$($_.ReasonCode)/$($_.AcctTerminate)/$($_.NasPortId)" }) -join ' ; ')
Check 'class pairing: reject belongs to jan' (@($c2 | Where-Object { $_.Result -eq 'Odmowa' }).Count -eq 1 -and @($c2 | Where-Object { $_.Result -eq 'Udzielono' }).Count -eq 0)
Check 'acct terminate cause' (@($c2 | Where-Object { $_.AcctTerminate -like 'Lost-Carrier*' }).Count -eq 1)
Check 'nas-port-id' (@($c2 | Where-Object { $_.NasPortId -eq 'GigabitEthernet1/0/12' }).Count -eq 1)
$it2 = New-HistoryItem ($c2 | Where-Object { $_.AcctTerminate } | Select-Object -First 1) 'Log' $q
Check 'history detail: session + cause' (($it2.Parts -join ' ') -like '*czas sesji 1 godz. 2 min*koniec sesji: Lost-Carrier*')

# --- Runspace w tle (jak w oknie): funkcje przekazane tekstem, loader i dane z zakładek ---
function Invoke-TestWorker([string]$Loader, $LoaderArgs, $Objects, $Q, [string]$Src) {
    $ps = [powershell]::Create()
    try {
        [void]$ps.AddScript($script:HistWorker).AddArgument((Get-HistoryWorkerFunctions)).AddArgument($Loader).AddArgument($LoaderArgs).AddArgument($Objects).AddArgument($Q).AddArgument($Src)
        $r = @($ps.Invoke())
        if ($ps.Streams.Error.Count) { "WORKER ERROR: $($ps.Streams.Error[0])" }
        return $r[$r.Count - 1]
    }
    finally { $ps.Dispose() }
}
$wt = Invoke-TestWorker '' $null @($evs) $q 'Ev'
Check 'worker (tab data): items built in runspace' ($wt.HistWorker -and $wt.Raw -eq $evs.Count -and $wt.Items.Count -eq 7 -and $wt.Items[0].Src -eq 'Ev')
$wl = Invoke-TestWorker ([string]$script:LogLoader) @(@($dir), 'IN*.log', $from, $to, 0, 'Mac', 'AABBCCDDEEFF', $false, $true) $null $q 'Log'
Check 'worker (.log loader): items + summary' ($wl.Items.Count -eq 2 -and $wl.Summary.IsSummary -and $wl.Items[0].Src -eq 'Log')

# --- Wydajność (dziesiątki tysięcy zdarzeń; wynik informacyjnie) ---
$big = New-Object System.Collections.Generic.List[object]
$t0 = [datetime]'2026-09-01 00:00:00'
for ($i = 0; $i -lt 20000; $i++) {
    $big.Add((New-Lg $t0.AddMinutes($i * 3) 'Info' 'Acct Interim' 'Acct' 'CONTOSO\jan' $mac '19' 'WLC-1' '1' '10.0.5.17'))
}
$sw = [Diagnostics.Stopwatch]::StartNew()
$wb = Invoke-TestWorker '' $null $big.ToArray() $q 'Log'
$tWorker = $sw.Elapsed.TotalSeconds
$sw.Restart(); $mb = Merge-HistoryItems @() $wb.Items; $tMerge = $sw.Elapsed.TotalSeconds
$sw.Restart(); $rb = Build-HistoryRows $mb $q @{ Group = $true; Markers = $true; Newest = $false }; $tRows = $sw.Elapsed.TotalSeconds
$sw.Restart(); $rb2 = Build-HistoryRows $mb $q @{ Group = $false; Markers = $true; Newest = $true }; $tRows2 = $sw.Elapsed.TotalSeconds
'PERF 20000 zdarzeń: runspace (dopasowanie + elementy) {0:N2} s | w oknie: scalenie {1:N2} s, wiersze z grupowaniem {2:N2} s, bez grupowania {3:N2} s' -f $tWorker, $tMerge, $tRows, $tRows2
Check 'perf: all items' ($mb.Count -eq 20000 -and $rb2.Rows.Count -ge 20000)

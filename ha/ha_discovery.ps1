<#
  ha_discovery.ps1  -  READ-ONLY look at your Home Assistant setup.

  It changes NOTHING in Home Assistant. It connects, reads, and writes one file:
      <your Downloads folder>\ha_discovery_output.json

  What it collects (and nothing else):
    - Your dashboards and the titles of their pages (views). The full layout is
      included ONLY for pages whose name contains "maint" or "batter".
    - Which custom cards are installed (e.g. auto-entities), and which
      integrations you use (just names + counts).
    - Battery entities.
    - Entities that look maintenance-related (filter, rinse aid, salt, toner,
      ink, descale, etc.) and "problem" alert sensors.
    - Every entity on appliance-type devices (dishwasher, fridge, printer,
      washer, dryer, vacuum, purifier, thermostat, ...), including disabled ones.

  Works in Windows PowerShell 5.1 and PowerShell 7.
#>
param(
    [string]$HaUrl,
    [string]$Token,
    [string]$OutFile = (Join-Path $env:USERPROFILE 'Downloads\ha_discovery_output.json')
)

$ErrorActionPreference = 'Stop'

# ---------- inputs ----------
if (-not $HaUrl) {
    $HaUrl = Read-Host 'Home Assistant address [press Enter for http://homeassistant.local:8123]'
    if ([string]::IsNullOrWhiteSpace($HaUrl)) { $HaUrl = 'http://homeassistant.local:8123' }
}
$HaUrl = $HaUrl.Trim().TrimEnd('/')
if ($HaUrl -notmatch '^https?://') { $HaUrl = 'http://' + $HaUrl }

$token = $Token
if (-not $token) {
    $secure = Read-Host 'Paste your Home Assistant long-lived access token (it will not be shown)' -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { $token = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}
$token = $token.Trim()
if (-not $token) { Write-Host 'No token entered. Stopping.' -ForegroundColor Red; exit 1 }

$wsUrl = ($HaUrl -replace '^http', 'ws') + '/api/websocket'   # http->ws, https->wss

# ---------- JSON helpers (5.1's ConvertFrom-Json chokes on big payloads) ----------
$isCore = $PSVersionTable.PSVersion.Major -ge 6
if (-not $isCore) {
    Add-Type -AssemblyName System.Web.Extensions
    $script:ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $script:ser.MaxJsonLength = [int]::MaxValue
    $script:ser.RecursionLimit = 1000
}
function ConvertFrom-JsonText([string]$text) {
    if ($isCore) { return ($text | ConvertFrom-Json -AsHashtable -Depth 200) }
    return $script:ser.DeserializeObject($text)
}
function Get-Val($dict, [string]$key) {
    if ($null -eq $dict -or -not ($dict -is [System.Collections.IDictionary])) { return $null }
    # Dictionary (5.1) has ContainsKey; Hashtable/OrderedHashtable (7.x) have Contains
    if ($dict.PSObject.Methods['ContainsKey']) { $has = $dict.ContainsKey($key) } else { $has = $dict.Contains($key) }
    if ($has) { return $dict[$key] }
    return $null
}

# ---------- websocket plumbing ----------
$ws = New-Object System.Net.WebSockets.ClientWebSocket
$ct = [Threading.CancellationToken]::None

function Send-HaRaw($obj) {
    $json = $obj | ConvertTo-Json -Depth 10 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $seg = New-Object 'System.ArraySegment[byte]' -ArgumentList (, $bytes)
    $ws.SendAsync($seg, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $ct).GetAwaiter().GetResult() | Out-Null
}
function Receive-HaMessage {
    $buffer = New-Object byte[] 65536
    $ms = New-Object System.IO.MemoryStream
    do {
        $seg = New-Object 'System.ArraySegment[byte]' -ArgumentList (, $buffer)
        $res = $ws.ReceiveAsync($seg, $ct).GetAwaiter().GetResult()
        if ($res.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) {
            throw 'Home Assistant closed the connection.'
        }
        $ms.Write($buffer, 0, $res.Count)
    } while (-not $res.EndOfMessage)
    return (ConvertFrom-JsonText ([Text.Encoding]::UTF8.GetString($ms.ToArray())))
}
$script:msgId = 0
function Invoke-Ha([hashtable]$cmd) {
    $script:msgId++
    $cmd['id'] = $script:msgId
    Send-HaRaw $cmd
    while ($true) {
        $m = Receive-HaMessage
        if ((Get-Val $m 'id') -eq $script:msgId -and (Get-Val $m 'type') -eq 'result') { return $m }
    }
}

# ---------- connect + authenticate ----------
Write-Host "Connecting to $wsUrl ..."
try { [void]$ws.ConnectAsync([Uri]$wsUrl, $ct).GetAwaiter().GetResult() }
catch {
    Write-Host "Could not connect to $HaUrl. Check the address (try the IP, e.g. http://192.168.1.50:8123)." -ForegroundColor Red
    Write-Host $_.Exception.Message
    exit 1
}
$null = Receive-HaMessage                      # auth_required
Send-HaRaw @{ type = 'auth'; access_token = $token }
$auth = Receive-HaMessage
if ((Get-Val $auth 'type') -ne 'auth_ok') {
    Write-Host 'Home Assistant rejected the token. Create a new long-lived token and try again.' -ForegroundColor Red
    exit 1
}
$haVersion = Get-Val $auth 'ha_version'
Write-Host "Connected. Home Assistant $haVersion" -ForegroundColor Green

# ---------- read everything we need ----------
Write-Host 'Reading entities and devices ...'
$states  = @(Get-Val (Invoke-Ha @{ type = 'get_states' }) 'result')
$entReg  = @(Get-Val (Invoke-Ha @{ type = 'config/entity_registry/list' }) 'result')
$devReg  = @(Get-Val (Invoke-Ha @{ type = 'config/device_registry/list' }) 'result')

$devById = @{}
foreach ($d in $devReg) { $devById[(Get-Val $d 'id')] = $d }
$regByEntity = @{}
foreach ($e in $entReg) { $regByEntity[(Get-Val $e 'entity_id')] = $e }
$stateByEntity = @{}
foreach ($s in $states) { $stateByEntity[(Get-Val $s 'entity_id')] = $s }

function Get-DeviceName($dev) {
    if (-not $dev) { return $null }
    $n = Get-Val $dev 'name_by_user'
    if (-not $n) { $n = Get-Val $dev 'name' }
    return $n
}

function Get-EntitySummary([string]$entityId) {
    $st  = $stateByEntity[$entityId]
    $reg = $regByEntity[$entityId]
    $a   = Get-Val $st 'attributes'
    $dev = $null
    if ($reg -and (Get-Val $reg 'device_id')) { $dev = $devById[(Get-Val $reg 'device_id')] }

    $o = [ordered]@{
        entity_id    = $entityId
        state        = if ($st) { Get-Val $st 'state' } else { '(disabled or not loaded)' }
        name         = Get-Val $a 'friendly_name'
        device_class = Get-Val $a 'device_class'
        unit         = Get-Val $a 'unit_of_measurement'
        device       = Get-DeviceName $dev
        manufacturer = Get-Val $dev 'manufacturer'
        model        = Get-Val $dev 'model'
        integration  = Get-Val $reg 'platform'
        category     = Get-Val $reg 'entity_category'
        disabled_by  = Get-Val $reg 'disabled_by'
    }
    if (-not $o.name -and $reg) {
        $o.name = Get-Val $reg 'name'
        if (-not $o.name) { $o.name = Get-Val $reg 'original_name' }
    }
    foreach ($k in 'event_type', 'event_types', 'options') {
        $v = Get-Val $a $k
        if ($null -ne $v) { $o[$k] = $v }
    }
    if ($st) { $o.last_changed = Get-Val $st 'last_changed' }
    return $o
}

# ---------- classify ----------
$maintTokens = @(
    'filter', 'filters', 'rinse', 'salt', 'toner', 'ink', 'cartridge', 'drum', 'fuser',
    'maintenance', 'consumable', 'consumables', 'detergent', 'descale', 'descaling', 'decalcify',
    'clean', 'cleaning', 'brush', 'mop', 'dust', 'bag', 'pad', 'blade', 'reservoir', 'tank',
    'replace', 'replacement', 'remaining', 'life', 'service', 'wear', 'supply', 'supplies',
    'hepa', 'softener', 'lint', 'lamp', 'sponge', 'scale', 'limescale', 'refill', 'empty',
    'low', 'problem', 'error', 'fault', 'alert', 'warning', 'waste', 'bin', 'sensor_dirty', 'dirty'
)
$maintSubstrings = 'toner|filter|maint|consumab|rinse|descal|cartridge|hepa|limescale|ink_level|inklevel'
$applianceRx = 'dishwasher|refrigerator|fridge|freezer|printer|laserjet|officejet|deskjet|envy|pixma|\bhp\b|' +
               'canon|brother|epson|xerox|lexmark|washer|dryer|laundry|vacuum|roomba|roborock|purifier|humidifier|' +
               'thermostat|ecobee|furnace|hvac|softener|coffee|espresso|oven|microwave|' +
               'ice ?maker|water heater|smoke|air quality|home ?connect|thinq|smartthings'

function Get-Tokens([string]$s) {
    if (-not $s) { return @() }
    return @(($s.ToLower() -split '[^a-z0-9]+') | Where-Object { $_ })
}
function Test-Maint([string]$entityId, [string]$name) {
    $objectId = ($entityId -split '\.', 2)[1]
    foreach ($t in (Get-Tokens $objectId) + (Get-Tokens $name)) {
        if ($maintTokens -contains $t) { return $true }
    }
    return (($objectId + ' ' + $name).ToLower() -match $maintSubstrings)
}

$battery = New-Object System.Collections.ArrayList
$maint   = New-Object System.Collections.ArrayList
foreach ($s in $states) {
    $eid = Get-Val $s 'entity_id'
    $a   = Get-Val $s 'attributes'
    $dc  = Get-Val $a 'device_class'
    $nm  = Get-Val $a 'friendly_name'
    $domain = ($eid -split '\.', 2)[0]
    if ($domain -notin @('sensor', 'binary_sensor', 'event', 'button', 'number', 'select', 'switch', 'todo', 'update', 'vacuum', 'input_datetime', 'input_number', 'input_boolean', 'counter', 'timer')) { continue }

    if ($dc -eq 'battery' -or $eid -match 'battery') {
        [void]$battery.Add((Get-EntitySummary $eid))
    }
    elseif (($domain -eq 'binary_sensor' -and $dc -eq 'problem') -or (Test-Maint $eid $nm)) {
        [void]$maint.Add((Get-EntitySummary $eid))
    }
}

# Every entity on appliance-type devices, and on any device that owns a
# maintenance candidate (enabled or not)
$maintDeviceIds = @{}
foreach ($m in $maint) {
    $reg = $regByEntity[$m.entity_id]
    if ($reg -and (Get-Val $reg 'device_id')) { $maintDeviceIds[(Get-Val $reg 'device_id')] = $true }
}
$appliances = New-Object System.Collections.ArrayList
foreach ($d in $devReg) {
    $label = @((Get-DeviceName $d), (Get-Val $d 'manufacturer'), (Get-Val $d 'model')) -join ' '
    if ($label.ToLower() -notmatch $applianceRx -and -not $maintDeviceIds.ContainsKey((Get-Val $d 'id'))) { continue }
    $devId = Get-Val $d 'id'
    $ents = @($entReg | Where-Object { (Get-Val $_ 'device_id') -eq $devId } | ForEach-Object { Get-EntitySummary (Get-Val $_ 'entity_id') })
    if ($ents.Count -eq 0) { continue }
    [void]$appliances.Add([ordered]@{
        device       = Get-DeviceName $d
        manufacturer = Get-Val $d 'manufacturer'
        model        = Get-Val $d 'model'
        entities     = $ents
    })
}

# Integrations in use (names + counts only)
$integrations = [ordered]@{}
$entReg | Group-Object { Get-Val $_ 'platform' } | Sort-Object Count -Descending |
    ForEach-Object { $integrations[[string]$_.Name] = $_.Count }

# ---------- dashboards ----------
Write-Host 'Reading dashboards ...'
$resources = @()
try {
    $r = Invoke-Ha @{ type = 'lovelace/resources' }
    if (Get-Val $r 'success') { $resources = @(Get-Val $r 'result' | ForEach-Object { Get-Val $_ 'url' }) }
} catch { }

$dashDefs = @([ordered]@{ url_path = $null; title = 'Overview (default)'; mode = 'storage' })
$dl = Invoke-Ha @{ type = 'lovelace/dashboards/list' }
foreach ($d in @(Get-Val $dl 'result')) {
    $dashDefs += [ordered]@{ url_path = Get-Val $d 'url_path'; title = Get-Val $d 'title'; mode = Get-Val $d 'mode' }
}

$pageRx = 'maint|batter'
$dashboards = New-Object System.Collections.ArrayList
foreach ($dd in $dashDefs) {
    $entry = [ordered]@{ url_path = $dd.url_path; title = $dd.title; mode = $dd.mode }
    $cfgResp = Invoke-Ha @{ type = 'lovelace/config'; url_path = $dd.url_path; force = $false }
    if (-not (Get-Val $cfgResp 'success')) {
        $entry.note = 'No saved layout (auto-generated or unavailable): ' + (Get-Val (Get-Val $cfgResp 'error') 'message')
        [void]$dashboards.Add($entry); continue
    }
    $cfg = Get-Val $cfgResp 'result'
    if (Get-Val $cfg 'strategy') { $entry.strategy = Get-Val $cfg 'strategy' }
    $views = @(Get-Val $cfg 'views')
    $entry.views = @($views | Where-Object { $_ } | ForEach-Object {
        [ordered]@{ title = Get-Val $_ 'title'; path = Get-Val $_ 'path'; type = Get-Val $_ 'type' }
    })
    $whole = (([string]$dd.title) + ' ' + ([string]$dd.url_path)) -match $pageRx
    $matched = @($views | Where-Object { $_ -and ((([string](Get-Val $_ 'title')) + ' ' + ([string](Get-Val $_ 'path'))) -match $pageRx) })
    if ($whole) { $entry.full_config = $cfg }
    elseif ($matched.Count -gt 0) { $entry.matching_views_full = $matched }
    [void]$dashboards.Add($entry)
}

try { [void]$ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, 'done', $ct).GetAwaiter().GetResult() } catch { }

# ---------- write output ----------
$out = [ordered]@{
    generated_at           = (Get-Date).ToString('s')
    ha_version             = $haVersion
    custom_card_resources  = $resources
    integrations           = $integrations
    dashboards             = $dashboards
    battery_entities       = $battery
    maintenance_candidates = $maint
    appliance_devices      = $appliances
}
$json = $out | ConvertTo-Json -Depth 60
[IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))

Write-Host ''
Write-Host 'Done. Nothing in Home Assistant was changed.' -ForegroundColor Green
Write-Host ("  Dashboards found:         {0}" -f $dashboards.Count)
Write-Host ("  Battery entities:         {0}" -f $battery.Count)
Write-Host ("  Maintenance candidates:   {0}" -f $maint.Count)
Write-Host ("  Appliance-type devices:   {0}" -f $appliances.Count)
Write-Host ''
Write-Host "Output file: $OutFile" -ForegroundColor Cyan

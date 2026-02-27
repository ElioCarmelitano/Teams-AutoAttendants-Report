<#
.SYNOPSIS
Generates an HTML documentation report for all Microsoft Teams Auto Attendants in the tenant.

.DESCRIPTION
This script enumerates all Auto Attendants (AAs) and produces a formatted HTML report that documents:
 - General AA settings (name, ID, time zone, language, voice input, operator).
 - Business hours schedule text.
 - Business hours call handling classified as Menu / Disconnect / Redirect using the 'Automatic' rule:
     * If there is one MenuOption with DtmfResponse = 'Automatic' and Action = DisconnectCall -> Disconnect
     * If there is one MenuOption with DtmfResponse = 'Automatic' and Action = TransferCallToTarget -> Redirect (shows friendly target)
     * If there are keypad options (Tone1..9, Star, Pound) -> Menu
 - Default menu greeting (from DefaultCallFlow.Menu.Prompts) shown as:
     * "Greeting: TTS: <text>" OR "Greeting: AudioFile" OR "Greeting: None"
 - Default call flow options (DTMF + Action + Target) for all keys 0–9/*/# (no per-option greeting).
 - After-hours call handling classified with the same logic (Menu / Disconnect / Redirect).
 - Holidays (date ranges, greeting, action, target).
 - Resource accounts and numbers bound to the AA.
 - Authorized users.

Targets are resolved into friendly names with audit-friendly brackets:
 - Call Queue / Auto Attendant / Group (Shared Voicemail) / User / PSTN / Resource Account (ApplicationEndpoint).
 - Resource Accounts are mapped to the Call Queue / Auto Attendant they’re attached to (when possible) and
   rendered with "attached to call queue/auto attendant: <Name>" for clarity.

REQUIREMENTS
 - Teams PowerShell (Teams/Skype for Business Online) module with access to:
     * Get-CsAutoAttendant
     * Get-CsCallQueue
     * Get-CsOnlineApplicationInstance
     * Get-CsOnlineUser
 - Optional (for better SharedVoicemail group names):
     * MicrosoftTeams: Get-Team
     * ExchangeOnlineManagement: Get-UnifiedGroup
     * Microsoft.Graph: Get-MgGroup
   The script will use these if present, otherwise it gracefully falls back.

.PARAMETER OutputPath
Full path (or relative path) to the HTML file that will be created.
Default: ".\Teams-AutoAttendants-Report-<timestamp>.html"

.PARAMETER Open
If specified, opens the generated HTML report in the default browser/viewer when complete.

.PARAMETER IncludeAfterHoursOptions
(Reserved for future use) Would list after-hours menu options similar to default call flow options.
Currently unused; included as a placeholder for easy extension.

.EXAMPLE
PS> .\Export-TeamsAAReport.ps1
Generates ".\Teams-AutoAttendants-Report-YYYYMMDD-HHmmss.html" and opens it if -Open is specified.

.EXAMPLE
PS> .\Export-TeamsAAReport.ps1 -OutputPath "C:\Temp\AA-Report.html" -Open
Writes the report to C:\Temp\AA-Report.html and opens it.

.NOTES
- If you see parser errors like "variable followed by colon": use ${variable}: inside double-quoted strings.
- This script is read-only; it does not change any tenant configuration.
#>

[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath ("Teams-AutoAttendants-Report-{0}.html" -f (Get-Date -Format "yyyyMMdd-HHmmss"))),
    [switch]$Open,
    [switch]$IncludeAfterHoursOptions  # Reserved for future use; kept for easy enhancement
)

# Fail fast on unhandled errors
$ErrorActionPreference = "Stop"

# ----------------------------- Utilities -----------------------------
function HtmlEncode {
    <#
    .SYNOPSIS
    HTML-encodes a string; null-safe.
    #>
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    [System.Net.WebUtility]::HtmlEncode($Text)
}

# Global caches to avoid repeated lookups (faster, less throttling)
$script:_ResolveCache = @{
    Users      = @{}   # id -> display (user)
    RAs        = @{}   # id -> RA display
    CQsById    = @{}   # CQ id -> CQ name
    CQByRA     = @{}   # RA id -> CQ { Id, Name }
    AAsById    = @{}   # AA id -> AA name
    AAByRA     = @{}   # RA id -> AA { Id, Name }
    Groups     = @{}   # M365 group id -> group display
    IndexBuilt = $false
}

function Build-TargetIndex {
    <#
    .SYNOPSIS
    Preloads cross-references to speed up target resolution (CQ/AA and RA mappings).
    .DESCRIPTION
    - Loads all Call Queues (names + RA bindings).
    - Indexes the passed Auto Attendants (names + RA bindings).
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$AllAAs
    )

    # Load all Call Queues; map CQ by Id and ResourceAccount -> CQ
    try {
        $allCqs = @(Get-CsCallQueue)
    } catch {
        $allCqs = @()
    }

    foreach ($q in $allCqs) {
        $qid = [string]$q.Identity
        $script:_ResolveCache.CQsById[$qid] = $q.Name
        foreach ($raId in @($q.ResourceAccounts)) {
            $rid = [string]$raId
            $script:_ResolveCache.CQByRA[$rid] = @{ Id = $qid; Name = $q.Name }
        }
    }

    # Index the provided AAs (names + RA bindings)
    foreach ($aa in $AllAAs) {
        $aid = [string]$aa.Identity
        $script:_ResolveCache.AAsById[$aid] = $aa.Name
        foreach ($raId in @($aa.ApplicationInstances)) {
            $rid = [string]$raId
            $script:_ResolveCache.AAByRA[$rid] = @{ Id = $aid; Name = $aa.Name }
        }
    }

    $script:_ResolveCache.IndexBuilt = $true
}

function Get-IdFromCallTarget {
    <#
    .SYNOPSIS
    Extracts a usable identifier from a CallTarget-like object across schema variations.
    .OUTPUTS
    System.String or $null
    #>
    param($CallTarget)

    if (-not $CallTarget) { return $null }

    if ($CallTarget.PSObject.Properties['Id']) {
        $v = [string]$CallTarget.Id
        if ($v) { return $v }
    }

    foreach ($p in @('Identity','ObjectId','GroupId','UserId','ResourceAccountId','TargetId','ApplicationId')) {
        if ($CallTarget.PSObject.Properties[$p]) {
            $v = [string]$CallTarget.$p
            if ($v) { return $v }
        }
    }

    if ($CallTarget.PSObject.Properties['SipUri']) {
        $sip = [string]$CallTarget.SipUri
        if ($sip) { return $sip }
    }
    if ($CallTarget.PSObject.Properties['PhoneNumber']) {
        $num = [string]$CallTarget.PhoneNumber
        if ($num) { return $num }
    }

    foreach ($prop in $CallTarget.PSObject.Properties) {
        $val = [string]$prop.Value
        if (-not [string]::IsNullOrWhiteSpace($val)) {
            if ($val -match '^[0-9a-fA-F-]{36}$') { return $val }                # GUID
            if ($val -match '^tel:\+?\d') { return $val }                       # tel:+E164
            if ($val -match '^\+?\d[\d\s\-()]{6,}$') { return $val }            # phone-ish
        }
    }
    return $null
}

function Normalize-TargetType {
    <#
    .SYNOPSIS
    Normalizes various Teams/AA target 'Type' values into canonical buckets.
    #>
    param([string]$Type)
    switch -Regex ($Type) {
        '^ConfigurationEndpoint$' { return 'CallQueue' }   # CQ targets often appear as ConfigurationEndpoint
        '^CallQueue$'             { return 'CallQueue' }
        '^Queue$'                 { return 'CallQueue' }
        '^ApplicationEndpoint$'   { return 'ResourceAccount' }
        '^VoiceApp$'              { return 'ResourceAccount' }
        '^ResourceAccount$'       { return 'ResourceAccount' }
        '^AutoAttendant$'         { return 'AutoAttendant' }
        '^User$'                  { return 'User' }
        '^(SharedVoicemail|Group|Team|Channel)$' { return 'Group' }
        '^(ExternalPstn|PhoneNumber|Pstn)$'      { return 'Pstn' }
        default { return $Type }
    }
}

function Get-HumanTypeLabel {
    <#
    .SYNOPSIS
    Provides a human-friendly label for a normalized target type.
    #>
    param([string]$NormalizedType)
    switch ($NormalizedType) {
        'CallQueue'       { 'call queue' }
        'AutoAttendant'   { 'auto attendant' }
        'ResourceAccount' { 'resource account' }
        'Group'           { 'group' }
        'User'            { 'user' }
        'Pstn'            { 'phone number' }
        default           { $NormalizedType }
    }
}

function Resolve-CallTargetFriendly {
    <#
    .SYNOPSIS
    Resolves a CallTarget object to a friendly label and attached mapping (for Resource Accounts).
    .DESCRIPTION
    Returns a hashtable with:
      RawType, Id, NormType, Friendly, AttachedType, AttachedName
    #>
    param($CallTarget)

    if (-not $CallTarget) { return @{ RawType=""; Id=""; NormType=""; Friendly=""; AttachedType=""; AttachedName="" } }

    $rawType = [string]$CallTarget.Type
    $id      = Get-IdFromCallTarget -CallTarget $CallTarget
    $type    = Normalize-TargetType -Type $rawType

    function _try { param([scriptblock]$s) try { & $s } catch { $null } }

    switch ($type) {

        'User' {
            $label = $id
            if ($id) {
                if (-not $script:_ResolveCache.Users.ContainsKey($id)) {
                    $u = _try { Get-CsOnlineUser -Identity $id -ErrorAction Stop }
                    $script:_ResolveCache.Users[$id] = ($u.DisplayName ?? $id)
                }
                $label = $script:_ResolveCache.Users[$id]
            }
            return @{ RawType=$rawType; Id=$id; NormType=$type; Friendly=$label; AttachedType=""; AttachedName="" }
        }

        'Group' {
            # Shared Voicemail is a Group; we try Teams/EXO/Graph for a friendly name.
            $label = $id
            if ($id) {
                if (-not $script:_ResolveCache.Groups.ContainsKey($id)) {
                    $gName = $null
                    if (Get-Command Get-Team -ErrorAction SilentlyContinue) {
                        $t = _try { Get-Team -GroupId $id }
                        if ($t) { $gName = $t.DisplayName }
                    }
                    if (-not $gName -and (Get-Command Get-UnifiedGroup -ErrorAction SilentlyContinue)) {
                        $g = _try { Get-UnifiedGroup -Identity $id -ErrorAction Stop }
                        if ($g) { $gName = $g.DisplayName }
                    }
                    if (-not $gName -and (Get-Command Get-MgGroup -ErrorAction SilentlyContinue)) {
                        $g = _try { Get-MgGroup -GroupId $id -Property DisplayName }
                        if ($g) { $gName = $g.DisplayName }
                    }
                    if (-not $gName) { $gName = $id }
                    $script:_ResolveCache.Groups[$id] = $gName
                }
                $label = $script:_ResolveCache.Groups[$id]
            }
            return @{ RawType=$rawType; Id=$id; NormType=$type; Friendly=$label; AttachedType=""; AttachedName="" }
        }

        'Pstn' {
            # Strip tel: prefix for display
            $label = if ($id) { ($id -replace '^tel:','') } else { "" }
            return @{ RawType=$rawType; Id=$id; NormType=$type; Friendly=$label; AttachedType=""; AttachedName="" }
        }

        'CallQueue' {
            # The Id may be a CQ Id, a Resource Account Id, or (rarely) an AA Id.
            if ($id) {
                if ($script:_ResolveCache.CQsById.ContainsKey($id)) {
                    return @{ RawType=$rawType; Id=$id; NormType='CallQueue'; Friendly=$script:_ResolveCache.CQsById[$id]; AttachedType=""; AttachedName="" }
                }
                if ($script:_ResolveCache.CQByRA.ContainsKey($id)) {
                    $cq = $script:_ResolveCache.CQByRA[$id]
                    return @{ RawType=$rawType; Id=$id; NormType='CallQueue'; Friendly=$cq.Name; AttachedType=""; AttachedName="" }
                }
                if ($script:_ResolveCache.AAsById.ContainsKey($id)) {
                    return @{ RawType=$rawType; Id=$id; NormType='AutoAttendant'; Friendly=$script:_ResolveCache.AAsById[$id]; AttachedType=""; AttachedName="" }
                }
                $cqTry = _try { Get-CsCallQueue -Identity $id -ErrorAction Stop }
                if ($cqTry) {
                    $script:_ResolveCache.CQsById[$id] = $cqTry.Name
                    return @{ RawType=$rawType; Id=$id; NormType='CallQueue'; Friendly=$cqTry.Name; AttachedType=""; AttachedName="" }
                }
                $aaTry = _try { Get-CsAutoAttendant -Identity $id -ErrorAction Stop }
                if ($aaTry) {
                    $script:_ResolveCache.AAsById[$id] = $aaTry.Name
                    return @{ RawType=$rawType; Id=$id; NormType='AutoAttendant'; Friendly=$aaTry.Name; AttachedType=""; AttachedName="" }
                }
            }
            return @{ RawType=$rawType; Id=$id; NormType='CallQueue'; Friendly=($id ?? ""); AttachedType=""; AttachedName="" }
        }

        'AutoAttendant' {
            $label = $id
            if ($id) {
                if (-not $script:_ResolveCache.AAsById.ContainsKey($id)) {
                    $aa = _try { Get-CsAutoAttendant -Identity $id -ErrorAction Stop }
                    $script:_ResolveCache.AAsById[$id] = ($aa.Name ?? $id)
                }
                $label = $script:_ResolveCache.AAsById[$id]
            }
            return @{ RawType=$rawType; Id=$id; NormType='AutoAttendant'; Friendly=$label; AttachedType=""; AttachedName="" }
        }

        'ResourceAccount' {
            # Resolve RA display and map to CQ/AA if bound. Friendly -> attached entity name where possible.
            $raName = $id
            if ($id -and -not $script:_ResolveCache.RAs.ContainsKey($id)) {
                $ra = _try { Get-CsOnlineApplicationInstance -Identity $id -ErrorAction Stop }
                $script:_ResolveCache.RAs[$id] = ($ra.DisplayName ?? $id)
            }
            if ($id) { $raName = $script:_ResolveCache.RAs[$id] }

            if ($id -and $script:_ResolveCache.CQByRA.ContainsKey($id)) {
                $cq = $script:_ResolveCache.CQByRA[$id]
                return @{ RawType=$rawType; Id=$id; NormType='ResourceAccount'; Friendly=$cq.Name; AttachedType='CallQueue'; AttachedName=$cq.Name }
            }
            if ($id -and $script:_ResolveCache.AAByRA.ContainsKey($id)) {
                $aa = $script:_ResolveCache.AAByRA[$id]
                return @{ RawType=$rawType; Id=$id; NormType='ResourceAccount'; Friendly=$aa.Name; AttachedType='AutoAttendant'; AttachedName=$aa.Name }
            }

            return @{ RawType=$rawType; Id=$id; NormType='ResourceAccount'; Friendly=$raName; AttachedType=""; AttachedName="" }
        }

        default {
            # Fallback: try RA heuristic; else echo ID
            if ($id -and -not $script:_ResolveCache.RAs.ContainsKey($id)) {
                $ra = _try { Get-CsOnlineApplicationInstance -Identity $id -ErrorAction Stop }
                if ($ra) { $script:_ResolveCache.RAs[$id] = $ra.DisplayName }
            }
            $friendly = if ($id) { ($script:_ResolveCache.RAs[$id] ?? $id) } else { "" }
            return @{ RawType=$rawType; Id=$id; NormType=$type; Friendly=$friendly; AttachedType=""; AttachedName="" }
        }
    }
}

function Format-TargetBracket {
    <#
    .SYNOPSIS
    Renders a muted bracket detail after a friendly target name.
    .DESCRIPTION
    (human label; RawType Id; [attached to call queue|auto attendant: Name])
    #>
    param($Resolved)

    $humanType = Get-HumanTypeLabel -NormalizedType $Resolved.NormType
    $rawPart   = if ($Resolved.Id) { "$($Resolved.RawType) $($Resolved.Id)" } else { $Resolved.RawType }

    $attach = ""
    if ($Resolved.NormType -eq 'ResourceAccount' -and $Resolved.AttachedType) {
        $attachHuman = Get-HumanTypeLabel -NormalizedType $Resolved.AttachedType
        if ($Resolved.AttachedName) {
            # Use ${attachHuman}: to avoid parser ambiguity
            $attach = "; attached to ${attachHuman}: $(HtmlEncode $Resolved.AttachedName)"
        } else {
            $attach = "; attached to ${attachHuman}"
        }
    }

    " <span style='color:#605E5C'>( $humanType; $(HtmlEncode $rawPart)$attach )</span>"
}

# ----------------------------- Greeting helpers -----------------------------
function Get-CallableEntitySummary {
    <#
    .SYNOPSIS
    Simple "Type : Id" HTML-encoded summary for operators etc.
    #>
    param($Entity)
    if (-not $Entity) { return "" }

    $t  = $Entity.Type
    $id = Get-IdFromCallTarget -CallTarget $Entity

    if ($t -and $id) { return "$(HtmlEncode $t) : $(HtmlEncode $id)" }
    if ($id) { return HtmlEncode $id }

    HtmlEncode (($Entity | Out-String).Trim())
}

function Get-PromptSummary {
    <#
    .SYNOPSIS
    Returns "TTS: <text>" or "AudioFile: <name>" or "Greeting present (unparsed)".
    #>
    param($Prompt)

    if (-not $Prompt) { return "No greeting" }

    if ($Prompt.TextToSpeechPrompt) {
        return "TTS: $(HtmlEncode $Prompt.TextToSpeechPrompt)"
    }

    if ($Prompt.AudioFilePrompt) {
        $af = $Prompt.AudioFilePrompt
        $name = $af.FileName
        if (-not $name) { $name = $af.Name }
        if (-not $name) { $name = $af.Id }
        if (-not $name) { $name = "Audio prompt (no name exposed)" }
        return "AudioFile: $(HtmlEncode $name)"
    }

    "Greeting present (unparsed)"
}

function Get-GreetingsSummary {
    <#
    .SYNOPSIS
    Joins prompt summaries with <br/>; returns "No greeting" if none.
    #>
    param($Greetings)
    if (-not $Greetings -or $Greetings.Count -eq 0) { return "No greeting" }
    ($Greetings | ForEach-Object { Get-PromptSummary -Prompt $_ }) -join "<br/>"
}

function Get-MenuPromptsSummary {
    <#
    .SYNOPSIS
    Summarizes DefaultCallFlow.Menu.Prompts as a single line for the table.
    .OUTPUTS
    "Greeting: TTS: <text>" | "Greeting: AudioFile" | "Greeting: None"
    #>
    param($Menu)
    if (-not $Menu -or -not $Menu.Prompts) { return "Greeting: None" }

    $p = $Menu.Prompts
    if ($p.ActiveType -eq 'TextToSpeech' -and $p.TextToSpeechPrompt) { return "Greeting: TTS: $(HtmlEncode $p.TextToSpeechPrompt)" }
    if ($p.ActiveType -eq 'AudioFile' -and $p.AudioFilePrompt)     { return "Greeting: AudioFile" }

    if ($p.TextToSpeechPrompt) { return "Greeting: TTS: $(HtmlEncode $p.TextToSpeechPrompt)" }
    if ($p.AudioFilePrompt)   { return "Greeting: AudioFile" }

    "Greeting: None"
}

# ----------------------------- Call flow classification -----------------------------
function Get-CallFlowModeSummary {
    <#
    .SYNOPSIS
    Classifies a call flow as Menu / Disconnect / Redirect using 'Automatic' vs keypad options.
    .DESCRIPTION
    - If exactly one option with DtmfResponse='Automatic' and Action=DisconnectCall -> Disconnect
    - If exactly one option with DtmfResponse='Automatic' and Action=TransferCallToTarget -> Redirect (shows target)
    - If there are keypad options (Tone1..9, Star, Pound) -> Menu
    Otherwise shows a reasonable fallback with greeting.
    #>
    param($CallFlow)

    if (-not $CallFlow) { return "Not configured" }
    $menu = $CallFlow.Menu
    $greeting = Get-GreetingsSummary -Greetings $CallFlow.Greetings

    if (-not $menu -or -not $menu.MenuOptions -or $menu.MenuOptions.Count -eq 0) {
        return "<b>Greeting:</b> $greeting<br/><b>Action:</b> Not configured"
    }

    $opts = @($menu.MenuOptions)
    $auto = $opts | Where-Object { $_.DtmfResponse -eq 'Automatic' }
    $keys = $opts | Where-Object { $_.DtmfResponse -ne 'Automatic' }

    if ($auto.Count -eq 1 -and $keys.Count -eq 0) {
        $o = $auto[0]
        switch ($o.Action) {
            'DisconnectCall' {
                return "<b>Greeting:</b> $greeting<br/><b>Action:</b> Disconnect"
            }
            'TransferCallToTarget' {
                if ($o.CallTarget) {
                    $resolved = Resolve-CallTargetFriendly -CallTarget $o.CallTarget
                    $id       = Get-IdFromCallTarget -CallTarget $o.CallTarget
                    $friendly = "$(HtmlEncode $resolved.Friendly)"
                    $rawLabel = if ($id) { "$(HtmlEncode $o.CallTarget.Type) $(HtmlEncode $id)" } else { "$(HtmlEncode $o.CallTarget.Type)" }
                    $bracket  = Format-TargetBracket -Resolved $resolved

                    $targetPart = if ($friendly) {
                        if ($rawLabel -and ($rawLabel -ne $friendly)) {
                            "<br/><b>Target:</b> $friendly <span style='color:#605E5C'>( $rawLabel )</span>$bracket"
                        } else {
                            "<br/><b>Target:</b> $friendly$bracket"
                        }
                    } elseif ($rawLabel) { "<br/><b>Target:</b> $rawLabel$bracket" } else { "" }

                    return "<b>Greeting:</b> $greeting<br/><b>Action:</b> Redirect$targetPart"
                } else {
                    return "<b>Greeting:</b> $greeting<br/><b>Action:</b> Redirect (no target returned)"
                }
            }
            default {
                $label = switch -Regex ($o.Action) {
                    "Announcement" { "Announcement" }
                    default        { [string]$o.Action }
                }
                return "<b>Greeting:</b> $greeting<br/><b>Action:</b> $(HtmlEncode $label)"
            }
        }
    }

    if ($keys.Count -gt 0) {
        return "<b>Greeting:</b> $greeting<br/><b>Action:</b> Menu"
    }

    # Fallback: multiple 'Automatic' (unusual)
    $first = $auto | Select-Object -First 1
    if ($first) {
        $act = switch -Regex ($first.Action) {
            "Disconnect"   { "Disconnect" }
            "Transfer"     { "Redirect" }
            "Announcement" { "Announcement" }
            default        { [string]$first.Action }
        }
        if ($act -eq 'Redirect' -and $first.CallTarget) {
            $resolved = Resolve-CallTargetFriendly -CallTarget $first.CallTarget
            $id       = Get-IdFromCallTarget -CallTarget $first.CallTarget
            $friendly = "$(HtmlEncode $resolved.Friendly)"
            $rawLabel = if ($id) { "$(HtmlEncode $first.CallTarget.Type) $(HtmlEncode $id)" } else { "$(HtmlEncode $first.CallTarget.Type)" }
            $bracket  = Format-TargetBracket -Resolved $resolved
            $targetPart = if ($friendly) {
                if ($rawLabel -and ($rawLabel -ne $friendly)) {
                    "<br/><b>Target:</b> $friendly <span style='color:#605E5C'>( $rawLabel )</span>$bracket"
                } else {
                    "<br/><b>Target:</b> $friendly$bracket"
                }
            } elseif ($rawLabel) { "<br/><b>Target:</b> $rawLabel$bracket" } else { "" }
            return "<b>Greeting:</b> $greeting<br/><b>Action:</b> $act$targetPart"
        } else {
            return "<b>Greeting:</b> $greeting<br/><b>Action:</b> $act"
        }
    }

    "<b>Greeting:</b> $greeting<br/><b>Action:</b> Unknown"
}

# ----------------------------- Misc formatting helpers -----------------------------
function Get-FixedScheduleRangesText {
    <#
    .SYNOPSIS
    Formats FixedSchedule.DateTimeRanges from an AA schedule into "dd/MM/yyyy HH:mm-dd/MM/yyyy HH:mm".
    #>
    param($Schedule)

    if (-not $Schedule) { return "" }
    if ($Schedule.Type -ne "Fixed") { return "" }
    if (-not $Schedule.FixedSchedule) { return "" }

    $ranges = $Schedule.FixedSchedule.DateTimeRanges
    if (-not $ranges) { return "" }

    $parts = @()

    foreach ($r in @($ranges)) {
        if ($null -eq $r) { continue }
        $start = $r.Start
        $end   = $r.End

        $sTxt = if ($start -is [datetime]) { $start.ToString("dd/MM/yyyy HH:mm") } else { [string]$start }
        $eTxt = if ($end   -is [datetime]) { $end.ToString("dd/MM/yyyy HH:mm") } else { [string]$end }

        $parts += "$sTxt-$eTxt"
    }

    if ($parts.Count -eq 0) { return "" }

    HtmlEncode ($parts -join "; ")
}

function Get-HolidayLines {
    <#
    .SYNOPSIS
    Renders AA holiday lines: "<Name>: <dates>, Greeting: <...>, Action: <...>, Target: <...>" (friendly + brackets).
    #>
    param(
        [Parameter(Mandatory)] $HolidayAssocs,
        [Parameter(Mandatory)] $SchedById,
        [Parameter(Mandatory)] $CfById
    )

    if (-not $HolidayAssocs -or $HolidayAssocs.Count -eq 0) { return "None configured" }

    $lines = foreach ($ha in $HolidayAssocs) {
        $sid = [string]$ha.ScheduleId
        $cid = [string]$ha.CallFlowId

        $hs = $SchedById[$sid]
        $hf = $CfById[$cid]

        $holidayName = if ($hs -and $hs.Name) { $hs.Name } else { $sid }

        # Dates
        $rangeText = ""
        if ($hs) { $rangeText = Get-FixedScheduleRangesText -Schedule $hs }
        if (-not $rangeText) { $rangeText = "Dates not found" }

        # Greeting (holiday call flow)
        $greetingText = "No greeting"
        if ($hf) { $greetingText = Get-GreetingsSummary -Greetings $hf.Greetings }

        # Action + Target (with enhanced brackets)
        $at = @{ Action="Unknown"; Target=""; TargetFriendly="" }
        if ($hf -and $hf.Menu) { $at = Get-ActionAndTargetFromMenu -Menu $hf.Menu }

        $targetPart = ""
        if ($at.TargetFriendly) {
            $targetPart = ", <b>Target:</b> $($at.TargetFriendly)"
        } elseif ($at.Target) {
            $targetPart = ", <b>Target:</b> $($at.Target)"
        }

        "<b>$(HtmlEncode $holidayName):</b> $rangeText, <b>Greeting:</b> $greetingText, <b>Action:</b> $(HtmlEncode $at.Action)$targetPart"
    }

    $lines -join "<br/>"
}

# Choose a representative target for single-line sections (Holidays, legacy one-liners)
function Get-ActionAndTargetFromMenu {
    <#
    .SYNOPSIS
    Picks a representative option from a menu, preferring Automatic transfers, and returns Action + Target labels.
    #>
    param($Menu)

    if (-not $Menu -or -not $Menu.MenuOptions -or $Menu.MenuOptions.Count -eq 0) {
        return @{ Action = "No menu"; Target = ""; TargetFriendly = "" }
    }

    $opt =
        $Menu.MenuOptions | Where-Object { $_.DtmfResponse -eq "Automatic" -and $_.Action -match 'Transfer' -and $_.CallTarget } | Select-Object -First 1
    if (-not $opt) {
        $opt = $Menu.MenuOptions | Where-Object { $_.DtmfResponse -eq "Automatic" -and $_.CallTarget } | Select-Object -First 1
    }
    if (-not $opt) {
        $opt = $Menu.MenuOptions | Where-Object { $_.CallTarget } | Select-Object -First 1
    }
    if (-not $opt) {
        $opt = $Menu.MenuOptions | Select-Object -First 1
    }

    $act = $opt.Action
    if (-not $act) { $act = "Unknown" }

    $actionLabel = $act
    switch -Regex ($act) {
        "Disconnect"   { $actionLabel = "Disconnect" }
        "Transfer"     { $actionLabel = "Forward/Transfer" }
        "Announcement" { $actionLabel = "Announcement" }
    }

    $targetLabel    = ""
    $targetFriendly = ""

    if ($opt.CallTarget) {
        $resolved = Resolve-CallTargetFriendly -CallTarget $opt.CallTarget
        $id       = Get-IdFromCallTarget -CallTarget $opt.CallTarget

        $targetFriendly = "$(HtmlEncode $resolved.Friendly)"
        $targetLabel    = if ($id) { "$(HtmlEncode $opt.CallTarget.Type) $(HtmlEncode $id)" } else { "$(HtmlEncode $opt.CallTarget.Type)" }

        # Append the bracket info right after friendly/raw (human type, raw Type/Id, RA attachment if any)
        $bracket = Format-TargetBracket -Resolved $resolved
        if ($targetFriendly) {
            $targetFriendly = "$targetFriendly$bracket"
        } elseif ($targetLabel) {
            $targetLabel = "$targetLabel$bracket"
        }
    }

    @{ Action = $actionLabel; Target = $targetLabel; TargetFriendly = $targetFriendly }
}

function Get-DefaultCallFlowCleanSummary {
    <#
    .SYNOPSIS
    Legacy one-liner for call flow greeting + representative action/target (still used in some sections).
    #>
    param($CallFlow)
    if (-not $CallFlow) { return "Not configured" }

    $greeting = Get-GreetingsSummary -Greetings $CallFlow.Greetings
    $at = Get-ActionAndTargetFromMenu -Menu $CallFlow.Menu

    $targetPart = ""
    if ($at.TargetFriendly) {
        $targetPart = "<br/><b>Target:</b> $($at.TargetFriendly)"
    } elseif ($at.Target) {
        $targetPart = "<br/><b>Target:</b> $($at.Target)"
    }

    "<b>Greeting:</b> $greeting<br/><b>Action:</b> $(HtmlEncode $at.Action)$targetPart"
}

# ----------------------------- Default menu options (no per-option greeting) -----------------------------
function Get-DefaultMenuOptionsLines {
    <#
    .SYNOPSIS
    Renders all default call flow options as "Option X: Action: <...>, Target: <...>" (no per-option greeting).
    #>
    param($Menu)

    if (-not $Menu -or -not $Menu.MenuOptions -or $Menu.MenuOptions.Count -eq 0) { return "No options configured" }

    function Get-DtmfLabel([string]$d) {
        switch ($d) {
            "Tone0" { "0" }
            "Tone1" { "1" }
            "Tone2" { "2" }
            "Tone3" { "3" }
            "Tone4" { "4" }
            "Tone5" { "5" }
            "Tone6" { "6" }
            "Tone7" { "7" }
            "Tone8" { "8" }
            "Tone9" { "9" }
            "Star"  { "*" }
            "Pound" { "#" }
            "Automatic" { "Auto" }
            default { $d }
        }
    }

    $lines = foreach ($opt in $Menu.MenuOptions) {
        $dtmf = Get-DtmfLabel -d ([string]$opt.DtmfResponse)
        $act  = if ($opt.Action) { [string]$opt.Action } else { "Unknown" }

        $actionLabel = $act
        switch -Regex ($act) {
            "Disconnect"   { $actionLabel = "Disconnect" }
            "Transfer"     { $actionLabel = "Forward/Transfer" }
            "Announcement" { $actionLabel = "Announcement" }
        }

        $targetLabel    = ""
        $targetFriendly = ""
        if ($opt.CallTarget) {
            $resolved = Resolve-CallTargetFriendly -CallTarget $opt.CallTarget
            $id       = Get-IdFromCallTarget -CallTarget $opt.CallTarget

            $targetFriendly = "$(HtmlEncode $resolved.Friendly)"
            $targetLabel    = if ($id) { "$(HtmlEncode $opt.CallTarget.Type) $(HtmlEncode $id)" } else { "$(HtmlEncode $opt.CallTarget.Type)" }

            # Add bracket detail after the friendly/raw
            $bracket = Format-TargetBracket -Resolved $resolved
            if ($targetFriendly) {
                $targetFriendly = "$targetFriendly$bracket"
            } elseif ($targetLabel) {
                $targetLabel = "$targetLabel$bracket"
            }
        }

        $targetPart = ""
        if ($targetFriendly) {
            if ($targetLabel -and ($targetLabel -ne $targetFriendly)) {
                $targetPart = ", <b>Target:</b> $targetFriendly <span style='color:#605E5C'>( $targetLabel )</span>"
            } else {
                $targetPart = ", <b>Target:</b> $targetFriendly"
            }
        } elseif ($targetLabel) {
            $targetPart = ", <b>Target:</b> $targetLabel"
        }

        # Use ${dtmf} because it's followed by a colon
        "<b>Option ${dtmf}:</b> <b>Action:</b> $(HtmlEncode $actionLabel)$targetPart"
    }

    $lines -join "<br/>"
}

# ----------------------------- Authorized users helper -----------------------------
function Resolve-AuthorizedUsers {
    <#
    .SYNOPSIS
    Resolves authorized user object IDs into recognizable strings (UPN / SIP / DisplayName).
    #>
    param([string[]]$ObjectIds, [hashtable]$Cache)

    if (-not $ObjectIds -or $ObjectIds.Count -eq 0) { return @() }
    if (-not $Cache) { $Cache = @{} }

    foreach ($id in $ObjectIds) {
        $idStr = [string]$id
        if ($Cache.ContainsKey($idStr)) { $Cache[$idStr]; continue }

        $label = $null
        try {
            $u = Get-CsOnlineUser -Identity $idStr -ErrorAction Stop
            $label = ($u.UserPrincipalName ?? $u.SipAddress ?? $u.DisplayName ?? $idStr)
        } catch {
            try {
                if (Get-Command Get-MgUser -ErrorAction SilentlyContinue) {
                    $g = Get-MgUser -UserId $idStr -Property DisplayName,UserPrincipalName -ErrorAction Stop
                    $label = ($g.UserPrincipalName ?? $g.DisplayName ?? $idStr)
                }
            } catch { $label = $idStr }
        }

        if (-not $label) { $label = $idStr }
        $Cache[$idStr] = $label
        $label
    }
}

# ----------------------------- HTML table builder -----------------------------
function New-ParamValueTableHtml {
    <#
    .SYNOPSIS
    Builds a <table> with "Parameter" / "Value" columns for a given section title.
    #>
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][System.Collections.Generic.List[object]]$Rows
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<h2>$(HtmlEncode $Title)</h2>")
    [void]$sb.AppendLine("<table>")
    [void]$sb.AppendLine("<thead><tr><th>Parameter</th><th>Value</th></tr></thead>")
    [void]$sb.AppendLine("<tbody>")

    foreach ($r in $Rows) {
        $p = HtmlEncode ([string]$r.Parameter)
        $v = [string]$r.Value   # Value may contain <br/>/<b> etc intentionally
        [void]$sb.AppendLine("<tr><td>$p</td><td>$v</td></tr>")
    }

    [void]$sb.AppendLine("</tbody></table>")
    $sb.ToString()
}

# ----------------------------- MAIN -----------------------------
try {
    # 1) Fetch all AAs up-front and build the cross-reference index (CQs, AAs, RA->CQ/AA)
    $aasAll = @(Get-CsAutoAttendant)
    if ($aasAll.Count -eq 0) { throw "No auto attendants found." }

    Build-TargetIndex -AllAAs $aasAll

    $authUserCache = @{}
    $sections = @()

    # 2) Iterate AAs and build sections
    foreach ($aaLite in $aasAll) {

        # Fetch full AA (ensures CallFlows/Schedules are present)
        $aa = Get-CsAutoAttendant -Identity $aaLite.Identity

        # Index schedules and call flows for quick lookup by Id
        $schedById = @{}
        foreach ($s in @($aa.Schedules)) { $schedById[[string]$s.Id] = $s }

        $cfById = @{}
        foreach ($cf in @($aa.CallFlows)) { $cfById[[string]$cf.Id] = $cf }

        $rows = New-Object "System.Collections.Generic.List[object]"

        # ---- Identity / Basics ----
        $rows.Add([pscustomobject]@{ Parameter="Name"; Value=HtmlEncode $aa.Name })
        $rows.Add([pscustomobject]@{ Parameter="Identity"; Value=HtmlEncode $aa.Identity })
        $rows.Add([pscustomobject]@{ Parameter="Time zone"; Value=HtmlEncode $aa.TimeZoneId })
        $rows.Add([pscustomobject]@{ Parameter="Language"; Value=HtmlEncode $aa.LanguageId })
        $rows.Add([pscustomobject]@{ Parameter="Voice input (EnableVoiceResponse)"; Value=($(if ($aa.VoiceResponseEnabled) { "Yes" } else { "No" })) })
        $rows.Add([pscustomobject]@{ Parameter="Voice / TTS voice"; Value=HtmlEncode ($aa.VoiceId ?? "") })
        $rows.Add([pscustomobject]@{ Parameter="Operator"; Value=(Get-CallableEntitySummary $aa.Operator) })

        # ---- Business hours schedule (text) ----
        $rows.Add([pscustomobject]@{ Parameter="Business hours"; Value=(Get-BusinessHoursText -AA $aa -SchedById $schedById) })

        # ---- Business hours call handling (Menu / Disconnect / Redirect) ----
        $rows.Add([pscustomobject]@{
            Parameter = "Business hours call handling"
            Value     = (Get-CallFlowModeSummary -CallFlow $aa.DefaultCallFlow)
        })

        # ---- Default menu greeting (from DefaultCallFlow.Menu.Prompts) ----
        if ($aa.DefaultCallFlow -and $aa.DefaultCallFlow.Menu) {
            $rows.Add([pscustomobject]@{
                Parameter = "Default menu greeting"
                Value     = (Get-MenuPromptsSummary -Menu $aa.DefaultCallFlow.Menu)
            })
        } else {
            $rows.Add([pscustomobject]@{ Parameter="Default menu greeting"; Value="Greeting: None" })
        }

        # ---- Default call flow options (DTMF + Action + Target; NO per-option greeting) ----
        if ($aa.DefaultCallFlow -and $aa.DefaultCallFlow.Menu) {
            $rows.Add([pscustomobject]@{
                Parameter = "Default call flow options"
                Value     = (Get-DefaultMenuOptionsLines -Menu $aa.DefaultCallFlow.Menu)
            })
        } else {
            $rows.Add([pscustomobject]@{ Parameter="Default call flow options"; Value="No options configured" })
        }

        # ---- After-hours call handling (classified; mirrors business hours logic) ----
        $afterAssoc = @($aa.CallHandlingAssociations | Where-Object { $_.Type -eq "AfterHours" } | Select-Object -First 1)
        if ($afterAssoc) {
            $afterFlow = $cfById[[string]$afterAssoc.CallFlowId]
            if ($afterFlow) {
                $rows.Add([pscustomobject]@{
                    Parameter = "After-hours call handling"
                    Value     = (Get-CallFlowModeSummary -CallFlow $afterFlow)
                })
            } else {
                $rows.Add([pscustomobject]@{ Parameter="After-hours call handling"; Value="After-hours call flow not found in AA object" })
            }
        } else {
            $rows.Add([pscustomobject]@{ Parameter="After-hours call handling"; Value="Not configured" })
        }

        # ---- Holidays ----
        $holidayAssocs = @($aa.CallHandlingAssociations | Where-Object { $_.Type -eq "Holiday" })
        $rows.Add([pscustomobject]@{ Parameter="Holidays"; Value=(Get-HolidayLines -HolidayAssocs $holidayAssocs -SchedById $schedById -CfById $cfById) })

        # ---- Resource accounts / numbers ----
        $raLines = @()
        if ($aa.ApplicationInstances) {
            foreach ($ai in $aa.ApplicationInstances) {
                try {
                    $ra = Get-CsOnlineApplicationInstance -Identity $ai -ErrorAction Stop
                    $line = "$(HtmlEncode $ra.DisplayName) ($(HtmlEncode $ra.ObjectId))"
                    if ($ra.PhoneNumber) { $line += " - $(HtmlEncode $ra.PhoneNumber)" }
                    $raLines += $line
                } catch {
                    $raLines += "$(HtmlEncode $ai) (unable to resolve resource account details)"
                }
            }
        }
        $rows.Add([pscustomobject]@{ Parameter="Resource accounts / numbers"; Value=($(if ($raLines.Count) { $raLines -join "<br/>" } else { "None / not returned by API" })) })

        # ---- Authorized users ----
        $authIds = @($aa.AuthorizedUsers)
        $authNames = Resolve-AuthorizedUsers -ObjectIds $authIds -Cache $authUserCache
        $rows.Add([pscustomobject]@{ Parameter="Authorized users"; Value=($(if ($authNames.Count) { ($authNames | ForEach-Object { HtmlEncode $_ }) -join "<br/>" } else { "None" })) })

        # Accumulate section
        $sections += (New-ParamValueTableHtml -Title $aa.Name -Rows $rows)
    }

    # ----------------------------- HTML Styles -----------------------------
    $style = @"
<style>
body { font-family: "Segoe UI Variable","Segoe UI",Arial,sans-serif; font-size: 13px; color: #323130; margin: 20px; }
h1 { margin-bottom: 10px; }
h2 { margin-top: 28px; }

table { 
    border-collapse: collapse; 
    width: 100%; 
    font-size: 12px;
    table-layout: auto;   /* allow dynamic sizing */
}

th, td { 
    border: 1px solid #ddd; 
    padding: 8px; 
    vertical-align: top; 
    word-wrap: break-word; 
}

th { 
    background: #f3f3f3; 
    text-align: left; 
    font-weight: 600;
}

/* Make left column narrower */
td:first-child, th:first-child {
    width: 220px;           /* adjust to taste (180–250px works well) */
    white-space: nowrap;    /* prevents wrapping of parameter names */
}

/* Let right column expand */
td:last-child, th:last-child {
    width: auto;
}

tr:nth-child(even) { background: #fafafa; }
</style>
"@

    # ----------------------------- Write HTML File -----------------------------
    @"
<html>
<head>
<meta charset="utf-8"/>
<title>Teams Auto Attendant Documentation</title>
$style
</head>
<body>
<h1>Teams Auto Attendant Documentation</h1>
<p>Generated: $(Get-Date)</p>
$($sections -join "`n")
</body>
</html>
"@ | Out-File -FilePath $OutputPath -Encoding UTF8

    Write-Host "Report generated: $OutputPath" -ForegroundColor Green
    if ($Open) { Invoke-Item $OutputPath }

} catch {
    Write-Error $_
    throw
}
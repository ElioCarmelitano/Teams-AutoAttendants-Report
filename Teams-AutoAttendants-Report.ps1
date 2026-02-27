<#
.SYNOPSIS
Generates an HTML documentation report for all Microsoft Teams Auto Attendants.

.DESCRIPTION
This read-only script enumerates all Auto Attendants (AAs) and produces a structured HTML report.
It resolves friendly names for targets and agents, and includes bracketed audit details such as
raw IDs and Resource Account attachments.

The report includes:
 - Identity & basics (Name, Identity, Time zone, Language, Voice input, Voice/TTS voice, Operator)
 - Business hours (weekly schedule text)
 - Business hours call handling classified as Menu / Disconnect / Redirect using the 'Automatic' rule:
     * Single MenuOption with DtmfResponse='Automatic' + Action=DisconnectCall -> Disconnect
     * Single MenuOption with DtmfResponse='Automatic' + Action=TransferCallToTarget -> Redirect (shows target)
     * Any keypad options (Tone1..9, Star, Pound) -> Menu
 - Default menu greeting (from DefaultCallFlow.Menu.Prompts):
     * "Greeting: TTS: <text>" OR "Greeting: AudioFile" OR "Greeting: None"
 - Default call flow options (DTMF + Action + Target) for all keys 0–9/*/# (no per-option greeting)
 - After-hours call handling classified with the same logic (Menu / Disconnect / Redirect)
 - Holidays (date ranges, greeting, action, target)
 - Resource accounts and numbers bound to the AA
 - Authorized users (friendly DisplayName)

TARGET RESOLUTION (best effort)
 - PSTN: strips 'tel:' and displays E.164
 - Call Queue: CQ name by Id (handles ConfigurationEndpoint)
 - Auto Attendant: AA name by Id
 - Resource Account (ApplicationInstance): RA display name and, if mapped, the attached CQ/AA name
 - Group (Shared Voicemail / Team / Channel): Team/M365 Group display name where possible
 - User: DisplayName (falls back to UPN/Id)
 - Otherwise: shows the raw Id as “Unknown”

REQUIREMENTS
 - Teams/Skype Online PowerShell with access to:
     * Get-CsAutoAttendant
     * Get-CsCallQueue (for target resolution)
     * Get-CsOnlineApplicationInstance (resource accounts)
     * Get-CsOnlineUser (authorized users)
 - Optional (for richer group names):
     * MicrosoftTeams: Get-Team
     * ExchangeOnlineManagement: Get-UnifiedGroup
     * Microsoft.Graph: Get-MgGroup

.PARAMETER OutputPath
Full or relative path to the HTML report file.
Default: .\Teams-AutoAttendants-Report-<timestamp>.html

.PARAMETER Open
If specified, opens the generated HTML report when complete.

.PARAMETER IncludeAfterHoursOptions
(Reserved for future use) Would list after-hours menu options similar to default call flow options.
Currently unused; included as a placeholder for easy extension.

.EXAMPLE
PS> .\Export-TeamsAAReport.ps1
Generates ".\Teams-AutoAttendants-Report-YYYYMMDD-HHmmss.html".

.EXAMPLE
PS> .\Export-TeamsAAReport.ps1 -OutputPath "C:\Temp\AA-Report.html" -Open
Writes the report to C:\Temp\AA-Report.html and opens it in the default browser.

.NOTES
 - Parser-safety: Hashtables use ';' between key/value pairs. Complex values are built in variables first.
 - Null-safety: All .ContainsKey / indexing operations guard against null/empty keys.
 - This script is read-only; it does not modify any tenant configuration.
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

# --------- Helpers to build a topology index for fast lookups ----------
function Build-TargetIndex {
    param(
        [Parameter(Mandatory=$true)]
        [array]$AllAAs
    )

    # Load all Call Queues; map CQ by Id and ResourceAccount -> CQ
    $allCqs = @()
    try { $allCqs = @(Get-CsCallQueue) } catch {}
    foreach ($q in $allCqs) {
        if ($null -eq $q) { continue }
        $qid = [string]$q.Identity
        if (-not [string]::IsNullOrWhiteSpace($qid)) {
            $script:_ResolveCache.CQsById[$qid] = $q.Name
        }
        foreach ($raId in @($q.ResourceAccounts)) {
            $rid = [string]$raId
            if (-not [string]::IsNullOrWhiteSpace($rid)) {
                $script:_ResolveCache.CQByRA[$rid] = @{ Id = $qid; Name = $q.Name }
            }
        }
    }

    # Index the provided AAs (names + RA bindings)
    foreach ($aa in $AllAAs) {
        if ($null -eq $aa) { continue }
        $aid = [string]$aa.Identity
        if (-not [string]::IsNullOrWhiteSpace($aid)) {
            $script:_ResolveCache.AAsById[$aid] = $aa.Name
        }
        foreach ($raId in @($aa.ApplicationInstances)) {
            $rid = [string]$raId
            if (-not [string]::IsNullOrWhiteSpace($rid)) {
                $script:_ResolveCache.AAByRA[$rid] = @{ Id = $aid; Name = $aa.Name }
            }
        }
    }

    $script:_ResolveCache.IndexBuilt = $true
}

# --------------------- Core extract/normalize/resolve ----------------------
function Get-IdFromCallTarget {
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
            if ($val -match '^[0-9a-fA-F-]{36}$') { return $val }     # GUID
            if ($val -match '^tel:\+?\d') { return $val }            # tel:+E164
            if ($val -match '^\+?\d[\d\s\-()]{6,}$') { return $val } # phone-ish
        }
    }
    return $null
}

function Normalize-TargetType {
    param([string]$Type)
    switch -Regex ($Type) {
        '^ConfigurationEndpoint$' { return 'CallQueue' }
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
      Returns:
      @{
        RawType; Id; NormType; Friendly; AttachedType; AttachedName
      }
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
            $label = if ($id) { ($id -replace '^tel:','') } else { "" }
            return @{ RawType=$rawType; Id=$id; NormType=$type; Friendly=$label; AttachedType=""; AttachedName="" }
        }
        'CallQueue' {
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
            $friendly = if ($id) { $id } else { "" }
            return @{ RawType=$rawType; Id=$id; NormType='CallQueue'; Friendly=$friendly; AttachedType=""; AttachedName="" }
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
            if ($id -and -not $script:_ResolveCache.RAs.ContainsKey($id)) {
                $ra = _try { Get-CsOnlineApplicationInstance -Identity $id -ErrorAction Stop }
                if ($ra) { $script:_ResolveCache.RAs[$id] = $ra.DisplayName }
            }
            $friendly = if ($id) { ($script:_ResolveCache.RAs[$id] ?? $id) } else { "" }
            return @{ RawType=$rawType; Id=$id; NormType=$type; Friendly=$friendly; AttachedType=""; AttachedName="" }
        }
    }
}

# NEW: Fallback resolver when CallTarget.Type-based resolution isn't enough
function Resolve-AnyIdFriendly {
    <#
      .SYNOPSIS
      Best-effort resolver given only an Id string (CQ -> AA -> RA (+attachment) -> Group -> User -> PSTN).
      .RETURNS hashtable with keys: RawType, Id, NormType, Friendly, AttachedType, AttachedName
    #>
    param([string]$Id)

    function _try { param([scriptblock]$s) try { & $s } catch { $null } }

    if ([string]::IsNullOrWhiteSpace($Id)) {
        return @{ RawType=""; Id=""; NormType=""; Friendly=""; AttachedType=""; AttachedName="" }
    }

    # PSTN
    if ($Id -match '^tel:\+?\d' -or $Id -match '^\+?\d[\d\s\-()]{6,}$') {
        $num = ($Id -replace '^tel:','')
        return @{ RawType='PhoneNumber'; Id=$Id; NormType='Pstn'; Friendly=$num; AttachedType=""; AttachedName="" }
    }

    # CQ
    if ($script:_ResolveCache.CQsById.ContainsKey($Id)) {
        return @{ RawType='CallQueue'; Id=$Id; NormType='CallQueue'; Friendly=$script:_ResolveCache.CQsById[$Id]; AttachedType=""; AttachedName="" }
    }
    $cq = _try { Get-CsCallQueue -Identity $Id -ErrorAction Stop }
    if ($cq) {
        $script:_ResolveCache.CQsById[$Id] = $cq.Name
        return @{ RawType='CallQueue'; Id=$Id; NormType='CallQueue'; Friendly=$cq.Name; AttachedType=""; AttachedName="" }
    }

    # AA
    if ($script:_ResolveCache.AAsById.ContainsKey($Id)) {
        return @{ RawType='AutoAttendant'; Id=$Id; NormType='AutoAttendant'; Friendly=$script:_ResolveCache.AAsById[$Id]; AttachedType=""; AttachedName="" }
    }
    $aa = _try { Get-CsAutoAttendant -Identity $Id -ErrorAction Stop }
    if ($aa) {
        $script:_ResolveCache.AAsById[$Id] = $aa.Name
        return @{ RawType='AutoAttendant'; Id=$Id; NormType='AutoAttendant'; Friendly=$aa.Name; AttachedType=""; AttachedName="" }
    }

    # RA (and RA -> CQ/AA)
    if (-not $script:_ResolveCache.RAs.ContainsKey($Id)) {
        $ra = _try { Get-CsOnlineApplicationInstance -Identity $Id -ErrorAction Stop }
        if ($ra) { $script:_ResolveCache.RAs[$Id] = $ra.DisplayName }
    }
    if ($script:_ResolveCache.RAs.ContainsKey($Id)) {
        if ($script:_ResolveCache.CQByRA.ContainsKey($Id)) {
            $name = $script:_ResolveCache.CQByRA[$Id].Name
            return @{ RawType='ApplicationEndpoint'; Id=$Id; NormType='ResourceAccount'; Friendly=$name; AttachedType='CallQueue'; AttachedName=$name }
        }
        if ($script:_ResolveCache.AAByRA.ContainsKey($Id)) {
            $name = $script:_ResolveCache.AAByRA[$Id].Name
            return @{ RawType='ApplicationEndpoint'; Id=$Id; NormType='ResourceAccount'; Friendly=$name; AttachedType='AutoAttendant'; AttachedName=$name }
        }
        return @{ RawType='ApplicationEndpoint'; Id=$Id; NormType='ResourceAccount'; Friendly=$script:_ResolveCache.RAs[$Id]; AttachedType=""; AttachedName="" }
    }

    # Group (Team/M365 Group)
    if (-not $script:_ResolveCache.Groups.ContainsKey($Id)) {
        $gName = $null
        if (Get-Command Get-Team -ErrorAction SilentlyContinue) {
            $t = _try { Get-Team -GroupId $Id }
            if ($t) { $gName = $t.DisplayName }
        }
        if (-not $gName -and (Get-Command Get-UnifiedGroup -ErrorAction SilentlyContinue)) {
            $g = _try { Get-UnifiedGroup -Identity $Id -ErrorAction Stop }
            if ($g) { $gName = $g.DisplayName }
        }
        if (-not $gName -and (Get-Command Get-MgGroup -ErrorAction SilentlyContinue)) {
            $g = _try { Get-MgGroup -GroupId $Id -Property DisplayName }
            if ($g) { $gName = $g.DisplayName }
        }
        if ($gName) { $script:_ResolveCache.Groups[$Id] = $gName }
    }
    if ($script:_ResolveCache.Groups.ContainsKey($Id)) {
        return @{ RawType='Group'; Id=$Id; NormType='Group'; Friendly=$script:_ResolveCache.Groups[$Id]; AttachedType=""; AttachedName="" }
    }

    # User
    if (-not $script:_ResolveCache.Users.ContainsKey($Id)) {
        $u = _try { Get-CsOnlineUser -Identity $Id -ErrorAction Stop }
        if ($u) { $script:_ResolveCache.Users[$Id] = ($u.DisplayName ?? $u.UserPrincipalName ?? $Id) }
    }
    if ($script:_ResolveCache.Users.ContainsKey($Id)) {
        return @{ RawType='User'; Id=$Id; NormType='User'; Friendly=$script:_ResolveCache.Users[$Id]; AttachedType=""; AttachedName="" }
    }

    # Unknown
    return @{ RawType=''; Id=$Id; NormType='Unknown'; Friendly=$Id; AttachedType=""; AttachedName="" }
}

function Format-TargetBracket {
    param($Resolved)
    $humanType = Get-HumanTypeLabel -NormalizedType $Resolved.NormType
    $rawPart   = if ($Resolved.Id) { "$($Resolved.RawType) $($Resolved.Id)" } else { $Resolved.RawType }
    $attach = ""
    if ($Resolved.NormType -eq 'ResourceAccount' -and $Resolved.AttachedType) {
        $attachHuman = Get-HumanTypeLabel -NormalizedType $Resolved.AttachedType
        if ($Resolved.AttachedName) { $attach = "; attached to ${attachHuman}: $(HtmlEncode $Resolved.AttachedName)" }
        else { $attach = "; attached to ${attachHuman}" }
    }
    " <span style='color:#605E5C'>( $humanType; $(HtmlEncode $rawPart)$attach )</span>"
}

# ---------------------- Greetings and summaries -----------------------------
function Get-CallableEntitySummary { # (kept, but not used for Operator anymore)
    param($Entity)
    if (-not $Entity) { return "" }
    $t  = $Entity.Type
    $id = Get-IdFromCallTarget -CallTarget $Entity
    if ($t -and $id) { return "$(HtmlEncode $t) : $(HtmlEncode $id)" }
    if ($id) { return HtmlEncode $id }
    HtmlEncode (($Entity | Out-String).Trim())
}

function Get-PromptSummary {
    param($Prompt)
    if (-not $Prompt) { return "No greeting" }
    if ($Prompt.TextToSpeechPrompt) { return "TTS: $(HtmlEncode $Prompt.TextToSpeechPrompt)" }
    if ($Prompt.AudioFilePrompt) {
        $af = $Prompt.AudioFilePrompt
        $name = $af.FileName; if (-not $name) { $name = $af.Name }; if (-not $name) { $name = $af.Id }; if (-not $name) { $name = "Audio prompt (no name exposed)" }
        return "AudioFile: $(HtmlEncode $name)"
    }
    "Greeting present (unparsed)"
}

function Get-GreetingsSummary {
    param($Greetings)
    if (-not $Greetings -or $Greetings.Count -eq 0) { return "No greeting" }
    ($Greetings | ForEach-Object { Get-PromptSummary -Prompt $_ }) -join "<br/>"
}

function Get-MenuPromptsSummary {
    param($Menu)
    if (-not $Menu -or -not $Menu.Prompts) { return "Greeting: None" }
    $p = $Menu.Prompts
    if ($p.ActiveType -eq 'TextToSpeech' -and $p.TextToSpeechPrompt) { return "Greeting: TTS: $(HtmlEncode $p.TextToSpeechPrompt)" }
    if ($p.ActiveType -eq 'AudioFile' -and $p.AudioFilePrompt) { return "Greeting: AudioFile" }
    if ($p.TextToSpeechPrompt) { return "Greeting: TTS: $(HtmlEncode $p.TextToSpeechPrompt)" }
    if ($p.AudioFilePrompt)   { return "Greeting: AudioFile" }
    "Greeting: None"
}

# Helper to build a friendly+bracket string from CallTarget with fallback resolver
function Build-FriendlyBlockFromCallTarget {
    param($CallTarget)

    if (-not $CallTarget) { return @{ Friendly=""; Bracket=""; RawLabel="" ; Res=$null } }

    $res = Resolve-CallTargetFriendly -CallTarget $CallTarget
    $id  = Get-IdFromCallTarget -CallTarget $CallTarget

    # Fallback resolution if friendly is blank or equals Id
    if (-not $res.Friendly -or ($id -and $res.Friendly -eq $id)) {
        $alt = Resolve-AnyIdFriendly -Id $id
        if ($alt.Friendly) { $res = $alt }
    }

    $friendly = "$(HtmlEncode $res.Friendly)"
    $bracket  = Format-TargetBracket -Resolved $res

    $rawLabel = ""
    if ($id) { $rawLabel = "$(HtmlEncode $CallTarget.Type) $(HtmlEncode $id)" }
    else { $rawLabel = "$(HtmlEncode $CallTarget.Type)" }

    @{ Friendly=$friendly; Bracket=$bracket; RawLabel=$rawLabel; Res=$res }
}

# ----------------------------- Call flow classification -----------------------------
function Get-CallFlowModeSummary {
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
            'DisconnectCall' { return "<b>Greeting:</b> $greeting<br/><b>Action:</b> Disconnect" }
            'TransferCallToTarget' {
                if ($o.CallTarget) {
                    $blk = Build-FriendlyBlockFromCallTarget -CallTarget $o.CallTarget
                    $targetPart = ""
                    if ($blk.Friendly) {
                        if ($blk.RawLabel -and ($blk.RawLabel -ne $blk.Friendly)) {
                            $targetPart = "<br/><b>Target:</b> $($blk.Friendly) <span style='color:#605E5C'>( $($blk.RawLabel) )</span>$($blk.Bracket)"
                        } else {
                            $targetPart = "<br/><b>Target:</b> $($blk.Friendly)$($blk.Bracket)"
                        }
                    } elseif ($blk.RawLabel) {
                        $targetPart = "<br/><b>Target:</b> $($blk.RawLabel)$($blk.Bracket)"
                    }
                    return "<b>Greeting:</b> $greeting<br/><b>Action:</b> Redirect$targetPart"
                } else {
                    return "<b>Greeting:</b> $greeting<br/><b>Action:</b> Redirect (no target returned)"
                }
            }
            default {
                $label = switch -Regex ($o.Action) { "Announcement" { "Announcement" } default { [string]$o.Action } }
                return "<b>Greeting:</b> $greeting<br/><b>Action:</b> $(HtmlEncode $label)"
            }
        }
    }

    if ($keys.Count -gt 0) {
        return "<b>Greeting:</b> $greeting<br/><b>Action:</b> Menu"
    }

    # Fallback: multiple Automatic
    $first = $auto | Select-Object -First 1
    if ($first) {
        $act = switch -Regex ($first.Action) { "Disconnect" { "Disconnect" } "Transfer" { "Redirect" } "Announcement" { "Announcement" } default { [string]$first.Action } }
        if ($act -eq 'Redirect' -and $first.CallTarget) {
            $blk = Build-FriendlyBlockFromCallTarget -CallTarget $first.CallTarget
            $targetPart = ""
            if ($blk.Friendly) {
                if ($blk.RawLabel -and ($blk.RawLabel -ne $blk.Friendly)) {
                    $targetPart = "<br/><b>Target:</b> $($blk.Friendly) <span style='color:#605E5C'>( $($blk.RawLabel) )</span>$($blk.Bracket)"
                } else {
                    $targetPart = "<br/><b>Target:</b> $($blk.Friendly)$($blk.Bracket)"
                }
            } elseif ($blk.RawLabel) { $targetPart = "<br/><b>Target:</b> $($blk.RawLabel)$($blk.Bracket)" }
            return "<b>Greeting:</b> $greeting<br/><b>Action:</b> $act$targetPart"
        } else {
            return "<b>Greeting:</b> $greeting<br/><b>Action:</b> $act"
        }
    }

    "<b>Greeting:</b> $greeting<br/><b>Action:</b> Unknown"
}

# ----------------------------- Misc formatting helpers -----------------------------
function Get-FixedScheduleRangesText {
    param($Schedule)
    if (-not $Schedule) { return "" }
    if ($Schedule.Type -ne "Fixed") { return "" }
    if (-not $Schedule.FixedSchedule) { return "" }

    $ranges = $Schedule.FixedSchedule.DateTimeRanges
    if (-not $ranges) { return "" }

    $parts = @()
    foreach ($r in @($ranges)) {
        if ($null -eq $r) { continue }
        $start = $r.Start; $end = $r.End
        $sTxt = if ($start -is [datetime]) { $start.ToString("dd/MM/yyyy HH:mm") } else { [string]$start }
        $eTxt = if ($end   -is [datetime]) { $end.ToString("dd/MM/yyyy HH:mm") } else { [string]$end }
        $parts += "$sTxt-$eTxt"
    }
    if ($parts.Count -eq 0) { return "" }
    HtmlEncode ($parts -join "; ")
}

function Get-WeeklyScheduleText {
    param($Weekly)
    if (-not $Weekly) { return "Schedule present (not weekly recurrent)" }
    $days = @("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($d in $days) {
        $prop  = "${d}Hours"
        $hours = $Weekly.$prop
        if ($hours -and $hours.Count -gt 0) {
            $ranges = foreach ($h in $hours) {
                $s = [string]$h.Start; $e = [string]$h.End
                "$s–$e"
            }
            [void]$lines.Add("<b>${d}:</b> $(HtmlEncode (($ranges -join ', ')))")
        }
    }
    if ($lines.Count -eq 0) { return "Not configured" }
    ($lines -join "<br/>")
}

function Get-BusinessHoursText {
    param(
        [Parameter(Mandatory)] $AA,
        [Parameter(Mandatory)] $SchedById
    )

    if (-not $AA) { return "Not configured" }

    $bhAssoc = $AA.CallHandlingAssociations | Where-Object { $_.Type -eq "BusinessHours" } | Select-Object -First 1
    if ($bhAssoc) {
        $bhSched = $SchedById[[string]$bhAssoc.ScheduleId]
        if ($bhSched -and $bhSched.WeeklyRecurrentSchedule) { return Get-WeeklyScheduleText -Weekly $bhSched.WeeklyRecurrentSchedule }
        elseif ($bhSched) { return "Schedule present (not weekly recurrent)" }
        else { return "Business-hours schedule not found in AA object" }
    }

    $weeklyCandidates = @($AA.Schedules | Where-Object { $_.WeeklyRecurrentSchedule })
    $nameHint = $weeklyCandidates | Where-Object { $_.Name -match '(business|work|working|hours)' } | Select-Object -First 1
    if ($nameHint) { return Get-WeeklyScheduleText -Weekly $nameHint.WeeklyRecurrentSchedule }

    $firstWeekly = $weeklyCandidates | Select-Object -First 1
    if ($firstWeekly) { return Get-WeeklyScheduleText -Weekly $firstWeekly.WeeklyRecurrentSchedule }

    "Not configured"
}

function Get-HolidayLines {
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

        $rangeText = ""; if ($hs) { $rangeText = Get-FixedScheduleRangesText -Schedule $hs }; if (-not $rangeText) { $rangeText = "Dates not found" }

        $greetingText = "No greeting"; if ($hf) { $greetingText = Get-GreetingsSummary -Greetings $hf.Greetings }

        $at = @{ Action="Unknown"; Target=""; TargetFriendly="" }
        if ($hf -and $hf.Menu) { $at = Get-ActionAndTargetFromMenu -Menu $hf.Menu }

        $targetPart = ""
        if ($at.TargetFriendly) { $targetPart = ", <b>Target:</b> $($at.TargetFriendly)" }
        elseif ($at.Target)     { $targetPart = ", <b>Target:</b> $($at.Target)" }

        "<b>$(HtmlEncode $holidayName):</b> $rangeText, <b>Greeting:</b> $greetingText, <b>Action:</b> $(HtmlEncode $at.Action)$targetPart"
    }

    $lines -join "<br/>"
}

function Get-ActionAndTargetFromMenu {
    param($Menu)

    if (-not $Menu -or -not $Menu.MenuOptions -or $Menu.MenuOptions.Count -eq 0) {
        return @{ Action = "No menu"; Target = ""; TargetFriendly = "" }
    }

    $opt =
        $Menu.MenuOptions | Where-Object { $_.DtmfResponse -eq "Automatic" -and $_.Action -match 'Transfer' -and $_.CallTarget } | Select-Object -First 1
    if (-not $opt) { $opt = $Menu.MenuOptions | Where-Object { $_.DtmfResponse -eq "Automatic" -and $_.CallTarget } | Select-Object -First 1 }
    if (-not $opt) { $opt = $Menu.MenuOptions | Where-Object { $_.CallTarget } | Select-Object -First 1 }
    if (-not $opt) { $opt = $Menu.MenuOptions | Select-Object -First 1 }

    $act = $opt.Action; if (-not $act) { $act = "Unknown" }

    $actionLabel = $act
    switch -Regex ($act) { "Disconnect" { $actionLabel = "Disconnect" }; "Transfer" { $actionLabel = "Forward/Transfer" }; "Announcement" { $actionLabel = "Announcement" } }

    $blk = $null
    if ($opt.CallTarget) { $blk = Build-FriendlyBlockFromCallTarget -CallTarget $opt.CallTarget }

    $targetLabel = ""; $targetFriendly = ""
    if ($blk) {
        $targetFriendly = $blk.Friendly
        $targetLabel = if ($blk.RawLabel) { $blk.RawLabel } else { "" }

        if ($targetFriendly) {
            $targetFriendly = "$targetFriendly$($blk.Bracket)"
        } elseif ($targetLabel) {
            $targetLabel = "$targetLabel$($blk.Bracket)"
        }
    }

    @{ Action = $actionLabel; Target = $targetLabel; TargetFriendly = $targetFriendly }
}

function Get-DefaultCallFlowCleanSummary {
    param($CallFlow)
    if (-not $CallFlow) { return "Not configured" }

    $greeting = Get-GreetingsSummary -Greetings $CallFlow.Greetings
    $at = Get-ActionAndTargetFromMenu -Menu $CallFlow.Menu

    $targetPart = ""
    if ($at.TargetFriendly) { $targetPart = "<br/><b>Target:</b> $($at.TargetFriendly)" }
    elseif ($at.Target)     { $targetPart = "<br/><b>Target:</b> $($at.Target)" }

    "<b>Greeting:</b> $greeting<br/><b>Action:</b> $(HtmlEncode $at.Action)$targetPart"
}

function Get-DefaultMenuOptionsLines {
    param($Menu)

    if (-not $Menu -or -not $Menu.MenuOptions -or $Menu.MenuOptions.Count -eq 0) { return "No options configured" }

    function Get-DtmfLabel([string]$d) {
        switch ($d) {
            "Tone0" { "0" } "Tone1" { "1" } "Tone2" { "2" } "Tone3" { "3" } "Tone4" { "4" }
            "Tone5" { "5" } "Tone6" { "6" } "Tone7" { "7" } "Tone8" { "8" } "Tone9" { "9" }
            "Star"  { "*" } "Pound" { "#" } "Automatic" { "Auto" } default { $d }
        }
    }

    $lines = foreach ($opt in $Menu.MenuOptions) {
        $dtmf = Get-DtmfLabel -d ([string]$opt.DtmfResponse)
        $act  = if ($opt.Action) { [string]$opt.Action } else { "Unknown" }

        $actionLabel = $act
        switch -Regex ($act) { "Disconnect" { $actionLabel = "Disconnect" } "Transfer" { $actionLabel = "Forward/Transfer" } "Announcement" { $actionLabel = "Announcement" } }

        $blk = $null; if ($opt.CallTarget) { $blk = Build-FriendlyBlockFromCallTarget -CallTarget $opt.CallTarget }
        $targetLabel = ""; $targetFriendly = ""
        if ($blk) {
            $targetFriendly = $blk.Friendly
            $targetLabel    = if ($blk.RawLabel) { $blk.RawLabel } else { "" }
            if ($targetFriendly) { $targetFriendly = "$targetFriendly$($blk.Bracket)" }
            elseif ($targetLabel) { $targetLabel = "$targetLabel$($blk.Bracket)" }
        }

        $targetPart = ""
        if ($targetFriendly) {
            if ($targetLabel -and ($targetLabel -ne $targetFriendly)) {
                $targetPart = ", <b>Target:</b> $targetFriendly <span style='color:#605E5C'>( $targetLabel )</span>"
            } else { $targetPart = ", <b>Target:</b> $targetFriendly" }
        } elseif ($targetLabel) { $targetPart = ", <b>Target:</b> $targetLabel" }

        "<b>Option ${dtmf}:</b> <b>Action:</b> $(HtmlEncode $actionLabel)$targetPart"
    }

    $lines -join "<br/>"
}

# ----------------------------- Authorized users helper (DisplayName-first) -----------------------------
function Resolve-AuthorizedUsers {
    param([string[]]$ObjectIds, [hashtable]$Cache)

    if (-not $ObjectIds -or $ObjectIds.Count -eq 0) { return @() }
    if (-not $Cache) { $Cache = @{} }

    foreach ($id in $ObjectIds) {
        $idStr = [string]$id
        if ($Cache.ContainsKey($idStr)) { $Cache[$idStr]; continue }

        $label = $null
        try {
            $u = Get-CsOnlineUser -Identity $idStr -ErrorAction Stop
            if ($u) {
                if ($u.DisplayName)       { $label = $u.DisplayName }
                elseif ($u.UserPrincipalName) { $label = $u.UserPrincipalName }
                elseif ($u.SipAddress)    { $label = $u.SipAddress }
            }
        } catch {
            try {
                if (Get-Command Get-MgUser -ErrorAction SilentlyContinue) {
                    $g = Get-MgUser -UserId $idStr -Property DisplayName,UserPrincipalName -ErrorAction Stop
                    if ($g) {
                        if     ($g.DisplayName)       { $label = $g.DisplayName }
                        elseif ($g.UserPrincipalName) { $label = $g.UserPrincipalName }
                    }
                }
            } catch { }
        }

        if (-not $label) { $label = $idStr }
        $Cache[$idStr] = $label
        $label
    }
}

# ----------------------------- HTML table builder -----------------------------
function New-ParamValueTableHtml {
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
        $v = [string]$r.Value
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

        $voiceIdVal = ""; if ($aa.VoiceId) { $voiceIdVal = $aa.VoiceId }
        $rows.Add([pscustomobject]@{ Parameter="Voice / TTS voice"; Value=HtmlEncode $voiceIdVal })

        # ---- Operator (friendly + brackets) ----
        $opVal = "None"
        if ($aa.Operator) {
            $blk = Build-FriendlyBlockFromCallTarget -CallTarget $aa.Operator
            if ($blk.Friendly)     { $opVal = "$($blk.Friendly)$($blk.Bracket)" }
            elseif ($blk.RawLabel) { $opVal = "$($blk.RawLabel)$($blk.Bracket)" }
        }
        $rows.Add([pscustomobject]@{ Parameter="Operator"; Value=$opVal })

        # ---- Business hours schedule (text) ----
        $rows.Add([pscustomobject]@{ Parameter="Business hours"; Value=(Get-BusinessHoursText -AA $aa -SchedById $schedById) })

        # ---- Business hours call handling (Menu / Disconnect / Redirect) ----
        $rows.Add([pscustomobject]@{ Parameter = "Business hours call handling"; Value = (Get-CallFlowModeSummary -CallFlow $aa.DefaultCallFlow) })

        # ---- Default menu greeting ----
        if ($aa.DefaultCallFlow -and $aa.DefaultCallFlow.Menu) {
            $rows.Add([pscustomobject]@{ Parameter = "Default menu greeting"; Value = (Get-MenuPromptsSummary -Menu $aa.DefaultCallFlow.Menu) })
        } else {
            $rows.Add([pscustomobject]@{ Parameter="Default menu greeting"; Value="Greeting: None" })
        }

        # ---- Default call flow options ----
        if ($aa.DefaultCallFlow -and $aa.DefaultCallFlow.Menu) {
            $rows.Add([pscustomobject]@{ Parameter = "Default call flow options"; Value = (Get-DefaultMenuOptionsLines -Menu $aa.DefaultCallFlow.Menu) })
        } else {
            $rows.Add([pscustomobject]@{ Parameter="Default call flow options"; Value="No options configured" })
        }

        # ---- After-hours call handling (classified) ----
        $afterAssoc = @($aa.CallHandlingAssociations | Where-Object { $_.Type -eq "AfterHours" } | Select-Object -First 1)
        if ($afterAssoc) {
            $afterFlow = $cfById[[string]$afterAssoc.CallFlowId]
            if ($afterFlow) {
                $rows.Add([pscustomobject]@{ Parameter = "After-hours call handling"; Value = (Get-CallFlowModeSummary -CallFlow $afterFlow) })
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

        # ---- Authorized users (DisplayName-first) ----
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
    table-layout: auto;
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

td:first-child, th:first-child { width: 220px; white-space: nowrap; }
td:last-child, th:last-child   { width: auto; }
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

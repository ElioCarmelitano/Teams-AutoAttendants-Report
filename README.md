Connect to Teams and ExchangeOnline PowerShell Modules before running.

# Teams-AutoAttendants-Report
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

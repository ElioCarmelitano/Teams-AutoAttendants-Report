Connect to Teams and ExchangeOnline PowerShell Modules before running.

# Document-TeamsAutoAttendants
.SYNOPSIS
Generates an HTML documentation report for all Microsoft Teams Auto Attendants.

.DESCRIPTION
Read-only documentation script that enumerates all Microsoft Teams Auto Attendants (AAs) and produces
a structured HTML report. The report includes resolved (friendly) names for common targets/agents where
possible, plus audit-friendly bracket details (raw types/IDs and Resource Account attachments).

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
 - Authorized users (friendly DisplayName where possible)

TARGET RESOLUTION (best effort)
 - PSTN: strips 'tel:' and displays E.164 where possible
 - Call Queue: CQ name by Id (handles ConfigurationEndpoint)
 - Auto Attendant: AA name by Id
 - Resource Account (ApplicationInstance): RA display name and, if mapped, attached CQ/AA name
 - Group (Shared Voicemail / Team / Channel): Team/M365 Group display name where possible
 - User: DisplayName (falls back to UPN/Id)
 - Otherwise: shows the raw Id as “Unknown”

REQUIREMENTS / PREREQUISITES
 - Teams/Skype Online PowerShell session with permissions to call:
     * Get-CsAutoAttendant
     * Get-CsCallQueue (for target resolution)
     * Get-CsOnlineApplicationInstance (resource accounts)
     * Get-CsOnlineUser (authorized users)
 - Optional modules/commands (used opportunistically if available to enrich group display names):
     * MicrosoftTeams: Get-Team
     * ExchangeOnlineManagement: Get-UnifiedGroup
     * Microsoft.Graph: Get-MgGroup

.OUTPUTS
Writes an HTML report to disk. The script emits a single success message with the output path,
and optionally opens the report in the default browser.

.PARAMETER OutputPath
Full or relative path to the HTML report file.
Default: .\Document-TeamsAutoAttendants-<timestamp>.html

.PARAMETER Open
If specified, opens the generated HTML report when complete.

.PARAMETER IncludeAfterHoursOptions
Reserved for future use. Intended to list after-hours menu options similar to the default call flow.
Currently unused; retained as a placeholder for easy extension.

.EXAMPLE
PS> .\Document-TeamsAutoAttendants.ps1
Generates ".\Document-TeamsAutoAttendants-YYYYMMDD-HHmmss.html".

.EXAMPLE
PS> .\Document-TeamsAutoAttendants.ps1 -OutputPath "C:\Temp\AA-Report.html" -Open
Writes the report to C:\Temp\AA-Report.html and opens it in the default browser.

.NOTES
 - Parser-safety: Hashtables use ';' between key/value pairs. Complex values are built in variables first.
 - Null-safety: .ContainsKey / indexing operations guard against null/empty keys where applicable.
 - This script is read-only; it does not modify any tenant configuration.
 - If optional modules (Teams/EXO/Graph) are not installed or not connected, group resolution will fall back to IDs.

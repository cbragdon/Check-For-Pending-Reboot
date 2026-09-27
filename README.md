# Check-For-Pending-Reboot

A PowerShell script (`Multi_Server_Pending_Reboot_1.3.3.ps1`) that remotely checks
one or more Windows Servers for pending-reboot conditions across every well-known
detection vector, reports the Windows Server OS version each check ran against,
and provides console, CSV/HTML export, and email reporting options.

## Supported OS range

Windows Server 2016, 2019, 2022, and 2025. All detection vectors use registry/WMI
locations that are unchanged across this entire range, so no OS-version-specific
branching is required in the detection logic — the OS is identified per server
purely for reporting context (`OSCaption`, `OSVersion`, `OSBuildNumber`).

## Detection vectors

| Result field | Source | What it means |
|---|---|---|
| `CBS_RebootPending` | `HKLM:\...\Component Based Servicing\RebootPending` | A component/role update is staged and requires a reboot to finish. |
| `CBS_PackagesPending` | `HKLM:\...\Component Based Servicing\PackagesPending` | One or more servicing packages are still mid-installation. |
| `CBS_RebootInProgress` | `HKLM:\...\Component Based Servicing\RebootInProgress` | The servicing stack has marked a reboot as already staged/underway. |
| `WUAU_RebootRequired` | `HKLM:\...\WindowsUpdate\Auto Update\RebootRequired` | Windows Update has installed updates awaiting a reboot to finalize. |
| `PendingFileRenameOperations_Exist` / `_Detail` | `HKLM:\SYSTEM\...\Session Manager\PendingFileRenameOperations` | Files in use during patching are queued for rename/delete at next boot. Filtered detail table shows only operationally significant entries (see [File rename/delete FAQ](#faq-does-a-file-marked-for-deletion-affect-system-behavior-before-the-reboot) below). |
| `PendingComputerRename` | `HKLM:\SYSTEM\...\ComputerName\ActiveComputerName` vs `ComputerName` | A computer rename has been requested but not yet applied. |
| `PendingDomainJoin` | `HKLM:\SYSTEM\...\Services\Netlogon` (`JoinDomain` / `AvoidSpnSet`) | A domain join/leave operation is staged and completes on next reboot. |
| `CCM_RebootPending_WMI` | `root\ccm\clientsdk:CCM_ClientUtilities` (if SCCM/MECM client installed) | SCCM client has signaled a reboot is required. |
| `RebootPending_Overall` | — | `$true` if any vector above is flagged. |
| `RecentHotfixes` | `Get-HotFix` (only populated when PFRO is flagged) | Best-effort list of the 5 most recently installed hotfixes/CUs, as candidate causes for the pending file operations — a timing correlation, **not** a guaranteed package link. |

## Usage

Configuration lives in the `USAGE` section at the bottom of the script:

```powershell
$ShowPendingFiles         = 0        # 1 = show filtered Rename/Delete file detail table
$OnlyShowFlagged          = 0        # 1 = only report servers with RebootPending_Overall = $true
$ComputerListPath         = ''       # optional path to a server list file (.txt one-per-line, or .csv with ComputerName column)
$DefaultComputerList      = @(...)   # fallback list used when $ComputerListPath is blank/missing
$RetryCount               = 2
$RetryDelaySeconds        = 5
$ConnectionTimeoutSeconds = 15
$ExportResults            = 0        # 1 = export CSV/HTML to $ExportFolder
$ExportFolder             = "C:\temp\Check-For-Pending-Reboot\Reports"
$ExportFormat             = 'Both'   # 'CSV', 'HTML', or 'Both'
$SendEmail                = 0        # 1 = email the report (requires SmtpServer/EmailFrom/EmailTo)
```

Run the script directly (`.\Multi_Server_Pending_Reboot_1.3.3.ps1`) with an account
that has WinRM/PSRemoting rights on the target servers.

## Requirements

- PowerShell 3.0+ (`#Requires -version 3`)
- WinRM/PowerShell Remoting enabled on target servers, with the executing account
  having remote management rights
- Optional: SCCM/MECM client installed on targets for the `CCM_RebootPending_WMI` check
- Optional: SMTP relay access for the email report feature

## FAQ

### Does a file marked for deletion affect system behavior before the reboot?

**No meaningful change occurs until the reboot processes the delete** — the file
just sits there, fully intact and functioning exactly as before.

Why:
- A file gets queued for delete (rather than deleted immediately) specifically
  *because* it's currently locked/in-use — some process already has it open in
  memory.
- That process keeps running against the version it already loaded. Deleting the
  file on disk wouldn't have changed what's in that process's memory anyway
  (Windows file locking prevents in-place deletion of open files, but doesn't
  unmap them from memory).
- Any *new* process that opens that same file path will still get the old,
  pre-delete version — the delete hasn't happened yet, so nothing has changed
  from the file system's perspective.
- No disk space is freed, no directory listing changes — it's a completely inert,
  no-op state until boot.

Practical implications:
- If the delete is part of an in-place upgrade/patch cleanup (e.g., removing a
  superseded DLL), the system is technically still running the **old** version
  of whatever component that file represents — it just hasn't been *removed*,
  so nothing is "half-updated" or broken. It's simply stale.
- The only real-world impact is: the patch/cleanup isn't fully complete, and
  depending on what it is (e.g., an old vulnerable binary that's part of a
  security fix's cleanup step), that old file could theoretically still be
  loaded by a new process until the reboot clears it out — worth checking if the
  CU release notes call out that specific file as security-relevant.
- Bottom line: functionally **inert until reboot**, not actively harmful, just
  incomplete.

### Is it best practice to reboot to clear files marked for deletion?

**Yes** — treating a pending reboot as something to remediate promptly (next
maintenance window) rather than deferring indefinitely is standard practice.

Reasoning:
- Patches aren't fully applied until the reboot completes the queued
  rename/delete — if it's part of a security fix, the vulnerability isn't
  actually remediated until the swap finishes (Microsoft's own reboot-pending
  detection, e.g. `MsiSystemRebootPending`, exists precisely because installers
  need to know work is unfinished).
- Windows Update itself will hold back installing further updates that require
  a restart until a pending restart is cleared (see Microsoft's "One restart a
  month for Windows updates" policy below) — so deferring reboots can delay
  your *next* patch cycle, not just the current one.
- `PendingFileRenameOperations` doesn't always mean something is broken while
  pending (see above — it's inert until reboot), but persistent/accumulating
  entries across multiple patch cycles are a sign of an unfinished, stacking
  servicing state that's best cleared rather than left to grow.
- Community/vendor guidance (Ivanti, Revenera) notes some entries can be benign
  (e.g., app log-rotation cleanup) — so not every `PendingFileRenameOperations`
  entry is critical, but system-file entries from CBS/WU cumulative updates are
  the ones worth prioritizing.

### Can I determine if a pending reboot is caused by a specific hotfix/CU?

`PendingFileRenameOperations` itself does **not** record which package queued a
file. For a definitive answer:
- Grep `C:\Windows\Logs\CBS\CBS.log` on the target server for the file name —
  this is the authoritative source for confirming which KB queued a given file.
- Cross-reference by timing: `Get-HotFix`, Windows Update history via the
  `Microsoft.Update.Session` COM object, or `DISM /Get-PackageInfo` install dates.
- WinSxS folder naming in the PFRO source path can be mapped back to a KB via
  `Get-WindowsPackage -Online`.

This script surfaces the 5 most recently installed hotfixes (`RecentHotfixes`)
as a best-effort timing correlation when file-rename operations are flagged —
not a guaranteed match.

## Sources

- Microsoft — [`MsiSystemRebootPending` (Win32 API docs)](https://learn.microsoft.com/en-us/windows/win32/msi/msisystemrebootpending)
- Microsoft Support — [One restart a month for Windows updates (KB5121772)](https://support.microsoft.com/en-us/servicing/os/windows/docs/2026/07/kb5121772-one-restart-a-month-for-windows-updates)
- Microsoft Support — [Troubleshoot problems updating Windows](https://support.microsoft.com/en-us/windows/deployment/updates-lifecycle/troubleshoot-problems-updating-windows)
- Microsoft Q&A — [Question regarding PendingFileRenameOperations caused by Microsoft Edge](https://learn.microsoft.com/en-gb/answers/questions/5981763/question-regarding-pendingfilerenameoperations-cau)
- Ivanti — [Troubleshoot Persistent PendingFileRenameOperations In Registry](https://hub.ivanti.com/s/article/Troubleshoot-Persistent-PendingFileRenameOperations-In-Registry)
- Revenera — [PendingFileRenameOperations Versus MsiSystemRebootPending](https://community.revenera.com/s/article/pendingfilerenameoperations-versus-msisystemrebootpending)

## License

[MIT](LICENSE)

## Additional history

See [`Session-Notes.md`](Session-Notes.md) for a detailed development log of how
this script evolved (retry/timeout handling, server-list loading, export/email
reporting, hotfix correlation, OS identification, and expanded detection vectors).

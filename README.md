# Check-For-Pending-Reboot

## Why this exists

If you're managing more than a couple of Windows Servers, you already know the
pain: patches get installed, but the server won't actually be "fully patched"
until it reboots. That reboot doesn't always happen right away — and there's
no single, obvious place to check whether one is waiting. Windows tracks
"pending reboot" state across several different, unrelated mechanisms
(Windows Update, the component servicing stack, in-use files queued for
rename/delete, computer renames, domain joins, and — if you use SCCM/MECM —
the configuration manager client). Missing any one of these means you might
think a server is clean when it's actually still carrying an unfinished patch,
an unapplied rename, or a stale/vulnerable file sitting on disk waiting to be
replaced.

Checking all of that by hand, one server at a time, doesn't scale. This script
solves that by remotely checking **every server you care about, against every
known pending-reboot indicator, in one pass**, and giving you a clear,
color-coded answer per server instead of a pile of registry paths to
cross-reference yourself. It also tells you *which* Windows Server version
each machine is running (2016 through 2025), and — when it finds files queued
for rename/delete — gives you a filtered, readable view of exactly which files
are involved and which recently-installed hotfixes might be responsible, so
you're not left guessing why a reboot is being requested.

In short: **one script, one run, a straight answer for every server on
whether it needs a reboot and why** — plus optional CSV/HTML reports and
email delivery if you want to hand this off to a recurring job.

## What it does

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
| `LastBootUpTime` | `Win32_OperatingSystem` | The server's last boot time — used to determine whether a pending reboot is new or leftover (see [FAQ](#faq-why-does-the-server-still-show-a-pending-reboot-right-after-i-already-rebooted-it) below). |
| `CBSLogLastWriteTime` / `CBSLogNewerThanBoot` | `C:\Windows\Logs\CBS\CBS.log` (file metadata) | Last-write time of CBS.log, and whether that's newer than `LastBootUpTime`. Distinguishes a fresh post-reboot pending state from a stale/leftover one that a reboot alone won't clear. |

### When a server can't be reached

If a server can't be connected to after retries, `ConnectionError` is set to
`$true` and `ConnectionErrorMessage` captures the actual error text. In that
case every detection field above (`RebootPending_Overall`, `CBS_*`,
`WUAU_RebootRequired`, `PendingFileRenameOperations_Exist`/`_Detail`,
`RecentHotfixes`, `PendingComputerRename`, `PendingDomainJoin`,
`CCM_RebootPending_WMI`, `LastBootUpTime`, `CBSLogLastWriteTime`,
`CBSLogNewerThanBoot`) is set to the string `"N/A"` rather than `$false`, and
`OSCaption`/`OSVersion`/`OSBuildNumber` are set to `"Unknown"` — so a failed
connection is never visually mistaken for a clean/false result. Failed
connections are also always listed in a dedicated "CONNECTION FAILURES"
section printed after the summary table.

## Usage

Configuration lives in the `USAGE` section at the bottom of the script:

```powershell
$ShowPendingFiles         = 0        # 1 = show filtered Rename/Delete file detail table
$OnlyShowFlagged          = 0        # 1 = only report servers with RebootPending_Overall = $true
$ServerNameListPath       = ''       # optional path to a server list file (.txt one-per-line, or .csv with ComputerName column)
$ComputerName             = @(...)   # fallback list used when $ServerNameListPath is blank/missing ($ServerNameListPath and $ComputerName are mutually exclusive — only one is used per run)
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

### Loading the functions without running the config/execution section

Dot-source the script to load every function (`Test-PendingReboot`,
`Write-PendingRebootSummary`, etc.) into your current session without
triggering the `USAGE` section's config/execution at the bottom — handy for
interactive testing or calling individual functions ad hoc:

```powershell
. .\Multi_Server_Pending_Reboot_1.3.3.ps1 2>$null

Test-PendingReboot -ComputerName "SERVER01", "SERVER02" | Format-List *
```

The script detects dot-sourcing (`$MyInvocation.InvocationName -eq '.'`) and
returns immediately after the function definitions, so `$ServerNameListPath`/
`$ComputerName` are never evaluated and the "no servers configured" guard
won't fire. Running the file normally (`.\Multi_Server_Pending_Reboot_1.3.3.ps1`)
is unaffected and still executes the full config/execution section.

## Requirements

- PowerShell 3.0+ (`#Requires -version 3`)
- WinRM/PowerShell Remoting enabled on target servers, with the executing account
  having remote management rights
- Optional: SCCM/MECM client installed on targets for the `CCM_RebootPending_WMI` check
- Optional: SMTP relay access for the email report feature

## FAQ

### If multiple fixes/hotfixes are chained/queued for installation, does one reboot satisfy all of them?

**Yes, generally one reboot clears the entire queue** — you don't need a
separate reboot per KB. This is exactly why `PendingFileRenameOperations` and
the CBS `RebootPending` state can (and often do) accumulate entries from
several different patches at once: Windows doesn't process each update's
file-swap/cleanup the moment it's installed if a reboot hasn't happened yet —
it just keeps staging more entries onto the same queue. The next reboot walks
the entire queue and finishes everything in it, regardless of how many
separate updates contributed to it.

This is also the underlying idea behind Microsoft's newer
["One restart a month"](https://support.microsoft.com/en-us/servicing/os/windows/docs/2026/07/kb5121772-one-restart-a-month-for-windows-updates)
behavior — Windows now deliberately holds driver, .NET, and firmware updates
back and installs them together with the monthly security update so they all
clear with a single restart, instead of prompting for a reboot after each one.

**One important exception:** Servicing Stack Updates (SSUs). Microsoft
recommends installing the latest SSU *before* installing a Latest Cumulative
Update (LCU) — the SSU updates the component responsible for installing
updates in the first place. If an SSU is involved, it can effectively need to
be applied (and sometimes rebooted) ahead of the other updates in the chain
rather than being satisfied by the same single reboot as everything else. See
Microsoft's
[Servicing Stack Updates (SSU): Frequently Asked Questions](https://support.microsoft.com/en-us/servicing/os/windows/2019/12/servicing-stack-updates-ssu-frequently-asked-questions)
for details.

**Practical takeaway for this script:** if `RebootPending_Overall` is flagged
by multiple vectors at once (e.g., both `CBS_RebootPending` and
`WUAU_RebootRequired`, or several `PendingFileRenameOperations_Detail`
entries tied to different KBs per `RecentHotfixes`), you almost always only
need to reboot the server **once** to clear all of them — you don't need to
reboot once per detected vector or once per pending hotfix.

### What if there's a mixture of Windows updates, SQL Server patches, and other products?

**Same answer: one clean reboot generally clears all of it.** Windows
Update, SQL Server setup, and most other MSI/MSP-based installers all funnel
their reboot requirement through the same OS-level mechanisms this script
already checks — `PendingFileRenameOperations`, the CBS `RebootPending` key,
and the WUAU `RebootRequired` key. It doesn't matter which product staged an
entry into that queue; a single reboot processes the whole queue regardless
of source.

**Exceptions worth knowing about:**
- **A failed or incomplete install** (e.g., SQL Server setup errored
  mid-patch) can leave stale/orphaned entries that survive the reboot —
  that's a genuine "still pending" problem, not just normal queued work. See
  the next FAQ entry for how to tell the difference.
- **Clustered SQL Server (AlwaysOn Availability Groups / FCI)** sometimes
  requires node-by-node reboots in sequence rather than one shared reboot,
  since only one node can be down at a time without impacting availability.
- If **new patches get staged during or right after the reboot** (an agent
  auto-install, SCCM re-triggering a install cycle at startup, etc.), that's
  a legitimately *new* pending state showing up — not the old one failing to
  clear.

### Why does the server still show a pending reboot right after I already rebooted it?

This almost always means one of two very different things happened, and
telling them apart matters:

1. **Something new got staged after the reboot completed.** If Windows
   Update/WSUS/SCCM is configured to run its detection-and-install cycle at
   startup, it can install another update within seconds or minutes of the
   reboot finishing — so what you're seeing is a *brand-new* pending state,
   not the old one failing to clear. Multi-stage servicing (large CUs,
   feature updates, .NET stacking) can also legitimately require two reboots
   in a row ("Configuring updates... X%"). An AV/EDR agent, backup agent, or
   the SCCM client itself can independently stage a rename/delete/reboot flag
   too.
2. **The servicing stack never actually cleared its own state.** Occasionally
   CBS fails to clear `RebootPending` due to a failed or interrupted
   servicing operation. Rebooting again won't fix this on its own — it needs
   `DISM /Online /Cleanup-Image /RestoreHealth` followed by `sfc /scannow`.

**The script checks this for you, with corroboration to avoid false positives.**
Every result includes `LastBootUpTime` (from `Win32_OperatingSystem`) plus
three independent "was this touched after the boot?" signals:
`CBSLogLastWriteTime` (the CBS.log file itself), `CBSKeyLastWriteTime` (the
actual `...\Component Based Servicing\RebootPending` registry key's own
last-write time, read via `RegQueryInfoKey` — not just the log file, which
can be touched by unrelated housekeeping), and `WUAUKeyLastWriteTime` (the
WUAU `RebootRequired` key's last-write time). For any server still flagged
with `RebootPending_Overall = True`, `Write-PendingRebootActivityCheck`
compares all available signals against `LastBootUpTime` and reports:

- **All available signals agree "newer than boot"** → new post-reboot
  activity (case 1 above), reported at **High** confidence if 2+ signals
  agree, **Medium** if only one signal was available. Expected, not a bug —
  reboot again once it settles.
- **All available signals agree "predates boot"** → stale/leftover state
  (case 2 above), same High/Medium confidence rule — a reboot alone won't
  fix it; run DISM/SFC on that server.
- **Signals disagree** → reported as **Low confidence / inconclusive**. Don't
  trust either conclusion — manually review `CBS.log` on that server instead.

**Naming the likely culprit — carefully.** When it's case 1, the script also
cross-references `RecentHotfixes` against `LastBootUpTime` and lists any
hotfix(es) installed after the reboot right under that server's entry (e.g.
"Likely candidate(s): KB5031234 ... (InstalledOn: ...)"). Two important
caveats are called out directly in the output:

- **Date-only granularity:** `InstalledOn` is frequently date-only (no
  time-of-day). A hotfix installed the *same calendar day* as the boot can't
  be reliably ordered against it, so those are listed separately as
  **"same-day (inconclusive)"** rather than being asserted as before/after.
- **Incomplete source:** `Get-HotFix` only sees OS-level Windows Update
  hotfixes — it will **not** show SQL Server, IIS, or other third-party
  product patches. An empty candidate list does not mean nothing was
  installed after the reboot; it just means nothing *Windows-Update-tracked*
  was. If `RecentHotfixes` wasn't populated at all for that server (it's only
  captured alongside the `PendingFileRenameOperations` check), no candidate
  list is shown and a manual CBS.log review is the fallback.

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

### Why can't the script see if a SQL Server patch/CU/hotfix was installed?

`RecentHotfixes` is populated via `Get-HotFix` (`Win32_QuickFixEngineering`),
which is a Windows Update Agent inventory — it only tracks OS-level Windows
Update patches. SQL Server CUs/hotfixes are typically applied via SQL's own
`setup.exe`/patch bundles that update SQL's own registry/version keys rather
than registering through Windows Update, so they never appear in
`QuickFixEngineering` — same for most other vendor/product patches.

This is a deliberate scope decision, not a technical limitation that could be
patched in trivially:

1. **Different discovery mechanism entirely.** SQL version/patch level comes
   from querying the SQL Server instance itself (e.g.
   `SELECT SERVERPROPERTY('ProductVersion')`, or
   `HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\...\Setup`), not from
   OS-level registry/WMI. That requires either a SQL connection (different
   auth/permissions than WinRM) or parsing SQL's own install registry
   path — a separate detection domain from "is a reboot pending."
2. **Instance-awareness problem.** A box can have zero, one, or many SQL
   instances (default + named instances, e.g. `SERVER08\INSTANCENAME` in the
   config example). Correlating a pending reboot to "which instance's which
   patch" multiplies the complexity — it's not a single flat check like
   `Get-HotFix`.
3. **The script's stated job is OS-level pending-reboot detection**,
   applicable uniformly across any Windows Server regardless of what's
   installed on it (SQL, IIS, nothing at all). Adding SQL-specific
   attribution logic would narrow that generality and couple it to one
   workload.
4. **It's diagnostic bonus info, not required to answer the core question.**
   The script already tells you a reboot is needed and why (rename/delete,
   CBS, WUAU, etc.); naming the exact SQL CU responsible isn't part of that
   yes/no reboot answer.

So the script can still tell you **that** a reboot is pending (via the
OS-level mechanisms any installer, SQL included, stages into), it just can't
*name* SQL as the specific cause the way it can for a Windows Update KB. See
the previous FAQ entry for how to confirm a SQL CU manually.

## Sources

- Microsoft — [`MsiSystemRebootPending` (Win32 API docs)](https://learn.microsoft.com/en-us/windows/win32/msi/msisystemrebootpending)
- Microsoft Support — [One restart a month for Windows updates (KB5121772)](https://support.microsoft.com/en-us/servicing/os/windows/docs/2026/07/kb5121772-one-restart-a-month-for-windows-updates)
- Microsoft Support — [Troubleshoot problems updating Windows](https://support.microsoft.com/en-us/windows/deployment/updates-lifecycle/troubleshoot-problems-updating-windows)
- Microsoft Q&A — [Question regarding PendingFileRenameOperations caused by Microsoft Edge](https://learn.microsoft.com/en-gb/answers/questions/5981763/question-regarding-pendingfilerenameoperations-cau)
- Ivanti — [Troubleshoot Persistent PendingFileRenameOperations In Registry](https://hub.ivanti.com/s/article/Troubleshoot-Persistent-PendingFileRenameOperations-In-Registry)
- Revenera — [PendingFileRenameOperations Versus MsiSystemRebootPending](https://community.revenera.com/s/article/pendingfilerenameoperations-versus-msisystemrebootpending)
- Microsoft Support — [Servicing Stack Updates (SSU): Frequently Asked Questions](https://support.microsoft.com/en-us/servicing/os/windows/2019/12/servicing-stack-updates-ssu-frequently-asked-questions)

## License

[MIT](LICENSE)

## Additional history

See [`Session-Notes.md`](Session-Notes.md) for a detailed development log of how
this script evolved (retry/timeout handling, server-list loading, export/email
reporting, hotfix correlation, OS identification, and expanded detection vectors).

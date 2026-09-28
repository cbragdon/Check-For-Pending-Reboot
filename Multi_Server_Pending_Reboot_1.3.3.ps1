#Requires -version 3

#==============================================================================
# Script  : Multi_Server_Pending_Reboot.ps1
# Version : 1.3.3
# Date    : 2026-03-25
#
# Purpose:
#   Remotely checks one or more servers for pending reboot conditions across
#   four detection vectors: Component Based Servicing (CBS), Windows Update
#   Auto Update (WUAU), PendingFileRenameOperations (Session Manager), and
#   the SCCM/CCM client WMI interface (if installed).
#
#   For servers flagged via PendingFileRenameOperations, the script parses
#   the raw REG_MULTI_SZ pairs, strips kernel namespace prefixes, filters
#   out installer/WU noise, and surfaces only operationally significant
#   Rename and Delete entries in a columnar table per server.
#
#   A definitions legend is printed at the end of the result set when
#   file operations are present, explaining Rename, Delete, and boot-phase
#   execution context.
#
# Functions:
#   Test-PendingReboot          - Core detection logic (pipeline-aware)
#   Write-PendingFileOperations - Tabular output of filtered file operations
#   Write-PendingRebootActivityCheck - CBS.log vs. last-boot-time comparison
#   Write-PendingRebootLegend   - Definitions footer (printed once, no params)
#
# Change Log:
#   1.0.0 - 2026-03-25 - Initial release / baseline
#   1.1.0 - 2026-03-25 - Moved definitions legend above summary table;
#                         added blank line separator between legend and table
#   1.2.0 - 2026-03-25 - Added $ShowPendingFiles control variable (default 0);
#                         file detail table and definitions legend suppressed
#                         unless $ShowPendingFiles is set to 1
#   1.2.1 - 2026-03-25 - Fixed: definitions legend now always displayed when
#                         file ops are flagged regardless of $ShowPendingFiles;
#                         $ShowPendingFiles gates file detail table only
#   1.3.0 - 2026-03-25 - Added RebootPending_Overall section to legend with
#                         per-vector definitions and per-server contributing
#                         flag breakdown; Write-PendingRebootLegend now accepts
#                         -Results parameter for dynamic server detail
#   1.3.1 - 2026-03-25 - Removed per-server contributing reasons block from
#                         legend; removed -Results parameter from
#                         Write-PendingRebootLegend — static definitions only
#   1.3.2 - 2026-03-25 - Left-justified REBOOT PENDING OVERALL vector entries
#                         to match indentation of other legend sections
#   1.3.3 - 2026-03-25 - Removed all indentation from legend; all sections
#                         fully left-justified for consistent readability
#   1.4.0 - 2026-09-26 - Added Get-PendingRebootComputerList (load server list
#                         from file or use default array); added connection
#                         retry/timeout support to Test-PendingReboot; added
#                         Write-PendingRebootSummary (color-coded console
#                         output); added Export-PendingRebootReport (CSV/HTML)
#                         and Send-PendingRebootEmail (optional email report);
#                         added $OnlyShowFlagged filter option
#   1.5.0 - 2026-09-26 - Added RecentHotfixes to Test-PendingReboot result:
#                         when PendingFileRenameOperations is flagged, the
#                         5 most recently installed hotfixes/CUs (Get-HotFix)
#                         are captured as candidate causes and displayed via
#                         Write-PendingFileOperations. This is a best-effort
#                         timing correlation, not a guaranteed package link —
#                         CBS.log on the target server is the authoritative
#                         source for confirming which KB queued a given file.
#   1.6.0 - 2026-09-26 - Added OS identification (OSCaption/OSVersion/
#                         OSBuildNumber via Win32_OperatingSystem) to every
#                         result so the script reports which Windows Server
#                         version (2016-2025) each check ran against. Added
#                         four additional detection vectors so the script
#                         covers every well-known pending-reboot location:
#                         CBS_PackagesPending, CBS_RebootInProgress,
#                         PendingComputerRename (ActiveComputerName vs
#                         ComputerName), and PendingDomainJoin (Netlogon
#                         JoinDomain/AvoidSpnSet). These registry/WMI
#                         locations are unchanged across Server 2016-2025, so
#                         no OS-version branching is needed in the detection
#                         logic itself. Write-PendingRebootSummary rewritten
#                         to compute columns dynamically; Export/Email
#                         column lists and the legend updated to match.
#   1.6.1 - 2026-09-27 - Fixed Write-PendingRebootSummary: rows are now built
#                         as a single Write-Host call per server instead of
#                         many small -NoNewline segments. PowerShell ISE's
#                         Start-Transcript logs each -NoNewline call on its
#                         own line, which garbled the summary table in saved
#                         transcripts (confirmed via real test run). Rows are
#                         now colored by overall severity (red = pending,
#                         magenta = connection error, gray = clean) instead of
#                         per-cell, trading granular cell color for reliable
#                         rendering across all hosts/logging methods.
#   1.7.0 - 2026-09-27 - Connection-failure rows now use explicit "N/A"
#                         (instead of False/$null) for every detection field
#                         (RebootPending_Overall, CBS_*, WUAU_RebootRequired,
#                         PendingFileRenameOperations_Exist/_Detail,
#                         RecentHotfixes, PendingComputerRename,
#                         PendingDomainJoin, CCM_RebootPending_WMI) and
#                         "Unknown" for OSCaption/OSVersion/OSBuildNumber, so
#                         a failed connection is never visually confused with
#                         a clean/False result. Added ConnectionErrorMessage
#                         field capturing the actual error text. Added
#                         Write-PendingRebootConnectionFailures — an always-
#                         visible, Write-Host-based report of servers that
#                         could not be reached, called automatically after
#                         the summary table. This replaces reliance on
#                         Write-Warning, whose output was not being captured
#                         by Start-Transcript in PowerShell ISE, so failed
#                         connections previously went unreported in saved
#                         transcripts. All truthiness checks against
#                         RebootPending_Overall/PendingFileRenameOperations_Exist
#                         (legend trigger, $OnlyShowFlagged filter, file-ops
#                         display, email flagged count) now explicitly compare
#                         "-eq $true" so the "N/A" placeholder string is never
#                         mistaken for a flagged/true result.
#   1.8.0 - 2026-09-27 - Added LastBootUpTime, CBSLogLastWriteTime, and
#                         CBSLogNewerThanBoot to Test-PendingReboot: compares
#                         C:\Windows\Logs\CBS\CBS.log's last-write time against
#                         the server's last boot time so a server that still
#                         shows a pending reboot right after rebooting can be
#                         told apart as either new post-reboot activity (a
#                         fresh update/agent/rename op staged afterward) or a
#                         stale/leftover CBS state that a reboot alone won't
#                         clear (needs DISM /RestoreHealth + sfc /scannow).
#                         Added Write-PendingRebootActivityCheck to surface
#                         this comparison for every flagged server. Added a
#                         matching FAQ entry to README.md.
#   1.8.1 - 2026-09-27 - Write-PendingRebootActivityCheck now cross-references
#                         RecentHotfixes against LastBootUpTime when CBS.log
#                         activity is newer than boot, listing any hotfix(es)
#                         installed after the reboot as the likely culprit(s)
#                         behind the fresh pending-reboot condition. Same
#                         best-effort heuristic caveat as RecentHotfixes
#                         elsewhere — not a guaranteed link.
#   1.8.2 - 2026-09-27 - False-positive hardening for the activity check:
#                         - Added CBSKeyLastWriteTime/WUAUKeyLastWriteTime
#                           (actual RebootPending/RebootRequired registry key
#                           timestamps via RegQueryInfoKey, not just the
#                           CBS.log file) as corroborating signals.
#                         - Write-PendingRebootActivityCheck now reports a
#                           High/Medium/Low confidence level based on how many
#                           of those signals agree, and calls out disagreement
#                           as inconclusive instead of asserting an answer.
#                         - Hotfix InstalledOn is often date-only; same-day
#                           hotfixes relative to boot are now labeled
#                           inconclusive instead of ordered as before/after.
#                         - Explicit note added wherever RecentHotfixes is
#                           shown that it only covers OS-level Windows Update
#                           hotfixes, not SQL Server/third-party patches.
#   1.8.3 - 2026-09-27 - Removed the Write-Warning calls in Test-PendingReboot
#                         for retry attempts and final connection failure.
#                         ConnectionErrorMessage is still captured on the
#                         result object, and Write-PendingRebootConnectionFailures
#                         already reports every unreachable server explicitly —
#                         the per-attempt/final warnings were redundant console
#                         noise on top of that.
#   1.8.4 - 2026-09-27 - Added a minimal single-line-per-host status message
#                         ("Connecting to X... Connected."/"Could not
#                         connect.") so the console shows live progress
#                         (avoiding the impression the script has hung)
#                         without reintroducing the removed per-retry/full
#                         WinRM error chatter.
#   1.8.5 - 2026-09-27 - Attempted fix: redirected Invoke-Command's error
#                         stream (2>$null) in the retry loop. Did NOT work —
#                         under PowerShell ISE the "PS>TerminatingError(): ..."
#                         trace for a caught -ErrorAction Stop failure is
#                         written directly to the host, bypassing stream
#                         redirection. Superseded by 1.8.6.
#   1.8.6 - 2026-09-27 - Replaced -ErrorAction Stop + try/catch around
#                         Invoke-Command with -ErrorAction SilentlyContinue +
#                         -ErrorVariable. Since the error is no longer
#                         terminating, ISE has nothing to trace — this is
#                         what actually eliminates the "PS>TerminatingError()"
#                         transcript/console chatter that 1.8.5 could not.
#   1.8.5 - 2026-09-27 - Redirected Invoke-Command's error stream (2>$null)
#                         in Test-PendingReboot's retry loop. Under
#                         PowerShell ISE + Start-Transcript, a caught
#                         -ErrorAction Stop failure was still writing a raw
#                         "PS>TerminatingError(): ..." engine trace line to
#                         the transcript even though the exception was
#                         already being handled in the catch block — this
#                         was the same kind of unwanted console/transcript
#                         chatter the 1.8.3 Write-Warning removal targeted,
#                         just from a different source.
#==============================================================================

function Test-PendingReboot {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        # Number of connection attempts per server before giving up.
        [int]$RetryCount = 2,

        # Delay between retry attempts.
        [int]$RetryDelaySeconds = 5,

        # Max time (seconds) to wait for the remote connection/command to respond.
        [int]$ConnectionTimeoutSeconds = 15
    )

    Process {
        foreach ($Computer in $ComputerName) {
            $Error.Clear()
            $attempt        = 0
            $remoteStatus   = $null
            $lastError      = $null
            $sessionOption  = New-PSSessionOption -OpenTimeout ($ConnectionTimeoutSeconds * 1000) `
                                                    -OperationTimeout ($ConnectionTimeoutSeconds * 1000)

            # Lightweight, single-line-per-host status — lets the person watching the
            # console know the script is alive and hasn't hung, without the noisy
            # per-retry/full-error chatter that used to print here.
            Write-Host "Connecting to $Computer... " -NoNewline

            while ($attempt -lt $RetryCount -and -not $remoteStatus) {
                $attempt++

                $scriptBlock = {
                    $result = [PSCustomObject]@{
                        ConnectionError                     = $false
                        ConnectionErrorMessage              = $null
                        OSCaption                           = $null
                        OSVersion                           = $null
                        OSBuildNumber                       = $null
                        LastBootUpTime                      = $null
                        CBSLogLastWriteTime                 = $null
                        CBSLogNewerThanBoot                 = $null
                        CBSKeyLastWriteTime                 = $null
                        CBSKeyNewerThanBoot                 = $null
                        WUAUKeyLastWriteTime                = $null
                        WUAUKeyNewerThanBoot                = $null
                        CBS_RebootPending                  = $false
                        CBS_PackagesPending                = $false
                        CBS_RebootInProgress               = $false
                        WUAU_RebootRequired                = $false
                        PendingFileRenameOperations_Exist  = $false
                        PendingFileRenameOperations_Detail = $null
                        RecentHotfixes                     = $null
                        PendingComputerRename              = $false
                        PendingDomainJoin                  = $false
                        CCM_RebootPending_WMI              = $false
                        RebootPending_Overall              = $false
                    }

                    # OS identification — works uniformly across Windows Server 2016-2025;
                    # the CBS/WUAU/PFRO/rename/domain-join mechanisms below are unchanged
                    # across that entire range, so no version-specific branching is needed
                    # for detection logic itself — this is captured for reporting/context.
                    try {
                        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
                        $result.OSCaption      = $os.Caption
                        $result.OSVersion      = $os.Version
                        $result.OSBuildNumber  = $os.BuildNumber
                        $result.LastBootUpTime = $os.LastBootUpTime
                    } catch {}

                    # CBS.log activity vs. last boot time — a server can still show a pending
                    # reboot right after rebooting for two very different reasons: (a) something
                    # staged a *new* update/operation after the reboot completed, or (b) the
                    # servicing stack never actually cleared its own state and further reboots
                    # alone won't fix it. Comparing CBS.log's last-write time to LastBootUpTime
                    # tells the two apart without having to open/parse the log itself.
                    try {
                        $cbsLog = Get-Item 'C:\Windows\Logs\CBS\CBS.log' -ErrorAction Stop
                        $result.CBSLogLastWriteTime = $cbsLog.LastWriteTime
                        if ($result.LastBootUpTime) {
                            $result.CBSLogNewerThanBoot = ($cbsLog.LastWriteTime -gt $result.LastBootUpTime)
                        }
                    } catch {}

                    # Check 1: Component Based Servicing (CBS) — reboot pending
                    if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA Ignore) {
                        $result.CBS_RebootPending = $true
                    }

                    # Registry-key LastWriteTime corroboration: CBS.log's file-write time alone
                    # can reflect benign housekeeping unrelated to a real reboot-required change,
                    # so we also read the actual RebootPending/RebootRequired key's own last-write
                    # timestamp (not exposed by Get-Item for the registry provider — requires
                    # RegQueryInfoKey via P/Invoke) as a second, independent signal. Agreement
                    # between this and CBSLogNewerThanBoot raises confidence; disagreement flags
                    # the result as inconclusive instead of asserting one way or the other.
                    if (-not ('PRRegistryTimestamp' -as [type])) {
                        try {
                            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
public class PRRegistryTimestamp {
    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern int RegQueryInfoKey(
        IntPtr hKey, StringBuilder lpClass, ref uint lpcbClass, IntPtr lpReserved,
        out uint lpcSubKeys, out uint lpcbMaxSubKeyLen, out uint lpcbMaxClassLen,
        out uint lpcValues, out uint lpcbMaxValueNameLen, out uint lpcbMaxValueLen,
        out uint lpcbSecurityDescriptor, out long lpftLastWriteTime);

    public static Nullable<DateTime> GetLastWriteTime(string subKeyPath) {
        using (RegistryKey key = Registry.LocalMachine.OpenSubKey(subKeyPath)) {
            if (key == null) { return null; }
            uint lpcbClass = 0, lpcSubKeys, lpcbMaxSubKeyLen, lpcbMaxClassLen, lpcValues,
                 lpcbMaxValueNameLen, lpcbMaxValueLen, lpcbSecurityDescriptor;
            long lastWriteTime;
            int hr = RegQueryInfoKey(key.Handle.DangerousGetHandle(), null, ref lpcbClass, IntPtr.Zero,
                out lpcSubKeys, out lpcbMaxSubKeyLen, out lpcbMaxClassLen, out lpcValues,
                out lpcbMaxValueNameLen, out lpcbMaxValueLen, out lpcbSecurityDescriptor, out lastWriteTime);
            if (hr != 0) { return null; }
            return DateTime.FromFileTime(lastWriteTime);
        }
    }
}
'@ -ErrorAction Stop
                        } catch {}
                    }

                    if ($result.CBS_RebootPending) {
                        try {
                            $result.CBSKeyLastWriteTime = [PRRegistryTimestamp]::GetLastWriteTime(
                                'Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending')
                            if ($result.LastBootUpTime -and $result.CBSKeyLastWriteTime) {
                                $result.CBSKeyNewerThanBoot = ($result.CBSKeyLastWriteTime -gt $result.LastBootUpTime)
                            }
                        } catch {}
                    }

                    # Check 1b: Component Based Servicing — packages still mid-installation
                    if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending" -EA Ignore) {
                        $result.CBS_PackagesPending = $true
                    }

                    # Check 1c: Component Based Servicing — reboot already in progress
                    if (Get-Item "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress" -EA Ignore) {
                        $result.CBS_RebootInProgress = $true
                    }

                    # Check 2: Windows Update Auto Update (WUAU)
                    if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) {
                        $result.WUAU_RebootRequired = $true
                        try {
                            $result.WUAUKeyLastWriteTime = [PRRegistryTimestamp]::GetLastWriteTime(
                                'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
                            if ($result.LastBootUpTime -and $result.WUAUKeyLastWriteTime) {
                                $result.WUAUKeyNewerThanBoot = ($result.WUAUKeyLastWriteTime -gt $result.LastBootUpTime)
                            }
                        } catch {}
                    }

                    # Check 3: PendingFileRenameOperations (Session Manager)
                    $pfro = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
                                -Name PendingFileRenameOperations -EA Ignore

                    if ($pfro) {
                        $result.PendingFileRenameOperations_Exist = $true

                        # --- Tunable filters ---------------------------------------------------
                        # Entries whose SOURCE path starts with any of these are considered significant.
                        # Extend this list as your environment requires.
                        $significantPrefixes = @(
                            'C:\Windows\System32',
                            'C:\Windows\SysWOW64',
                            'C:\Windows\WinSxS',
                            'C:\Windows\assembly',
                            'C:\Program Files\',
                            'C:\Program Files (x86)\'
                        )

                        # Even within significant paths, these patterns are installer/WU housekeeping
                        # and do not represent a functional dependency on the reboot.
                        $noisyPatterns = @(
                            '*\Temp\*',
                            '*\Windows\Installer\*',
                            '*\Windows\WER\*',
                            '*.tmp',
                            '*.log'
                        )
                        # -----------------------------------------------------------------------

                        $entries = $pfro.PendingFileRenameOperations
                        $pairs   = [System.Collections.Generic.List[PSCustomObject]]::new()

                        for ($i = 0; $i -lt $entries.Count; $i += 2) {
                            # Strip everything up to and including the \??\ kernel namespace delimiter.
                            # Raw values vary: \??\C:\... | *\??\C:\... | *1\??\C:\...
                            # The flags byte (e.g. '1') between the retry marker and \??\ makes
                            # enumerating prefix patterns brittle — anchor on \??\ instead.
                            $source      = $entries[$i]      -replace '^.*\\\?\?\\', ''
                            $destination = if ($i + 1 -lt $entries.Count) {
                                               $entries[$i + 1] -replace '^.*\\\?\?\\', ''
                                           } else { '' }

                            # Step 1: must match at least one significant prefix
                            $isSignificant = $significantPrefixes | Where-Object { $source -like "$_*" }
                            if (-not $isSignificant) { continue }

                            # Step 2: exclude known noisy patterns even within significant paths
                            $isNoisy = $noisyPatterns | Where-Object { $source -like $_ }
                            if ($isNoisy) { continue }

                            $pairs.Add([PSCustomObject]@{
                                Source      = $source
                                Destination = if ($destination) { $destination } else { '[DELETE]' }
                                Action      = if ($destination) { 'Rename' }    else { 'Delete'   }
                            })
                        }

                        # Only populate detail if filtered list has entries worth reporting.
                        # Note: PendingFileRenameOperations_Exist remains TRUE regardless — the reboot
                        # is still pending. The filter is for operational visibility only.
                        $result.PendingFileRenameOperations_Detail = if ($pairs.Count -gt 0) { $pairs } else { $null }

                        # Best-effort correlation: PendingFileRenameOperations does not itself record
                        # which package queued a file, so we surface the most recently installed
                        # hotfixes/CUs as candidates. Match by proximity of InstalledOn to "now" —
                        # not a guaranteed link, but the common practical heuristic. For a definitive
                        # answer, grep C:\Windows\Logs\CBS\CBS.log on the server for the file name.
                        try {
                            $result.RecentHotfixes = Get-HotFix -ErrorAction Stop |
                                Sort-Object InstalledOn -Descending |
                                Select-Object -First 5 HotFixID, Description, InstalledOn
                        } catch {}
                    }

                    # Check 4: SCCM/CCM Client (optional — only if client is installed)
                    try {
                        $util   = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities"
                        $status = $util.DetermineIfRebootPending()
                        if (($status -ne $null) -and $status.RebootPending) {
                            $result.CCM_RebootPending_WMI = $true
                        }
                    } catch {}

                    # Check 5: Pending computer rename — ActiveComputerName vs ComputerName
                    # differ when a rename has been requested but not yet applied (takes
                    # effect on next boot).
                    try {
                        $active = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName" -Name ComputerName -EA Stop).ComputerName
                        $pending = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName" -Name ComputerName -EA Stop).ComputerName
                        if ($active -ne $pending) {
                            $result.PendingComputerRename = $true
                        }
                    } catch {}

                    # Check 6: Pending domain join/leave — Netlogon stages these values and
                    # removes them after the reboot that completes the operation.
                    $netlogon = "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon"
                    if ((Get-ItemProperty $netlogon -Name JoinDomain    -EA Ignore) -or
                        (Get-ItemProperty $netlogon -Name AvoidSpnSet  -EA Ignore)) {
                        $result.PendingDomainJoin = $true
                    }

                    # Overall flag
                    if ($result.CBS_RebootPending -or $result.CBS_PackagesPending -or $result.CBS_RebootInProgress -or
                        $result.WUAU_RebootRequired -or $result.PendingFileRenameOperations_Exist -or
                        $result.PendingComputerRename -or $result.PendingDomainJoin -or
                        $result.CCM_RebootPending_WMI) {
                        $result.RebootPending_Overall = $true
                    }

                    return $result
                }

                $icError = $null
                $remoteStatus = Invoke-Command -ComputerName $Computer -ScriptBlock $scriptBlock `
                                    -SessionOption $sessionOption -ErrorAction SilentlyContinue -ErrorVariable icError

                # Deliberately non-terminating (-ErrorAction SilentlyContinue + -ErrorVariable)
                # rather than -ErrorAction Stop + try/catch: under PowerShell ISE, a caught
                # terminating error from Invoke-Command still writes a raw
                # "PS>TerminatingError(): ..." trace directly to the host/transcript,
                # bypassing normal stream redirection (2>$null doesn't touch it). Checking
                # $icError avoids triggering that trace in the first place.
                if ($icError) {
                    $lastError    = $icError[0]
                    $remoteStatus = $null
                    if ($attempt -lt $RetryCount) {
                        Start-Sleep -Seconds $RetryDelaySeconds
                    }
                }
            }

            if ($remoteStatus) {
                Write-Host "Connected." -ForegroundColor Green
                $remoteStatus | Select-Object @{Name = 'ComputerName'; Expression = {$Computer}}, *
            }
            else {
                Write-Host "Could not connect." -ForegroundColor Red
                # No Write-Warning here by design: the summary table / CONNECTION FAILURES
                # section (Write-PendingRebootConnectionFailures) already reports every
                # unreachable server, including this same error message, via ConnectionErrorMessage
                # below — so a console warning here would just be redundant noise.
                [PSCustomObject]@{
                    ComputerName                       = $Computer
                    ConnectionError                    = $true
                    ConnectionErrorMessage             = "$lastError"
                    OSCaption                           = 'Unknown'
                    OSVersion                           = 'Unknown'
                    OSBuildNumber                       = 'Unknown'
                    LastBootUpTime                      = 'N/A'
                    CBSLogLastWriteTime                 = 'N/A'
                    CBSLogNewerThanBoot                 = 'N/A'
                    CBSKeyLastWriteTime                 = 'N/A'
                    CBSKeyNewerThanBoot                 = 'N/A'
                    WUAUKeyLastWriteTime                = 'N/A'
                    WUAUKeyNewerThanBoot                = 'N/A'
                    RebootPending_Overall              = 'N/A'
                    CBS_RebootPending                  = 'N/A'
                    CBS_PackagesPending                = 'N/A'
                    CBS_RebootInProgress               = 'N/A'
                    WUAU_RebootRequired                = 'N/A'
                    PendingFileRenameOperations_Exist  = 'N/A'
                    PendingFileRenameOperations_Detail = 'N/A'
                    RecentHotfixes                     = 'N/A'
                    PendingComputerRename              = 'N/A'
                    PendingDomainJoin                  = 'N/A'
                    CCM_RebootPending_WMI              = 'N/A'
                }
            }
        }
    }
}


function Write-PendingFileOperations {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results
    )

    $flagged = $Results | Where-Object { $_.PendingFileRenameOperations_Exist -eq $true }
    if (-not $flagged) { return }

    foreach ($server in $flagged) {
        Write-Host ""
        Write-Host "[$($server.ComputerName)] Pending File Rename Operations:" -ForegroundColor Cyan

        if (-not $server.PendingFileRenameOperations_Detail) {
            Write-Host "   No significant operations after filtering (reboot still required)." -ForegroundColor DarkYellow
            continue
        }

        # Output as a formatted table matching the original columnar layout
        $server.PendingFileRenameOperations_Detail | Format-Table Source, Destination, Action -AutoSize

        if ($server.RecentHotfixes) {
            Write-Host "   Recently installed hotfixes/CUs on this server (possible cause — not a guaranteed match;" -ForegroundColor DarkCyan
            Write-Host "   OS-level Windows Update hotfixes only, does not include SQL Server/third-party patches):" -ForegroundColor DarkCyan
            $server.RecentHotfixes | Format-Table HotFixID, Description, InstalledOn -AutoSize
            Write-Host "   For a definitive link, grep C:\Windows\Logs\CBS\CBS.log on $($server.ComputerName) for the file name above." -ForegroundColor DarkGray
        }
    }
}


function Write-PendingRebootActivityCheck {
    <#
        For servers currently flagged with a pending reboot, compares
        C:\Windows\Logs\CBS\CBS.log's last-write time against the server's
        last boot time (LastBootUpTime), corroborated by the actual
        RebootPending/RebootRequired registry keys' own LastWriteTime (not
        just the log file), to answer "why am I still seeing a pending
        reboot right after rebooting?" with a stated confidence level rather
        than a flat assertion:

          - NEWER than the last boot -> something (Windows Update/WSUS/SCCM,
            an agent installer, a rename/domain-join op) staged a fresh
            pending-reboot condition *after* the reboot completed. This is a
            new state, not a leftover — expected, not a bug.

          - OLDER than (or equal to) the last boot -> the servicing stack
            never actually cleared its own RebootPending state during that
            reboot. Further reboots alone likely won't resolve it;
            investigate with 'DISM /Online /Cleanup-Image /RestoreHealth' and
            'sfc /scannow' on the target server.

        Confidence is High when all available signals (CBS.log file, CBS
        RebootPending key, WUAU RebootRequired key) agree, Medium when only
        one signal is available, and Low/"inconclusive" when available
        signals disagree — in which case a manual CBS.log review on the
        server is the recommended fallback rather than trusting either
        conclusion.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results
    )

    $flagged = $Results | Where-Object {
        $_.RebootPending_Overall -eq $true -and $_.CBSLogNewerThanBoot -is [bool]
    }
    if (-not $flagged) { return }

    Write-Host ""
    Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "CBS.LOG ACTIVITY vs. LAST BOOT TIME" -ForegroundColor White
    Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray

    foreach ($server in $flagged) {
        Write-Host ""
        Write-Host "[$($server.ComputerName)]" -ForegroundColor Cyan
        Write-Host "   Last boot time            : $($server.LastBootUpTime)" -ForegroundColor Yellow
        Write-Host "   CBS.log last activity     : $($server.CBSLogLastWriteTime)" -ForegroundColor Yellow
        if ($server.CBSKeyLastWriteTime -is [datetime]) {
            Write-Host "   CBS RebootPending key     : $($server.CBSKeyLastWriteTime)" -ForegroundColor Yellow
        }
        if ($server.WUAUKeyLastWriteTime -is [datetime]) {
            Write-Host "   WUAU RebootRequired key   : $($server.WUAUKeyLastWriteTime)" -ForegroundColor Yellow
        }

        # Gather every available "newer than boot" signal. Only signals that were
        # actually captured (not $null/'N/A') count toward agreement/confidence.
        $signals = @()
        if ($server.CBSLogNewerThanBoot -is [bool])  { $signals += $server.CBSLogNewerThanBoot }
        if ($server.CBSKeyNewerThanBoot -is [bool])  { $signals += $server.CBSKeyNewerThanBoot }
        if ($server.WUAUKeyNewerThanBoot -is [bool]) { $signals += $server.WUAUKeyNewerThanBoot }

        $newerCount = ($signals | Where-Object { $_ -eq $true }).Count
        $olderCount = ($signals | Where-Object { $_ -eq $false }).Count
        $agree      = ($newerCount -eq 0 -or $olderCount -eq 0)   # all available signals agree
        $confidence = if ($signals.Count -le 1) { 'Medium' } elseif ($agree) { 'High' } else { 'Low' }

        if (-not $agree) {
            Write-Host "   >> Signals DISAGREE ($newerCount newer / $olderCount older than boot) —" -ForegroundColor Magenta
            Write-Host "      inconclusive. Don't trust either conclusion below; manually review" -ForegroundColor Magenta
            Write-Host "      C:\Windows\Logs\CBS\CBS.log on $($server.ComputerName) instead." -ForegroundColor Magenta
            Write-Host "   Confidence: $confidence" -ForegroundColor Magenta
            continue
        }

        if ($server.CBSLogNewerThanBoot) {
            Write-Host "   >> Activity is NEWER than the last boot — this looks like a fresh" -ForegroundColor DarkYellow
            Write-Host "      pending-reboot condition staged after the reboot completed" -ForegroundColor DarkYellow
            Write-Host "      (new update/agent/rename op), not a leftover from before it." -ForegroundColor DarkYellow
            Write-Host "   Confidence: $confidence ($($signals.Count) signal(s) checked)" -ForegroundColor DarkYellow

            # Best-effort: name the likely culprit(s) by cross-referencing RecentHotfixes
            # (captured earlier, see Check 1) against LastBootUpTime. Caveats:
            #  - Get-HotFix only sees OS-level Windows Update hotfixes — it will NOT show
            #    SQL Server, IIS, or other third-party product patches, so an empty/absent
            #    list here does not mean nothing was installed after the reboot.
            #  - InstalledOn is frequently DATE-ONLY (midnight, no time-of-day). A hotfix
            #    installed the same calendar day as the boot can't be reliably ordered
            #    against it, so those are called out separately as inconclusive rather
            #    than asserted as before/after.
            if ($server.RecentHotfixes -and $server.RecentHotfixes -ne 'N/A') {
                $sameDay = $server.RecentHotfixes | Where-Object {
                    $_.InstalledOn -and $server.LastBootUpTime -is [datetime] -and
                    $_.InstalledOn.Date -eq $server.LastBootUpTime.Date
                }
                $postBootHotfixes = $server.RecentHotfixes | Where-Object {
                    $_.InstalledOn -and $server.LastBootUpTime -is [datetime] -and
                    $_.InstalledOn.Date -gt $server.LastBootUpTime.Date
                }
                if ($postBootHotfixes) {
                    Write-Host "      Likely candidate(s) — hotfix(es) installed after the last boot:" -ForegroundColor DarkYellow
                    $postBootHotfixes | ForEach-Object {
                        Write-Host "        - $($_.HotFixID)  $($_.Description)  (InstalledOn: $($_.InstalledOn))" -ForegroundColor DarkYellow
                    }
                }
                if ($sameDay) {
                    Write-Host "      Same-day (inconclusive — InstalledOn is date-only, so before/after" -ForegroundColor DarkGray
                    Write-Host "      the boot can't be determined) hotfix(es):" -ForegroundColor DarkGray
                    $sameDay | ForEach-Object {
                        Write-Host "        - $($_.HotFixID)  $($_.Description)  (InstalledOn: $($_.InstalledOn))" -ForegroundColor DarkGray
                    }
                }
            }
            Write-Host "      Note: candidate list only covers OS-level Windows Update hotfixes —" -ForegroundColor DarkGray
            Write-Host "      SQL Server/third-party patches won't appear here even if responsible." -ForegroundColor DarkGray
        }
        else {
            Write-Host "   >> Activity PREDATES the last boot — the servicing stack never" -ForegroundColor Red
            Write-Host "      actually cleared its RebootPending state. Another reboot alone" -ForegroundColor Red
            Write-Host "      likely won't resolve this — investigate with" -ForegroundColor Red
            Write-Host "      'DISM /Online /Cleanup-Image /RestoreHealth' and 'sfc /scannow'." -ForegroundColor Red
            Write-Host "   Confidence: $confidence ($($signals.Count) signal(s) checked)" -ForegroundColor Red
        }
    }
    Write-Host ""
}


function Write-PendingRebootLegend {

    Write-Host ""
    Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "PENDING FILE OPERATION DEFINITIONS" -ForegroundColor White
    Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray

    Write-Host ""
    Write-Host "ACTION: Rename" -ForegroundColor White
    Write-Host "A file is currently locked by the OS or a running process and" -ForegroundColor Yellow
    Write-Host "cannot be replaced in-place. Windows will move the new version" -ForegroundColor Yellow
    Write-Host "into the target path on next boot before any services start." -ForegroundColor Yellow
    Write-Host "Common causes: DLL/EXE updates, driver replacements, in-use" -ForegroundColor Yellow
    Write-Host "system binaries patched by Windows Update or an installer." -ForegroundColor Yellow

    Write-Host ""
    Write-Host "ACTION: Delete" -ForegroundColor White
    Write-Host "A file is queued for removal but is currently locked or still" -ForegroundColor Yellow
    Write-Host "in use. Windows will delete it on next boot before user-space" -ForegroundColor Yellow
    Write-Host "processes load. Common causes: old DLL versions left behind" -ForegroundColor Yellow
    Write-Host "after an in-place upgrade, orphaned installer staging files," -ForegroundColor Yellow
    Write-Host "or superseded driver binaries pending cleanup." -ForegroundColor Yellow

    Write-Host ""
    Write-Host "NOTE:" -ForegroundColor White
    Write-Host "Both operations execute during the early boot phase (Session" -ForegroundColor Yellow
    Write-Host "Manager initialization) before the system is fully online." -ForegroundColor Yellow
    Write-Host "A server flagged for either reason is NOT fully patched or" -ForegroundColor Yellow
    Write-Host "updated until the reboot completes successfully." -ForegroundColor Yellow

    # ── RebootPending_Overall breakdown ───────────────────────────────────────
    Write-Host ""
    Write-Host "REBOOT PENDING — OVERALL:" -ForegroundColor White
    Write-Host "RebootPending_Overall is True when one or more of the following" -ForegroundColor Yellow
    Write-Host "detection vectors is flagged on the server:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "CBS_RebootPending           : Component Based Servicing" -ForegroundColor Yellow
    Write-Host "A Windows component or role update has been staged but" -ForegroundColor Yellow
    Write-Host "requires a reboot to complete installation." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "CBS_PackagesPending         : Component Based Servicing (packages)" -ForegroundColor Yellow
    Write-Host "One or more servicing packages are still mid-installation and" -ForegroundColor Yellow
    Write-Host "have not yet been finalized — typically resolved by a reboot." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "CBS_RebootInProgress        : Component Based Servicing (in progress)" -ForegroundColor Yellow
    Write-Host "The servicing stack has marked a reboot as already underway/" -ForegroundColor Yellow
    Write-Host "required to finish applying staged changes." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "WUAU_RebootRequired         : Windows Update Auto Update" -ForegroundColor Yellow
    Write-Host "One or more Windows Updates have been installed and are" -ForegroundColor Yellow
    Write-Host "waiting on a reboot to finalize." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "PendingFileRenameOperations : Session Manager (PFRO)" -ForegroundColor Yellow
    Write-Host "Files that were in use during patching are queued for" -ForegroundColor Yellow
    Write-Host "rename or deletion at next boot. See detail above." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "PendingComputerRename       : Session Manager (ComputerName)" -ForegroundColor Yellow
    Write-Host "A computer rename has been requested but ActiveComputerName" -ForegroundColor Yellow
    Write-Host "and ComputerName registry values differ until the next boot." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "PendingDomainJoin           : Netlogon (domain join/leave)" -ForegroundColor Yellow
    Write-Host "A domain join or leave operation has been staged and will" -ForegroundColor Yellow
    Write-Host "complete on the next reboot." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "CCM_RebootPending_WMI       : SCCM/CCM Client" -ForegroundColor Yellow
    Write-Host "The SCCM client has signaled a reboot is required," -ForegroundColor Yellow
    Write-Host "typically following a software deployment or patch cycle." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "APPLICABILITY:" -ForegroundColor White
    Write-Host "All detection vectors above use registry/WMI locations that are" -ForegroundColor Yellow
    Write-Host "unchanged across Windows Server 2016, 2019, 2022, and 2025 — no" -ForegroundColor Yellow
    Write-Host "OS-version-specific logic is required. OSCaption/OSVersion/" -ForegroundColor Yellow
    Write-Host "OSBuildNumber are captured per server purely for reporting." -ForegroundColor Yellow

    Write-Host ""
    Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}


function Get-PendingRebootComputerList {
    <#
        Resolves the list of servers to check.
        - If $Path is supplied and exists, loads server names from it:
            * .csv  -> must contain a "ComputerName" column
            * any other extension -> one server name per line (blank/'#' lines ignored)
        - Otherwise falls back to $DefaultList.
    #>
    param(
        [string]  $Path,
        [string[]]$DefaultList
    )

    if ($Path -and (Test-Path -LiteralPath $Path)) {
        if ($Path -match '\.csv$') {
            return (Import-Csv -Path $Path).ComputerName | Where-Object { $_ }
        }
        else {
            return Get-Content -Path $Path |
                Where-Object { $_.Trim() -and -not $_.Trim().StartsWith('#') } |
                ForEach-Object { $_.Trim() }
        }
    }

    return $DefaultList
}


function Write-PendingRebootSummary {
    <#
        Color-coded, per-server console summary. Columns are computed dynamically
        so new detection vectors/OS info can be added without hardcoding widths.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results
    )

    # Name = property on the result object, Header = column title, Flag = colorize as YES/no
    $columns = @(
        @{Name = 'ComputerName';                       Header = 'ComputerName'; Flag = $false}
        @{Name = 'OSCaption';                           Header = 'OS';           Flag = $false}
        @{Name = 'RebootPending_Overall';               Header = 'Overall';      Flag = $true}
        @{Name = 'CBS_RebootPending';                   Header = 'CBS';          Flag = $true}
        @{Name = 'CBS_PackagesPending';                 Header = 'PkgsPend';     Flag = $true}
        @{Name = 'CBS_RebootInProgress';                Header = 'RbtInProg';    Flag = $true}
        @{Name = 'WUAU_RebootRequired';                 Header = 'WUAU';         Flag = $true}
        @{Name = 'PendingFileRenameOperations_Exist';   Header = 'PFRO';         Flag = $true}
        @{Name = 'PendingComputerRename';               Header = 'CRename';      Flag = $true}
        @{Name = 'PendingDomainJoin';                   Header = 'DomJoin';      Flag = $true}
        @{Name = 'CCM_RebootPending_WMI';               Header = 'CCM';          Flag = $true}
    )

    # Compute each column's width from the header and the widest value in the data.
    foreach ($col in $columns) {
        $dataWidth = ($Results | ForEach-Object {
            $v = $_.($col.Name)
            if ($col.Flag) { 3 } else { "$v".Length }
        } | Measure-Object -Maximum).Maximum
        $col.Width = [Math]::Max($col.Header.Length, [int]$dataWidth)
    }

    $headerLine = ($columns | ForEach-Object { "{0,-$($_.Width)}" -f $_.Header }) -join '  '
    Write-Host ""
    Write-Host $headerLine -ForegroundColor White
    Write-Host ('-' * $headerLine.Length) -ForegroundColor DarkGray

    foreach ($r in $Results) {
        # Build the entire row as one string and issue a single Write-Host call.
        # (Using many small -NoNewline segments renders correctly live in a normal
        # console, but PowerShell ISE's Start-Transcript logs each -NoNewline call
        # on its own line, garbling the table in the saved transcript. A single
        # call per row avoids that regardless of host.)
        $cells = foreach ($col in $columns) {
            $value = $r.($col.Name)
            if ($col.Flag) {
                # Booleans render as YES/no; anything else (e.g. the 'N/A' string used
                # for connection-error rows) is displayed verbatim.
                $text = if ($value -is [bool]) { if ($value) { 'YES' } else { 'no' } } else { "$value" }
            }
            else {
                $text = "$value"
            }
            "{0,-$($col.Width)}" -f $text
        }
        $rowLine = $cells -join '  '

        $rowColor = if ($r.ConnectionError)              { 'Magenta' }
                    elseif ($r.RebootPending_Overall -eq $true) { 'Red' }
                    else                                  { 'DarkGray' }

        Write-Host $rowLine -ForegroundColor $rowColor

        if ($r.ConnectionError) {
            Write-Host ("   >> Connection/execution error — see Connection Failures section below.") -ForegroundColor Magenta
        }
    }
    Write-Host ""
}


function Write-PendingRebootConnectionFailures {
    <#
        Explicit, always-visible report of servers the script could not reach.
        Deliberately uses Write-Host rather than Write-Warning: Write-Warning
        output is not reliably captured by Start-Transcript in every host
        (notably PowerShell ISE), which previously caused failed connections
        to go unreported in saved transcripts even though a warning was raised.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results
    )

    $failed = $Results | Where-Object { $_.ConnectionError }
    if (-not $failed) { return }

    Write-Host ""
    Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "CONNECTION FAILURES" -ForegroundColor White
    Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray

    foreach ($f in $failed) {
        Write-Host ""
        Write-Host "[$($f.ComputerName)] Could not connect / execute after retries." -ForegroundColor Magenta
        if ($f.ConnectionErrorMessage) {
            Write-Host "   Error: $($f.ConnectionErrorMessage)" -ForegroundColor DarkGray
        }
    }
    Write-Host ""
}


function Export-PendingRebootReport {
    <#
        Exports summary results (flat columns only — file-op detail is skipped
        since it doesn't flatten cleanly into CSV/HTML) to CSV and/or HTML.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results,

        [Parameter(Mandatory = $true)]
        [string]$OutputFolder,

        [ValidateSet('CSV','HTML','Both')]
        [string]$Format = 'Both'
    )

    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $flat = $Results | Select-Object ComputerName,
                                     OSCaption,
                                     OSVersion,
                                     OSBuildNumber,
                                     LastBootUpTime,
                                     CBSLogLastWriteTime,
                                     CBSLogNewerThanBoot,
                                     CBSKeyNewerThanBoot,
                                     WUAUKeyNewerThanBoot,
                                     RebootPending_Overall,
                                     CBS_RebootPending,
                                     CBS_PackagesPending,
                                     CBS_RebootInProgress,
                                     WUAU_RebootRequired,
                                     PendingFileRenameOperations_Exist,
                                     PendingComputerRename,
                                     PendingDomainJoin,
                                     CCM_RebootPending_WMI,
                                     ConnectionError,
                                     ConnectionErrorMessage

    $paths = @{}

    if ($Format -in 'CSV','Both') {
        $csvPath = Join-Path $OutputFolder "PendingReboot_$timestamp.csv"
        $flat | Export-Csv -Path $csvPath -NoTypeInformation
        $paths.CSV = $csvPath
    }

    if ($Format -in 'HTML','Both') {
        $htmlPath = Join-Path $OutputFolder "PendingReboot_$timestamp.html"
        $style = "<style>table{border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px;} " +
                 "th,td{border:1px solid #ccc;padding:4px 8px;text-align:left;} " +
                 "th{background:#333;color:#fff;} tr:nth-child(even){background:#f4f4f4;}</style>"
        $flat | ConvertTo-Html -Title "Pending Reboot Report - $timestamp" -Head $style |
            Out-File -FilePath $htmlPath -Encoding UTF8
        $paths.HTML = $htmlPath
    }

    return [PSCustomObject]$paths
}


function Send-PendingRebootEmail {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results,

        [Parameter(Mandatory = $true)]
        [string]$SmtpServer,

        [Parameter(Mandatory = $true)]
        [string]$From,

        [Parameter(Mandatory = $true)]
        [string[]]$To,

        [string]$Subject = "Pending Reboot Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm')",

        [string]$AttachmentPath
    )

    $flaggedCount = ($Results | Where-Object { $_.RebootPending_Overall -eq $true }).Count
    $style = "<style>table{border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px;} " +
             "th,td{border:1px solid #ccc;padding:4px 8px;text-align:left;} " +
             "th{background:#333;color:#fff;} tr:nth-child(even){background:#f4f4f4;}</style>"

    $body = $Results | Select-Object ComputerName,
                                     OSCaption,
                                     OSVersion,
                                     OSBuildNumber,
                                     LastBootUpTime,
                                     CBSLogLastWriteTime,
                                     CBSLogNewerThanBoot,
                                     CBSKeyNewerThanBoot,
                                     WUAUKeyNewerThanBoot,
                                     RebootPending_Overall,
                                     CBS_RebootPending,
                                     CBS_PackagesPending,
                                     CBS_RebootInProgress,
                                     WUAU_RebootRequired,
                                     PendingFileRenameOperations_Exist,
                                     PendingComputerRename,
                                     PendingDomainJoin,
                                     CCM_RebootPending_WMI,
                                     ConnectionError,
                                     ConnectionErrorMessage |
        ConvertTo-Html -Head $style -Body "<h2>Pending Reboot Report</h2><p>$flaggedCount of $($Results.Count) server(s) flagged.</p>" |
        Out-String

    $mailParams = @{
        SmtpServer = $SmtpServer
        From       = $From
        To         = $To
        Subject    = $Subject
        Body       = $body
        BodyAsHtml = $true
    }
    if ($AttachmentPath -and (Test-Path -LiteralPath $AttachmentPath)) {
        $mailParams.Attachments = $AttachmentPath
    }

    Send-MailMessage @mailParams
}


# =============================================================================
# USAGE
# =============================================================================

# Set to 1 to display pending Rename/Delete file details per server.
# Default is 0 — summary table only.
$ShowPendingFiles = 0

# Set to 1 to show only servers where RebootPending_Overall is True
# (in the console summary, export, and email). Default 0 shows all servers.
$OnlyShowFlagged = 0

# Optional: path to a file listing servers to check — one name per line (# = comment),
# or a CSV with a "ComputerName" column. Leave blank ('') to use $DefaultComputerList below.
$ComputerListPath = ''

# Fallback / default server list used when $ComputerListPath is blank or not found.
# Replace these placeholder names with your own server(s) — additional example
# entries are left commented out to show the array format for more than one server.
$DefaultComputerList = @(
    "SERVER01"
    #"SERVER02"
    #"SERVER03"
    #"SERVER04"
    #"SERVER05"
    #"SERVER06"
    #"SERVER07"
    #"SERVER08\INSTANCENAME"
)

# Connection resiliency
$RetryCount               = 2
$RetryDelaySeconds         = 5
$ConnectionTimeoutSeconds  = 15

# Set to 1 to export the summary results to CSV/HTML.
$ExportResults = 0
$ExportFolder  = "C:\temp\Check-For-Reboot\Reports"
$ExportFormat  = 'Both'   # 'CSV', 'HTML', or 'Both'

# Set to 1 to email the report. Requires $SmtpServer/$EmailFrom/$EmailTo to be set.
$SendEmail   = 0
$SmtpServer  = ''
$EmailFrom   = ''
$EmailTo     = @()

# -----------------------------------------------------------------------------

$computerList = Get-PendingRebootComputerList -Path $ComputerListPath -DefaultList $DefaultComputerList

$results = Test-PendingReboot -ComputerName $computerList `
                               -RetryCount $RetryCount `
                               -RetryDelaySeconds $RetryDelaySeconds `
                               -ConnectionTimeoutSeconds $ConnectionTimeoutSeconds

if ($OnlyShowFlagged -eq 1) {
    # Keep servers that are actually flagged, plus any that failed to connect —
    # both need attention, whereas 'N/A' placeholder values on error rows should
    # never be mistaken for a real "flagged" result.
    $results = $results | Where-Object { $_.RebootPending_Overall -eq $true -or $_.ConnectionError }
}

# Definitions legend — always printed when file ops are flagged, regardless of ShowPendingFiles
if ($results | Where-Object { $_.PendingFileRenameOperations_Exist -eq $true }) {
    Write-PendingRebootLegend
}

# Color-coded summary — one row per server
Write-PendingRebootSummary -Results $results

# Explicit connection-failure report — uses Write-Host (not Write-Warning) so it
# is guaranteed to appear in Start-Transcript output regardless of host
# (PowerShell ISE in particular does not reliably capture the warning stream).
Write-PendingRebootConnectionFailures -Results $results

# File operation result set — one block per server, Rename then Delete
# Only shown when ShowPendingFiles = 1
if ($ShowPendingFiles -eq 1) {
    Write-PendingFileOperations -Results $results
}

# CBS.log activity vs. last boot time — helps explain "still pending right
# after a reboot" by telling new post-reboot activity apart from a stale/
# leftover servicing state. Shown for any currently-flagged server.
Write-PendingRebootActivityCheck -Results $results

# Export to CSV/HTML
$exportedPaths = $null
if ($ExportResults -eq 1) {
    $exportedPaths = Export-PendingRebootReport -Results $results -OutputFolder $ExportFolder -Format $ExportFormat
    Write-Host "Report exported: $($exportedPaths | Out-String)" -ForegroundColor Cyan
}

# Email report
if ($SendEmail -eq 1) {
    if (-not ($SmtpServer -and $EmailFrom -and $EmailTo)) {
        Write-Warning "SendEmail is enabled but SmtpServer/EmailFrom/EmailTo are not fully configured. Skipping email."
    }
    else {
        $attachment = if ($exportedPaths -and $exportedPaths.CSV) { $exportedPaths.CSV } else { $null }
        Send-PendingRebootEmail -Results $results -SmtpServer $SmtpServer -From $EmailFrom -To $EmailTo -AttachmentPath $attachment
    }
}
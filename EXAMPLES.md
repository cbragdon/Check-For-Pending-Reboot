# Examples

Working, copy-pasteable PowerShell constructs for common ways to run
`Multi_Server_Pending_Reboot_1.3.3.ps1`. Replace the placeholder server names
below (`SERVER01`, `SERVER02`, `SERVER03`) with your own.

> **Note:** `$ServerNameListPath` and `$ComputerName` are mutually exclusive —
> only one is ever used per run (the path wins if it's set and the file
> exists). Both are session variables: if you set one in your PowerShell
> console before running the script, that value is used instead of the
> script's own default — but it also means a value set in an earlier command
> in the same session will persist and get reused unless you clear it
> (`Remove-Variable ComputerName -ErrorAction SilentlyContinue`).

## 1. Load functions only, without running the config/execution section

Useful for interactive testing — dot-source the script to load every function
(`Test-PendingReboot`, `Write-PendingRebootSummary`, etc.) without triggering
the `USAGE` section at the bottom.

```powershell
# Load the functions without running the USAGE section at the bottom
. C:\temp\Check-For-Pending-Reboot\Multi_Server_Pending_Reboot_1.3.3.ps1 2>$null

# Run against specific servers and see everything
Test-PendingReboot -ComputerName "SERVER01", "SERVER02", "SERVER03" | Format-List *
```

## 2. Set `$ComputerName` directly, then run the script with a transcript

```powershell
$ComputerName = @(
    "SERVER01", "SERVER02", "SERVER03"
)

Start-Transcript -Path C:\temp\Check-For-Pending-Reboot\TestRun.txt -Force
.\Multi_Server_Pending_Reboot_1.3.3.ps1
Stop-Transcript
```

`TestRun.txt` is git-ignored (see `.gitignore`) since transcripts capture real
hostnames — safe to use for local testing, never committed.

## 3. Use a server list file instead of `$ComputerName` (console output only)

```powershell
cd C:\temp\Check-For-Pending-Reboot

# Clear any leftover $ComputerName from earlier testing (session variables
# persist until you close the console or clear them — this is NOT hardcoded
# in the script)
Remove-Variable ComputerName -ErrorAction SilentlyContinue

# Create the server list file (one-time setup — edit the server names as needed)
@"
SERVER01
SERVER02
SERVER03
"@ | Set-Content -Path C:\temp\Check-For-Pending-Reboot\servers.txt

# Confirm it actually exists before relying on it
Test-Path C:\temp\Check-For-Pending-Reboot\servers.txt

# Point the script at the file
$ServerNameListPath = "C:\temp\Check-For-Pending-Reboot\servers.txt"

# Run it — output prints straight to the console, nothing written to disk
.\Multi_Server_Pending_Reboot_1.3.3.ps1
```

Expect to see `Using server list from $ServerNameListPath: ...` in the output
(not a fallback warning) confirming the file — not a leftover session
variable — was used.

### Server list file format options

**Plain text, one server per line** (`#` = comment, blank lines ignored):

```
# servers.txt
SERVER01
SERVER02
SERVER03
# SERVER04  <- commented out, skipped
SERVER08\INSTANCENAME
```

**CSV with a `ComputerName` column** (header name must match exactly):

```
ComputerName
SERVER01
SERVER02
SERVER03
```

Both `servers.txt` and any `.csv` server list are git-ignored — real server
names used locally are never committed.

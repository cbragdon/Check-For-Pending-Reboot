# Examples

Five working, copy-pasteable PowerShell constructs for running
`Multi_Server_Pending_Reboot_1.3.3.ps1`, each tested and confirmed working.
Replace the placeholder server names below (`SERVER01`, `SERVER02`,
`SERVER03`) with your own.

> **Note:** `$ServerNameListPath` and `$ComputerName` are mutually exclusive —
> only one is ever used per run (the path wins if it's set and the file
> exists). Both are session variables: if you set one in your PowerShell
> console before running the script, that value is used instead of the
> script's own default — but it also means a value set in an earlier command
> in the same session will persist and get reused unless you clear it
> (`Remove-Variable ComputerName -ErrorAction SilentlyContinue`).

## 1. Load functions only, then call `Test-PendingReboot` directly

**What it does:** Dot-sources the script to load every function
(`Test-PendingReboot`, `Write-PendingRebootSummary`, etc.) into your current
session *without* running the `USAGE` section's config/execution at the
bottom. Use this for ad hoc/interactive testing when you want to call a
function yourself with specific servers, instead of relying on
`$ServerNameListPath`/`$ComputerName`.

```powershell
# Load the functions without running the USAGE section at the bottom
. C:\temp\Check-For-Pending-Reboot\Multi_Server_Pending_Reboot_1.3.3.ps1 2>$null

# Run against specific servers and see everything
Test-PendingReboot -ComputerName "SERVER01", "SERVER02", "SERVER03" | Format-List *
```

The script detects dot-sourcing (`$MyInvocation.InvocationName -eq '.'`) and
returns immediately after the function definitions, so no server list needs
to be configured at all for this use case.

## 2. Set `$ComputerName` directly, then run the script with a transcript

**What it does:** Sets the server list in-session (no file needed), then runs
the script normally so it goes through the full flow (summary table, legend,
connection failures, etc.) while `Start-Transcript`/`Stop-Transcript` capture
everything to a log file for later review.

```powershell
$ComputerName = @(
    "SERVER01", "SERVER02", "SERVER03"
)

Start-Transcript -Path C:\temp\Check-For-Pending-Reboot\TestRun.txt -Force
.\Multi_Server_Pending_Reboot_1.3.3.ps1
Stop-Transcript
```

`TestRun.txt` is git-ignored (see `.gitignore`).

## 3. Use a server list file, with a transcript

**What it does:** Instead of setting `$ComputerName` in-session, this creates
a reusable server list file on disk and points `$ServerNameListPath` at it —
useful when you want the list to persist across sessions instead of
retyping it every time. Output is captured to a transcript, same as example 2.

```powershell
cd C:\temp\Check-For-Pending-Reboot

# Create the server list file (one-time setup — edit the server names as needed)
@"
SERVER01
SERVER02
SERVER03
"@ | Set-Content -Path C:\temp\Check-For-Pending-Reboot\servers.txt

# Point the script at that file
$ServerNameListPath = "C:\temp\Check-For-Pending-Reboot\servers.txt"

Start-Transcript -Path C:\temp\Check-For-Pending-Reboot\TestRun.txt -Force
.\Multi_Server_Pending_Reboot_1.3.3.ps1
Stop-Transcript
```

`servers.txt` is also git-ignored (see `.gitignore`).

## 4. Use a server list file, console output only (with clean session state)

**What it does:** Same server-list-file approach as example 3, but skips the
transcript entirely so output just prints live to the console — and adds two
safety steps worth knowing about: clearing any leftover `$ComputerName` from
an earlier command in the same session (it is a **session variable**, not
something hardcoded in the script, so it persists until cleared or the
console is closed), and confirming the file actually exists with `Test-Path`
before running — since if `$ServerNameListPath` points at a file that isn't
found, the script warns and silently falls back to `$ComputerName` instead of
failing loudly.

```powershell
cd C:\temp\Check-For-Pending-Reboot

# Clear any leftover $ComputerName from earlier testing
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
(not a fallback warning), confirming the file — not a leftover session
variable — was actually used.

## 5. Point to a server list file you already created yourself

**What it does:** Skips the file-creation step entirely — for when you (or
someone else) already maintain a server list file somewhere (any path,
any of your own naming), and just need to tell the script where it is. Same
mutual-exclusivity rule applies: as long as the path is valid, it's used
instead of `$ComputerName`.

```powershell
cd C:\temp\Check-For-Pending-Reboot

# Point at an existing server list file — replace with your actual path
$ServerNameListPath = "D:\Ops\Inventory\my-servers.txt"

.\Multi_Server_Pending_Reboot_1.3.3.ps1
```

The file just needs to follow one of the two formats below (plain text or
CSV) — see the next section.

## Server list file format options

Applies to examples 3, 4, and 5 above (`$ServerNameListPath`).

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

Both `servers.txt` and any `.csv` server list are git-ignored.

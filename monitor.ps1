#Requires -Version 5.1
<#
.SYNOPSIS
  Monitors tem-ord-cards GitHub Actions and runs PDF generator tests against
  integration whenever a new deployment completes.

.NOTES
  Trigger:   Windows Task Scheduler, every 15 min (runs only while laptop is on/logged in)
  Signal:    "CI/CD validate-pipeline" workflow on master — conclusion=success
             + job "systems_integ_integrate" conclusion=success (per Confluence blueprint docs)
  Optional:  /healthcheck/heartbeat probe on VPN (best-effort; skipped if not reachable)
  Tests:     KioskTest, PdfFormatTest, PdfGenericTest, PdfNamesTest, RegressionTest

  NOTE: Confluence docs call this pipeline "CI/CD int-deploy-and-validation" and use
  the job name "systems_int_integrate". In tem-ord-cards it is actually named
  "CI/CD validate-pipeline" / "systems_integ_integrate". Same blueprint, different name.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Configuration -------------------------------------------------------
$REPO              = 'fs-eng/tem-ord-cards'
$WORKFLOW_NAME     = 'CI/CD validate-pipeline'
$INTEG_JOB_GLOB    = 'systems_integ_integrate'   # confirmed from GitHub; Confluence calls it systems_int_integrate
$HEALTHCHECK_URL   = 'https://cards.temple.service.integ.us-east-1.dev.fslocal.org/healthcheck/heartbeat'
$HEALTHCHECK_SEC   = 6     # timeout — if VPN is not up this just skips

$SCRIPT_DIR    = Split-Path -Parent $MyInvocation.MyCommand.Path
$STATE_FILE    = Join-Path $SCRIPT_DIR 'state.json'
$LOG_FILE      = Join-Path $SCRIPT_DIR 'monitor.log'
$RESULTS_DIR   = Join-Path $SCRIPT_DIR 'test-results'
$PROJECT_DIR   = 'C:\Users\josehoon\OneDrive - Church of Jesus Christ\Documents\git\tem-test-pdf-generator'
$TESTS         = 'KioskTest,PdfFormatTest,PdfGenericTest,PdfNamesTest,RegressionTest'
$MVN           = 'C:\Tools\apache-maven-3.9.5\bin\mvn.cmd'
$LOG_MAX_LINES = 5000
# -------------------------------------------------------------------------

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  [$Level]  $Message"
    Add-Content -Path $LOG_FILE -Value $line -Encoding UTF8
    Write-Host $line
}

function Trim-Log {
    if (-not (Test-Path $LOG_FILE)) { return }
    $lines = Get-Content $LOG_FILE
    if ($lines.Count -gt $LOG_MAX_LINES) {
        $lines | Select-Object -Last $LOG_MAX_LINES | Set-Content $LOG_FILE -Encoding UTF8
    }
}

function Get-GitHubToken {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class WinCredMgr {
    [DllImport("advapi32.dll", EntryPoint="CredReadW", CharSet=CharSet.Unicode, SetLastError=true)]
    static extern bool CredRead(string target, int type, int flags, out IntPtr pCred);
    [DllImport("advapi32.dll")]
    static extern void CredFree(IntPtr pCred);
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
    struct CREDENTIAL {
        public int Flags, Type;
        public string TargetName, Comment;
        public long LastWritten;
        public int CredentialBlobSize;
        public IntPtr CredentialBlob;
        public int Persist, AttributeCount;
        public IntPtr Attributes;
        public string TargetAlias, UserName;
    }
    public static string Read(string target) {
        IntPtr p;
        if (!CredRead(target, 1, 0, out p)) return null;
        var c = Marshal.PtrToStructure<CREDENTIAL>(p);
        var b = new byte[c.CredentialBlobSize];
        Marshal.Copy(c.CredentialBlob, b, 0, c.CredentialBlobSize);
        CredFree(p);
        return Encoding.Unicode.GetString(b);
    }
}
'@ -ErrorAction SilentlyContinue
    return [WinCredMgr]::Read('git:https://github.com')
}

function Invoke-GitHubApi {
    param([string]$Path, [string]$Token)
    $h = @{
        'Authorization'        = "Bearer $Token"
        'Accept'               = 'application/vnd.github+json'
        'X-GitHub-Api-Version' = '2022-11-28'
    }
    return Invoke-RestMethod -Uri "https://api.github.com$Path" -Headers $h -ErrorAction Stop
}

function Get-LatestValidateRun {
    param([string]$Token)
    $runs = Invoke-GitHubApi -Path "/repos/$REPO/actions/runs?branch=master&status=success&per_page=20" -Token $Token
    return $runs.workflow_runs |
        Where-Object { $_.name -eq $WORKFLOW_NAME } |
        Sort-Object { [datetime]$_.updated_at } -Descending |
        Select-Object -First 1
}

function Confirm-IntegDeployed {
    # Step 1: verify systems_integ_integrate job passed inside the workflow run.
    # This is the blueprint "systems_int_integrate" stage (Confluence CI/CD int-deploy-and-validation doc).
    param([string]$RunId, [string]$Token)
    try {
        $jobs = Invoke-GitHubApi -Path "/repos/$REPO/actions/runs/$RunId/jobs?per_page=30" -Token $Token
        $integJob = $jobs.jobs | Where-Object { $_.name -like "*$INTEG_JOB_GLOB*" } | Select-Object -First 1
        if (-not $integJob) {
            Write-Log "Job '$INTEG_JOB_GLOB' not found in run $RunId — proceeding anyway" 'WARN'
            return $true
        }
        if ($integJob.conclusion -ne 'success') {
            Write-Log "Job '$($integJob.name)' conclusion=$($integJob.conclusion) — skipping tests" 'WARN'
            return $false
        }
        Write-Log "Job '$($integJob.name)' confirmed success"
        return $true
    } catch {
        Write-Log "Could not fetch jobs for run $RunId : $_ — proceeding anyway" 'WARN'
        return $true
    }
}

function Test-Healthcheck {
    # Step 2 (optional): probe the integ heartbeat endpoint.
    # Requires VPN; silently skips if not reachable.
    # Hostname pattern: *.service.integ.us-east-1.dev.fslocal.org (per Confluence deployment docs)
    try {
        $r = Invoke-WebRequest -Uri $HEALTHCHECK_URL -TimeoutSec $HEALTHCHECK_SEC -UseBasicParsing -ErrorAction Stop
        if ($r.StatusCode -eq 200) {
            Write-Log "Healthcheck OK (HTTP 200): $HEALTHCHECK_URL"
            return $true
        }
        Write-Log "Healthcheck returned HTTP $($r.StatusCode) — proceeding" 'WARN'
        return $true
    } catch [System.Net.WebException] {
        if ($_.Exception.Status -eq 'Timeout' -or $_.Exception.Status -eq 'NameResolutionFailure') {
            Write-Log "Healthcheck skipped (VPN not reachable): $HEALTHCHECK_URL"
        } else {
            Write-Log "Healthcheck error ($($_.Exception.Status)) — proceeding" 'WARN'
        }
        return $true   # not a blocking failure — VPN may simply be off
    } catch {
        Write-Log "Healthcheck probe failed — proceeding: $_" 'WARN'
        return $true
    }
}

function Show-ToastNotification {
    param([string]$Title, [string]$Body)
    try {
        [Windows.UI.Notifications.ToastNotificationManager,Windows.UI.Notifications,ContentType=WindowsRuntime] | Out-Null
        $tpl = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent(
            [Windows.UI.Notifications.ToastTemplateType]::ToastText02)
        $tpl.GetElementsByTagName('text')[0].AppendChild($tpl.CreateTextNode($Title)) | Out-Null
        $tpl.GetElementsByTagName('text')[1].AppendChild($tpl.CreateTextNode($Body))  | Out-Null
        $n = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('tem-ord-cards Monitor')
        $n.Show([Windows.UI.Notifications.ToastNotification]::new($tpl))
    } catch { }
}

function Run-PdfTests {
    param([string]$Sha)

    $shortSha = $Sha.Substring(0, 8)
    Write-Log "Running PDF tests for sha $shortSha"

    New-Item -ItemType Directory -Force $RESULTS_DIR | Out-Null
    $reportDir = Join-Path $RESULTS_DIR $shortSha
    New-Item -ItemType Directory -Force $reportDir | Out-Null

    $mvnArgs = @(
        'test',
        '-Dacceptance.testing',
        "-Dtest=$TESTS",
        '-f', "`"$PROJECT_DIR\pom.xml`""
    )

    $startTime = Get-Date
    $proc = Start-Process -FilePath $MVN -ArgumentList $mvnArgs `
        -WorkingDirectory $PROJECT_DIR `
        -RedirectStandardOutput (Join-Path $reportDir 'stdout.txt') `
        -RedirectStandardError  (Join-Path $reportDir 'stderr.txt') `
        -PassThru -NoNewWindow -Wait

    $elapsed = [int](New-TimeSpan -Start $startTime -End (Get-Date)).TotalMinutes

    $surefireDir = Join-Path $PROJECT_DIR 'target\surefire-reports'
    if (Test-Path $surefireDir) {
        Copy-Item "$surefireDir\*.xml" $reportDir -ErrorAction SilentlyContinue
    }

    $passed = 0; $failed = 0; $errors = 0; $skipped = 0
    function ToInt($v) { if ($v) { [int]$v } else { 0 } }
    Get-ChildItem $reportDir -Filter '*.xml' | ForEach-Object {
        [xml]$xml = Get-Content $_.FullName
        $ts = $xml.testsuite
        if ($ts) {
            $f = ToInt $ts.failures; $e = ToInt $ts.errors; $s = ToInt $ts.skipped
            $passed  += (ToInt $ts.tests) - $f - $e - $s
            $failed  += $f
            $errors  += $e
            $skipped += $s
        }
    }

    if ($proc.ExitCode -eq 0) {
        $summary = "PASSED: $passed tests  (${elapsed}m)  sha=$shortSha"
        Write-Log $summary 'PASS'
        Show-ToastNotification 'tem-ord-cards PDF Tests PASSED' $summary
    } else {
        $summary = "FAILED: $failed failures, $errors errors, $passed passed  (${elapsed}m)  sha=$shortSha"
        Write-Log $summary 'FAIL'
        Show-ToastNotification 'tem-ord-cards PDF Tests FAILED' $summary
        Write-Log "Reports: $reportDir" 'FAIL'
    }

    return $proc.ExitCode -eq 0
}

# =========================================================================
# Main
# =========================================================================

Trim-Log
Write-Log 'Monitor check started'

# 1. Load state
$state = if (Test-Path $STATE_FILE) {
    Get-Content $STATE_FILE -Raw | ConvertFrom-Json
} else {
    [PSCustomObject]@{ lastTestedSha = ''; lastKnownSha = '' }
}

# 2. GitHub token
$token = $null
try { $token = Get-GitHubToken } catch {}
if (-not $token) {
    Write-Log 'Could not read GitHub token from Windows Credential Manager.' 'WARN'
    exit 1
}

# 3. Latest successful validate-pipeline run on master
$run = $null
try {
    $run = Get-LatestValidateRun -Token $token
} catch {
    Write-Log "GitHub API error: $_" 'ERROR'
    exit 1
}

if (-not $run) {
    Write-Log "No completed '$WORKFLOW_NAME' run found on master."
    exit 0
}

$latestSha  = $run.head_sha
$shortSha   = $latestSha.Substring(0, 8)
$state.lastKnownSha = $latestSha

Write-Log "Latest deploy: sha=$shortSha  completed=$($run.updated_at)"

# 4. Already tested?
if ($latestSha -eq $state.lastTestedSha) {
    Write-Log "SHA $shortSha already tested — nothing to do."
    $state | ConvertTo-Json | Set-Content $STATE_FILE -Encoding UTF8
    exit 0
}

# 5. Confirm systems_integ_integrate job succeeded (blueprint validation stage)
if (-not (Confirm-IntegDeployed -RunId $run.id -Token $token)) {
    Write-Log 'Integ job check failed — skipping test run.' 'WARN'
    $state | ConvertTo-Json | Set-Content $STATE_FILE -Encoding UTF8
    exit 0
}

# 6. Best-effort healthcheck probe (requires VPN; non-blocking if unreachable)
Test-Healthcheck | Out-Null

# 7. Run PDF tests
$testsPassed = Run-PdfTests -Sha $latestSha

# 8. Persist state
$state.lastTestedSha = $latestSha
$state | ConvertTo-Json | Set-Content $STATE_FILE -Encoding UTF8

Write-Log 'Monitor check complete'
exit $(if ($testsPassed) { 0 } else { 1 })

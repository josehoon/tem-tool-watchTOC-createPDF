# tem-tool-watchTOC-createPDF

This tool watches tem-ord-cards build and deploy status to integration and launches
[tem-tool-pdf-generator](https://github.com/fs-eng/tem-tool-pdf-generator) to create
new PDF ordinance cards to validate the new tem-ord-cards checkin.

## How it works

1. **monitor.ps1** runs every 15 minutes via Windows Task Scheduler (while the laptop is on)
2. It queries the GitHub Actions API for the `CI/CD validate-pipeline` workflow on the
   `fs-eng/tem-ord-cards` master branch
3. When a new successful deployment to integration is detected (confirmed by the
   `systems_integ_integrate` job), it runs the PDF generator acceptance tests
4. Results are stored in `test-results/<sha>/` and a Windows toast notification fires
5. **status-window.ps1** provides a small floating desktop window showing live monitor state

## Files

| File | Purpose |
|---|---|
| `monitor.ps1` | Main watcher — polls GitHub, detects new deploys, runs tests |
| `status-window.ps1` | WPF floating status window (IDLE / RUNNING / PENDING / PASSED / FAILED) |
| `install-scheduler.ps1` | One-time Task Scheduler setup (run as Administrator) |

## Setup

### 1. Prerequisites

- Windows with PowerShell 5.1+
- A GitHub HTTPS credential stored in Windows Credential Manager as `git:https://github.com`
  (present automatically if you have used `git clone` over HTTPS with the Windows Git Credential Manager)
- Maven at `C:\Tools\apache-maven-3.9.5\bin\mvn.cmd`
- [tem-tool-pdf-generator](https://github.com/fs-eng/tem-tool-pdf-generator) checked out at
  `C:\Users\josehoon\OneDrive - Church of Jesus Christ\Documents\git\tem-test-pdf-generator`
- VPN (optional — healthcheck probe is best-effort and silently skipped if VPN is not up)

### 2. Register the scheduled task (once, as Administrator)

```powershell
powershell -ExecutionPolicy Bypass -File install-scheduler.ps1
```

### 3. Open the status window

```powershell
powershell -ExecutionPolicy Bypass -File status-window.ps1
```

To create a Desktop shortcut:

```powershell
powershell -ExecutionPolicy Bypass -File status-window.ps1 -CreateShortcut
```

## Tests run

`KioskTest`, `PdfFormatTest`, `PdfGenericTest`, `PdfNamesTest`, `RegressionTest`

## Pipeline detection

| GitHub name | Confluence doc name |
|---|---|
| `CI/CD validate-pipeline` | `CI/CD int-deploy-and-validation` |
| job: `systems_integ_integrate` | job: `systems_int_integrate` |

The Confluence blueprint and the actual GitHub workflow use slightly different names;
monitor.ps1 targets the actual GitHub names.

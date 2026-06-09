---
name: sharepoint-cert-auth-recovery
description: 'Diagnose and fix SharePoint automation authentication issues in PowerShell/Python projects. Use when WebLogin warnings persist, certificate auth is not applied, PnP connection fails, or batch/script path issues break execution.'
argument-hint: 'Provide project path, runner script, and failing log snippet.'
user-invocable: true
---

# SharePoint Certificate Auth Recovery

## Goal
Create a repeatable workflow to migrate/fix SharePoint automation auth from legacy WebLogin to certificate-based app-only auth, with strong troubleshooting for common failures.

## When To Use
- Error contains `UseWebLogin`, `WebLogin`, or interactive login warnings.
- Pipeline claims to use certificate but still authenticates by browser login.
- Batch runner cannot find script paths (`-File ... does not exist`).
- PowerShell parser errors due to version incompatibilities.
- Mixed Python + PowerShell stacks where auth mode is configured in env files.

## Inputs Required
- Entry runner file (usually `.bat` or `.ps1`).
- SharePoint upload script (`Connect-PnPOnline` call site).
- Auth config source (`.env` or pipeline variables).
- Error logs from latest execution.

## Procedure
1. Map auth flow end-to-end.
- Find where strategy/mode is selected (`AUTH_MODE`, `AZURE_MODE`, etc).
- Find actual upload connection command (`Connect-PnPOnline ...`).
- Confirm which runner starts the flow (`.bat` -> `.ps1` -> Python).

2. Remove legacy WebLogin usage.
- Replace `Connect-PnPOnline -UseWebLogin` with modern options:
  - Preferred: `-ClientId -Tenant -CertificatePath [-CertificatePassword]`
  - Fallback: `-DeviceLogin`
- Keep a single connection helper function to avoid duplicated auth logic.

3. Ensure certificate mode is truly selected.
- Read env from process and `.env`.
- Respect both `AUTH_MODE` and `AZURE_MODE`.
- Auto-detect cert mode if cert variables exist.
- Required variables:
  - `AZURE_CLIENT_ID`
  - `AZURE_TENANT_ID`
  - `AZURE_CERT_PFX_PATH`
  - `AZURE_CERT_PFX_PASSWORD` (if protected)

4. Normalize runner paths.
- In top-level `.bat`, resolve project root from `%~dp0` carefully.
- Validate `powershell -File <path>` points to an existing script.
- Avoid fragile relative paths that depend on launch directory.

5. Validate PowerShell version compatibility.
- If environment is Windows PowerShell 5.1, avoid newer operators like `??`.
- Replace with compatible syntax:
  - From: `($value ?? '')`
  - To: cast + null-safe trim sequence.

6. Verify prerequisites before full run.
- `Test-Path` for PFX file.
- `Get-Module -ListAvailable -Name PnP.PowerShell`.
- Confirm env values are loaded before upload step.

7. Run end-to-end and inspect logs.
- Success criteria:
  - Log shows `PnP Certificate App-Only` connection path.
  - No `UseWebLogin` references.
  - Upload command completes and returns success.

## Quick Checks
Use these commands in PowerShell:

```powershell
# Check certificate file and PnP module
$pfx = '.certs/your-cert.pfx'
"PFX_EXISTS=$((Test-Path $pfx))"
$mod = Get-Module -ListAvailable -Name PnP.PowerShell | Select-Object -First 1
if ($null -eq $mod) { 'PNP_MODULE=NOT_FOUND' } else { "PNP_MODULE=FOUND $($mod.Version)" }
```

```powershell
# Search for legacy login usage in scripts
Get-ChildItem -Path scripts -Recurse -File | Select-String -Pattern 'UseWebLogin|WebLogin|Connect-PnPOnline'
```

## Known Error Signatures And Fixes
- `The argument '.\scripts\...ps1' to the -File parameter does not exist`
  - Fix: correct `.bat` root resolution from `%~dp0` and re-test file existence.

- `Unexpected token '??' in expression or statement`
  - Fix: remove PowerShell 7-only operators for PowerShell 5.1 compatibility.

- `WebLogin warning` or browser prompt still appears
  - Fix: remove `-UseWebLogin`, route through certificate helper, and verify mode/env values.

- `Arquivo de certificado nao encontrado`
  - Fix: correct `AZURE_CERT_PFX_PATH` and resolve relative path against project root.

## Reference Implementation Pattern
Use one helper for SharePoint connection:
1. Read mode/env config.
2. If cert config exists, connect with app-only certificate.
3. Else fallback to `-DeviceLogin`.
4. Keep upload/move/archive logic separate from auth logic.

## Definition Of Done
- No legacy WebLogin path in codebase.
- Cert auth variables documented and loaded from env.
- Runner path issues fixed.
- Script compatible with target PowerShell version.
- End-to-end run proves SharePoint publish works with certificate path.

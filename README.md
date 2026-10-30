# DUODeploy

A one-window tool that deploys **Duo Windows Logon** and the **Duo Authentication Proxy** on a domain controller, then
distributes Duo Windows Logon to every domain machine through a Group Policy startup script.

Built as a PowerShell script (`DeployDuo.ps1`) compiled to an exe with [PS2EXE](https://github.com/MScholtes/PS2EXE).

> **Independent project.** DUODeploy was built by TR Computer Sales as a convenience for our own deployments. It is
> **unofficial** and is **not affiliated with, endorsed by, or supported by Duo Security or Cisco**. It automates the
> steps in Duo's own documentation, [Duo Authentication for Windows Logon: Group Policy](https://duo.com/docs/winlogon-gpo)
> and [Duo Directory Sync / Active Directory](https://duo.com/docs/adsync), and Duo's names and services are
> referenced only to describe what is being deployed. Questions or problems with this tool belong with TR Computer
> Sales, not Duo support. Review what it does and test it before using it in production.

## What you need

- A domain controller, run **as administrator** (the exe requests elevation).
- Internet access from the DC (downloads the Duo installers and the GPO backup from this repo).
- The Duo **Microsoft RDP** application values: `HOST`, `IKEY`, `SKEY`.
- An `authproxy.cfg` for the client. The filename does not matter (`authproxy (13).cfg` is fine); you pick it in the window.
- RSAT modules on the DC: ActiveDirectory and GroupPolicy (present on any DC).

## Using it

1. Run `DeployDuo.exe` (or `DeployDuo.ps1` from an elevated PowerShell).
2. Fill in `HOST`, `IKEY`, `SKEY`, then **Browse...** to the client's `authproxy.cfg`.
3. Click **START**. Progress and results show in the colour-coded log pane.

Everything the tool creates or downloads lives in `C:\ProgramData\TRCS\Duo`, including a timestamped
`DeployDuo_<timestamp>.log` per run. Your original `authproxy.cfg` is never modified; a working copy is made there and
deleted once it has been deployed.

Optional parameters pre-fill the window: `-DUO_HOST`, `-IKEY`, `-SKEY`, `-CfgFile`, and `-AutoStart` to begin
immediately when all four values are valid.

## What it does (13 steps)

| # | Step |
|---|------|
| 1 | Check preconditions; copy the chosen `authproxy.cfg` to a working copy |
| 2 | Download the Duo Windows Logon bundle, the Auth Proxy installer, and the GPO backup |
| 3 | Extract the downloads |
| 4 | Copy the ADMX/ADML templates (local `PolicyDefinitions`, plus the domain Central Store if one exists) and the MSIs to the real NETLOGON folder |
| 5 | Build one transform per architecture: `duo1.mst` (x64), `duo1-arm64.mst`, `duo1-x86.mst`; verify all six files are visible at `\\<server>\NETLOGON` |
| 6 | Install the Duo Authentication Proxy |
| 7 | Create the `duo_sync` service account and `DuoProtected` group; set `TR_NOC` email and fill blank Domain Admin phone numbers |
| 8 | Update `authproxy.cfg`, deploy it, and encrypt the service password with `authproxy_passwd` |
| 9 | Remove any existing GPO with "duo" in its name, then create, delegate and link **DUO Windows Login** at the domain root |
| 10 | Import the GPO backup |
| 11 | Write `HOST` / `IKEY` / `SKEY` into the GPO |
| 12 | Deploy the startup script to the GPO and run it on this server, so the DC is installed the same way as every other machine |
| 13 | Start the `duoauthproxy` service and open the Proxy Manager |

### The startup script

Duo Windows Logon is installed by a GPO **startup script** rather than GPO Software Installation, so upgrades are just
"drop the new MSI in NETLOGON" and removal goes through your RMM.

- Picks the MSI/MST by CPU architecture: x64, ARM64 (Snapdragon), or x86. Any other architecture is skipped.
- Reads the share from the machine's own domain (`\\<domain>\NETLOGON`), so nothing is hard-coded per client.
- Uses Windows Installer's own state check, so it does **not** reinstall on every boot. Already current: "nothing to do".
  Older build installed: upgrades. Not installed: installs.
- Clears stale registrations (for example a leftover Software Installation advertisement from an older deployment)
  that would otherwise make the MSI return success while installing nothing, and verifies the product is really
  installed afterwards.
- Logs to `C:\ProgramData\TRCS\Duo\startup.log` on each machine, with the MSI log next to it.

Startup scripts run as SYSTEM, so the machine's PowerShell execution policy does not matter.

### Accounts and groups

- **`duo_sync`**: a read-only directory-sync account for the Auth Proxy. Standard Domain Users membership only, password
  never expires, user cannot change it, removed from any privileged group, and denied interactive and Remote Desktop
  logon in the GPO. If another GPO already defines those deny-logon rights, the tool leaves them alone and warns, so
  add `duo_sync` there by hand. A new random password is generated each run; it only ever exists in `authproxy.cfg`
  and is encrypted in place.
- **`DuoProtected`**: contains the Domain Admins group. `duo_sync` and the legacy `DuoAdmin` are kept out, because every
  protected user counts toward Duo billing.
- **`DuoAdmin`** (older versions): if found it is disabled, not deleted.
- **`TR_NOC`**: gets its email set to `trhelpdesk@trcomputersales.com`. Its phone numbers are left alone.
- **Domain Admins** (enabled users, other than `TR_NOC`): blank `telephoneNumber` / `mobile` fields are filled with
  the NOC number. Existing numbers are never overwritten.

## Running it on a domain that already has Duo

Safe to run as an overlay. Any existing GPO with "duo" in its name is **backed up** to
`C:\ProgramData\TRCS\Duo\RemovedGpoBackups\<timestamp>`, unlinked everywhere, and deleted, then recreated fresh. The name
match is deliberately broad, so check the log if a client has an unrelated GPO that happens to contain "duo".

## Notes and known limits

- GPMC caches link state. If a new link doesn't show, close and reopen `gpmc.msc` (F5 is not enough).
- Central Store changes replicate through SYSVOL; give them a few minutes to appear in GPMC.
- ARM64 and x86 clients use their own MSI/MST pair, but those paths have not yet been tested on real hardware.
- No control file is used. To remove Duo from machines, uninstall it through your RMM.

## Building the exe

```powershell
Invoke-PS2EXE .\DeployDuo.ps1 .\DeployDuo.exe -noConsole -requireAdmin -DPIAware -STA -title "Duo Deployment"
```

`-noConsole` is required: the script writes nothing to the console streams and never calls `Read-Host` or a message box,
so PS2EXE doesn't pop dialogs. Keep the script name `DeployDuo.ps1`; git history carries the versions.

## Repo contents

- `DeployDuo.ps1`: the tool.
- `DeployDuo.exe`: compiled build of `DeployDuo.ps1`.
- `DUO_Windows_Login_GPO.zip`: GPO backup the tool downloads and imports at runtime. The tool pulls it from this repo,
  so don't rename or move it.

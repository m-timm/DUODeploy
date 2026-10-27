<#
.SYNOPSIS
    Duo Windows Logon + Authentication Proxy deployment for a domain controller, with a single GUI status window.

.DESCRIPTION
    GUI rewrite of DeployDuo_2026v6.ps1, built for PS2EXE -noConsole.
    Nothing is written to the output/warning/error/host streams and nothing calls Read-Host or
    MessageBox, so PS2EXE never pops a dialog. Duo values are entered in the window; all progress
    goes to one color-coded log pane plus DeployDuo_<timestamp>.log in the working directory.

    Working directory: C:\ProgramData\TRCS\Duo (created automatically).
    authproxy.cfg is picked in the window (Browse...). The original file is never modified; a working
    copy is made in the working directory and removed once it has been deployed.

    The deployment runs in a background runspace so the window stays responsive during
    downloads, MSI installs and GPO imports.

    Steps:
      1. Preconditions (authproxy.cfg selected, copied to the working directory)
      2. Download Duo Windows Logon bundle, Auth Proxy, GPO backup
      3. Extract ZIPs
      4. Copy ADMX/ADML to PolicyDefinitions, MSIs to SYSVOL\scripts (NETLOGON)
      5. Generate duo1.mst (original MSI untouched)
      6. Install Duo Authentication Proxy
      7. Create duo_sync service account and DuoProtected group; set TR_NOC email and Domain Admin phone numbers (not TR_NOC)
      8. Update/deploy authproxy.cfg, run authproxy_passwd
      9. Create, delegate and link GPO "DUO Windows Login" (an old Software Installation GPO is removed first)
     10. Import GPO backup
     11. Inject HOST / IKEY / SKEY into the GPO
     12. Add the Duo install/upgrade startup script to the GPO, then run it on this server
         (Duo Windows Logon is installed here the same way every other machine gets it)
     13. Start the Duo Authentication Proxy service, then launch the Proxy Manager

.PARAMETER IKEY
    Pre-fills the Microsoft RDP Application IKEY field.
.PARAMETER SKEY
    Pre-fills the Microsoft RDP Application SKEY field.
.PARAMETER DUO_HOST
    Pre-fills the Microsoft RDP Application HOST field.
.PARAMETER CfgFile
    Pre-fills the authproxy.cfg path.
.PARAMETER AutoStart
    Start immediately when all four values are supplied and valid.

.NOTES
    Build:
      Invoke-PS2EXE .\DeployDuo.ps1 .\DeployDuo.exe -noConsole -requireAdmin -DPIAware -STA -title "Duo Deployment"
#>
[CmdletBinding()]
param(
    [string]$IKEY,
    [string]$SKEY,
    [string]$DUO_HOST,
    [string]$CfgFile,
    [switch]$AutoStart
)

$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'
$WarningPreference     = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

#region ---------- Settings ----------
$script:DuoPath    = Join-Path $env:ProgramData 'TRCS\Duo'
$script:CfgPath    = Join-Path $script:DuoPath 'authproxy.working.cfg'   # working copy of the file picked in the UI
$script:TotalSteps = 13
$script:Job        = $null
$script:IsAdmin    = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
                        [Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not (Test-Path -LiteralPath $script:DuoPath)) { New-Item -Path $script:DuoPath -ItemType Directory -Force | Out-Null }
$script:LogFile = Join-Path $script:DuoPath ('DeployDuo_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
#endregion

#region ---------- Worker (runs in background runspace) ----------
$Worker = {
    param($sync)

    $ErrorActionPreference = 'Stop'
    $ProgressPreference    = 'SilentlyContinue'
    $WarningPreference     = 'SilentlyContinue'
    $ConfirmPreference     = 'None'

    $IKEY     = $sync.IKEY
    $SKEY     = $sync.SKEY
    $DUO_HOST = $sync.DUO_HOST
    $duoPath  = $sync.DuoPath
    $cfgPath  = $sync.CfgPath
    $srcCfg   = $sync.SourceCfg
    $logFile  = $sync.LogFile
    $failures = New-Object System.Collections.Generic.List[string]

    # ---------- messaging ----------
    function Write-Log ([string]$Message, [string]$Level = 'INFO') {
        $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        try { Add-Content -LiteralPath $logFile -Value $line -ErrorAction Stop } catch { }
        $sync.Queue.Enqueue([pscustomobject]@{ Kind = 'log'; Level = $Level; Text = $Message })
    }
    function Start-Phase ([string]$Text) {
        Write-Log $Text 'HEADER'
        $sync.Queue.Enqueue([pscustomobject]@{ Kind = 'status'; Text = $Text })
    }
    function Complete-Phase { $sync.Queue.Enqueue([pscustomobject]@{ Kind = 'step' }) }
    function Add-Failure ([string]$Tag) { $failures.Add($Tag) }

    # Runs an exe with no window; returns exit code + captured output.
    # -InputText is written to the process's stdin (for tools that prompt, e.g. authproxy_passwd).
    function Invoke-Quiet ([string]$FilePath, [string]$Arguments, [int]$TimeoutSec = 1800, [string]$InputText) {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = $FilePath
        $psi.Arguments              = $Arguments
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.RedirectStandardInput  = $PSBoundParameters.ContainsKey('InputText')
        $p = [System.Diagnostics.Process]::Start($psi)
        $out = $p.StandardOutput.ReadToEndAsync()
        $err = $p.StandardError.ReadToEndAsync()
        if ($psi.RedirectStandardInput) {
            try {
                $p.StandardInput.NewLine = "`r`n"
                foreach ($line in ($InputText -split "`r?`n")) { $p.StandardInput.WriteLine($line) }
                $p.StandardInput.Flush()
                $p.StandardInput.Close()
            } catch { }
        }
        if (-not $p.WaitForExit($TimeoutSec * 1000)) {
            try { $p.Kill() } catch { }
            throw ('{0} timed out after {1}s' -f [IO.Path]::GetFileName($FilePath), $TimeoutSec)
        }
        $p.WaitForExit()
        [pscustomobject]@{ ExitCode = $p.ExitCode; Output = ($out.Result + $err.Result).Trim() }
    }

    # ---------- MSI helpers ----------
    function Release-ComObject ($Object) {
        if ($null -ne $Object) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } catch { } }
    }

    function Test-MsiTableExists ($Database, [string]$TableName, $Installer) {
        $view = $null; $record = $null; $result = $null
        try {
            $view = $Database.OpenView("SELECT ``Name`` FROM ``_Tables`` WHERE ``Name``=?")
            $record = $Installer.CreateRecord(1)
            $record.StringData(1) = $TableName
            $view.Execute($record)
            $result = $view.Fetch()
            return ($null -ne $result)
        }
        finally { Release-ComObject $result; Release-ComObject $record; Release-ComObject $view }
    }

    function Set-MsiProperty ($Database, [string]$PropertyName, [string]$PropertyValue, $Installer) {
        $selectView = $null; $selectRecord = $null; $existing = $null; $modifyView = $null; $modifyRecord = $null
        try {
            $selectView = $Database.OpenView("SELECT ``Property``,``Value`` FROM ``Property`` WHERE ``Property``=?")
            $selectRecord = $Installer.CreateRecord(1)
            $selectRecord.StringData(1) = $PropertyName
            $selectView.Execute($selectRecord)
            $existing = $selectView.Fetch()

            $modifyRecord = $Installer.CreateRecord(2)
            if ($existing) {
                Write-Log "Updating MSI property '$PropertyName'"
                $modifyView = $Database.OpenView("UPDATE ``Property`` SET ``Value``=? WHERE ``Property``=?")
                $modifyRecord.StringData(1) = $PropertyValue
                $modifyRecord.StringData(2) = $PropertyName
            }
            else {
                Write-Log "Inserting MSI property '$PropertyName'"
                $modifyView = $Database.OpenView("INSERT INTO ``Property`` (``Property``,``Value``) VALUES (?,?)")
                $modifyRecord.StringData(1) = $PropertyName
                $modifyRecord.StringData(2) = $PropertyValue
            }
            $modifyView.Execute($modifyRecord)
        }
        finally {
            Release-ComObject $existing; Release-ComObject $selectRecord; Release-ComObject $selectView
            Release-ComObject $modifyRecord; Release-ComObject $modifyView
        }
    }

    function Get-MsiLastError ($Installer) {
        $record = $null
        try {
            $record = $Installer.LastErrorRecord()
            if ($null -eq $record) { return $null }
            $parts = @()
            for ($i = 0; $i -lt 20; $i++) {
                try {
                    $value = $record.StringData($i)
                    if (-not [string]::IsNullOrWhiteSpace($value)) { $parts += $value }
                } catch { break }
            }
            if ($parts.Count -gt 0) { return ($parts -join ' | ') }
            return $null
        }
        catch { return $null }
        finally { Release-ComObject $record }
    }

    function New-DuoTransform ([string]$SourceMsi, [string]$OutputMst) {
        $workingMsi = Join-Path $env:TEMP ('DuoWindowsLogon64_WORKING_{0}.msi' -f ([guid]::NewGuid().ToString('N')))
        $installer = $null; $originalDb = $null; $workingDb = $null
        try {
            Write-Log 'Copying source MSI to temp working copy for MST generation'
            Copy-Item -LiteralPath $SourceMsi -Destination $workingMsi -Force
            if (Test-Path -LiteralPath $OutputMst) {
                Write-Log "Removing existing MST: $OutputMst"
                Remove-Item -LiteralPath $OutputMst -Force
            }

            $installer  = New-Object -ComObject WindowsInstaller.Installer
            $originalDb = $installer.OpenDatabase($SourceMsi, 0)
            $workingDb  = $installer.OpenDatabase($workingMsi, 1)

            if (-not (Test-MsiTableExists -Database $workingDb -TableName 'Property' -Installer $installer)) {
                throw 'Property table does not exist in MSI.'
            }

            Set-MsiProperty -Database $workingDb -PropertyName 'IKEY' -PropertyValue $IKEY     -Installer $installer
            Set-MsiProperty -Database $workingDb -PropertyName 'SKEY' -PropertyValue $SKEY     -Installer $installer
            Set-MsiProperty -Database $workingDb -PropertyName 'HOST' -PropertyValue $DUO_HOST -Installer $installer

            Write-Log 'Committing changes to temp MSI'
            $workingDb.Commit()

            Write-Log 'Generating duo1.mst transform'
            $generated = $workingDb.GenerateTransform($originalDb, $OutputMst)
            if (-not $generated) { throw 'No transform was generated. The databases may not differ.' }

            Write-Log 'Creating transform summary information'
            $workingDb.CreateTransformSummaryInfo($originalDb, $OutputMst, 0, 0)

            if (-not (Test-Path -LiteralPath $OutputMst)) { throw "Transform was not created: $OutputMst" }
            Write-Log "MST generated at $OutputMst" 'OK'
        }
        catch {
            $detail = $null
            if ($installer) { $detail = Get-MsiLastError -Installer $installer }
            if ($detail) { throw "$($_.Exception.Message) | Windows Installer detail: $detail" }
            throw
        }
        finally {
            Release-ComObject $workingDb; Release-ComObject $originalDb; Release-ComObject $installer
            if (Test-Path -LiteralPath $workingMsi) { Remove-Item -LiteralPath $workingMsi -Force -ErrorAction SilentlyContinue }
        }
    }

    $extractDir    = Join-Path $duoPath 'DuoExtracted'
    function Find-Files ([string]$Pattern, [string]$Filter) {
        if (-not (Test-Path -LiteralPath $extractDir)) { return @() }
        $all = @(Get-ChildItem -LiteralPath $extractDir -Recurse -File -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -notmatch '^(?:\._|__MACOSX)' -and $_.FullName -notmatch '__MACOSX' })
        $hits = @($all | Where-Object { $_.Name -match $Pattern })
        if ($hits.Count -eq 0) { $hits = @($all | Where-Object { $_.Name -like $Filter }) }
        $hits
    }

    function Copy-Matches ([object[]]$Files, [string]$Destination, [string]$Label, [string]$Tag) {
        if ($Files.Count -eq 0) {
            Write-Log "No $Label files found in extracted ZIP" 'ERROR'
            Add-Failure "missing_$Tag"
            return
        }
        foreach ($m in $Files) {
            try {
                Copy-Item -LiteralPath $m.FullName -Destination $Destination -Force -ErrorAction Stop
                Write-Log ('Copied {0}: {1} -> {2}' -f $Label, $m.Name, $Destination) 'OK'
            }
            catch {
                Write-Log ('Failed copying {0} {1}: {2}' -f $Label, $m.Name, $_.Exception.Message) 'ERROR'
                Add-Failure "copy_$Tag"
            }
        }
    }

    function Ensure-Directory ([string]$Path, [string]$Tag) {
        if (Test-Path -LiteralPath $Path) { return }
        try { New-Item -Path $Path -ItemType Directory -Force | Out-Null; Write-Log "Created $Path" 'OK' }
        catch { Write-Log ('Failed to create {0}: {1}' -f $Path, $_.Exception.Message) 'ERROR'; Add-Failure "mkdir:$Tag" }
    }

    # Fallback for removing Authenticated Users from a GPO's security filtering via the GPMC COM object (no prompts)
    function Remove-GpoAuthenticatedUsers ([string]$GpoName) {
        $dom      = Get-ADDomain
        $gpoObj   = Get-GPO -Name $GpoName -ErrorAction Stop
        $gpm      = New-Object -ComObject GPMgmt.GPM
        $const    = $gpm.GetConstants()
        $gpmDom   = $gpm.GetDomain($dom.DNSRoot, '', $const.UseAnyDC)
        $gpmGpo   = $gpmDom.GetGPO('{' + $gpoObj.Id.ToString().ToUpper() + '}')
        $sd       = $gpmGpo.GetSecurityInfo()
        $targets  = @(foreach ($perm in $sd) { if ($perm.Trustee.TrusteeSid -like '*S-1-5-11') { $perm } })
        if ($targets.Count -eq 0) { return }
        foreach ($perm in $targets) { $sd.Remove($perm) }
        $gpmGpo.SetSecurityInfo($sd)
    }

    # Fallback for deleting a GPO via the GPMC COM object (no prompts)
    function Remove-GpoViaGpm ([string]$GpoName) {
        $dom    = Get-ADDomain
        $gpoObj = Get-GPO -Name $GpoName -ErrorAction Stop
        $gpm    = New-Object -ComObject GPMgmt.GPM
        $const  = $gpm.GetConstants()
        $gpmDom = $gpm.GetDomain($dom.DNSRoot, '', $const.UseAnyDC)
        $gpmGpo = $gpmDom.GetGPO('{' + $gpoObj.Id.ToString().ToUpper() + '}')
        $gpmGpo.Delete()
    }

    try {
        Write-Log ('DeployDuo started by {0}\{1} on {2}' -f $env:USERDOMAIN, $env:USERNAME, $env:COMPUTERNAME)
        Write-Log "Microsoft RDP Application IKEY: $IKEY"
        Write-Log "Microsoft RDP Application HOST: $DUO_HOST"

        Import-Module ActiveDirectory -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        Import-Module GroupPolicy     -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

        # ---------- 1. Preconditions ----------
        Start-Phase 'Step 1 of 13: Checking preconditions'
        if (-not $srcCfg -or -not (Test-Path -LiteralPath $srcCfg -PathType Leaf)) {
            Write-Log "authproxy.cfg not found: $srcCfg. Deployment stopped." 'ERROR'
            Add-Failure 'cfg_missing'
            return
        }
        try {
            Copy-Item -LiteralPath $srcCfg -Destination $cfgPath -Force -ErrorAction Stop
            Write-Log "Working directory: $duoPath" 'OK'
            Write-Log "Using authproxy.cfg from $srcCfg (original is left untouched)" 'OK'
        }
        catch {
            Write-Log ('Could not copy authproxy.cfg to the working directory: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'cfg_copy_failed'
            return
        }
        Complete-Phase

        # ---------- 2. Download ----------
        Start-Phase 'Step 2 of 13: Downloading Duo packages'
        $urls = @(
            'https://dl.duosecurity.com/DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip',
            'https://dl.duosecurity.com/duoauthproxy-latest.exe',
            'https://raw.githubusercontent.com/m-timm/DUODeploy/main/DUO_Windows_Login_GPO.zip'
        )
        foreach ($u in $urls) {
            $fileName = Split-Path $u -Leaf
            $dest = Join-Path $duoPath $fileName
            $sync.Queue.Enqueue([pscustomobject]@{ Kind = 'status'; Text = "Step 2 of 13: Downloading $fileName..." })
            Write-Log "Downloading $fileName"
            try {
                Invoke-WebRequest -Uri $u -OutFile $dest -UseBasicParsing -ErrorAction Stop
                $mb = [Math]::Round((Get-Item -LiteralPath $dest).Length / 1MB, 1)
                Write-Log "Downloaded $fileName ($mb MB)" 'OK'
            }
            catch {
                Write-Log ('Failed to download {0}: {1}' -f $fileName, $_.Exception.Message) 'ERROR'
                Add-Failure "download:$fileName"
            }
        }
        Complete-Phase

        # ---------- 3. Extract ----------
        Start-Phase 'Step 3 of 13: Extracting archives'
        $zips = @(
            @{ Zip = (Join-Path $duoPath 'DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip'); Dir = $extractDir;                    Label = 'Duo Windows Logon ZIP'; Tag = 'zip' },
            @{ Zip = (Join-Path $duoPath 'DUO_Windows_Login_GPO.zip');                              Dir = (Join-Path $duoPath 'GP');      Label = 'GPO backup ZIP';        Tag = 'gpo_zip' }
        )
        foreach ($z in $zips) {
            if (-not (Test-Path -LiteralPath $z.Zip)) {
                Write-Log "$($z.Label) not found: $($z.Zip)" 'ERROR'
                Add-Failure "$($z.Tag)_missing"
                continue
            }
            try {
                if (Test-Path -LiteralPath $z.Dir) {
                    Remove-Item -LiteralPath $z.Dir -Recurse -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Milliseconds 200
                }
                Expand-Archive -LiteralPath $z.Zip -DestinationPath $z.Dir -Force -ErrorAction Stop
                Write-Log "Extracted $($z.Label) -> $($z.Dir)" 'OK'
            }
            catch {
                Write-Log ('{0} extraction failed: {1}' -f $z.Label, $_.Exception.Message) 'ERROR'
                Add-Failure "$($z.Tag)_extract_failed"
            }
        }
        $gpoExtractDir = Join-Path $duoPath 'GP'
        Complete-Phase

        # ---------- 4. Copy ADMX / ADML / MSI ----------
        Start-Phase 'Step 4 of 13: Copying policy templates and MSIs'
        $policyRoot    = 'C:\Windows\PolicyDefinitions'
        $enUS          = Join-Path $policyRoot 'en-us'
        # Ask the server where its NETLOGON / SYSVOL shares really live. C:\Windows\SYSVOL\domain is only a junction to
        # C:\Windows\SYSVOL\sysvol\<domain> and is missing or stale on some servers (and SYSVOL can be relocated off C:).
        $adDom = Get-ADDomain
        $sysvolScripts = $null
        $sysvolShareRoot = $null
        try { $nl = (Get-SmbShare -Name 'NETLOGON' -ErrorAction Stop).Path; if ($nl -and (Test-Path -LiteralPath $nl)) { $sysvolScripts = $nl } } catch { }
        try { $sv = (Get-SmbShare -Name 'SYSVOL'   -ErrorAction Stop).Path; if ($sv -and (Test-Path -LiteralPath $sv)) { $sysvolShareRoot = $sv } } catch { }
        if (-not $sysvolScripts) {
            $sysvolScripts = 'C:\Windows\SYSVOL\sysvol\{0}\scripts' -f $adDom.DNSRoot
            if (-not (Test-Path -LiteralPath $sysvolScripts)) { $sysvolScripts = 'C:\Windows\SYSVOL\domain\scripts' }
            Write-Log "NETLOGON share path not readable; falling back to $sysvolScripts" 'WARN'
        }
        else { Write-Log "NETLOGON share is $sysvolScripts" }
        Ensure-Directory $policyRoot    'PolicyDefinitions'
        Ensure-Directory $enUS          'en-us'
        Ensure-Directory $sysvolScripts 'SYSVOL'

        Copy-Matches -Files (Find-Files '(?i)duowindowslogon.*\.admx$' '*.admx') -Destination $policyRoot    -Label 'ADMX' -Tag 'admx'
        Copy-Matches -Files (Find-Files '(?i)duowindowslogon.*\.adml$' '*.adml') -Destination $enUS          -Label 'ADML' -Tag 'adml'
        Copy-Matches -Files (Find-Files '(?i)duowindowslogon.*\.msi$'  '*.msi')  -Destination $sysvolScripts -Label 'MSI'  -Tag 'msi'

        # If the domain has a Central Store, GPMC reads templates ONLY from there and ignores the local folder,
        # so the Duo templates must be in it too. The local copy above stays as the fallback for domains without one.
        $centralStore = if ($sysvolShareRoot) { Join-Path $sysvolShareRoot ('{0}\Policies\PolicyDefinitions' -f $adDom.DNSRoot) }
                        else { 'C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions' }
        if (Test-Path -LiteralPath $centralStore) {
            Write-Log "Central Store found at $centralStore; copying Duo templates there too"
            $centralEn = Join-Path $centralStore 'en-US'
            Ensure-Directory $centralEn 'Central Store en-US'
            Copy-Matches -Files (Find-Files '(?i)duowindowslogon.*\.admx$' '*.admx') -Destination $centralStore -Label 'ADMX (Central Store)' -Tag 'admx'
            Copy-Matches -Files (Find-Files '(?i)duowindowslogon.*\.adml$' '*.adml') -Destination $centralEn    -Label 'ADML (Central Store)' -Tag 'adml'
        }
        Complete-Phase

        # ---------- 5. MST ----------
        Start-Phase 'Step 5 of 13: Generating transforms (duo1.mst + ARM64 / x86 variants)'
        # A transform only fits the MSI it was generated from, so each architecture gets its own
        foreach ($t in @(
            @{ Msi = 'DuoWindowsLogon64.msi';    Mst = 'duo1.mst';       Tag = 'x64'   },
            @{ Msi = 'DuoWindowsLogonARM64.msi'; Mst = 'duo1-arm64.mst'; Tag = 'arm64' },
            @{ Msi = 'DuoWindowsLogon32.msi';    Mst = 'duo1-x86.mst';   Tag = 'x86'   }
        )) {
            try {
                $src = Join-Path $sysvolScripts $t.Msi
                if (-not (Test-Path -LiteralPath $src)) { throw "Source MSI not found: $src" }
                Write-Log "Building $($t.Mst) from $($t.Msi)"
                New-DuoTransform -SourceMsi $src -OutputMst (Join-Path $sysvolScripts $t.Mst)
            }
            catch {
                Write-Log ('MST generation failed for {0}: {1}' -f $t.Msi, $_.Exception.Message) 'ERROR'
                Add-Failure ('mst_generation_failed:{0}' -f $t.Tag)
            }
        }

        # Sanity check: clients read these over the network, so confirm they are visible through the NETLOGON share itself
        $unc = '\\{0}\NETLOGON' -f $env:COMPUTERNAME
        $missingNl = @()
        foreach ($f in 'DuoWindowsLogon64.msi', 'duo1.mst', 'DuoWindowsLogonARM64.msi', 'duo1-arm64.mst', 'DuoWindowsLogon32.msi', 'duo1-x86.mst') {
            if (-not (Test-Path -LiteralPath (Join-Path $unc $f))) { $missingNl += $f }
        }
        if ($missingNl.Count -eq 0) { Write-Log "All installers and transforms are visible at $unc" 'OK' }
        else {
            Write-Log ('Not visible at {0}: {1}' -f $unc, ($missingNl -join ', ')) 'ERROR'
            Add-Failure 'netlogon_files_missing'
        }
        Complete-Phase

        # ---------- 6. Auth Proxy ----------
        Start-Phase 'Step 6 of 13: Installing Duo Authentication Proxy'
        $proxyInstaller = Join-Path $duoPath 'duoauthproxy-latest.exe'
        if (Test-Path -LiteralPath $proxyInstaller) {
            try {
                $r = Invoke-Quiet $proxyInstaller '/S /V /qn'
                if ($r.ExitCode -ne 0) { throw "installer exit code $($r.ExitCode)" }
                Write-Log 'Auth Proxy installed' 'OK'
            }
            catch {
                Write-Log ('Auth Proxy installation failed: {0}' -f $_.Exception.Message) 'ERROR'
                Add-Failure 'authproxy_install_failed'
            }
        }
        else {
            Write-Log "Auth Proxy installer not found at $proxyInstaller" 'ERROR'
            Add-Failure 'authproxy_installer_missing'
        }
        Complete-Phase

        # ---------- 7. duo_sync + DuoProtected ----------
        Start-Phase 'Step 7 of 13: Configuring duo_sync and DuoProtected'
        # Random per-run password. It only ever lives in authproxy.cfg, where authproxy_passwd encrypts it (step 9).
        $pwChars = 'abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!*+-'
        $rng     = [Security.Cryptography.RandomNumberGenerator]::Create()
        $pwBytes = New-Object byte[] 24
        $rng.GetBytes($pwBytes)
        $rng.Dispose()
        $randPass = 'Aa1!' + (-join ($pwBytes | ForEach-Object { $pwChars[$_ % $pwChars.Length] }))
        try {
            $secure = ConvertTo-SecureString $randPass -AsPlainText -Force
            # duo_sync: plain Domain Users member (read-only to the directory), password never expires, user cannot change it
            $svcDesc = 'Duo Authentication Proxy directory sync (read-only service account)'
            $existingUser = Get-ADUser -Filter "SamAccountName -eq 'duo_sync'" -ErrorAction SilentlyContinue
            if (-not $existingUser) {
                New-ADUser -Name 'duo_sync' -SamAccountName 'duo_sync' -Description $svcDesc -AccountPassword $secure -Enabled $true -PasswordNeverExpires $true -CannotChangePassword $true -ErrorAction Stop
                Write-Log 'duo_sync created' 'OK'
            }
            else {
                Set-ADAccountPassword -Identity 'duo_sync' -Reset -NewPassword $secure -ErrorAction Stop
                Enable-ADAccount -Identity 'duo_sync' -ErrorAction SilentlyContinue
                Set-ADUser -Identity 'duo_sync' -Description $svcDesc -PasswordNeverExpires $true -CannotChangePassword $true -ErrorAction SilentlyContinue
                Write-Log 'duo_sync already existed; password reset and account verified' 'OK'
            }
            # Read-only means no privileged groups
            foreach ($grp in 'Domain Admins', 'Enterprise Admins', 'Schema Admins', 'Administrators', 'Account Operators', 'Backup Operators', 'Server Operators', 'Print Operators') {
                try {
                    $isMember = Get-ADGroupMember -Identity $grp -ErrorAction Stop | Where-Object { $_.SamAccountName -eq 'duo_sync' }
                    if ($isMember) {
                        Remove-ADGroupMember -Identity $grp -Members 'duo_sync' -Confirm:$false -ErrorAction Stop
                        Write-Log ('Removed duo_sync from {0}' -f $grp) 'OK'
                    }
                }
                catch { }
            }
            # Overlay on a deployment from an older version: the old DuoAdmin account is no longer used, so disable it (not deleted)
            $legacy = Get-ADUser -Filter "SamAccountName -eq 'DuoAdmin'" -ErrorAction SilentlyContinue
            if ($legacy) {
                try {
                    Remove-ADGroupMember -Identity 'Domain Admins' -Members 'DuoAdmin' -Confirm:$false -ErrorAction SilentlyContinue
                    Disable-ADAccount -Identity 'DuoAdmin' -ErrorAction Stop
                    Write-Log 'Legacy DuoAdmin account found and disabled (replaced by duo_sync)' 'OK'
                }
                catch { Write-Log ('Could not disable legacy DuoAdmin: {0}' -f $_.Exception.Message) 'WARN' }
            }
        }
        catch {
            Write-Log ('Failed to create or update duo_sync: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'create_duosync_failed'
        }

        try {
            if (-not (Get-ADGroup -Filter "Name -eq 'DuoProtected'" -ErrorAction SilentlyContinue)) {
                New-ADGroup -Name 'DuoProtected' -GroupScope Global -GroupCategory Security -Path ('CN=Users,' + (Get-ADDomain).DistinguishedName) -ErrorAction Stop
                Write-Log 'Created DuoProtected group' 'OK'
            }
            else { Write-Log 'DuoProtected group already exists' }

            # duo_sync is deliberately NOT a DuoProtected member (every protected user counts toward Duo billing)
            try {
                Add-ADGroupMember -Identity 'DuoProtected' -Members 'Domain Admins' -ErrorAction Stop
                Write-Log 'Domain Admins is a member of DuoProtected' 'OK'
            }
            catch { Write-Log 'Domain Admins already in DuoProtected (or could not be added); continuing' }
            try {
                foreach ($acct in 'duo_sync', 'DuoAdmin') {
                    Remove-ADGroupMember -Identity 'DuoProtected' -Members $acct -Confirm:$false -ErrorAction SilentlyContinue
                }
                Write-Log 'Ensured duo_sync and DuoAdmin are not in DuoProtected' 'OK'
            }
            catch { Write-Log 'Could not update DuoProtected membership for service accounts; continuing' }
        }
        catch {
            Write-Log ('DuoProtected group creation/add failed: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'duoprotected_failed'
        }

        # TR_NOC email + a callable number on Domain Admin accounts (so Duo enrollment/notifications reach the NOC)
        $nocEmail = 'trhelpdesk@trcomputersales.com'
        $nocPhone = '6124409419'
        try {
            $noc = Get-ADUser -Filter "SamAccountName -eq 'TR_NOC'" -Properties EmailAddress -ErrorAction Stop
            if ($noc) {
                if ($noc.EmailAddress -ne $nocEmail) {
                    Set-ADUser -Identity $noc -EmailAddress $nocEmail -ErrorAction Stop
                    Write-Log "TR_NOC email set to $nocEmail" 'OK'
                }
                else { Write-Log 'TR_NOC email already set' }
            }
            else { Write-Log 'No TR_NOC account found; skipping email' }
        }
        catch {
            Write-Log ('Failed setting TR_NOC email: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'tr_noc_email_failed'
        }

        try {
            $admins = @(Get-ADGroupMember -Identity 'Domain Admins' -Recursive -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' })
            if ($admins.Count -eq 0) { Write-Log 'No Domain Admins user accounts found' 'WARN' }
            foreach ($a in $admins) {
                try {
                    if ($a.SamAccountName -ieq 'TR_NOC') { Write-Log 'TR_NOC: phone numbers left untouched' ; continue }
                    $u = Get-ADUser -Identity $a.distinguishedName -Properties telephoneNumber, mobile -ErrorAction Stop
                    if (-not $u.Enabled) { continue }
                    # Only fill blanks; existing numbers are left alone
                    $changes = @{}
                    if (-not $u.telephoneNumber) { $changes['telephoneNumber'] = $nocPhone }
                    if (-not $u.mobile)          { $changes['mobile']          = $nocPhone }
                    if ($changes.Count) {
                        Set-ADUser -Identity $u -Replace $changes -ErrorAction Stop
                        Write-Log ('Added {0} to {1} ({2})' -f $nocPhone, $u.SamAccountName, ($changes.Keys -join ', ')) 'OK'
                    }
                    else { Write-Log "$($u.SamAccountName) already has phone numbers; left unchanged" }
                }
                catch {
                    Write-Log ('Failed updating phone for {0}: {1}' -f $a.SamAccountName, $_.Exception.Message) 'ERROR'
                    Add-Failure 'admin_phone_failed'
                }
            }
        }
        catch {
            Write-Log ('Failed enumerating Domain Admins: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'admin_phone_failed'
        }
        Complete-Phase

        # ---------- 8. authproxy.cfg ----------
        Start-Phase 'Step 8 of 13: Deploying authproxy.cfg'
        try {
            $cfg = Get-Content -LiteralPath $cfgPath -ErrorAction Stop
            # Works on a fresh template ("; service_account_...") and on a cfg that was already deployed once
            # (duo_sync gets a new random password every run, so a stale encrypted password line must go)
            $cfg = @($cfg | Where-Object { $_ -notmatch '^\s*service_account_password_protected\s*=' })
            $cfg = $cfg -replace '^\s*;?\s*service_account_username\s*=.*', 'service_account_username=duo_sync'
            $cfg = $cfg -replace '^\s*;?\s*service_account_password\s*=.*', "service_account_password=$randPass"
            Set-Content -LiteralPath $cfgPath -Value $cfg -Encoding ASCII -ErrorAction Stop
            Write-Log 'authproxy.cfg service account lines updated' 'OK'
        }
        catch {
            Write-Log ('Failed updating authproxy.cfg: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'cfg_update_failed'
        }

        $cfgDest = 'C:\Program Files\Duo Security Authentication Proxy\conf\authproxy.cfg'
        try {
            $cfgDir = Split-Path $cfgDest
            if (-not (Test-Path -LiteralPath $cfgDir)) { New-Item -Path $cfgDir -ItemType Directory -Force | Out-Null }
            Copy-Item -LiteralPath $cfgPath -Destination $cfgDest -Force -ErrorAction Stop
            Write-Log "Deployed authproxy.cfg -> $cfgDest" 'OK'
        }
        catch {
            Write-Log ('Failed deploying cfg: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'cfg_deploy_failed'
        }

        $passwdExe = 'C:\Program Files\Duo Security Authentication Proxy\bin\authproxy_passwd.exe'
        if (Test-Path -LiteralPath $passwdExe) {
            try {
                # --yes bypasses the "Are you sure you want to proceed? [Yes/no]" prompt (proxy 5.3.0+;
                # older proxies used --yes-overwrite). "yes" is also piped to stdin as a fallback.
                Write-Log 'Running authproxy_passwd --whole-config --yes'
                $r = Invoke-Quiet $passwdExe '--whole-config --yes' -TimeoutSec 120 -InputText 'yes'
                if ($r.Output) { $r.Output -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log "  $_" } }
                if ($r.ExitCode -ne 0) { throw "exit code $($r.ExitCode)" }

                # Verify the plaintext password was actually encrypted in the deployed config
                $deployed = Get-Content -LiteralPath $cfgDest -Raw -ErrorAction Stop
                if ($deployed -match '(?m)^\s*service_account_password\s*=') {
                    throw 'plaintext service_account_password is still in the deployed authproxy.cfg; run authproxy_passwd --whole-config manually.'
                }
                Write-Log 'authproxy_passwd completed; deployed config contains no plaintext service account password' 'OK'
            }
            catch {
                Write-Log ('authproxy_passwd failed: {0}' -f $_.Exception.Message) 'ERROR'
                Add-Failure 'authproxy_passwd_failed'
            }
        }
        else {
            Write-Log "authproxy_passwd not found: $passwdExe" 'ERROR'
            Add-Failure 'authproxy_passwd_missing'
        }

        # The working copy holds the plaintext service-account password; don't leave it behind
        try {
            if (Test-Path -LiteralPath $cfgPath) { Remove-Item -LiteralPath $cfgPath -Force -ErrorAction Stop; Write-Log 'Removed working copy of authproxy.cfg' }
        }
        catch { Write-Log ('Could not remove working copy {0}: {1}' -f $cfgPath, $_.Exception.Message) 'WARN' }
        Complete-Phase

        # ---------- 9. GPO create / delegate / link ----------
        $gpoName = 'DUO Windows Login'
        Start-Phase "Step 9 of 13: Creating GPO '$gpoName'"

        # Overlay on an existing deployment: any GPO with "duo" in its name (typically the old Software Installation one)
        # is backed up, unlinked and deleted so this run starts clean. The backup is kept in the work folder.
        try {
            $odom   = Get-ADDomain
            $oldGpos = @(Get-GPO -All -ErrorAction Stop | Where-Object { $_.DisplayName -match '(?i)duo' })
            if ($oldGpos.Count -eq 0) { Write-Log "No existing GPO with 'duo' in its name; nothing to remove" }
            $backupRoot = Join-Path $duoPath ('RemovedGpoBackups\' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
            foreach ($old in $oldGpos) {
                $oname = $old.DisplayName
                $oguid = '{' + $old.Id.ToString().ToUpper() + '}'
                Write-Log ("Existing GPO '{0}' {1} matches 'duo'; removing it so this run can start fresh" -f $oname, $oguid) 'WARN'
                try {
                    New-Item -Path $backupRoot -ItemType Directory -Force | Out-Null
                    Backup-GPO -Guid $old.Id -Path $backupRoot -ErrorAction Stop | Out-Null
                    Write-Log ("Backed up '{0}' to {1}" -f $oname, $backupRoot) 'OK'
                }
                catch { Write-Log ("Could not back up '{0}': {1}; continuing" -f $oname, $_.Exception.Message) 'WARN' }

                # Remove every link (domain root and OUs) before deleting
                try {
                    $linked = @(Get-ADObject -LDAPFilter ('(gPLink=*{0}*)' -f $oguid) -SearchBase $odom.DistinguishedName -SearchScope Subtree -Server $odom.PDCEmulator -ErrorAction SilentlyContinue)
                    foreach ($lk in $linked) {
                        try { Remove-GPLink -Guid $old.Id -Target $lk.DistinguishedName -ErrorAction Stop | Out-Null; Write-Log ("Unlinked '{0}' from {1}" -f $oname, $lk.DistinguishedName) 'OK' }
                        catch { }
                    }
                }
                catch { }

                try { Remove-GPO -Guid $old.Id -Confirm:$false -ErrorAction Stop }
                catch {
                    Write-Log ('Remove-GPO failed ({0}); trying GPMC COM' -f $_.Exception.Message)
                    Remove-GpoViaGpm -GpoName $oname
                }
                Write-Log ("Removed GPO '{0}'" -f $oname) 'OK'
            }
        }
        catch {
            Write-Log ('Could not remove existing Duo GPO(s): {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'gpo_cleanup_failed'
        }
        try {
            if (-not (Get-GPO -Name $gpoName -ErrorAction SilentlyContinue)) {
                New-GPO -Name $gpoName -ErrorAction Stop | Out-Null
                Write-Log "GPO $gpoName created" 'OK'
            }
            else { Write-Log "GPO $gpoName already exists" }
        }
        catch {
            Write-Log ('GPO creation failed: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'gpo_failed'
        }

        # Grant the computer groups first so removing Authenticated Users never leaves the GPO unreadable
        $grantsOk = $false
        try {
            foreach ($g in 'Domain Computers', 'Domain Controllers') {
                Set-GPPermission -Name $gpoName -TargetName $g -TargetType Group -PermissionLevel GpoApply -Confirm:$false -ErrorAction Stop | Out-Null
                Write-Log "Granted GpoApply to $g" 'OK'
            }
            $grantsOk = $true
        }
        catch {
            Write-Log ('GPO permission delegation failed: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'gpo_permissions_failed'
        }

        if ($grantsOk) {
            try {
                Set-GPPermission -Name $gpoName -TargetName 'Authenticated Users' -TargetType Group -PermissionLevel None -Confirm:$false -ErrorAction Stop | Out-Null
                Write-Log 'Removed Authenticated Users from GPO permissions' 'OK'
            }
            catch {
                Write-Log 'Set-GPPermission prompts for confirmation, which this window cannot answer; using GPMC COM instead'
                try {
                    Remove-GpoAuthenticatedUsers -GpoName $gpoName
                    Write-Log 'Removed Authenticated Users from GPO permissions (GPMC COM)' 'OK'
                }
                catch {
                    Write-Log ('Removing Authenticated Users failed: {0}' -f $_.Exception.Message) 'ERROR'
                    Add-Failure 'gpo_authusers_failed'
                }
            }
        }

        try {
            $domainDN = (Get-ADDomain).DistinguishedName
            try   { New-GPLink -Name $gpoName -Target $domainDN -Enforced No -ErrorAction Stop | Out-Null }
            catch { Set-GPLink -Name $gpoName -Target $domainDN -LinkEnabled Yes -ErrorAction Stop | Out-Null }
            Write-Log 'GPO linked at domain root' 'OK'
        }
        catch {
            Write-Log ('GPO link failed: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'gpo_link_failed'
        }
        Complete-Phase

        # ---------- 10. Import GPO backup ----------
        Start-Phase 'Step 10 of 13: Importing GPO backup'
        try {
            if (-not (Test-Path -LiteralPath $gpoExtractDir)) { throw "GPO backup path not found: $gpoExtractDir" }
            Import-Module GroupPolicy -ErrorAction Stop -WarningAction SilentlyContinue
            if (-not (Get-Command Import-GPO -ErrorAction SilentlyContinue)) {
                throw 'Import-GPO is not available. Install the Group Policy Management feature and try again.'
            }

            $backupXml = Get-ChildItem -LiteralPath $gpoExtractDir -Recurse -Filter 'bkupInfo.xml' -File -ErrorAction Stop | Select-Object -First 1
            if (-not $backupXml) { throw "No bkupInfo.xml found under $gpoExtractDir" }
            Write-Log "Found GPO backup metadata: $($backupXml.FullName)"
            [xml]$backupInfoXml = Get-Content -LiteralPath $backupXml.FullName -ErrorAction Stop

            $nameNode = $backupInfoXml.SelectSingleNode("//*[local-name()='GPODisplayName' or local-name()='DisplayName' or local-name()='GPOName']")
            $idNode   = $backupInfoXml.SelectSingleNode("//*[local-name()='ID' or local-name()='BackupId' or local-name()='BackupID']")
            $backupGpoName = if ($nameNode -and $nameNode.InnerText) { $nameNode.InnerText.Trim() }
            $backupId      = if ($idNode -and $idNode.InnerText) { $idNode.InnerText.Trim('{} ').Trim() }

            if ($backupGpoName) { Write-Log "Backup GPO name: $backupGpoName" } else { Write-Log 'Could not resolve backup GPO name from bkupInfo.xml' 'WARN' }
            if ($backupId)      { Write-Log "Backup ID: $backupId" }           else { Write-Log 'Could not resolve backup ID from bkupInfo.xml' 'WARN' }

            $imported = $false
            if ($backupGpoName) {
                try {
                    Import-GPO -Path $gpoExtractDir -BackupGpoName $backupGpoName -TargetName $gpoName -CreateIfNeeded -ErrorAction Stop | Out-Null
                    Write-Log "Imported backup '$backupGpoName' into '$gpoName'" 'OK'
                    $imported = $true
                }
                catch { Write-Log ('Import by BackupGpoName failed: {0}' -f $_.Exception.Message) 'WARN' }
            }
            if (-not $imported -and $backupId) {
                Import-GPO -Path $gpoExtractDir -BackupId $backupId -TargetName $gpoName -CreateIfNeeded -ErrorAction Stop | Out-Null
                Write-Log "Imported backup ID '$backupId' into '$gpoName'" 'OK'
                $imported = $true
            }
            if (-not $imported) { throw 'Neither BackupGpoName nor BackupId could be used to import the backup.' }
        }
        catch {
            Write-Log ('GPO backup import failed: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'gpo_import_failed'
        }
        Complete-Phase

        # ---------- 11. Inject values ----------
        Start-Phase 'Step 11 of 13: Injecting HOST / IKEY / SKEY into GPO'
        try {
            $regKey = 'HKLM\SOFTWARE\Policies\Duo Security\DuoCredProv'
            Set-GPRegistryValue -Name $gpoName -Key $regKey -ValueName 'HOST' -Type String -Value $DUO_HOST -ErrorAction Stop | Out-Null
            Set-GPRegistryValue -Name $gpoName -Key $regKey -ValueName 'IKEY' -Type String -Value $IKEY     -ErrorAction Stop | Out-Null
            Set-GPRegistryValue -Name $gpoName -Key $regKey -ValueName 'SKEY' -Type String -Value $SKEY     -ErrorAction Stop | Out-Null
            Write-Log 'Duo policy values injected' 'OK'

            $verifyReport = Join-Path $duoPath 'Verify_DuoGPO.xml'
            Get-GPOReport -Name $gpoName -ReportType Xml -Path $verifyReport -ErrorAction Stop
            Write-Log "Verification GPO report written to $verifyReport" 'OK'
        }
        catch {
            Write-Log ('Duo GPO value injection failed: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'gpo_value_injection_failed'
        }
        Complete-Phase

        # ---------- 12. Startup script ----------
        Start-Phase 'Step 12 of 13: Adding Duo startup script to GPO and installing here'
        $startupScript = @'
<#
    DuoWindowsLogon-Startup.ps1 (deployed by DeployDuo; runs at computer startup as SYSTEM via the "DUO Windows Login" GPO)

    Installs or upgrades Duo Windows Logon from \\<domain>\NETLOGON when Duo is missing or older than the MSI there.
    Does nothing when the installed version is current.
    Upgrade: re-run DeployDuo (it refreshes the MSIs and the per-architecture MSTs in NETLOGON).
    Remove:  unlink/disable the GPO (or remove this startup script) FIRST, then uninstall through the RMM.
             Otherwise the next boot sees Duo missing and reinstalls it.
    Log: C:\ProgramData\TRCS\Duo\startup.log
    -Share overrides the NETLOGON share (DeployDuo uses it to install on the DC it runs on).
#>
param([string]$Share)
$ErrorActionPreference = 'Stop'
$logDir  = Join-Path $env:ProgramData 'TRCS\Duo'
$logFile = Join-Path $logDir 'startup.log'

function Write-Log ([string]$Message) {
    try { Add-Content -LiteralPath $logFile -Value ('{0:yyyy-MM-dd HH:mm:ss} {1}' -f (Get-Date), $Message) } catch { }
}

try {
    if (-not (Test-Path -LiteralPath $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
    if ((Test-Path -LiteralPath $logFile) -and ((Get-Item -LiteralPath $logFile).Length -gt 1MB)) { Remove-Item -LiteralPath $logFile -Force }

    $arch = if ($env:PROCESSOR_ARCHITEW6432) { $env:PROCESSOR_ARCHITEW6432 } else { $env:PROCESSOR_ARCHITECTURE }
    switch ($arch) {
        'AMD64' { $msiName = 'DuoWindowsLogon64.msi';    $mstName = 'duo1.mst' }
        'ARM64' { $msiName = 'DuoWindowsLogonARM64.msi'; $mstName = 'duo1-arm64.mst' }
        'x86'   { $msiName = 'DuoWindowsLogon32.msi';    $mstName = 'duo1-x86.mst' }
        default { Write-Log "Unsupported architecture '$arch'; skipping"; exit 0 }
    }

    if ($Share) { $share = $Share }
    else {
        $domain = (Get-CimInstance -ClassName Win32_ComputerSystem).Domain
        $share  = "\\$domain\NETLOGON"
    }
    $msi    = Join-Path $share $msiName
    $mst    = Join-Path $share $mstName

    # The network can still be settling early in boot
    $ready = $false
    for ($i = 0; $i -lt 6; $i++) {
        if ((Test-Path -LiteralPath $msi) -and (Test-Path -LiteralPath $mst)) { $ready = $true; break }
        Start-Sleep -Seconds 10
    }
    if (-not $ready) { Write-Log "MSI/MST not reachable at $share; will retry next boot"; exit 2 }

    # Read ProductVersion, ProductCode and UpgradeCode from the MSI in NETLOGON
    $wi = New-Object -ComObject WindowsInstaller.Installer
    $db = $wi.OpenDatabase($msi, 0)
    function Get-MsiProp ([string]$Name) {
        $v = $db.OpenView(('SELECT `Value` FROM `Property` WHERE `Property`=''{0}''' -f $Name))
        [void]$v.Execute()
        $r = $v.Fetch()
        $out = $null
        if ($r) { $out = [string]$r.StringData(1) }
        [void]$v.Close()
        $out
    }
    $msiVersion  = [version]([string](@(Get-MsiProp 'ProductVersion')[0]))
    $productCode = [string](@(Get-MsiProp 'ProductCode')[0])
    $upgradeCode = [string](@(Get-MsiProp 'UpgradeCode')[0])

    # Remove a leftover registration (e.g. an old Software Installation advertisement) that stops the MSI from installing
    function Clear-StaleProduct ([string]$Code) {
        Write-Log "Removing stale registration $Code"
        $q = Start-Process -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList ('/x {0} /qn /norestart' -f $Code) -Wait -PassThru -WindowStyle Hidden
        Write-Log "msiexec /x $Code returned $($q.ExitCode)"
        if ($wi.ProductState($Code) -eq 5 -or $wi.ProductState($Code) -eq -1) { return }
        # Still registered but not installed: delete the registration keys directly
        $g = [guid]$Code
        $packed = -join (($g.ToString('N').Substring(0,8)).ToCharArray()[7..0]) + `
                  -join (($g.ToString('N').Substring(8,4)).ToCharArray()[3..0]) + `
                  -join (($g.ToString('N').Substring(12,4)).ToCharArray()[3..0])
        $tail = $g.ToString('N').Substring(16)
        for ($i = 0; $i -lt $tail.Length; $i += 2) { $packed += $tail[$i+1] + '' + $tail[$i] }
        foreach ($k in @(
            "HKLM:\SOFTWARE\Classes\Installer\Products\$packed",
            "HKLM:\SOFTWARE\Classes\Installer\Features\$packed",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$packed",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Managed\S-1-5-18\Installer\Products\$packed")) {
            if (Test-Path -LiteralPath $k) {
                try { Remove-Item -LiteralPath $k -Recurse -Force -ErrorAction Stop; Write-Log "Deleted $k" }
                catch { Write-Log "Could not delete ${k}: $($_.Exception.Message)" }
            }
        }
        Write-Log "State of $Code is now $($wi.ProductState($Code))"
    }

    # Ask Windows Installer itself (state 5 = fully installed). Independent of Add/Remove names and 32/64-bit registry views.
    if ($wi.ProductState($productCode) -eq 5) {
        Write-Log "Duo Windows Logon $msiVersion is installed (product $productCode); nothing to do"
        exit 0
    }

    # Not this exact build: look for an older build that is really installed, and flag stale registrations
    $installedVersion = $null
    foreach ($pc in $wi.RelatedProducts($upgradeCode)) {
        $st = $wi.ProductState($pc)
        if ($st -ne 5) {
            Write-Log "Related product $pc is registered but not installed (state $st); stale registration"
            Clear-StaleProduct $pc
            continue
        }
        $iv = $null
        try { $iv = [version]($wi.ProductInfo($pc, 'VersionString')) } catch { }
        if ($iv -and (-not $installedVersion -or $iv -gt $installedVersion)) { $installedVersion = $iv }
    }
    if ($installedVersion) {
        if ($installedVersion -ge $msiVersion) {
            Write-Log "Duo Windows Logon $installedVersion is installed and not older than the MSI ($msiVersion); nothing to do"
            exit 0
        }
        Write-Log "Duo Windows Logon $installedVersion found, MSI is $msiVersion; upgrading"
    }
    else { Write-Log "Duo Windows Logon not installed; installing $msiVersion" }

    $msiLog  = Join-Path $logDir 'DuoWindowsLogon_Startup_Install.log'
    $msiArgs = '/i "{0}" TRANSFORMS="{1}" FAILOPEN=#1 RDPONLY=#0 /qn /norestart /L*v "{2}"' -f $msi, $mst, $msiLog
    $p = Start-Process -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList $msiArgs -Wait -PassThru -WindowStyle Hidden
    switch ($p.ExitCode) {
        0       { Write-Log 'Install completed'; break }
        3010    { Write-Log 'Install completed; reboot required (3010)'; break }
        1641    { Write-Log 'Install completed; reboot initiated (1641)'; break }
        default { Write-Log "msiexec exit code $($p.ExitCode). See $msiLog"; exit $p.ExitCode }
    }

    # msiexec can return 0 without installing anything (e.g. a stale registration of an older build makes it skip all features)
    if ($wi.ProductState($productCode) -ne 5) {
        Write-Log "ERROR: msiexec returned $($p.ExitCode) but Duo Windows Logon is not installed (state $($wi.ProductState($productCode))). See $msiLog"
        exit 3
    }
}
catch {
    Write-Log ('ERROR: {0}' -f $_.Exception.Message)
    exit 1
}
'@
        try {
            $dom     = Get-ADDomain
            $pdc     = $dom.PDCEmulator
            $gpoObj  = Get-GPO -Name $gpoName -ErrorAction Stop
            $guid    = '{' + $gpoObj.Id.ToString().ToUpper() + '}'
            $gpoDN   = 'CN={0},CN=Policies,CN=System,{1}' -f $guid, $dom.DistinguishedName
            $gptRoot = '\\{0}\SYSVOL\{1}\Policies\{2}' -f $pdc, $dom.DNSRoot, $guid
            if (-not (Test-Path -LiteralPath $gptRoot)) { throw "GPO folder not found in SYSVOL: $gptRoot" }

            # 1. Script + psscripts.ini in the GPO's Machine\Scripts folder
            $scriptsRoot = Join-Path $gptRoot 'Machine\Scripts'
            $startupDir  = Join-Path $scriptsRoot 'Startup'
            $scriptName  = 'DuoWindowsLogon-Startup.ps1'
            New-Item -Path $startupDir -ItemType Directory -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $startupDir $scriptName) -Value $startupScript -Encoding UTF8 -ErrorAction Stop
            $ini = "`r`n[Startup]`r`n0CmdLine=$scriptName`r`n0Parameters=`r`n"
            Set-Content -LiteralPath (Join-Path $scriptsRoot 'psscripts.ini') -Value $ini -Encoding Unicode -ErrorAction Stop
            Write-Log "Startup script written to $startupDir" 'OK'

            # 1b. Deny duo_sync interactive and Remote Desktop logon (User Rights Assignment in GptTmpl.inf).
            #     User rights are replaced, not merged, by the highest-precedence GPO that defines them, so if any OTHER
            #     GPO already defines these two rights we leave them alone and say so instead of overriding it.
            $secCse = '{827D319E-6EAC-11D2-A4EA-00C04F79F83A}'
            $secTool = '{803E14A0-B4FB-11D0-A0D0-00A0C90F574B}'
            $denyApplied = $false
            try {
                $svcSid = (Get-ADUser -Identity 'duo_sync' -Server $pdc -ErrorAction Stop).SID.Value
                $polRoot = '\\{0}\SYSVOL\{1}\Policies' -f $pdc, $dom.DNSRoot
                $conflicts = @()
                foreach ($d in Get-ChildItem -LiteralPath $polRoot -Directory -ErrorAction SilentlyContinue) {
                    if ($d.Name -eq $guid) { continue }
                    $otherInf = Join-Path $d.FullName 'Machine\Microsoft\Windows NT\SecEdit\GptTmpl.inf'
                    if ((Test-Path -LiteralPath $otherInf) -and (Select-String -LiteralPath $otherInf -Pattern '^\s*(SeDenyInteractiveLogonRight|SeDenyRemoteInteractiveLogonRight)\s*=' -Quiet)) { $conflicts += $d.Name }
                }
                if ($conflicts.Count -gt 0) {
                    Write-Log ('Not setting logon restrictions for duo_sync: another GPO already defines the deny-logon rights ({0}). Add duo_sync there manually.' -f ($conflicts -join ', ')) 'WARN'
                }
                else {
                    $secDir = Join-Path $gptRoot 'Machine\Microsoft\Windows NT\SecEdit'
                    New-Item -Path $secDir -ItemType Directory -Force | Out-Null
                    $infPath = Join-Path $secDir 'GptTmpl.inf'
                    $lines = New-Object 'System.Collections.Generic.List[string]'
                    if (Test-Path -LiteralPath $infPath) { foreach ($l in (Get-Content -LiteralPath $infPath)) { $lines.Add($l) } }
                    else {
                        $lines.Add('[Unicode]'); $lines.Add('Unicode=yes')
                        $lines.Add('[Version]'); $lines.Add('signature="$CHICAGO$"'); $lines.Add('Revision=1')
                    }
                    $hdr = $lines.FindIndex({ param($x) $x.Trim() -ieq '[Privilege Rights]' })
                    if ($hdr -lt 0) { $lines.Add('[Privilege Rights]'); $hdr = $lines.Count - 1 }
                    foreach ($right in 'SeDenyInteractiveLogonRight', 'SeDenyRemoteInteractiveLogonRight') {
                        $idx = -1
                        for ($i = $hdr + 1; $i -lt $lines.Count -and $lines[$i].Trim() -notmatch '^\['; $i++) {
                            if ($lines[$i] -match ('^\s*' + $right + '\s*=')) { $idx = $i; break }
                        }
                        if ($idx -ge 0) {
                            if ($lines[$idx] -notmatch [regex]::Escape($svcSid)) { $lines[$idx] = $lines[$idx].TrimEnd() + ',*' + $svcSid }
                        }
                        else { $lines.Insert($hdr + 1, ('{0} = *{1}' -f $right, $svcSid)) }
                    }
                    Set-Content -LiteralPath $infPath -Value $lines -Encoding Unicode -ErrorAction Stop
                    $denyApplied = $true
                    Write-Log 'duo_sync denied interactive and Remote Desktop logon in the GPO' 'OK'
                }
            }
            catch { Write-Log ('Could not set logon restrictions for duo_sync: {0}' -f $_.Exception.Message) 'WARN' }

            # 2. Register the Scripts client-side extension on the GPO (gPCMachineExtensionNames, sorted)
            $cse  = '{42B5FAAE-6536-11D2-AE5A-0000F87571E3}'
            $tool = '{40B6664F-4972-11D1-A7CA-0000F87571E3}'
            $obj  = Get-ADObject -Identity $gpoDN -Server $pdc -Properties gPCMachineExtensionNames, versionNumber -ErrorAction Stop

            $map = New-Object 'System.Collections.Generic.SortedDictionary[string,object]' -ArgumentList ([StringComparer]::Ordinal)
            foreach ($m in [regex]::Matches([string]$obj.gPCMachineExtensionNames, '\[([^\]]+)\]')) {
                $g = @([regex]::Matches($m.Groups[1].Value, '\{[0-9A-Fa-f-]+\}') | ForEach-Object { $_.Value.ToUpper() })
                if ($g.Count -eq 0) { continue }
                if (-not $map.ContainsKey($g[0])) { $map[$g[0]] = New-Object 'System.Collections.Generic.List[string]' }
                for ($k = 1; $k -lt $g.Count; $k++) { if (-not $map[$g[0]].Contains($g[$k])) { $map[$g[0]].Add($g[$k]) } }
            }
            if ($denyApplied) {
                if (-not $map.ContainsKey($secCse)) { $map[$secCse] = New-Object 'System.Collections.Generic.List[string]' }
                if (-not $map[$secCse].Contains($secTool)) { $map[$secCse].Add($secTool) }
            }
            if (-not $map.ContainsKey($cse)) { $map[$cse] = New-Object 'System.Collections.Generic.List[string]' }
            if (-not $map[$cse].Contains($tool)) { $map[$cse].Add($tool) }

            $extNames = ''
            foreach ($key in $map.Keys) {
                $tools = $map[$key]
                $tools.Sort([StringComparer]::Ordinal)
                $extNames += '[' + $key + ($tools -join '') + ']'
            }

            # 3. Bump the GPO version so clients re-read it. AD 'versionNumber' and GPT.INI 'Version' hold the same
            #    combined value (user version in the high 16 bits, computer version in the low 16 bits).
            $iniPath = Join-Path $gptRoot 'GPT.INI'
            $iniText = Get-Content -LiteralPath $iniPath -Raw -ErrorAction Stop
            $gptVer  = 0
            if ($iniText -match '(?m)^Version=(\d+)') { $gptVer = [int64]$Matches[1] }
            $newVer  = [Math]::Max($gptVer, [int64]$obj.versionNumber) + 1
            if ($iniText -match '(?m)^Version=\d+') { $iniText = $iniText -replace '(?m)^Version=\d+', "Version=$newVer" }
            else { $iniText = $iniText.TrimEnd() + "`r`nVersion=$newVer`r`n" }
            Set-Content -LiteralPath $iniPath -Value $iniText -Encoding ASCII -NoNewline -ErrorAction Stop

            Set-ADObject -Identity $gpoDN -Server $pdc -Replace @{ gPCMachineExtensionNames = $extNames; versionNumber = [int]$newVer } -ErrorAction Stop
            Write-Log "Registered Scripts extension on GPO (version $newVer)" 'OK'
        }
        catch {
            Write-Log ('Startup script deployment failed: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'startup_script_failed'
        }

        # Run the same script on this server now, so the DC is installed exactly the way every other machine will be
        try {
            $localScript = Join-Path $duoPath 'DuoWindowsLogon-Startup.ps1'
            Set-Content -LiteralPath $localScript -Value $startupScript -Encoding UTF8 -ErrorAction Stop
            $startupLog = Join-Path $env:ProgramData 'TRCS\Duo\startup.log'
            $before = 0
            if (Test-Path -LiteralPath $startupLog) { $before = @(Get-Content -LiteralPath $startupLog).Count }

            # 64-bit PowerShell even if this exe runs as a 32-bit process (so the 64-bit uninstall keys are visible)
            $psExe = if ([Environment]::Is64BitProcess) { "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" }
                     else { "$env:SystemRoot\SysNative\WindowsPowerShell\v1.0\powershell.exe" }
            $localShare = '\\{0}\NETLOGON' -f $env:COMPUTERNAME
            Write-Log "Running the startup script on this server (source: $localShare)"
            $r = Invoke-Quiet $psExe ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -Share "{1}"' -f $localScript, $localShare)
            if (Test-Path -LiteralPath $startupLog) {
                Get-Content -LiteralPath $startupLog | Select-Object -Skip $before | ForEach-Object { Write-Log "  $_" }
            }
            if ($r.ExitCode -ne 0) { throw "startup script exit code $($r.ExitCode). See $startupLog" }
            Write-Log 'Duo Windows Logon install/upgrade via the startup script completed' 'OK'
        }
        catch {
            Write-Log ('Duo Windows Logon installation failed: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'duowinlogon_install_failed'
        }
        Complete-Phase

        # ---------- 13. Proxy Manager ----------
        Start-Phase 'Step 13 of 13: Starting Auth Proxy service and launching Proxy Manager'
        # Start the service (or restart it so it re-reads the new config) and make sure it comes up at boot
        try {
            $svcName = 'duoauthproxy'
            $svc = Get-Service -Name $svcName -ErrorAction Stop
            Set-Service -Name $svcName -StartupType Automatic -ErrorAction Stop
            if ($svc.Status -eq 'Running') { Restart-Service -Name $svcName -Force -ErrorAction Stop }
            else { Start-Service -Name $svcName -ErrorAction Stop }
            (Get-Service -Name $svcName).WaitForStatus('Running', [TimeSpan]::FromSeconds(30))
            Write-Log 'Duo Authentication Proxy service is running' 'OK'
        }
        catch {
            Write-Log ('Duo Authentication Proxy service did not start: {0}. Use Validate in the Proxy Manager to check the config.' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'authproxy_service_failed'
        }

        $proxyGui = 'C:\Program Files\Duo Security Authentication Proxy\bin\local_proxy_manager-win32-x64\Duo_Authentication_Proxy_Manager.exe'
        try {
            if (-not (Test-Path -LiteralPath $proxyGui)) { throw "not found at $proxyGui" }
            Start-Process -FilePath $proxyGui -ErrorAction Stop
            Write-Log 'Proxy Manager launched' 'OK'
        }
        catch {
            Write-Log ('Failed to launch Proxy Manager: {0}' -f $_.Exception.Message) 'ERROR'
            Add-Failure 'launch_proxy_gui_failed'
        }
        Complete-Phase
    }
    catch {
        Write-Log ('Unexpected error: {0}' -f $_.Exception.Message) 'ERROR'
        Add-Failure 'unexpected_error'
    }
    finally {
        $sync.Failures = $failures.ToArray()
        $sync.Done = $true
    }
}
#endregion

#region ---------- UI ----------
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Duo Deployment'
$form.ClientSize    = New-Object System.Drawing.Size(860, 660)
$form.MinimumSize   = New-Object System.Drawing.Size(640, 480)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

# --- Inputs ---
$inputs = New-Object System.Windows.Forms.TableLayoutPanel
$inputs.Dock         = 'Top'
$inputs.AutoSize     = $true
$inputs.ColumnCount  = 3
$inputs.Padding      = New-Object System.Windows.Forms.Padding(12, 12, 12, 0)
[void]$inputs.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$inputs.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$inputs.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))

function New-InputRow ([string]$Label, [int]$Row) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text      = $Label
    $l.AutoSize  = $true
    $l.Anchor    = 'Left'
    $l.Margin    = New-Object System.Windows.Forms.Padding(0, 6, 10, 6)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Dock      = 'Fill'
    $t.Margin    = New-Object System.Windows.Forms.Padding(0, 3, 6, 3)
    $t.Font      = New-Object System.Drawing.Font('Consolas', 10)
    $inputs.Controls.Add($l, 0, $Row)
    $inputs.Controls.Add($t, 1, $Row)
    $t
}
$txtIkey = New-InputRow 'Microsoft RDP Application IKEY' 0
$txtSkey = New-InputRow 'Microsoft RDP Application SKEY' 1
$txtHost = New-InputRow 'Microsoft RDP Application HOST' 2
$txtCfg  = New-InputRow 'Authentication Proxy config (authproxy.cfg)' 3
$txtSkey.UseSystemPasswordChar = $true

$chkShow = New-Object System.Windows.Forms.CheckBox
$chkShow.Text     = 'Show'
$chkShow.AutoSize = $true
$chkShow.Anchor   = 'Left'
$inputs.Controls.Add($chkShow, 2, 1)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text     = 'Browse...'
$btnBrowse.AutoSize = $true
$btnBrowse.Anchor   = 'Left'
$inputs.Controls.Add($btnBrowse, 2, 3)

$lblInputErr = New-Object System.Windows.Forms.Label
$lblInputErr.AutoSize  = $true
$lblInputErr.ForeColor = [System.Drawing.Color]::Firebrick
$lblInputErr.Margin    = New-Object System.Windows.Forms.Padding(0, 2, 0, 4)
$inputs.Controls.Add($lblInputErr, 1, 4)
$inputs.SetColumnSpan($lblInputErr, 2)

# --- Status + progress ---
$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Dock    = 'Top'
$statusPanel.Height  = 58
$statusPanel.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 8)

$status = New-Object System.Windows.Forms.Label
$status.Dock   = 'Top'
$status.Height = 24
$status.Font   = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$status.Text   = 'Enter the Microsoft RDP Application values, choose authproxy.cfg, then click Start.'

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Dock    = 'Top'
$progress.Height  = 20
$progress.Maximum = $script:TotalSteps

$statusPanel.Controls.Add($progress)
$statusPanel.Controls.Add($status)

# --- Log ---
$log = New-Object System.Windows.Forms.RichTextBox
$log.Dock        = 'Fill'
$log.ReadOnly    = $true
$log.BorderStyle = 'None'
$log.WordWrap    = $false
$log.DetectUrls  = $false
$log.BackColor   = [System.Drawing.Color]::FromArgb(30, 30, 30)
$log.ForeColor   = [System.Drawing.Color]::Gainsboro
$log.Font        = New-Object System.Drawing.Font('Consolas', 9.5)

# --- Buttons ---
$bottom = New-Object System.Windows.Forms.FlowLayoutPanel
$bottom.Dock          = 'Bottom'
$bottom.Height        = 58
$bottom.FlowDirection = 'RightToLeft'
$bottom.Padding       = New-Object System.Windows.Forms.Padding(8)

function New-Button ([string]$Text) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Width = 110; $b.Height = 28
    $b
}
$btnClose  = New-Button 'Close'
$btnStart  = New-Button 'START'
$btnLog    = New-Button 'Open Log'
$btnFolder = New-Button 'Open Work Folder'
$bottom.Controls.AddRange([System.Windows.Forms.Control[]]@($btnClose, $btnStart, $btnLog, $btnFolder))

# START: big, bold, green when ready; grey when disabled (flat buttons don't grey out on their own)
$btnStart.Width  = 160
$btnStart.Height = 40
$btnStart.FlatStyle = 'Flat'
$btnStart.FlatAppearance.BorderSize = 0
$btnStart.Font   = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$btnStart.Cursor = [System.Windows.Forms.Cursors]::Hand
function Set-StartStyle {
    if ($btnStart.Enabled) {
        $btnStart.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
        $btnStart.ForeColor = [System.Drawing.Color]::White
    }
    else {
        $btnStart.BackColor = [System.Drawing.Color]::FromArgb(205, 205, 205)
        $btnStart.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    }
}
$btnStart.Add_EnabledChanged({ Set-StartStyle })
Set-StartStyle

# Dock order: highest index docks first (Bottom, then inputs on top, then status), Fill takes the rest
$form.Controls.Add($log)
$form.Controls.Add($statusPanel)
$form.Controls.Add($inputs)
$form.Controls.Add($bottom)

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 150
#endregion

#region ---------- UI helpers ----------
function Write-UiLog ([string]$Message, [string]$Level = 'INFO', [switch]$NoFile) {
    $color = switch ($Level) {
        'OK'     { 'LightGreen' }
        'WARN'   { 'Gold' }
        'ERROR'  { 'Tomato' }
        'HEADER' { 'DeepSkyBlue' }
        default  { 'Gainsboro' }
    }
    $line = if ($Level -eq 'HEADER') { "`r`n--- $Message ---" }
            else { '{0:HH:mm:ss} {1,-5} {2}' -f (Get-Date), $Level, $Message }

    $log.SelectionStart  = $log.TextLength
    $log.SelectionLength = 0
    $log.SelectionColor  = [System.Drawing.Color]::FromName($color)
    $log.AppendText($line + "`r`n")
    $log.ScrollToCaret()

    if (-not $NoFile) {
        $fileLine = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        Add-Content -LiteralPath $script:LogFile -Value $fileLine -ErrorAction SilentlyContinue
    }
}

function Set-Status ([string]$Text, [System.Drawing.Color]$Color = [System.Drawing.SystemColors]::ControlText) {
    $status.Text      = $Text
    $status.ForeColor = $Color
}

function Get-NormalizedHost ([string]$Value) {
    $h = "$Value".Trim()
    $h = $h -replace '^(?i)https?://', ''
    $h = ($h -split '/')[0]
    $h.ToLower()
}

function Get-InputErrors {
    $i = $txtIkey.Text.Trim()
    $s = $txtSkey.Text.Trim()
    $h = Get-NormalizedHost $txtHost.Text
    $errs = @()
    if (-not $i)                             { $errs += 'IKEY is required' }
    elseif ($i -notmatch '^DI[A-Z0-9]{18}$') { $errs += 'IKEY should be 20 characters starting with DI' }
    if (-not $s)                             { $errs += 'SKEY is required' }
    elseif ($s.Length -ne 40)                { $errs += "SKEY should be 40 characters (currently $($s.Length))" }
    if (-not $h)                             { $errs += 'HOST is required' }
    elseif ($h -notmatch '^api-[a-z0-9]+\.duo(security|federal)\.com$') { $errs += 'HOST should look like api-xxxxxxxx.duosecurity.com' }
    $c = $txtCfg.Text.Trim().Trim('"')
    if (-not $c)                                           { $errs += 'authproxy.cfg is required (click Browse)' }
    elseif (-not (Test-Path -LiteralPath $c -PathType Leaf)) { $errs += 'authproxy.cfg file not found' }
    $errs
}

function Update-InputState {
    $errs = @(Get-InputErrors)
    $lblInputErr.Text = if ($errs.Count) { $errs -join '   |   ' } else { '' }
    $btnStart.Enabled = $script:IsAdmin -and -not $script:Job -and $errs.Count -eq 0
}

function Set-InputsEnabled ([bool]$Enabled) {
    $txtIkey.ReadOnly = -not $Enabled
    $txtSkey.ReadOnly = -not $Enabled
    $txtHost.ReadOnly = -not $Enabled
    $txtCfg.ReadOnly  = -not $Enabled
    $btnBrowse.Enabled = $Enabled
}

function Start-Deployment {
    if ($script:Job) { return }
    if (@(Get-InputErrors).Count) { Update-InputState; return }

    $txtHost.Text = Get-NormalizedHost $txtHost.Text
    $txtCfg.Text  = $txtCfg.Text.Trim().Trim('"')
    Set-InputsEnabled $false
    $btnStart.Enabled = $false
    $btnClose.Enabled = $false
    $progress.Value   = 0
    Set-Status 'Starting deployment...'

    $sync = [hashtable]::Synchronized(@{
        Queue    = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
        IKEY     = $txtIkey.Text.Trim()
        SKEY     = $txtSkey.Text.Trim()
        DUO_HOST = $txtHost.Text
        DuoPath  = $script:DuoPath
        CfgPath  = $script:CfgPath
        SourceCfg = $txtCfg.Text
        LogFile  = $script:LogFile
        Failures = @()
        Done     = $false
    })

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($Worker).AddArgument($sync)

    $script:Job = @{ PS = $ps; RS = $rs; Sync = $sync; Handle = $ps.BeginInvoke() }
    $timer.Start()
}

function Complete-Deployment {
    $timer.Stop()
    $job = $script:Job
    try { $job.PS.EndInvoke($job.Handle) | Out-Null } catch { Write-UiLog "Worker error: $($_.Exception.Message)" 'ERROR' }
    if (-not $job.Sync.Done) {
        foreach ($e in $job.PS.Streams.Error) { Write-UiLog "Worker error: $($e.Exception.Message)" 'ERROR' }
    }
    $failures = @($job.Sync.Failures)
    if (-not $job.Sync.Done -and $failures.Count -eq 0) { $failures = @('worker_aborted') }
    try { $job.PS.Dispose(); $job.RS.Close(); $job.RS.Dispose() } catch { }
    $script:Job = $null

    Write-UiLog 'Summary' 'HEADER'
    if ($failures.Count) {
        Write-UiLog ('Deployment completed with errors: {0}' -f ($failures -join '; ')) 'ERROR'
        Set-Status "Finished with $($failures.Count) failure(s). Review the log." ([System.Drawing.Color]::Firebrick)
    }
    else {
        $progress.Value = $progress.Maximum
        Write-UiLog 'Deployment completed successfully' 'OK'
        Set-Status 'Deployment complete. Finish setup in the Proxy Manager.' ([System.Drawing.Color]::ForestGreen)
    }
    Write-UiLog "Log file: $($script:LogFile)"

    Set-InputsEnabled $true
    $btnStart.Text    = 'RUN AGAIN'
    $btnClose.Enabled = $true
    Update-InputState
}
#endregion

#region ---------- Wire-up ----------
if ($IKEY)     { $txtIkey.Text = $IKEY }
if ($SKEY)     { $txtSkey.Text = $SKEY }
if ($DUO_HOST) { $txtHost.Text = $DUO_HOST }
if ($CfgFile)  { $txtCfg.Text  = $CfgFile }

$timer.Add_Tick({
    $job = $script:Job
    if (-not $job) { $timer.Stop(); return }
    $item = $null
    while ($job.Sync.Queue.TryDequeue([ref]$item)) {
        switch ($item.Kind) {
            'log'    { Write-UiLog $item.Text $item.Level -NoFile }
            'status' { Set-Status $item.Text }
            'step'   { if ($progress.Value -lt $progress.Maximum) { $progress.Value++ } }
        }
    }
    if ($job.Handle.IsCompleted) {
        while ($job.Sync.Queue.TryDequeue([ref]$item)) {
            if ($item.Kind -eq 'log') { Write-UiLog $item.Text $item.Level -NoFile }
        }
        Complete-Deployment
    }
})

foreach ($tb in $txtIkey, $txtSkey, $txtHost, $txtCfg) { $tb.Add_TextChanged({ Update-InputState }) }
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title            = 'Select authproxy.cfg'
    $dlg.Filter           = 'Authentication Proxy config (*.cfg)|*.cfg|All files (*.*)|*.*'
    $dlg.CheckFileExists  = $true
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $txtCfg.Text = $dlg.FileName }
    $dlg.Dispose()
})
$chkShow.Add_CheckedChanged({ $txtSkey.UseSystemPasswordChar = -not $chkShow.Checked })
$btnStart.Add_Click({ Start-Deployment })
$btnClose.Add_Click({ $form.Close() })
$btnLog.Add_Click({
    if (Test-Path -LiteralPath $script:LogFile) {
        Start-Process -FilePath "$env:SystemRoot\notepad.exe" -ArgumentList "`"$($script:LogFile)`""
    }
})
$btnFolder.Add_Click({ Start-Process -FilePath "$env:SystemRoot\explorer.exe" -ArgumentList "`"$($script:DuoPath)`"" })
$form.Add_FormClosing({ if ($script:Job) { $_.Cancel = $true } })

$form.Add_Shown({
    $form.Activate()
    Write-UiLog 'DeployDuo ready' 'HEADER'

    if (-not $script:IsAdmin) {
        Write-UiLog 'This tool must run as Administrator (build with -requireAdmin or Run as administrator).' 'ERROR'
        Set-Status 'Administrator rights required. Re-run elevated.' ([System.Drawing.Color]::Firebrick)
    }
    Write-UiLog "Working directory: $($script:DuoPath)"

    foreach ($m in 'ActiveDirectory', 'GroupPolicy') {
        if (-not (Get-Module -ListAvailable -Name $m)) { Write-UiLog "$m PowerShell module not found; AD/GPO steps will fail on this machine." 'WARN' }
    }

    Update-InputState
    if ($btnStart.Enabled) { $form.AcceptButton = $btnStart }
    if ($script:AutoStart -and $btnStart.Enabled) { Start-Deployment }
    elseif (-not $txtIkey.Text) { $txtIkey.Focus() | Out-Null }
})

[void]$form.ShowDialog()
$timer.Dispose()
$form.Dispose()
#endregion

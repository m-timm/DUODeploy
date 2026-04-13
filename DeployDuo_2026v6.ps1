# DeployDuo_Final_v4.ps1
# Comprehensive Duo Windows Login deployment script
# - Logging to C:\Duo
# - No normal popups (only error popups)
# - Robust extraction & file copy
# - Creates DuoAdmin user and group
# - Installs Duo Auth Proxy and deploys config
# - Imports DUO Windows Login GPO backup
# - Generates duo1.mst without modifying the original MSI
# - Installs Duo Windows Logon from SYSVOL scripts
# - Injects HOST / IKEY / SKEY into the imported GPO

[CmdletBinding()]
param(
    [string]$IKEY,
    [string]$SKEY,
    [string]$DUO_HOST
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = 'Stop'

# try to load required modules (silently)
Import-Module ActiveDirectory -ErrorAction SilentlyContinue
Import-Module GroupPolicy -ErrorAction SilentlyContinue

# -------------------------
# Logging + helper function
# -------------------------
$duoPath = 'C:\Duo'
if (-not (Test-Path $duoPath)) { New-Item -Path $duoPath -ItemType Directory -Force | Out-Null }
$logFile = "C:\Duo\DeployDuo_{0}.log" -f (Get-Date -Format yyyyMMdd_HHmmss)

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "$timestamp [$Level] $Message"
    Write-Host $line
    try { Add-Content -Path $logFile -Value $line } catch { Write-Host "WARN: failed to write log: $($_.Exception.Message)" }
    if ($Level -eq 'ERROR') {
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
            [System.Windows.Forms.MessageBox]::Show($Message, 'DeployDuo Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } catch {
            Write-Host "ERROR popup failed: $($_.Exception.Message)"
        }
    }
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } catch {}
    }
}

function Test-MsiTableExists {
    param(
        [Parameter(Mandatory)]$Database,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)]$Installer
    )

    $view = $null
    $record = $null
    $result = $null
    try {
        $sql = "SELECT `Name` FROM `_Tables` WHERE `Name`=?"
        $view = $Database.OpenView($sql)
        $record = $Installer.CreateRecord(1)
        $record.StringData(1) = $TableName
        $view.Execute($record)
        $result = $view.Fetch()
        return ($null -ne $result)
    }
    finally {
        Release-ComObject $result
        Release-ComObject $record
        Release-ComObject $view
    }
}

function Set-MsiProperty {
    param(
        [Parameter(Mandatory)]$Database,
        [Parameter(Mandatory)][string]$PropertyName,
        [Parameter(Mandatory)][string]$PropertyValue,
        [Parameter(Mandatory)]$Installer
    )

    $selectView = $null
    $selectRecord = $null
    $existing = $null
    $modifyView = $null
    $modifyRecord = $null

    try {
        $selectSql = "SELECT `Property`,`Value` FROM `Property` WHERE `Property`=?"
        $selectView = $Database.OpenView($selectSql)
        $selectRecord = $Installer.CreateRecord(1)
        $selectRecord.StringData(1) = $PropertyName
        $selectView.Execute($selectRecord)
        $existing = $selectView.Fetch()

        if ($existing) {
            Write-Log "Updating MSI property '$PropertyName'"
            $updateSql = "UPDATE `Property` SET `Value`=? WHERE `Property`=?"
            $modifyView = $Database.OpenView($updateSql)
            $modifyRecord = $Installer.CreateRecord(2)
            $modifyRecord.StringData(1) = $PropertyValue
            $modifyRecord.StringData(2) = $PropertyName
            $modifyView.Execute($modifyRecord)
        }
        else {
            Write-Log "Inserting MSI property '$PropertyName'"
            $insertSql = "INSERT INTO `Property` (`Property`,`Value`) VALUES (?,?)"
            $modifyView = $Database.OpenView($insertSql)
            $modifyRecord = $Installer.CreateRecord(2)
            $modifyRecord.StringData(1) = $PropertyName
            $modifyRecord.StringData(2) = $PropertyValue
            $modifyView.Execute($modifyRecord)
        }
    }
    finally {
        Release-ComObject $existing
        Release-ComObject $selectRecord
        Release-ComObject $selectView
        Release-ComObject $modifyRecord
        Release-ComObject $modifyView
    }
}

function Get-MsiLastError {
    param([Parameter(Mandatory)]$Installer)

    $record = $null
    try {
        $record = $Installer.LastErrorRecord()
        if ($null -eq $record) { return $null }

        $parts = @()
        for ($i = 0; $i -lt 20; $i++) {
            try {
                $value = $record.StringData($i)
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $parts += $value
                }
            } catch { break }
        }

        if ($parts.Count -gt 0) { return ($parts -join ' | ') }
        return $null
    }
    catch { return $null }
    finally { Release-ComObject $record }
}

function New-DuoTransform {
    param(
        [Parameter(Mandatory)][string]$SourceMsi,
        [Parameter(Mandatory)][string]$OutputMst,
        [Parameter(Mandatory)][string]$IKEY,
        [Parameter(Mandatory)][string]$SKEY,
        [Parameter(Mandatory)][string]$DUO_HOST
    )

    $workingMsi = Join-Path $env:TEMP ("DuoWindowsLogon64_WORKING_{0}.msi" -f ([guid]::NewGuid().ToString('N')))
    $installer = $null
    $originalDb = $null
    $workingDb = $null

    try {
        Write-Log "Copying source MSI to temp working copy for MST generation"
        Copy-Item -Path $SourceMsi -Destination $workingMsi -Force

        if (Test-Path $OutputMst) {
            Write-Log "Removing existing MST: $OutputMst"
            Remove-Item -Path $OutputMst -Force
        }

        $installer = New-Object -ComObject WindowsInstaller.Installer
        $originalDb = $installer.OpenDatabase($SourceMsi, 0)
        $workingDb  = $installer.OpenDatabase($workingMsi, 1)

        if (-not (Test-MsiTableExists -Database $workingDb -TableName 'Property' -Installer $installer)) {
            throw 'Property table does not exist in MSI.'
        }

        Set-MsiProperty -Database $workingDb -PropertyName 'IKEY' -PropertyValue $IKEY -Installer $installer
        Set-MsiProperty -Database $workingDb -PropertyName 'SKEY' -PropertyValue $SKEY -Installer $installer
        Set-MsiProperty -Database $workingDb -PropertyName 'HOST' -PropertyValue $DUO_HOST -Installer $installer

        Write-Log 'Committing changes to temp MSI'
        $workingDb.Commit()

        Write-Log 'Generating duo1.mst transform'
        $generated = $workingDb.GenerateTransform($originalDb, $OutputMst)
        if (-not $generated) {
            throw 'No transform was generated. The databases may not differ.'
        }

        Write-Log 'Creating transform summary information'
        $workingDb.CreateTransformSummaryInfo($originalDb, $OutputMst, 0, 0)

        if (-not (Test-Path $OutputMst)) {
            throw "Transform was not created: $OutputMst"
        }

        Write-Log "MST generated successfully at $OutputMst"
    }
    catch {
        $detail = $null
        if ($installer) { $detail = Get-MsiLastError -Installer $installer }
        if ($detail) {
            throw "$($_.Exception.Message)`nWindows Installer detail: $detail"
        } else {
            throw
        }
    }
    finally {
        Release-ComObject $workingDb
        Release-ComObject $originalDb
        Release-ComObject $installer
        if (Test-Path $workingMsi) {
            Remove-Item -Path $workingMsi -Force -ErrorAction SilentlyContinue
        }
    }
}

$failures = @()
Write-Log '===== DeployDuo started ====='

# -------------------------
# Step 1: Preconditions
# -------------------------
$cfgPath = Join-Path $duoPath 'authproxy.cfg'
if (-not (Test-Path $cfgPath)) {
    Write-Log "Required file missing: $cfgPath" 'ERROR'
    exit 1
}
Write-Log 'Preconditions OK: C:\Duo exists and authproxy.cfg present'

# Collect Duo values once and reuse them for MST + GPO injection
if ([string]::IsNullOrWhiteSpace($IKEY)) { $IKEY = Read-Host 'Enter IKEY' }
if ([string]::IsNullOrWhiteSpace($SKEY)) { $SKEY = Read-Host 'Enter SKEY' }
if ([string]::IsNullOrWhiteSpace($DUO_HOST)) { $DUO_HOST = Read-Host 'Enter HOST' }

if ([string]::IsNullOrWhiteSpace($IKEY) -or [string]::IsNullOrWhiteSpace($SKEY) -or [string]::IsNullOrWhiteSpace($DUO_HOST)) {
    Write-Log 'IKEY, SKEY, and HOST are all required.' 'ERROR'
    exit 1
}
Write-Log 'Captured Duo application values for MST generation and GPO injection'

# -------------------------
# Step 2: Download files
# -------------------------
$urls = @(
    'https://dl.duosecurity.com/DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip',
    'https://dl.duosecurity.com/duoauthproxy-latest.exe',
    'https://raw.githubusercontent.com/m-timm/DUODeploy/main/DUO_Windows_Login_GPO.zip'
)

foreach ($u in $urls) {
    $fileName = Split-Path $u -Leaf
    $dest = Join-Path $duoPath $fileName
    Write-Log "Downloading $fileName from $u"
    try {
        Invoke-WebRequest -Uri $u -OutFile $dest -UseBasicParsing -ErrorAction Stop
        Write-Log "Downloaded $fileName -> $dest"
    } catch {
        Write-Log ("Failed to download {0}: {1}" -f $fileName, $_.Exception.Message) 'ERROR'
        $failures += "download:$fileName"
    }
}

# -------------------------
# Step 3: Extract ZIPs
# -------------------------
$zipFile = Join-Path $duoPath 'DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip'
$extractDir = Join-Path $duoPath 'DuoExtracted'
$gpoZipFile = Join-Path $duoPath 'DUO_Windows_Login_GPO.zip'
$gpoExtractDir = Join-Path $duoPath 'GP'

if (-not (Test-Path $zipFile)) {
    Write-Log "ZIP file not found: $zipFile" 'ERROR'
    $failures += 'zip_missing'
} else {
    try {
        if (Test-Path $extractDir) {
            Remove-Item -Path $extractDir -Recurse -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 200
        }
        Write-Log "Extracting $zipFile -> $extractDir"
        Expand-Archive -Path $zipFile -DestinationPath $extractDir -Force -ErrorAction Stop
        Write-Log 'Duo ZIP extraction complete'
    } catch {
        Write-Log ("Duo ZIP extraction failed: {0}" -f $_.Exception.Message) 'ERROR'
        $failures += 'zip_extract_failed'
    }
}

if (-not (Test-Path $gpoZipFile)) {
    Write-Log "GPO backup ZIP file not found: $gpoZipFile" 'ERROR'
    $failures += 'gpo_zip_missing'
} else {
    try {
        if (Test-Path $gpoExtractDir) {
            Remove-Item -Path $gpoExtractDir -Recurse -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 200
        }
        Write-Log "Extracting $gpoZipFile -> $gpoExtractDir"
        Expand-Archive -Path $gpoZipFile -DestinationPath $gpoExtractDir -Force -ErrorAction Stop
        Write-Log 'GPO backup ZIP extraction complete'
    } catch {
        Write-Log ("GPO backup ZIP extraction failed: {0}" -f $_.Exception.Message) 'ERROR'
        $failures += 'gpo_zip_extract_failed'
    }
}

# -------------------------
# Step 4: Locate & copy files (recursive, robust)
# -------------------------
function Find-Files {
    param([string]$pattern)
    if (-not (Test-Path $extractDir)) { return @() }
    Get-ChildItem -Path $extractDir -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { ($_.Name -notmatch '^(?:\._|__MACOSX)') -and ($_.Name -match $pattern) }
}

$policyRoot = 'C:\Windows\PolicyDefinitions'
$enUS = Join-Path $policyRoot 'en-us'
if (-not (Test-Path $policyRoot)) {
    try { New-Item -Path $policyRoot -ItemType Directory -Force | Out-Null; Write-Log "Created $policyRoot" }
    catch { Write-Log ("Failed to create {0}: {1}" -f $policyRoot,$_.Exception.Message) 'ERROR'; $failures += 'mkdir:PolicyDefinitions' }
}
if (-not (Test-Path $enUS)) {
    try { New-Item -Path $enUS -ItemType Directory -Force | Out-Null; Write-Log "Created $enUS" }
    catch { Write-Log ("Failed to create {0}: {1}" -f $enUS,$_.Exception.Message) 'ERROR'; $failures += 'mkdir:en-us' }
}

$sysvolScripts = 'C:\Windows\SYSVOL\domain\scripts'
if (-not (Test-Path $sysvolScripts)) {
    try { New-Item -Path $sysvolScripts -ItemType Directory -Force | Out-Null; Write-Log "Created $sysvolScripts" }
    catch { Write-Log ("Failed to create {0}: {1}" -f $sysvolScripts,$_.Exception.Message) 'ERROR'; $failures += 'mkdir:SYSVOL' }
}

$admxMatches = Find-Files '(?i)duowindowslogon.*\.admx$'
if ($admxMatches.Count -eq 0) {
    $admxMatches = Get-ChildItem -Path $extractDir -Recurse -File -Filter '*.admx' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '^(?:\._|__MACOSX)' }
}
if ($admxMatches.Count -eq 0) {
    Write-Log 'No ADMX files found in extracted ZIP' 'ERROR'
    $failures += 'missing_admx'
} else {
    foreach ($m in $admxMatches) {
        try {
            Copy-Item -Path $m.FullName -Destination $policyRoot -Force -ErrorAction Stop
            Write-Log ("Copied ADMX: {0} -> {1}" -f $m.FullName, $policyRoot)
        } catch {
            Write-Log ("Failed copying ADMX {0} : {1}" -f $m.FullName, $_.Exception.Message) 'ERROR'
            $failures += 'copy_admx'
        }
    }
}

$admlMatches = Find-Files '(?i)duowindowslogon.*\.adml$'
if ($admlMatches.Count -eq 0) {
    $admlMatches = Get-ChildItem -Path $extractDir -Recurse -File -Filter '*.adml' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '^(?:\._|__MACOSX)' }
}
if ($admlMatches.Count -eq 0) {
    Write-Log 'No ADML files found in extracted ZIP' 'ERROR'
    $failures += 'missing_adml'
} else {
    foreach ($m in $admlMatches) {
        try {
            Copy-Item -Path $m.FullName -Destination $enUS -Force -ErrorAction Stop
            Write-Log ("Copied ADML: {0} -> {1}" -f $m.FullName, $enUS)
        } catch {
            Write-Log ("Failed copying ADML {0} : {1}" -f $m.FullName, $_.Exception.Message) 'ERROR'
            $failures += 'copy_adml'
        }
    }
}

$msiMatches = Find-Files '(?i)duowindowslogon.*\.msi$'
if ($msiMatches.Count -eq 0) {
    $msiMatches = Get-ChildItem -Path $extractDir -Recurse -File -Filter '*.msi' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '^(?:\._|__MACOSX)' }
}
if ($msiMatches.Count -eq 0) {
    Write-Log 'No MSI files found in extracted ZIP' 'ERROR'
    $failures += 'missing_msi'
} else {
    foreach ($m in $msiMatches) {
        try {
            Copy-Item -Path $m.FullName -Destination $sysvolScripts -Force -ErrorAction Stop
            Write-Log ("Copied MSI: {0} -> {1}" -f $m.FullName, $sysvolScripts)
        } catch {
            Write-Log ("Failed copying MSI {0} : {1}" -f $m.FullName, $_.Exception.Message) 'ERROR'
            $failures += 'copy_msi'
        }
    }
}

# -------------------------
# Step 4d: Generate Duo MST from the copied MSI using the captured values
# -------------------------
try {
    $sourceMsi = Join-Path $sysvolScripts 'DuoWindowsLogon64.msi'
    $outputMst = Join-Path $sysvolScripts 'duo1.mst'

    if (-not (Test-Path $sourceMsi)) {
        throw "Source MSI not found for MST generation: $sourceMsi"
    }

    New-DuoTransform -SourceMsi $sourceMsi -OutputMst $outputMst -IKEY $IKEY -SKEY $SKEY -DUO_HOST $DUO_HOST
} catch {
    Write-Log ("Duo MST generation failed: {0}" -f $_.Exception.Message) 'ERROR'
    $failures += 'mst_generation_failed'
}

# -------------------------
# Step 5: Install Duo Windows Logon from SYSVOL scripts (silently)
# -------------------------
try {
    $duoWindowsLogonMsi = Join-Path $sysvolScripts 'DuoWindowsLogon64.msi'
    if (-not (Test-Path $duoWindowsLogonMsi)) {
        throw "Duo Windows Logon MSI not found at $duoWindowsLogonMsi"
    }

    $duoMsiLog = Join-Path $duoPath 'DuoWindowsLogon_Install.log'
    Write-Log 'Installing Duo Windows Logon silently from SYSVOL scripts'
    $duoInstallArgs = @(
        '/i'
        "`"$duoWindowsLogonMsi`""
        "IKEY=`"$IKEY`""
        "SKEY=`"$SKEY`""
        "HOST=`"$DUO_HOST`""
        'FAILOPEN=#1'
        'RDPONLY=#0'
        '/qn'
        '/L*v'
        "`"$duoMsiLog`""
    )
    $duoInstall = Start-Process -FilePath 'msiexec.exe' -ArgumentList $duoInstallArgs -Wait -PassThru -ErrorAction Stop
    if ($duoInstall.ExitCode -ne 0) {
        throw "Duo Windows Logon install failed with exit code $($duoInstall.ExitCode)"
    }
    Write-Log 'Duo Windows Logon installation completed'
} catch {
    Write-Log ("Duo Windows Logon installation failed: {0}" -f $_.Exception.Message) 'ERROR'
    $failures += 'duowinlogon_install_failed'
}

# -------------------------
# Step 6: Install Duo Authentication Proxy (silently)
# -------------------------
$installer = Join-Path $duoPath 'duoauthproxy-latest.exe'
if (Test-Path $installer) {
    try {
        Write-Log 'Installing Duo Authentication Proxy (silent)'
        Start-Process -FilePath $installer -ArgumentList '/S','/V','/qn' -Wait -ErrorAction Stop
        Write-Log 'Auth Proxy installation completed'
    } catch {
        Write-Log ("Auth Proxy installation failed: {0}" -f $_.Exception.Message) 'ERROR'
        $failures += 'authproxy_install_failed'
    }
} else {
    Write-Log "Auth Proxy installer not found at $installer" 'ERROR'
    $failures += 'authproxy_installer_missing'
}

# -------------------------
# Step 7: Create DuoAdmin & create DuoProtected group
# -------------------------
$randPass = 'thisIStheDEEyouOHpwd1!'

try {
    Write-Log 'Ensuring AD user DuoAdmin exists with the requested password'
    $secure = ConvertTo-SecureString $randPass -AsPlainText -Force
    $existingUser = Get-ADUser -Filter "SamAccountName -eq 'DuoAdmin'" -ErrorAction SilentlyContinue

    if (-not $existingUser) {
        New-ADUser -Name 'DuoAdmin' -SamAccountName 'DuoAdmin' -AccountPassword $secure -Enabled $true -PasswordNeverExpires $true -ErrorAction Stop
        Write-Log 'DuoAdmin created'
    } else {
        Set-ADAccountPassword -Identity 'DuoAdmin' -Reset -NewPassword $secure -ErrorAction Stop
        Enable-ADAccount -Identity 'DuoAdmin' -ErrorAction SilentlyContinue
        Set-ADUser -Identity 'DuoAdmin' -PasswordNeverExpires $true -ErrorAction SilentlyContinue
        Write-Log 'DuoAdmin already existed; password reset and account verified'
    }

    try {
        Remove-ADGroupMember -Identity 'Domain Admins' -Members 'DuoAdmin' -Confirm:$false -ErrorAction Stop
        Write-Log 'Removed DuoAdmin from Domain Admins'
    } catch {
        Write-Log 'DuoAdmin was not a member of Domain Admins or could not be removed cleanly; continuing' 'WARN'
    }
} catch {
    Write-Log ("Failed to create or update DuoAdmin: {0}" -f $_.Exception.Message) 'ERROR'
    $failures += 'create_duoadmin_failed'
}

try {
    Write-Log 'Ensuring DuoProtected group exists and required members are present'
    if (-not (Get-ADGroup -Filter "Name -eq 'DuoProtected'" -ErrorAction SilentlyContinue)) {
        New-ADGroup -Name 'DuoProtected' -GroupScope Global -GroupCategory Security -Path ('CN=Users,' + (Get-ADDomain).DistinguishedName) -ErrorAction Stop
        Write-Log 'Created DuoProtected group'
    } else {
        Write-Log 'DuoProtected group already exists'
    }

    foreach ($member in @('DuoAdmin','Domain Admins')) {
        try {
            Add-ADGroupMember -Identity 'DuoProtected' -Members $member -ErrorAction Stop
            Write-Log "Ensured $member is a member of DuoProtected"
        } catch {
            Write-Log "$member is already a member of DuoProtected or could not be added cleanly; continuing" 'WARN'
        }
    }
} catch {
    Write-Log ("DuoProtected group creation/add failed: {0}" -f $_.Exception.Message) 'ERROR'
    $failures += 'duoprotected_failed'
}

# -------------------------
# Step 8: Update authproxy.cfg, deploy it, then run authproxy_passwd
# -------------------------
try {
    Write-Log 'Updating authproxy.cfg service account lines'
    $cfg = Get-Content -Path $cfgPath -ErrorAction Stop
    $cfg = $cfg -replace '; service_account_username=.*', 'service_account_username=DuoAdmin'
    $cfg = $cfg -replace '; service_account_password=.*', "service_account_password=$randPass"
    Set-Content -Path $cfgPath -Value $cfg -Encoding ASCII -ErrorAction Stop
    Write-Log 'authproxy.cfg updated in C:\Duo'
} catch {
    Write-Log ("Failed updating authproxy.cfg: {0}" -f $_.Exception.Message) 'ERROR'
    $failures += 'cfg_update_failed'
}

$cfgDest = 'C:\Program Files\Duo Security Authentication Proxy\conf\authproxy.cfg'
try {
    if (-not (Test-Path (Split-Path $cfgDest))) { New-Item -Path (Split-Path $cfgDest) -ItemType Directory -Force | Out-Null }
    Copy-Item -Path $cfgPath -Destination $cfgDest -Force -ErrorAction Stop
    Write-Log "Deployed authproxy.cfg -> $cfgDest"
} catch {
    Write-Log ("Failed deploying cfg: {0}" -f $_.Exception.Message) 'ERROR'
    $failures += 'cfg_deploy_failed'
}

$passwdExe = 'C:\Program Files\Duo Security Authentication Proxy\bin\authproxy_passwd.exe'
if (Test-Path $passwdExe) {
    try {
        Write-Log 'Running authproxy_passwd --whole-config'
        Start-Process -FilePath $passwdExe -ArgumentList '--whole-config' -Wait -ErrorAction Stop
        Write-Log 'authproxy_passwd completed'
    } catch {
        Write-Log ("authproxy_passwd failed: {0}" -f $_.Exception.Message) 'ERROR'
        $failures += 'authproxy_passwd_failed'
    }
} else {
    Write-Log "authproxy_passwd not found: $passwdExe" 'ERROR'
    $failures += 'authproxy_passwd_missing'
}

# -------------------------
# Step 9: Create and delegate GPO "DUO Windows Login"
# -------------------------
$gpoName = 'DUO Windows Login'
try {
    Write-Log "Creating/configuring GPO '$gpoName'"
    if (-not (Get-GPO -Name $gpoName -ErrorAction SilentlyContinue)) {
        New-GPO -Name $gpoName -ErrorAction Stop | Out-Null
        Write-Log "GPO $gpoName created"
    } else {
        Write-Log "GPO $gpoName exists"
    }

    Set-GPPermissions -Name $gpoName -TargetName 'Authenticated Users' -TargetType Group -PermissionLevel None -Confirm:$false -ErrorAction Stop | Out-Null
    Write-Log 'Removed Authenticated Users from GPO permissions'

    foreach ($g in 'Domain Computers','Domain Controllers') {
        Set-GPPermissions -Name $gpoName -TargetName $g -TargetType Group -PermissionLevel GpoApply -Confirm:$false -ErrorAction Stop | Out-Null
        Write-Log "Set GPO permissions for $g"
    }

    $domainDN = (Get-ADDomain).DistinguishedName
    try {
        New-GPLink -Name $gpoName -Target ("LDAP://" + $domainDN) -Enforced:$false -ErrorAction Stop | Out-Null
    } catch {
        Set-GPLink -Name $gpoName -Target ("LDAP://" + $domainDN) -LinkEnabled Yes -ErrorAction Stop | Out-Null
    }
    Write-Log 'GPO configured and linked at domain root'
} catch {
    Write-Log ("GPO configuration failed: {0}" -f $_.Exception.Message) 'ERROR'
    $failures += 'gpo_failed'
}

# -------------------------
# Step 9b: Import GPO backup into 'DUO Windows Login'
# -------------------------
try {
    Write-Log "Importing GPO backup into '$gpoName' from C:\Duo\GP"

    if (-not (Test-Path $gpoExtractDir)) {
        throw "GPO backup path not found: $gpoExtractDir"
    }

    Import-Module GroupPolicy -ErrorAction Stop

    if (-not (Get-Command Import-GPO -ErrorAction SilentlyContinue)) {
        throw 'Import-GPO is not available on this system. Install the Group Policy Management tools/feature and try again.'
    }

    $backupXml = Get-ChildItem -Path $gpoExtractDir -Recurse -Filter 'bkupInfo.xml' -File -ErrorAction Stop | Select-Object -First 1
    if (-not $backupXml) {
        throw "No bkupInfo.xml file was found under $gpoExtractDir"
    }

    Write-Log ("Found GPO backup metadata file at {0}" -f $backupXml.FullName)
    [xml]$backupInfoXml = Get-Content -Path $backupXml.FullName -ErrorAction Stop

    $backupNameNode = $backupInfoXml.SelectSingleNode("//*[local-name()='GPODisplayName' or local-name()='DisplayName' or local-name()='GPOName']")
    $backupIdNode   = $backupInfoXml.SelectSingleNode("//*[local-name()='ID' or local-name()='BackupId' or local-name()='BackupID']")

    $backupGpoName = $null
    $backupId = $null
    if ($backupNameNode -and $backupNameNode.InnerText) { $backupGpoName = $backupNameNode.InnerText.Trim() }
    if ($backupIdNode -and $backupIdNode.InnerText) { $backupId = $backupIdNode.InnerText.Trim('{} ').Trim() }

    if ($backupGpoName) { Write-Log ("Resolved backup GPO name: {0}" -f $backupGpoName) } else { Write-Log 'Could not resolve backup GPO name from bkupInfo.xml' 'WARN' }
    if ($backupId) { Write-Log ("Resolved backup ID: {0}" -f $backupId) } else { Write-Log 'Could not resolve backup ID from bkupInfo.xml' 'WARN' }

    $imported = $false
    if ($backupGpoName) {
        try {
            Import-GPO -Path $gpoExtractDir -BackupGpoName $backupGpoName -TargetName $gpoName -CreateIfNeeded -ErrorAction Stop | Out-Null
            Write-Log ("Imported GPO backup '{0}' into '{1}' using BackupGpoName" -f $backupGpoName, $gpoName)
            $imported = $true
        } catch {
            Write-Log ("Import by BackupGpoName failed: {0}" -f $_.Exception.Message) 'WARN'
        }
    }

    if (-not $imported -and $backupId) {
        Import-GPO -Path $gpoExtractDir -BackupId $backupId -TargetName $gpoName -CreateIfNeeded -ErrorAction Stop | Out-Null
        Write-Log ("Imported GPO backup ID '{0}' into '{1}'" -f $backupId, $gpoName)
        $imported = $true
    }

    if (-not $imported) {
        throw 'Unable to import the GPO backup. Neither BackupGpoName nor BackupId could be used successfully.'
    }
} catch {
    Write-Log ("GPO backup import failed: {0}" -f $_.Exception.Message) 'ERROR'
    $failures += 'gpo_import_failed'
}

# -------------------------
# Step 9c: Inject Duo HOST / IKEY / SKEY into the imported GPO
# -------------------------
try {
    Write-Log "Injecting Duo policy values into GPO '$gpoName'"
    $regKey = 'HKLM\SOFTWARE\Policies\Duo Security\DuoCredProv'

    Set-GPRegistryValue -Name $gpoName -Key $regKey -ValueName 'HOST' -Type String -Value $DUO_HOST -ErrorAction Stop
    Set-GPRegistryValue -Name $gpoName -Key $regKey -ValueName 'IKEY' -Type String -Value $IKEY -ErrorAction Stop
    Set-GPRegistryValue -Name $gpoName -Key $regKey -ValueName 'SKEY' -Type String -Value $SKEY -ErrorAction Stop

    Write-Log 'Duo policy values injected successfully'

    $verifyReport = Join-Path $duoPath 'Verify_DuoGPO.xml'
    Get-GPOReport -Name $gpoName -ReportType Xml -Path $verifyReport -ErrorAction Stop
    Write-Log "Verification GPO report written to $verifyReport"
} catch {
    Write-Log ("Duo GPO value injection failed: {0}" -f $_.Exception.Message) 'ERROR'
    $failures += 'gpo_value_injection_failed'
}

# -------------------------
# Step 10: Launch Duo Authentication Proxy Manager GUI (interactive)
# -------------------------
try {
    Write-Log 'Launching Duo Authentication Proxy Manager GUI'
    Start-Process 'C:\Program Files\Duo Security Authentication Proxy\bin\local_proxy_manager-win32-x64\Duo_Authentication_Proxy_Manager.exe'
    Write-Log 'Proxy Manager launched'
} catch {
    Write-Log ("Failed to launch Proxy Manager: {0}" -f $_.Exception.Message) 'ERROR'
    $failures += 'launch_proxy_gui_failed'
}

# -------------------------
# Final: Summary & Exit
# -------------------------
if ($failures.Count -gt 0) {
    Write-Log ("Deployment completed with errors: {0}" -f ($failures -join '; ')) 'ERROR'
    exit 1
} else {
    Write-Log 'Deployment completed successfully'
    exit 0
}

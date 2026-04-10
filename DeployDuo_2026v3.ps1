# DeployDuo_Final.ps1
# Comprehensive Duo Windows Login deployment script
# - Logging to C:\Duo
# - No normal popups (only error popups)
# - Robust extraction & file copy
# - Creates DuoAdmin user with static password, groups, GPO, installs Auth Proxy and deploys config

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

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
        [string]$Level = "INFO"
    )
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "$timestamp [$Level] $Message"
    # echo to console and append to log
    Write-Host $line
    try { Add-Content -Path $logFile -Value $line } catch { Write-Host "WARN: failed to write log: $($_.Exception.Message)" }
    if ($Level -eq "ERROR") {
        # show only error popups
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
            [System.Windows.Forms.MessageBox]::Show($Message, "DeployDuo Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } catch {
            Write-Host "ERROR popup failed: $($_.Exception.Message)"
        }
    }
}

$failures = @()
Write-Log "===== DeployDuo started ====="

# -------------------------
# Step 1: Preconditions
# -------------------------
# Ensure Duo folder exists and authproxy.cfg is present
$cfgPath = Join-Path $duoPath 'authproxy.cfg'
if (-not (Test-Path $cfgPath)) {
    Write-Log "Required file missing: $cfgPath" "ERROR"
    exit 1
}
Write-Log "Preconditions OK: C:\Duo exists and authproxy.cfg present"

# -------------------------
# Step 2: Download files
# -------------------------
$urls = @(
    'https://dl.duosecurity.com/DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip',
    'https://dl.duosecurity.com/duo-win-login-latest.exe',
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
        Write-Log ("Failed to download {0}: {1}" -f $fileName, $_.Exception.Message) "ERROR"
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
    Write-Log "ZIP file not found: $zipFile" "ERROR"
    $failures += "zip_missing"
} else {
    try {
        if (Test-Path $extractDir) {
            Remove-Item -Path $extractDir -Recurse -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 200
        }
        Write-Log "Extracting $zipFile -> $extractDir"
        Expand-Archive -Path $zipFile -DestinationPath $extractDir -Force -ErrorAction Stop
        Write-Log "Duo ZIP extraction complete"
    } catch {
        Write-Log ("Duo ZIP extraction failed: {0}" -f $_.Exception.Message) "ERROR"
        $failures += "zip_extract_failed"
    }
}

if (-not (Test-Path $gpoZipFile)) {
    Write-Log "GPO backup ZIP file not found: $gpoZipFile" "ERROR"
    $failures += "gpo_zip_missing"
} else {
    try {
        if (Test-Path $gpoExtractDir) {
            Remove-Item -Path $gpoExtractDir -Recurse -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 200
        }
        Write-Log "Extracting $gpoZipFile -> $gpoExtractDir"
        Expand-Archive -Path $gpoZipFile -DestinationPath $gpoExtractDir -Force -ErrorAction Stop
        Write-Log "GPO backup ZIP extraction complete"
    } catch {
        Write-Log ("GPO backup ZIP extraction failed: {0}" -f $_.Exception.Message) "ERROR"
        $failures += "gpo_zip_extract_failed"
    }
}

# -------------------------
# Step 4: Locate & copy files (recursive, robust)
# -------------------------
# helper: find files under $extractDir while excluding macOS artifacts and '._' files
function Find-Files {
    param([string]$pattern)
    if (-not (Test-Path $extractDir)) { return @() }
    Get-ChildItem -Path $extractDir -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { ($_.Name -notmatch '^(?:\._|__MACOSX)') -and ($_.Name -match $pattern) }
}

# Ensure destination dirs
$policyRoot = 'C:\Windows\PolicyDefinitions'
$enUS = Join-Path $policyRoot 'en-us'
if (-not (Test-Path $policyRoot)) { 
    try { New-Item -Path $policyRoot -ItemType Directory -Force | Out-Null; Write-Log "Created $policyRoot" } 
    catch { Write-Log ("Failed to create {0}: {1}" -f $policyRoot,$_.Exception.Message) "ERROR"; $failures += "mkdir:PolicyDefinitions" }
}
if (-not (Test-Path $enUS)) { 
    try { New-Item -Path $enUS -ItemType Directory -Force | Out-Null; Write-Log "Created $enUS" } 
    catch { Write-Log ("Failed to create {0}: {1}" -f $enUS,$_.Exception.Message) "ERROR"; $failures += "mkdir:en-us" }
}

# SYSVOL scripts destination (pick primary SYSVOL or fallback)
$sysvolScripts = if (Test-Path 'C:\Windows\SYSVOL') { 'C:\Windows\SYSVOL\domain\scripts' } else { 'C:\Windows\SYSVOL_DFSR\domain\scripts' }
if (-not (Test-Path $sysvolScripts)) {
    try { New-Item -Path $sysvolScripts -ItemType Directory -Force | Out-Null; Write-Log "Created $sysvolScripts" } 
    catch { Write-Log ("Failed to create {0}: {1}" -f $sysvolScripts,$_.Exception.Message) "ERROR"; $failures += "mkdir:SYSVOL" }
}

# 4a: Copy ADMX files (case-insensitive match)
$admxMatches = Find-Files '(?i)duowindowslogon.*\.admx$'
if ($admxMatches.Count -eq 0) {
    # fallback: any admx
    $admxMatches = Get-ChildItem -Path $extractDir -Recurse -File -Filter '*.admx' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '^(?:\._|__MACOSX)' }
}
if ($admxMatches.Count -eq 0) {
    Write-Log "No ADMX files found in extracted ZIP" "ERROR"
    $failures += "missing_admx"
} else {
    foreach ($m in $admxMatches) {
        try {
            Copy-Item -Path $m.FullName -Destination $policyRoot -Force -ErrorAction Stop
            Write-Log ("Copied ADMX: {0} -> {1}" -f $m.FullName, $policyRoot)
        } catch {
            Write-Log ("Failed copying ADMX {0} : {1}" -f $m.FullName, $_.Exception.Message) "ERROR"
            $failures += "copy_admx"
        }
    }
}

# 4b: Copy ADML files to en-us
$admlMatches = Find-Files '(?i)duowindowslogon.*\.adml$'
if ($admlMatches.Count -eq 0) {
    $admlMatches = Get-ChildItem -Path $extractDir -Recurse -File -Filter '*.adml' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '^(?:\._|__MACOSX)' }
}
if ($admlMatches.Count -eq 0) {
    Write-Log "No ADML files found in extracted ZIP" "ERROR"
    $failures += "missing_adml"
} else {
    foreach ($m in $admlMatches) {
        try {
            Copy-Item -Path $m.FullName -Destination $enUS -Force -ErrorAction Stop
            Write-Log ("Copied ADML: {0} -> {1}" -f $m.FullName, $enUS)
        } catch {
            Write-Log ("Failed copying ADML {0} : {1}" -f $m.FullName, $_.Exception.Message) "ERROR"
            $failures += "copy_adml"
        }
    }
}

# 4c: Copy MSI files to SYSVOL scripts
$msiMatches = Find-Files '(?i)duowindowslogon.*\.msi$'
if ($msiMatches.Count -eq 0) {
    $msiMatches = Get-ChildItem -Path $extractDir -Recurse -File -Filter '*.msi' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '^(?:\._|__MACOSX)' }
}
if ($msiMatches.Count -eq 0) {
    Write-Log "No MSI files found in extracted ZIP" "ERROR"
    $failures += "missing_msi"
} else {
    foreach ($m in $msiMatches) {
        try {
            Copy-Item -Path $m.FullName -Destination $sysvolScripts -Force -ErrorAction Stop
            Write-Log ("Copied MSI: {0} -> {1}" -f $m.FullName, $sysvolScripts)
        } catch {
            Write-Log ("Failed copying MSI {0} : {1}" -f $m.FullName, $_.Exception.Message) "ERROR"
            $failures += "copy_msi"
        }
    }
}

# -------------------------
# Step 5: Install Duo Authentication Proxy (silently)
# -------------------------
$installer = Join-Path $duoPath 'duoauthproxy-latest.exe'
if (Test-Path $installer) {
    try {
        Write-Log "Installing Duo Authentication Proxy (silent)"
        Start-Process -FilePath $installer -ArgumentList '/S','/V','/qn' -Wait -ErrorAction Stop
        Write-Log "Auth Proxy installation completed"
    } catch {
        Write-Log ("Auth Proxy installation failed: {0}" -f $_.Exception.Message) "ERROR"
        $failures += "authproxy_install_failed"
    }
} else {
    Write-Log "Auth Proxy installer not found at $installer" "ERROR"
    $failures += "authproxy_installer_missing"
}

# -------------------------
# Step 6: Create DuoAdmin & create DuoProtected group
# -------------------------
# Static password per request
$randPass = 'thisIStheDEEyouOHpwd1!'

try {
    Write-Log "Creating AD user DuoAdmin"
    $secure = ConvertTo-SecureString $randPass -AsPlainText -Force
    New-ADUser -Name 'DuoAdmin' -SamAccountName 'DuoAdmin' -AccountPassword $secure -Enabled $true -PasswordNeverExpires $true -ErrorAction Stop
    Write-Log "DuoAdmin created"
} catch {
    Write-Log ("Failed to create DuoAdmin: {0}" -f $_.Exception.Message) "ERROR"
    $failures += "create_duoadmin_failed"
}

# REMOVED: Adding DuoAdmin to Domain Admins group

try {
    Write-Log "Ensuring DuoProtected group exists and adding members"
    if (-not (Get-ADGroup -Filter "Name -eq 'DuoProtected'" -ErrorAction SilentlyContinue)) {
        New-ADGroup -Name 'DuoProtected' -GroupScope Global -GroupCategory Security -Path ('CN=Users,' + (Get-ADDomain).DistinguishedName) -ErrorAction Stop
        Write-Log "Created DuoProtected group"
    } else {
        Write-Log "DuoProtected group already exists"
    }
    Add-ADGroupMember -Identity 'DuoProtected' -Members @('DuoAdmin','Domain Admins') -ErrorAction Stop
    Write-Log "Added DuoAdmin and Domain Admins to DuoProtected"
} catch {
    Write-Log ("DuoProtected group creation/add failed: {0}" -f $_.Exception.Message) "ERROR"
    $failures += "duoprotected_failed"
}

# Save credentials
try {
    Write-Log "Saving DuoAdmin credentials to delete_me.txt"
    "Username: DuoAdmin`nPassword: $randPass" | Out-File -FilePath (Join-Path $duoPath 'delete_me.txt') -Encoding ASCII -ErrorAction Stop
    Write-Log "Credentials saved to delete_me.txt"
} catch {
    Write-Log ("Failed saving credentials: {0}" -f $_.Exception.Message) "ERROR"
    $failures += "save_creds_failed"
}

# -------------------------
# Step 7: Update authproxy.cfg, deploy it, then run authproxy_passwd
# -------------------------
try {
    Write-Log "Updating authproxy.cfg service account lines"
    $cfg = Get-Content -Path $cfgPath -ErrorAction Stop
    $cfg = $cfg -replace '; service_account_username=.*', 'service_account_username=DuoAdmin'
    $cfg = $cfg -replace '; service_account_password=.*', "service_account_password=$randPass"
    Set-Content -Path $cfgPath -Value $cfg -Encoding ASCII -ErrorAction Stop
    Write-Log "authproxy.cfg updated in C:\Duo"
} catch {
    Write-Log ("Failed updating authproxy.cfg: {0}" -f $_.Exception.Message) "ERROR"
    $failures += "cfg_update_failed"
}

# Deploy edited cfg to Duo conf folder
$cfgDest = 'C:\Program Files\Duo Security Authentication Proxy\conf\authproxy.cfg'
try {
    if (-not (Test-Path (Split-Path $cfgDest))) { New-Item -Path (Split-Path $cfgDest) -ItemType Directory -Force | Out-Null }
    Copy-Item -Path $cfgPath -Destination $cfgDest -Force -ErrorAction Stop
    Write-Log "Deployed authproxy.cfg -> $cfgDest"
} catch {
    Write-Log ("Failed deploying cfg: {0}" -f $_.Exception.Message) "ERROR"
    $failures += "cfg_deploy_failed"
}

# Run authproxy_passwd --whole-config
$passwdExe = 'C:\Program Files\Duo Security Authentication Proxy\bin\authproxy_passwd.exe'
if (Test-Path $passwdExe) {
    try {
        Write-Log "Running authproxy_passwd --whole-config"
        Start-Process -FilePath $passwdExe -ArgumentList '--whole-config' -Wait -ErrorAction Stop
        Write-Log "authproxy_passwd completed"
    } catch {
        Write-Log ("authproxy_passwd failed: {0}" -f $_.Exception.Message) "ERROR"
        $failures += "authproxy_passwd_failed"
    }
} else {
    Write-Log "authproxy_passwd not found: $passwdExe" "ERROR"
    $failures += "authproxy_passwd_missing"
}

# -------------------------
# Step 8: Create and delegate GPO "DUO Windows Login"
# -------------------------
try {
    Write-Log "Creating/configuring GPO 'DUO Windows Login'"
    $gpoName = 'DUO Windows Login'
    if (-not (Get-GPO -Name $gpoName -ErrorAction SilentlyContinue)) {
        New-GPO -Name $gpoName -ErrorAction Stop | Out-Null
        Write-Log "GPO $gpoName created"
    } else {
        Write-Log "GPO $gpoName exists"
    }
    # Remove Authenticated Users from delegation (suppress confirmation)
    Set-GPPermissions -Name $gpoName -TargetName 'Authenticated Users' -TargetType Group -PermissionLevel None -Confirm:$false -ErrorAction Stop | Out-Null
    Write-Log "Removed Authenticated Users from GPO permissions"
    
    # Grant Domain Computers and Domain Controllers Read & Apply (GpoApply) - suppress confirmation
    foreach ($g in 'Domain Computers','Domain Controllers') {
        Set-GPPermissions -Name $gpoName -TargetName $g -TargetType Group -PermissionLevel GpoApply -Confirm:$false -ErrorAction Stop | Out-Null
        Write-Log "Set GPO permissions for $g"
    }
    
    # Link at domain root
    $domainDN = (Get-ADDomain).DistinguishedName
    try {
        New-GPLink -Name $gpoName -Target ("LDAP://" + $domainDN) -Enforced:$false -ErrorAction Stop | Out-Null
    } catch {
        Set-GPLink -Name $gpoName -Target ("LDAP://" + $domainDN) -LinkEnabled Yes -ErrorAction Stop | Out-Null
    }
    Write-Log "GPO configured and linked at domain root"
} catch {
    Write-Log ("GPO configuration failed: {0}" -f $_.Exception.Message) "ERROR"
    $failures += "gpo_failed"
}

# -------------------------
# Step 8b: Import GPO backup into 'DUO Windows Login'
# -------------------------
try {
    Write-Log "Importing GPO backup into 'DUO Windows Login' from C:\Duo\GP"

    if (-not (Test-Path $gpoExtractDir)) {
        throw "GPO backup path not found: $gpoExtractDir"
    }

    Import-Module GroupPolicy -ErrorAction Stop

    if (-not (Get-Command Import-GPO -ErrorAction SilentlyContinue)) {
        throw "Import-GPO is not available on this system. Install the Group Policy Management tools/feature and try again."
    }

    $backupXml = Get-ChildItem -Path $gpoExtractDir -Recurse -Filter 'bkupInfo.xml' -File -ErrorAction Stop |
        Select-Object -First 1

    if (-not $backupXml) {
        throw "No bkupInfo.xml file was found under $gpoExtractDir"
    }

    Write-Log ("Found GPO backup metadata file at {0}" -f $backupXml.FullName)

    [xml]$backupInfoXml = Get-Content -Path $backupXml.FullName -ErrorAction Stop

    $backupNameNode = $backupInfoXml.SelectSingleNode("//*[local-name()='GPODisplayName' or local-name()='DisplayName' or local-name()='GPOName']")
    $backupIdNode   = $backupInfoXml.SelectSingleNode("//*[local-name()='ID' or local-name()='BackupId' or local-name()='BackupID']")

    $backupGpoName = $null
    $backupId = $null

    if ($backupNameNode -and $backupNameNode.InnerText) {
        $backupGpoName = $backupNameNode.InnerText.Trim()
    }

    if ($backupIdNode -and $backupIdNode.InnerText) {
        $backupId = $backupIdNode.InnerText.Trim('{} ').Trim()
    }

    if ($backupGpoName) {
        Write-Log ("Resolved backup GPO name: {0}" -f $backupGpoName)
    } else {
        Write-Log "Could not resolve backup GPO name from bkupInfo.xml" "WARN"
    }

    if ($backupId) {
        Write-Log ("Resolved backup ID: {0}" -f $backupId)
    } else {
        Write-Log "Could not resolve backup ID from bkupInfo.xml" "WARN"
    }

    $imported = $false

    if ($backupGpoName) {
        try {
            Import-GPO -Path $gpoExtractDir -BackupGpoName $backupGpoName -TargetName $gpoName -CreateIfNeeded -ErrorAction Stop | Out-Null
            Write-Log ("Imported GPO backup '{0}' into '{1}' using BackupGpoName" -f $backupGpoName, $gpoName)
            $imported = $true
        } catch {
            Write-Log ("Import by BackupGpoName failed: {0}" -f $_.Exception.Message) "WARN"
        }
    }

    if (-not $imported -and $backupId) {
        Import-GPO -Path $gpoExtractDir -BackupId $backupId -TargetName $gpoName -CreateIfNeeded -ErrorAction Stop | Out-Null
        Write-Log ("Imported GPO backup ID '{0}' into '{1}'" -f $backupId, $gpoName)
        $imported = $true
    }

    if (-not $imported) {
        throw "Unable to import the GPO backup. Neither BackupGpoName nor BackupId could be used successfully."
    }
} catch {
    Write-Log ("GPO backup import failed: {0}" -f $_.Exception.Message) "ERROR"
    $failures += "gpo_import_failed"
}

# -------------------------
# Step 9: Launch Duo Authentication Proxy Manager GUI (interactive)
# -------------------------
try {
    Write-Log "Launching Duo Authentication Proxy Manager GUI"
    Start-Process 'C:\Program Files\Duo Security Authentication Proxy\bin\local_proxy_manager-win32-x64\Duo_Authentication_Proxy_Manager.exe'
    Write-Log "Proxy Manager launched"
} catch {
    Write-Log ("Failed to launch Proxy Manager: {0}" -f $_.Exception.Message) "ERROR"
    $failures += "launch_proxy_gui_failed"
}

# -------------------------
# Final: Summary & Exit
# -------------------------
if ($failures.Count -gt 0) {
    Write-Log ("Deployment completed with errors: {0}" -f ($failures -join '; ')) "ERROR"
    exit 1
} else {
    Write-Log "Deployment completed successfully"
    exit 0
}
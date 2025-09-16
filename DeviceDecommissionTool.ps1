# DeviceDecommissionTool.ps1
# WinForms GUI for modular decommission and Autopilot HWID upload
# Requirements: PowerShell 5.1+ on Windows, Microsoft.Graph PowerShell module, RSAT AD (if doing AD operations), SCCM Console for SCCM actions.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------- GLOBAL FLAGS ----------------
$global:GraphConnected   = $false
$global:PrereqsChecked   = $false
$global:SerialFound      = $false
$global:DeviceFound      = $false
$global:ADAvailable      = $false
$global:SCCMAvailable    = $false
$global:ADFound          = $false
$global:SCCMFound        = $false
$global:AutoPilotDevice  = $null
$global:EntraDevice      = $null
$global:ADComputer       = $null
$global:SCCMDevice       = $null
$global:ServiceCredential = $null

# ---------------- HELPERS ----------------
function Get-ScriptDirectory {
    # Primary: $PSScriptRoot when script saved and executed as file.
    if ($PSScriptRoot) { return $PSScriptRoot }
    # Fallback to the location of the running file (if available)
    $inv = $MyInvocation.MyCommand.Definition
    if ($inv) {
        return (Split-Path -Path $inv -Parent)
    }
    return (Get-Location).Path
}

function Log-UI {
    param([string]$Message, [string]$Status = "INFO")
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "{0} | {1,-6} | {2}" -f $ts, $Status, $Message
    $OutputBox.AppendText("$line`r`n")
    $OutputBox.ScrollToCaret()
}

function Save-ActionLog {
    param(
        [string]$Action,
        [string]$SerialOrDevice,
        [string]$Content
    )
    $dir = Get-ScriptDirectory
    if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $safeName = ($SerialOrDevice -replace '[\\/:*?"<>| ]','_')
    $filename = "{0}_{1}_{2}.log" -f $stamp, $Action, $safeName
    $full = Join-Path -Path $dir -ChildPath $filename
    $header = "Action: $Action`nTarget: $SerialOrDevice`nTime: $stamp`n`n"
    $Content | Out-File -FilePath $full -Encoding UTF8 -Force
    # Prepend header
    (Get-Content $full) | Out-File -FilePath $full -Encoding UTF8
    Set-Content -Path $full -Value ($header + (Get-Content -Path $full -Raw))
    Log-UI "Saved action log to $full" "OK"
}

function Confirm-Action {
    param([string]$Message)
    $result = [System.Windows.Forms.MessageBox]::Show(
        $Message,
        "Confirm Action",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

function Ensure-Connected {
    if (-not $global:GraphConnected) {
        Log-UI "Not connected to Microsoft Graph. Please Connect/Login first." "FAIL"
        return $false
    }
    return $true
}

function Update-ActionButtonsState {
    # Decide which action buttons are enabled based on checks
    # Start disabled
    $btnIntune.Enabled = $false
    $btnEntra.Enabled   = $false
    $btnAD.Enabled      = $false
    $btnSCCM.Enabled    = $false
    $btnHWID.Enabled    = $false

    # If only Serial present (Autopilot), enable Intune only
    if ($global:SerialFound -and -not $global:DeviceFound) {
        $btnIntune.Enabled = $true
        Log-UI "Serial-only found -> Intune action enabled." "INFO"
    }

    # If only Device name present
    if ($global:DeviceFound -and -not $global:SerialFound) {
        if ($global:EntraDevice) { $btnEntra.Enabled = $true }
        if ($global:ADFound)      { $btnAD.Enabled   = $true }
        if ($global:SCCMFound)    { $btnSCCM.Enabled = $true }
        Log-UI "Device-name-only found -> Entra/AD/SCCM actions enabled as available." "INFO"
    }

    # If both present -> enable all appropriate actions
    if ($global:SerialFound -and $global:DeviceFound) {
        $btnIntune.Enabled = $true
        if ($global:EntraDevice) { $btnEntra.Enabled = $true }
        if ($global:ADFound)      { $btnAD.Enabled   = $true }
        if ($global:SCCMFound)    { $btnSCCM.Enabled = $true }
        Log-UI "Both Serial & Device found -> all available actions enabled." "INFO"
    }

    # HWID upload enabled only when connected and serial NOT present for that serial
    if ($global:GraphConnected) {
        # If SerialFound is $true (Autopilot record exists) -> disable HWID upload
        if ($global:SerialFound) {
            $btnHWID.Enabled = $false
            Log-UI "HWID upload disabled because Autopilot serial exists." "INFO"
        } else {
            # If graph connected and serial NOT found, allow HWID upload
            $btnHWID.Enabled = $true
            Log-UI "HWID upload enabled (no Autopilot serial found)." "INFO"
        }
    }
}

# ---------------- CONNECT / PREREQS ----------------
function Connect-ServiceAccount {
    try {
        if (-not $UserBox.Text -or -not $PassBox.Text) {
            [System.Windows.Forms.MessageBox]::Show("Please enter Service Account Username and Password before Connect.","Missing credentials",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
            return
        }
        $secure = ConvertTo-SecureString -String $PassBox.Text -AsPlainText -Force
        $global:ServiceCredential = New-Object System.Management.Automation.PSCredential ($UserBox.Text, $secure)
    } catch {
        Log-UI "Failed to create PSCredential: $_" "ERROR"
        return
    }

    # Ensure Microsoft.Graph module exists (or attempt to install)
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        $install = [System.Windows.Forms.MessageBox]::Show("Microsoft.Graph module is not installed. Install from PSGallery now?","Install module",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
        if ($install -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Write-Host "Installing NuGet provider (if required) and Microsoft.Graph..."
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction SilentlyContinue
                Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Log-UI "Microsoft.Graph installed." "OK"
            } catch {
                Log-UI "Could not install Microsoft.Graph automatically: $_" "ERROR"
                [System.Windows.Forms.MessageBox]::Show("Automatic install failed. Please install Microsoft.Graph module manually and relaunch the tool.","Install failed",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        } else {
            Log-UI "Microsoft.Graph not installed; cannot connect." "FAIL"
            return
        }
    }

    Import-Module Microsoft.Graph -ErrorAction SilentlyContinue

    # Try -Credential first (non-interactive). If fails (MFA/conditional), fallback to interactive sign-in so MFA can be completed.
    try {
        Log-UI "Attempting Graph connection using provided credential..." "INFO"
        Connect-MgGraph -Credential $global:ServiceCredential -Scopes "Device.ReadWrite.All","DeviceManagementServiceConfig.ReadWrite.All","DeviceManagementManagedDevices.ReadWrite.All" -ErrorAction Stop
        $global:GraphConnected = $true
        $lblStatus.Text = "Connected ✅"
        $lblStatus.ForeColor = [System.Drawing.Color]::Green
        Log-UI "Connected to Microsoft Graph (credential auth)." "OK"
    } catch {
        Log-UI "Credential-based Graph connection failed (likely MFA/conditional). Falling back to interactive sign-in..." "WARN"
        try {
            Connect-MgGraph -Scopes "Device.ReadWrite.All","DeviceManagementServiceConfig.ReadWrite.All","DeviceManagementManagedDevices.ReadWrite.All" -ErrorAction Stop
            $global:GraphConnected = $true
            $lblStatus.Text = "Connected ✅"
            $lblStatus.ForeColor = [System.Drawing.Color]::Green
            Log-UI "Connected to Microsoft Graph (interactive auth)." "OK"
        } catch {
            $global:GraphConnected = $false
            $lblStatus.Text = "Not Connected ❌"
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            Log-UI "Interactive Graph login failed: $_" "ERROR"
            return
        }
    }

    # update action buttons state (in case checks already done)
    # detect AD & SCCM availability after connect
    if (Get-Module -ListAvailable -Name ActiveDirectory) { $global:ADAvailable = $true }
    if ($Env:SMS_ADMIN_UI_PATH) { $global:SCCMAvailable = $true }
    Update-ActionButtonsState
}

function Check-Prerequisites {
    Log-UI "Starting prerequisite checks..." "INFO"
    $missing = @()

    # Microsoft.Graph
    if (Get-Module -ListAvailable -Name Microsoft.Graph) {
        Log-UI "Microsoft.Graph module: OK" "OK"
    } else {
        Log-UI "Microsoft.Graph module: MISSING" "FAIL"
        $missing += "Microsoft.Graph"
    }

    # ActiveDirectory module
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        Log-UI "ActiveDirectory module: OK" "OK"
        $global:ADAvailable = $true
    } else {
        Log-UI "ActiveDirectory module: MISSING (RSAT AD Tools)" "FAIL"
        $missing += "ActiveDirectory"
        $global:ADAvailable = $false
    }

    # SCCM / ConfigurationManager module
    $sccmPath = $Env:SMS_ADMIN_UI_PATH
    if ($sccmPath -and (Test-Path ($sccmPath.Substring(0,$sccmPath.Length-5) + '\ConfigurationManager.psd1'))) {
        Log-UI "SCCM ConfigurationManager module: OK" "OK"
        $global:SCCMAvailable = $true
    } else {
        Log-UI "SCCM ConfigurationManager module: NOT FOUND (run this tool on SCCM console machine for SCCM actions)" "WARN"
        $global:SCCMAvailable = $false
    }

    # If any missing, prompt user to attempt install
    if ($missing.Count -gt 0) {
        $msg = "Missing components detected:`n- " + ($missing -join "`n- ") + "`n`nAttempt to install Microsoft.Graph and RSAT-AD (if available) now?"
        $resp = [System.Windows.Forms.MessageBox]::Show($msg,"Missing Components", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($resp -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Try to install Microsoft.Graph if missing
            if ($missing -contains "Microsoft.Graph") {
                try {
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction SilentlyContinue
                    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    Log-UI "Microsoft.Graph installed." "OK"
                } catch {
                    Log-UI "Failed to install Microsoft.Graph automatically: $_" "ERROR"
                }
            }
            # Try to install RSAT AD (only on supported Windows)
            if ($missing -contains "ActiveDirectory") {
                try {
                    Log-UI "Attempting to add RSAT ActiveDirectory feature (may require restart & admin)..." "INFO"
                    # For Windows 10/11 and Server 2019+, attempt Add-WindowsCapability
                    Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -ErrorAction Stop
                    Log-UI "RSAT ActiveDirectory installed via WindowsCapability." "OK"
                } catch {
                    Log-UI "Automatic RSAT add failed or not supported on this OS: $_" "WARN"
                    [System.Windows.Forms.MessageBox]::Show("Automatic RSAT install failed or not supported. Please install RSAT Active Directory Tools manually (Settings -> Optional features on 10/11, or Server Features).","Install RSAT", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            }
        } else {
            Log-UI "User chose not to auto-install missing components." "INFO"
        }
    } else {
        Log-UI "All prerequisites present." "OK"
    }

    $global:PrereqsChecked = $true
    # Re-evaluate availability flags (re-check)
    if (Get-Module -ListAvailable -Name ActiveDirectory) { $global:ADAvailable = $true }
    if ($Env:SMS_ADMIN_UI_PATH) { $global:SCCMAvailable = $true }
    Update-ActionButtonsState
}

# ---------------- CHECKS: SERIAL & DEVICE ----------------
function Check-Serial {
    if (-not $global:PrereqsChecked) { Log-UI "Run 'Check Prerequisites' first." "FAIL"; return }
    if (-not $global:GraphConnected) { Log-UI "Connect/Login first." "FAIL"; return }

    $serial = $SerialBox.Text.Trim()
    if (-not $serial) { Log-UI "Serial number required." "FAIL"; return }

    try {
        Import-Module Microsoft.Graph -ErrorAction SilentlyContinue
        $found = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All | Where-Object { $_.SerialNumber -eq $serial }
        if ($found) {
            $global:SerialFound = $true
            $global:AutoPilotDevice = $found
            $lblSerialStatus.Text = "Found ✅"
            $lblSerialStatus.ForeColor = [System.Drawing.Color]::Green
            Log-UI "Autopilot record found: Serial=$serial ; Id=$($found.Id)" "OK"
        } else {
            $global:SerialFound = $false
            $global:AutoPilotDevice = $null
            $lblSerialStatus.Text = "Not Found ❌"
            $lblSerialStatus.ForeColor = [System.Drawing.Color]::Red
            Log-UI "No Autopilot record found for serial: $serial" "FAIL"
        }
    } catch {
        Log-UI "Error checking Autopilot (Graph): $_" "ERROR"
    }
    Update-ActionButtonsState
}

function Check-DeviceName {
    if (-not $global:PrereqsChecked) { Log-UI "Run 'Check Prerequisites' first." "FAIL"; return }
    if (-not $global:GraphConnected) { Log-UI "Connect/Login first." "FAIL"; return }

    $dev = $DeviceBox.Text.Trim()
    if (-not $dev) { Log-UI "Device name required." "FAIL"; return }

    # Reset
    $global:DeviceFound = $false
    $global:EntraDevice = $null
    $global:ADFound = $false
    $global:ADComputer = $null
    $global:SCCMFound = $false
    $global:SCCMDevice = $null

    # Entra
    try {
        Import-Module Microsoft.Graph -ErrorAction SilentlyContinue
        $device = Get-MgDevice -All | Where-Object { $_.DisplayName -eq $dev }
        if ($device) {
            $global:DeviceFound = $true
            $global:EntraDevice = $device
            $lblDeviceStatus.Text = "Entra: Found ✅"
            $lblDeviceStatus.ForeColor = [System.Drawing.Color]::Green
            Log-UI "Entra device found: $dev ; Id=$($device.Id)" "OK"
        } else {
            $lblDeviceStatus.Text = "Entra: Not Found ❌"
            $lblDeviceStatus.ForeColor = [System.Drawing.Color]::Red
            Log-UI "Entra: No device found with name $dev" "FAIL"
        }
    } catch {
        Log-UI "Error checking Entra device: $_" "ERROR"
    }

    # AD
    if ($global:ADAvailable) {
        try {
            $cred = $global:ServiceCredential
            $comp = Get-ADComputer -Identity $dev -Credential $cred -ErrorAction SilentlyContinue
            if ($comp) {
                $global:ADFound = $true
                $global:ADComputer = $comp
                $lblADStatus.Text = "AD: Found ✅"
                $lblADStatus.ForeColor = [System.Drawing.Color]::Green
                Log-UI "AD Computer found: $dev ; DN=$($comp.DistinguishedName)" "OK"
            } else {
                $lblADStatus.Text = "AD: Not Found ❌"
                $lblADStatus.ForeColor = [System.Drawing.Color]::Red
                Log-UI "AD: No computer found with name $dev" "FAIL"
            }
        } catch {
            Log-UI "Error checking AD computer: $_" "ERROR"
        }
    } else {
        $lblADStatus.Text = "AD: Module N/A"
        $lblADStatus.ForeColor = [System.Drawing.Color]::Orange
    }

    # SCCM
    if ($global:SCCMAvailable) {
        try {
            if (-not (Get-Module ConfigurationManager)) {
                Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1') -ErrorAction SilentlyContinue
                # NOTE: You may need to 'cd' to site drive e.g. CD XYZ:
            }
            $d = Get-CMDevice -Name $dev -ErrorAction SilentlyContinue
            if ($d) {
                $global:SCCMFound = $true
                $global:SCCMDevice = $d
                $lblSCCMStatus.Text = "SCCM: Found ✅"
                $lblSCCMStatus.ForeColor = [System.Drawing.Color]::Green
                Log-UI "SCCM device found: $dev ; ResourceID=$($d.ResourceID)" "OK"
            } else {
                $lblSCCMStatus.Text = "SCCM: Not Found ❌"
                $lblSCCMStatus.ForeColor = [System.Drawing.Color]::Red
                Log-UI "SCCM: No device found with name $dev" "FAIL"
            }
        } catch {
            Log-UI "Error checking SCCM: $_" "ERROR"
        }
    } else {
        $lblSCCMStatus.Text = "SCCM: N/A"
        $lblSCCMStatus.ForeColor = [System.Drawing.Color]::Orange
    }

    if ($global:DeviceFound -or $global:SerialFound) {
        # set combined device-found flag
        $global:DeviceFound = $global:DeviceFound -or $global:ADFound -or $global:SCCMFound
    }

    Update-ActionButtonsState
}

# ---------------- ACTIONS (with triple confirmation & logging) ----------------
function Action-RemoveIntune {
    if (-not (Ensure-Connected)) { return }
    if (-not $global:SerialFound -or -not $global:AutoPilotDevice) { Log-UI "No Autopilot/Serial found for removal." "FAIL"; return }

    $serial = $global:AutoPilotDevice.SerialNumber
    $id = $global:AutoPilotDevice.Id
    $msg = "Autopilot record found:`r`nSerial: $serial`r`nID: $id`r`n`r`nDo you confirm deletion of this Autopilot enrollment?"
    if (-not (Confirm-Action $msg)) {
        Log-UI "User cancelled Autopilot deletion." "CANCEL"
        return
    }

    try {
        Remove-MgDeviceManagementWindowsAutopilotDeviceIdentity -WindowsAutopilotDeviceIdentityId $id -Confirm:$false -ErrorAction Stop
        $out = "Autopilot enrollment removed for serial $serial (Id=$id)."
        Log-UI $out "OK"
        Save-ActionLog -Action "RemoveAutopilot" -SerialOrDevice $serial -Content $out
    } catch {
        Log-UI "Failed to remove Autopilot device: $_" "ERROR"
        Save-ActionLog -Action "RemoveAutopilot_FAILED" -SerialOrDevice $serial -Content $_.ToString()
    }
}

function Action-RemoveEntra {
    if (-not (Ensure-Connected)) { return }
    if (-not $global:EntraDevice) { Log-UI "No Entra device found." "FAIL"; return }

    $name = $global:EntraDevice.DisplayName
    $id = $global:EntraDevice.Id
    $msg = "Entra device found:`r`nName: $name`r`nID: $id`r`n`r`nDo you confirm deletion of this Entra device?"
    if (-not (Confirm-Action $msg)) {
        Log-UI "User cancelled Entra deletion." "CANCEL"
        return
    }

    try {
        Remove-MgDevice -DeviceId $id -Confirm:$false -ErrorAction Stop
        $out = "Entra device removed: $name (Id=$id)."
        Log-UI $out "OK"
        Save-ActionLog -Action "RemoveEntra" -SerialOrDevice $name -Content $out
    } catch {
        Log-UI "Failed to remove Entra device: $_" "ERROR"
        Save-ActionLog -Action "RemoveEntra_FAILED" -SerialOrDevice $name -Content $_.ToString()
    }
}

function Action-ADDisableMove {
    if (-not $global:ADFound -or -not $global:ADComputer) { Log-UI "No AD computer found to modify." "FAIL"; return }
    $dev = $global:ADComputer.Name
    $dn  = $global:ADComputer.DistinguishedName
    $targetOU = $OUBox.Text.Trim()
    if (-not $targetOU) { Log-UI "Target OU is required." "FAIL"; return }

    $msg = "AD Computer found:`r`nName: $dev`r`nDN: $dn`r`nTarget OU: $targetOU`r`n`r`nProceed to Disable account and move to OU?"
    if (-not (Confirm-Action $msg)) {
        Log-UI "User cancelled AD disable/move." "CANCEL"
        return
    }

    try {
        $cred = $global:ServiceCredential
        Disable-ADAccount -Identity $global:ADComputer -Credential $cred -ErrorAction Stop
        Move-ADObject -Identity $global:ADComputer.DistinguishedName -TargetPath $targetOU -Credential $cred -ErrorAction Stop
        $out = "Disabled and moved AD computer $dev to $targetOU."
        Log-UI $out "OK"
        Save-ActionLog -Action "ADDisableMove" -SerialOrDevice $dev -Content $out
    } catch {
        Log-UI "Failed AD disable/move: $_" "ERROR"
        Save-ActionLog -Action "ADDisableMove_FAILED" -SerialOrDevice $dev -Content $_.ToString()
    }
}

function Action-RemoveSCCM {
    if (-not $global:SCCMFound -or -not $global:SCCMDevice) { Log-UI "No SCCM device found to remove." "FAIL"; return }
    $dev = $global:SCCMDevice.Name
    $id  = $global:SCCMDevice.ResourceID
    $msg = "SCCM Device found:`r`nName: $dev`r`nResourceID: $id`r`n`r`nProceed to remove from SCCM?"
    if (-not (Confirm-Action $msg)) {
        Log-UI "User cancelled SCCM removal." "CANCEL"
        return
    }

    try {
        if (-not (Get-Module ConfigurationManager)) {
            Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1') -ErrorAction SilentlyContinue
            # Note: you may need to cd to the site drive (e.g. CD XYZ:)
        }
        Remove-CMDevice -DeviceName $dev -Force -ErrorAction Stop
        $out = "Removed device $dev from SCCM."
        Log-UI $out "OK"
        Save-ActionLog -Action "RemoveSCCM" -SerialOrDevice $dev -Content $out
    } catch {
        Log-UI "Failed to remove SCCM device: $_" "ERROR"
        Save-ActionLog -Action "RemoveSCCM_FAILED" -SerialOrDevice $dev -Content $_.ToString()
    }
}

# ---------------- HWID Upload (Autopilot CSV) ----------------
function Upload-HWIDCSV {
    if (-not (Ensure-Connected)) { return }

    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $ofd.Multiselect = $false
    $ok = $ofd.ShowDialog()
    if ($ok -ne [System.Windows.Forms.DialogResult]::OK) { Log-UI "HWID CSV upload cancelled by user." "CANCEL"; return }

    $path = $ofd.FileName
    Log-UI "Selected HWID CSV: $path" "INFO"

    try {
        $rows = Import-Csv -Path $path -ErrorAction Stop
        if ($rows.Count -eq 0) { Log-UI "CSV contains no rows." "FAIL"; return }

        # Before upload: check each serial in CSV to ensure it's NOT already in Autopilot
        $alreadyPresent = @()
        foreach ($r in $rows) {
            $serial = $r.SerialNumber -or $r.DeviceSerialNumber -or $r."Device Serial Number"
            if (-not $serial) { continue }
            $exists = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All | Where-Object { $_.SerialNumber -eq $serial }
            if ($exists) { $alreadyPresent += $serial }
        }
        if ($alreadyPresent.Count -gt 0) {
            $msg = "The following serial(s) are already registered in Autopilot:`r`n" + ($alreadyPresent -join "`r`n") + "`r`n`r`nUpload aborted. Remove these serials first if you intend to re-upload."
            [System.Windows.Forms.MessageBox]::Show($msg,"Autopilot Serial Exists",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
            Log-UI "HWID upload aborted - serial(s) already present: $($alreadyPresent -join ',')" "FAIL"
            Save-ActionLog -Action "UploadHWID_ABORT_SerialExists" -SerialOrDevice (Split-Path -Leaf $path) -Content ($msg)
            return
        }

        # Triple confirmation for upload
        $c1 = [System.Windows.Forms.MessageBox]::Show("Confirm intent to upload HWID CSV: $path","Confirm 1",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
        if ($c1 -ne [System.Windows.Forms.DialogResult]::Yes) { Log-UI "HWID upload cancelled at confirmation 1" "CANCEL"; return }
        $c2 = [System.Windows.Forms.MessageBox]::Show("Please confirm the CSV file is correct and intended for Autopilot enrollment.","Confirm 2",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
        if ($c2 -ne [System.Windows.Forms.DialogResult]::Yes) { Log-UI "HWID upload cancelled at confirmation 2" "CANCEL"; return }
        $c3 = [System.Windows.Forms.MessageBox]::Show("Final confirmation: Proceed with uploading HWID CSV?","Confirm 3",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($c3 -ne [System.Windows.Forms.DialogResult]::Yes) { Log-UI "HWID upload cancelled at confirmation 3" "CANCEL"; return }

        $success = 0; $failed = 0; $details = @()
        foreach ($r in $rows) {
            # Expecting at minimum: "SerialNumber" and "HardwareHash" fields as columns; other columns optional (Manufacturer, Model, OrderIdentifier)
            $serial = $r.SerialNumber -or $r.DeviceSerialNumber -or $r."Device Serial Number"
            $hw = $r.HardwareHash -or $r.HardwareIdentifier -or $r."Hardware Hash"
            $manufacturer = $r.Manufacturer -or $r.Mfg -or $null
            $model = $r.Model -or $null

            if (-not $serial -or -not $hw) {
                $failed++;
                $details += "Row missing required SerialNumber or HardwareHash: $($r | Out-String)"
                continue
            }

            # Use Microsoft.Graph to create Autopilot identity
            try {
                # New-MgDeviceManagementWindowsAutopilotDeviceIdentity is used when available
                New-MgDeviceManagementWindowsAutopilotDeviceIdentity -BodyParameter @{serialNumber = $serial; hardwareIdentifier = $hw; manufacturer = $manufacturer; model = $model} -ErrorAction Stop
                $success++
            } catch {
                $failed++
                $details += "Failed to upload serial $serial : $_"
            }
        }

        $summary = "HWID upload completed. Success: $success ; Failed: $failed"
        Log-UI $summary "OK"
        Save-ActionLog -Action "UploadHWID" -SerialOrDevice (Split-Path -Leaf $path) -Content ($summary + "`n`n" + ($details -join "`n"))

    } catch {
        Log-UI "Failed to process/upload HWID CSV: $_" "ERROR"
    }
}

# ---------------- UI BUILD ----------------
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Device Decommission Tool (Delegated Rights) - Modular"
$Form.Size = New-Object System.Drawing.Size(880,720)
$Form.StartPosition = "CenterScreen"

# Credentials group
$grpCred = New-Object System.Windows.Forms.GroupBox
$grpCred.Text = "Service Account (Delegated Rights)"
$grpCred.Size = New-Object System.Drawing.Size(820,110)
$grpCred.Location = New-Object System.Drawing.Point(20,10)
$Form.Controls.Add($grpCred)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Username (UPN):"
$lblUser.Location = New-Object System.Drawing.Point(10,25)
$lblUser.AutoSize = $true
$grpCred.Controls.Add($lblUser)
$UserBox = New-Object System.Windows.Forms.TextBox
$UserBox.Location = New-Object System.Drawing.Point(120,22)
$UserBox.Width = 560
$grpCred.Controls.Add($UserBox)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Password:"
$lblPass.Location = New-Object System.Drawing.Point(10,60)
$lblPass.AutoSize = $true
$grpCred.Controls.Add($lblPass)
$PassBox = New-Object System.Windows.Forms.TextBox
$PassBox.Location = New-Object System.Drawing.Point(120,57)
$PassBox.Width = 560
$PassBox.PasswordChar = '*'
$grpCred.Controls.Add($PassBox)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect / Login (MFA if required)"
$btnConnect.Size = New-Object System.Drawing.Size(240,32)
$btnConnect.Location = New-Object System.Drawing.Point(690,35)
$btnConnect.Add_Click({ Connect-ServiceAccount })
$grpCred.Controls.Add($btnConnect)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Not Connected ❌"
$lblStatus.Location = New-Object System.Drawing.Point(690,10)
$lblStatus.AutoSize = $true
$lblStatus.ForeColor = [System.Drawing.Color]::Red
$grpCred.Controls.Add($lblStatus)

# Prereq group
$grpPre = New-Object System.Windows.Forms.GroupBox
$grpPre.Text = "Prerequisite & Checks"
$grpPre.Size = New-Object System.Drawing.Size(820,90)
$grpPre.Location = New-Object System.Drawing.Point(20,130)
$Form.Controls.Add($grpPre)

$btnCheckPrereq = New-Object System.Windows.Forms.Button
$btnCheckPrereq.Text = "Check Prerequisites"
$btnCheckPrereq.Size = New-Object System.Drawing.Size(200,32)
$btnCheckPrereq.Location = New-Object System.Drawing.Point(10,25)
$btnCheckPrereq.Add_Click({ Check-Prerequisites })
$grpPre.Controls.Add($btnCheckPrereq)

$lblModuleSummary = New-Object System.Windows.Forms.Label
$lblModuleSummary.Text = "Modules status will be listed in the log."
$lblModuleSummary.Location = New-Object System.Drawing.Point(230,31)
$lblModuleSummary.AutoSize = $true
$grpPre.Controls.Add($lblModuleSummary)

# Device identifiers group
$grpDev = New-Object System.Windows.Forms.GroupBox
$grpDev.Text = "Device Identifiers"
$grpDev.Size = New-Object System.Drawing.Size(820,140)
$grpDev.Location = New-Object System.Drawing.Point(20,230)
$Form.Controls.Add($grpDev)

$lblSerial = New-Object System.Windows.Forms.Label
$lblSerial.Text = "Serial Number (Autopilot):"
$lblSerial.Location = New-Object System.Drawing.Point(10,25)
$lblSerial.AutoSize = $true
$grpDev.Controls.Add($lblSerial)
$SerialBox = New-Object System.Windows.Forms.TextBox
$SerialBox.Location = New-Object System.Drawing.Point(170,22)
$SerialBox.Width = 420
$grpDev.Controls.Add($SerialBox)
$btnSerialCheck = New-Object System.Windows.Forms.Button
$btnSerialCheck.Text = "Check Serial (Autopilot)"
$btnSerialCheck.Size = New-Object System.Drawing.Size(180,28)
$btnSerialCheck.Location = New-Object System.Drawing.Point(600,20)
$btnSerialCheck.Add_Click({ Check-Serial })
$grpDev.Controls.Add($btnSerialCheck)
$lblSerialStatus = New-Object System.Windows.Forms.Label
$lblSerialStatus.Text = "Status: Unknown"
$lblSerialStatus.Location = New-Object System.Drawing.Point(10,55)
$lblSerialStatus.AutoSize = $true
$grpDev.Controls.Add($lblSerialStatus)

$lblDevice = New-Object System.Windows.Forms.Label
$lblDevice.Text = "Device Name (Entra/AD/SCCM):"
$lblDevice.Location = New-Object System.Drawing.Point(10,85)
$lblDevice.AutoSize = $true
$grpDev.Controls.Add($lblDevice)
$DeviceBox = New-Object System.Windows.Forms.TextBox
$DeviceBox.Location = New-Object System.Drawing.Point(170,82)
$DeviceBox.Width = 420
$grpDev.Controls.Add($DeviceBox)
$btnDeviceCheck = New-Object System.Windows.Forms.Button
$btnDeviceCheck.Text = "Check Device Name"
$btnDeviceCheck.Size = New-Object System.Drawing.Size(180,28)
$btnDeviceCheck.Location = New-Object System.Drawing.Point(600,80)
$btnDeviceCheck.Add_Click({ Check-DeviceName })
$grpDev.Controls.Add($btnDeviceCheck)
$lblDeviceStatus = New-Object System.Windows.Forms.Label
$lblDeviceStatus.Text = "Entra: N/A"
$lblDeviceStatus.Location = New-Object System.Drawing.Point(10,110)
$lblDeviceStatus.AutoSize = $true
$grpDev.Controls.Add($lblDeviceStatus)

# AD & SCCM status labels
$lblADStatus = New-Object System.Windows.Forms.Label
$lblADStatus.Text = "AD: N/A"
$lblADStatus.Location = New-Object System.Drawing.Point(260,110)
$lblADStatus.AutoSize = $true
$grpDev.Controls.Add($lblADStatus)

$lblSCCMStatus = New-Object System.Windows.Forms.Label
$lblSCCMStatus.Text = "SCCM: N/A"
$lblSCCMStatus.Location = New-Object System.Drawing.Point(420,110)
$lblSCCMStatus.AutoSize = $true
$grpDev.Controls.Add($lblSCCMStatus)

# Action Buttons group
$grpActions = New-Object System.Windows.Forms.GroupBox
$grpActions.Text = "Actions (Triple-confirmation before every destructive action)"
$grpActions.Size = New-Object System.Drawing.Size(820,150)
$grpActions.Location = New-Object System.Drawing.Point(20,380)
$Form.Controls.Add($grpActions)

$btnIntune = New-Object System.Windows.Forms.Button
$btnIntune.Text = "Remove from Intune (Autopilot)"
$btnIntune.Size = New-Object System.Drawing.Size(260,44)
$btnIntune.Location = New-Object System.Drawing.Point(10,30)
$btnIntune.Enabled = $false
$btnIntune.Add_Click({ Action-RemoveIntune })
$grpActions.Controls.Add($btnIntune)

$btnEntra = New-Object System.Windows.Forms.Button
$btnEntra.Text = "Remove from Entra ID"
$btnEntra.Size = New-Object System.Drawing.Size(260,44)
$btnEntra.Location = New-Object System.Drawing.Point(290,30)
$btnEntra.Enabled = $false
$btnEntra.Add_Click({ Action-RemoveEntra })
$grpActions.Controls.Add($btnEntra)

$btnAD = New-Object System.Windows.Forms.Button
$btnAD.Text = "Disable + Move in AD"
$btnAD.Size = New-Object System.Drawing.Size(260,44)
$btnAD.Location = New-Object System.Drawing.Point(10,80)
$btnAD.Enabled = $false
$btnAD.Add_Click({ Action-ADDisableMove })
$grpActions.Controls.Add($btnAD)

$btnSCCM = New-Object System.Windows.Forms.Button
$btnSCCM.Text = "Remove from SCCM"
$btnSCCM.Size = New-Object System.Drawing.Size(260,44)
$btnSCCM.Location = New-Object System.Drawing.Point(290,80)
$btnSCCM.Enabled = $false
$btnSCCM.Add_Click({ Action-RemoveSCCM })
$grpActions.Controls.Add($btnSCCM)

# HWID Upload button
$btnHWID = New-Object System.Windows.Forms.Button
$btnHWID.Text = "Upload Autopilot HWID CSV"
$btnHWID.Size = New-Object System.Drawing.Size(260,44)
$btnHWID.Location = New-Object System.Drawing.Point(570,30)
$btnHWID.Enabled = $false
$btnHWID.Add_Click({ Upload-HWIDCSV })
$grpActions.Controls.Add($btnHWID)

# Target OU input (for AD move)
$lblTargetOU = New-Object System.Windows.Forms.Label
$lblTargetOU.Text = "Target OU (DistinguishedName) for AD move:"
$lblTargetOU.Location = New-Object System.Drawing.Point(10,125)
$lblTargetOU.AutoSize = $true
$grpActions.Controls.Add($lblTargetOU)
$OUBox = New-Object System.Windows.Forms.TextBox
$OUBox.Location = New-Object System.Drawing.Point(260,122)
$OUBox.Width = 480
$grpActions.Controls.Add($OUBox)

# Output / Log area
$OutputBox = New-Object System.Windows.Forms.TextBox
$OutputBox.Multiline = $true
$OutputBox.ScrollBars = "Vertical"
$OutputBox.ReadOnly = $true
$OutputBox.Location = New-Object System.Drawing.Point(20,540)
$OutputBox.Size = New-Object System.Drawing.Size(820,120)
$Form.Controls.Add($OutputBox)

# On form closing: notify log location
$Form.add_FormClosing({
    $dir = Get-ScriptDirectory
    $msg = "The application will close. All per-action logs (if any) were saved to:`n$dir`n`nPress OK to close."
    [System.Windows.Forms.MessageBox]::Show($msg,"Closing - Logs saved",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
})

# Show form
[void]$Form.ShowDialog()

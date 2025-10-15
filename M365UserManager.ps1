# M365UserManager.ps1
# Microsoft 365 Copilot User Management Tool
# 
# A Windows Forms GUI application for provisioning and managing Microsoft 365 users
# with automatic Copilot licensing in your Azure tenant.
#
# Features:
# - Create users with M365 E5 + Copilot licenses
# - View all tenant users with license details
# - Bulk deprovision selected users
# - Real-time logging and status updates
# - Secure Microsoft Graph authentication
#
# Prerequisites:
# - Windows PowerShell 5.1+ or PowerShell 7+
# - Microsoft.Graph PowerShell module
# - Microsoft 365 tenant admin privileges
# - Available M365 E5 and Copilot licenses
#
# Usage:
# 1. Run .\Setup-Environment.ps1 (first time only)
# 2. Run .\Setup-Configuration.ps1 (configure tenant)
# 3. Launch .\M365UserManager.ps1
#
# For quick setup: .\Quick-Start.ps1
#

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$global:ConfigPath = "$ScriptRoot\Config\settings.json"
$global:LogPath = "$ScriptRoot\Logs\$(Get-Date -Format 'yyyy-MM-dd-HHmmss')_Complete_UserManagement.log"
$global:IsConnected = $false

# Ensure Logs directory exists
if (-not (Test-Path "$ScriptRoot\Logs")) {
    New-Item -Path "$ScriptRoot\Logs" -ItemType Directory -Force | Out-Null
}

function Write-GuiLog {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Add to GUI output
    $outputBox.AppendText("$logEntry`r`n")
    $outputBox.ScrollToCaret()
    
    # Add to log file
    if (Test-Path (Split-Path $global:LogPath)) {
        Add-Content -Path $global:LogPath -Value $logEntry
    }
    
    # Update UI
    [System.Windows.Forms.Application]::DoEvents()
}

function Connect-ToMicrosoftGraph {
    try {
        if (-not $global:IsConnected) {
            $config = Get-Content $global:ConfigPath | ConvertFrom-Json
            Write-GuiLog "Connecting to Microsoft Graph..."
            
            # Define required scopes
            $graphScopes = @(
                'User.ReadWrite.All',
                'Directory.ReadWrite.All',
                'Directory.Read.All',
                'User.Read.All',
                'Organization.Read.All'
            )
            
            # Connect to Microsoft Graph
            Connect-MgGraph -Scopes $graphScopes -TenantId $config.tenant.tenantId -NoWelcome
            $graphContext = Get-MgContext
            
            if ($graphContext) {
                $global:IsConnected = $true
                
                # Update all connect buttons
                $connectButton.Text = "Connected"
                $connectButton.BackColor = [System.Drawing.Color]::LightGreen
                $mgmtConnectButton.Text = "Connected"
                $mgmtConnectButton.BackColor = [System.Drawing.Color]::LightGreen
                
                # Enable buttons
                $provisionButton.Enabled = $true
                $refreshUsersButton.Enabled = $true
                $deprovisionSelectedButton.Enabled = $true
                
                Write-GuiLog "Successfully connected to Microsoft Graph"
                Write-GuiLog "Account: $($graphContext.Account)"
                Write-GuiLog "Tenant: $($graphContext.TenantId)"
                return $true
            } else {
                Write-GuiLog "Connection failed" -Level "ERROR"
                return $false
            }
        } else {
            Write-GuiLog "Already connected to Microsoft Graph"
            return $true
        }
    } catch {
        Write-GuiLog "Connection error: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Start-UserProvisioning {
    try {
        # Disable provision button during processing
        $provisionButton.Enabled = $false
        $provisionButton.Text = "Processing..."
        
        # Validate required fields
        if ([string]::IsNullOrWhiteSpace($firstNameBox.Text) -or [string]::IsNullOrWhiteSpace($lastNameBox.Text) -or [string]::IsNullOrWhiteSpace($usernameBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("First Name, Last Name, and Username are required!", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        Write-GuiLog "=== STARTING USER PROVISIONING ===" -Level "INFO"
        
        # Load configuration
        $config = Get-Content $global:ConfigPath | ConvertFrom-Json
        
        # Create user parameters
        $userParams = @{
            DisplayName = "$($firstNameBox.Text) $($lastNameBox.Text)"
            UserPrincipalName = "$($usernameBox.Text)@$($config.tenant.domain)"
            MailNickname = $usernameBox.Text
            UsageLocation = "US"
            AccountEnabled = $true
            PasswordProfile = @{
                ForceChangePasswordNextSignIn = $true
                Password = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count 12 | ForEach-Object {[char]$_}) + "!"
            }
        }

        # Create user account
        Write-GuiLog "Creating user account for: $($userParams.DisplayName)"
        $newUser = New-MgUser @userParams
        $temporaryPassword = $userParams.PasswordProfile.Password
        Write-GuiLog "User created successfully: $($newUser.DisplayName)"
        Write-GuiLog "User ID: $($newUser.Id)"
        
        # Assign licenses
        Write-GuiLog "Attempting to assign M365 E5 + Copilot licenses..."
        $licenseSuccess = $false
        try {
            $e5Sku = Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -eq $config.licensing.m365E5Sku}
            $copilotSku = Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -eq $config.licensing.copilotSku}
            
            if ($e5Sku -and $copilotSku) {
                $e5Available = $e5Sku.PrepaidUnits.Enabled - $e5Sku.ConsumedUnits
                $copilotAvailable = $copilotSku.PrepaidUnits.Enabled - $copilotSku.ConsumedUnits
                
                Write-GuiLog "E5 licenses available: $e5Available"
                Write-GuiLog "Copilot licenses available: $copilotAvailable"
                
                if ($e5Available -gt 0 -and $copilotAvailable -gt 0) {
                    $licenseAssignments = @(
                        @{ SkuId = $e5Sku.SkuId; DisabledPlans = @() }
                        @{ SkuId = $copilotSku.SkuId; DisabledPlans = @() }
                    )
                    
                    Set-MgUserLicense -UserId $newUser.Id -AddLicenses $licenseAssignments -RemoveLicenses @()
                    Write-GuiLog "M365 E5 + Copilot licenses assigned successfully!"
                    $licenseSuccess = $true
                } else {
                    Write-GuiLog "Insufficient licenses available (E5: $e5Available, Copilot: $copilotAvailable)" -Level "WARN"
                }
            } else {
                Write-GuiLog "License SKUs not found in tenant" -Level "WARN"
            }
        } catch {
            Write-GuiLog "License assignment failed: $($_.Exception.Message)" -Level "WARN"
        }
        
        # Final status summary
        Write-GuiLog "=== PROVISIONING COMPLETED ===" -Level "INFO"
        Write-GuiLog "User Account: $($newUser.UserPrincipalName)"
        Write-GuiLog "Display Name: $($newUser.DisplayName)"
        Write-GuiLog "Temporary Password: $temporaryPassword"
        Write-GuiLog "M365 Copilot Status: $(if ($licenseSuccess) { 'ENABLED' } else { 'PENDING' })"
        
        # Show completion message
        $completionMessage = "User provisioning completed successfully!`n`n"
        $completionMessage += "User: $($newUser.DisplayName)`n"
        $completionMessage += "Email: $($newUser.UserPrincipalName)`n"
        $completionMessage += "Temporary Password: $temporaryPassword`n"
        $completionMessage += "M365 Copilot: $(if ($licenseSuccess) { 'ENABLED' } else { 'Licensing pending' })`n`n"
        $completionMessage += "You can now go to the 'Manage Users' tab to see this user in the list."
        
        [System.Windows.Forms.MessageBox]::Show($completionMessage, "Provisioning Completed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Clear form for next user
        Clear-ProvisionForm
        
        # Auto-refresh user list if management tab is active
        if ($userListView.Items.Count -gt 0) {
            Write-GuiLog "Auto-refreshing user list to show new user..."
            Load-UserList
        }
        
    } catch {
        Write-GuiLog "Provisioning failed: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Provisioning failed: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        # Re-enable provision button
        $provisionButton.Enabled = $true
        $provisionButton.Text = "Provision User"
    }
}

function Clear-ProvisionForm {
    $firstNameBox.Text = ""
    $lastNameBox.Text = ""
    $usernameBox.Text = ""
}

function Load-UserList {
    try {
        if (-not $global:IsConnected) {
            [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first", "Connection Required", "OK", "Warning")
            return
        }

        Write-GuiLog "Loading user list from tenant..."
        $userListView.BeginUpdate()
        $userListView.Items.Clear()
        
        # Show loading indicator
        $loadingItem = New-Object System.Windows.Forms.ListViewItem("Loading users...")
        $loadingItem.SubItems.Add("Please wait...")
        $loadingItem.SubItems.Add("")
        $loadingItem.SubItems.Add("")
        [void]$userListView.Items.Add($loadingItem)
        $userListView.EndUpdate()
        [System.Windows.Forms.Application]::DoEvents()

        # Get all users from the tenant
        Write-GuiLog "Querying all users in tenant..."
        $users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled,UserType,CreatedDateTime | 
                 Where-Object { 
                    $_.UserPrincipalName -and 
                    $_.UserPrincipalName -notlike "*#EXT#*" -and
                    $_.DisplayName -and
                    $_.DisplayName -ne "Guest" -and
                    $_.Id
                 } |
                 Sort-Object DisplayName

        Write-GuiLog "Retrieved $($users.Count) users from tenant"

        # Clear loading and rebuild list
        $userListView.BeginUpdate()
        $userListView.Items.Clear()

        foreach ($user in $users) {
            try {
                # Get license information
                $licenseInfo = "Checking..."
                try {
                    $userLicenses = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction SilentlyContinue
                    if ($userLicenses) {
                        $licenseNames = @()
                        foreach ($license in $userLicenses) {
                            $sku = Get-MgSubscribedSku | Where-Object {$_.SkuId -eq $license.SkuId} | Select-Object -First 1
                            if ($sku) {
                                $licenseNames += $sku.SkuPartNumber
                            }
                        }
                        if ($licenseNames.Count -gt 0) {
                            $licenseInfo = $licenseNames -join ", "
                        } else {
                            $licenseInfo = "No licenses"
                        }
                    } else {
                        $licenseInfo = "No licenses"
                    }
                } catch {
                    $licenseInfo = "Error checking"
                }

                # Create list item with proper null handling
                $displayName = if ($user.DisplayName) { $user.DisplayName } else { "No Name" }
                $upn = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { "No UPN" }
                $enabled = if ($user.AccountEnabled -ne $null) { $user.AccountEnabled.ToString() } else { "Unknown" }

                $listItem = New-Object System.Windows.Forms.ListViewItem($displayName)
                $listItem.SubItems.Add($upn)
                $listItem.SubItems.Add($enabled)
                $listItem.SubItems.Add($licenseInfo)
                $listItem.Tag = $user
                
                [void]$userListView.Items.Add($listItem)

                # Update UI periodically
                if ($userListView.Items.Count % 10 -eq 0) {
                    [System.Windows.Forms.Application]::DoEvents()
                }

            } catch {
                Write-GuiLog "Error processing user $($user.UserPrincipalName): $($_.Exception.Message)" -Level "WARN"
            }
        }

        $userListView.EndUpdate()
        Write-GuiLog "Successfully loaded $($userListView.Items.Count) users"
        
    } catch {
        $userListView.EndUpdate()
        Write-GuiLog "Error loading user list: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error loading users: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

function Start-SelectedUserDeprovisioning {
    try {
        # Get selected users
        $selectedUsers = @()
        foreach ($item in $userListView.CheckedItems) {
            $selectedUsers += $item.Tag
        }

        if ($selectedUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one user to deprovision.", "No Users Selected", "OK", "Warning")
            return
        }

        # Confirm deprovisioning
        $userList = ($selectedUsers | ForEach-Object { "- $($_.DisplayName) ($($_.UserPrincipalName))" }) -join "`n"
        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to deprovision $($selectedUsers.Count) user(s)?`n`n$userList`n`nThis will:`n- Remove all licenses`n- Delete user accounts`n- Remove mailbox access`n`nThis action cannot be undone!",
            "Confirm Bulk Deprovisioning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($confirmation -eq [System.Windows.Forms.DialogResult]::No) {
            Write-GuiLog "Bulk deprovisioning cancelled by user"
            return
        }

        # Disable buttons during processing
        $deprovisionSelectedButton.Enabled = $false
        $deprovisionSelectedButton.Text = "Processing..."
        $refreshUsersButton.Enabled = $false

        Write-GuiLog "=== STARTING BULK DEPROVISIONING ===" -Level "INFO"
        Write-GuiLog "Processing $($selectedUsers.Count) selected users..."

        $successCount = 0
        $errorCount = 0

        foreach ($user in $selectedUsers) {
            Write-GuiLog "Processing: $($user.DisplayName) ($($user.UserPrincipalName))"
            Write-GuiLog "  -> User ID: $($user.Id)" -Level "INFO"

            try {
                # Remove licenses first (like we proved works)
                Write-GuiLog "  -> Removing licenses..."
                
                # Ensure we have a valid user ID
                if (-not $user.Id) {
                    Write-GuiLog "  -> ERROR: User ID is missing or empty" -Level "ERROR"
                    $errorCount++
                    continue
                }
                
                $licenses = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction SilentlyContinue
                if ($licenses) {
                    $licensesToRemove = $licenses | ForEach-Object { $_.SkuId }
                    Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses $licensesToRemove
                    Write-GuiLog "  -> Removed $($licenses.Count) licenses"
                    Start-Sleep -Seconds 1
                } else {
                    Write-GuiLog "  -> No licenses to remove"
                }

                # Delete user account (this is the exact method that worked in CLI)
                Write-GuiLog "  -> Deleting user account..."
                Remove-MgUser -UserId $user.Id -Confirm:$false
                Write-GuiLog "  -> SUCCESS: User account deleted" -Level "INFO"
                $successCount++

            } catch {
                Write-GuiLog "  -> FAILED: $($_.Exception.Message)" -Level "ERROR"
                $errorCount++
            }
            
            Write-GuiLog "  ----------------------------------------"
        }

        Write-GuiLog "=== BULK DEPROVISIONING COMPLETED ===" -Level "INFO"
        Write-GuiLog "Successfully deprovisioned: $successCount users" -Level "INFO"
        if ($errorCount -gt 0) {
            Write-GuiLog "Failed to deprovision: $errorCount users" -Level "WARN"
        }

        # Refresh the user list to show results
        Write-GuiLog "Refreshing user list to show current state..."
        Load-UserList

        # Show completion summary
        $summaryMessage = "Bulk deprovisioning completed!`n`n"
        $summaryMessage += "Successfully deprovisioned: $successCount users`n"
        if ($errorCount -gt 0) {
            $summaryMessage += "Failed to deprovision: $errorCount users`n`n"
            $summaryMessage += "Check the output log for detailed error information."
        } else {
            $summaryMessage += "`nAll selected users were successfully removed from the tenant."
        }

        [System.Windows.Forms.MessageBox]::Show(
            $summaryMessage,
            "Deprovisioning Completed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

    } catch {
        Write-GuiLog "Bulk deprovisioning failed: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Bulk deprovisioning failed: $($_.Exception.Message)", "Error", "OK", "Error")
    } finally {
        # Re-enable buttons
        $deprovisionSelectedButton.Enabled = $true
        $deprovisionSelectedButton.Text = "Deprovision Selected"
        $refreshUsersButton.Enabled = $true
    }
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "M365 Copilot User Management - Complete"
$form.Size = New-Object System.Drawing.Size(900, 750)
$form.StartPosition = "CenterScreen"

# Create TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(860, 680)
$form.Controls.Add($tabControl)

# Tab 1: User Provisioning
$provisionTab = New-Object System.Windows.Forms.TabPage
$provisionTab.Text = "Create Users"
$provisionTab.BackColor = [System.Drawing.Color]::White
$tabControl.TabPages.Add($provisionTab)

# Tab 2: User Management
$managementTab = New-Object System.Windows.Forms.TabPage
$managementTab.Text = "Manage Users"
$managementTab.BackColor = [System.Drawing.Color]::White
$tabControl.TabPages.Add($managementTab)

# === PROVISION TAB CONTROLS ===

# Connection Section
$connectionGroup = New-Object System.Windows.Forms.GroupBox
$connectionGroup.Text = "Connection Status"
$connectionGroup.Location = New-Object System.Drawing.Point(20, 20)
$connectionGroup.Size = New-Object System.Drawing.Size(800, 60)
$provisionTab.Controls.Add($connectionGroup)

$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Text = "Connect to Microsoft Graph"
$connectButton.Location = New-Object System.Drawing.Point(20, 25)
$connectButton.Size = New-Object System.Drawing.Size(200, 25)
$connectButton.Add_Click({ Connect-ToMicrosoftGraph })
$connectionGroup.Controls.Add($connectButton)

# User Information Section
$userGroup = New-Object System.Windows.Forms.GroupBox
$userGroup.Text = "User Information"
$userGroup.Location = New-Object System.Drawing.Point(20, 100)
$userGroup.Size = New-Object System.Drawing.Size(800, 120)
$provisionTab.Controls.Add($userGroup)

# First Name
$firstNameLabel = New-Object System.Windows.Forms.Label
$firstNameLabel.Text = "First Name:"
$firstNameLabel.Location = New-Object System.Drawing.Point(20, 30)
$firstNameLabel.Size = New-Object System.Drawing.Size(80, 20)
$userGroup.Controls.Add($firstNameLabel)

$firstNameBox = New-Object System.Windows.Forms.TextBox
$firstNameBox.Location = New-Object System.Drawing.Point(110, 27)
$firstNameBox.Size = New-Object System.Drawing.Size(150, 20)
$userGroup.Controls.Add($firstNameBox)

# Last Name
$lastNameLabel = New-Object System.Windows.Forms.Label
$lastNameLabel.Text = "Last Name:"
$lastNameLabel.Location = New-Object System.Drawing.Point(280, 30)
$lastNameLabel.Size = New-Object System.Drawing.Size(80, 20)
$userGroup.Controls.Add($lastNameLabel)

$lastNameBox = New-Object System.Windows.Forms.TextBox
$lastNameBox.Location = New-Object System.Drawing.Point(370, 27)
$lastNameBox.Size = New-Object System.Drawing.Size(150, 20)
$userGroup.Controls.Add($lastNameBox)

# Username
$usernameLabel = New-Object System.Windows.Forms.Label
$usernameLabel.Text = "Username:"
$usernameLabel.Location = New-Object System.Drawing.Point(20, 65)
$usernameLabel.Size = New-Object System.Drawing.Size(80, 20)
$userGroup.Controls.Add($usernameLabel)

$usernameBox = New-Object System.Windows.Forms.TextBox
$usernameBox.Location = New-Object System.Drawing.Point(110, 62)
$usernameBox.Size = New-Object System.Drawing.Size(150, 20)
$userGroup.Controls.Add($usernameBox)

# Button Group
$buttonGroup = New-Object System.Windows.Forms.GroupBox
$buttonGroup.Text = "Actions"
$buttonGroup.Location = New-Object System.Drawing.Point(20, 240)
$buttonGroup.Size = New-Object System.Drawing.Size(800, 60)
$provisionTab.Controls.Add($buttonGroup)

$provisionButton = New-Object System.Windows.Forms.Button
$provisionButton.Text = "Provision User"
$provisionButton.Location = New-Object System.Drawing.Point(20, 25)
$provisionButton.Size = New-Object System.Drawing.Size(120, 25)
$provisionButton.BackColor = [System.Drawing.Color]::LightBlue
$provisionButton.Enabled = $false
$provisionButton.Add_Click({ Start-UserProvisioning })
$buttonGroup.Controls.Add($provisionButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "Clear Form"
$clearButton.Location = New-Object System.Drawing.Point(150, 25)
$clearButton.Size = New-Object System.Drawing.Size(100, 25)
$clearButton.Add_Click({ Clear-ProvisionForm })
$buttonGroup.Controls.Add($clearButton)

# Output Section
$outputGroup = New-Object System.Windows.Forms.GroupBox
$outputGroup.Text = "Output Log"
$outputGroup.Location = New-Object System.Drawing.Point(20, 320)
$outputGroup.Size = New-Object System.Drawing.Size(800, 320)
$provisionTab.Controls.Add($outputGroup)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$outputBox.Location = New-Object System.Drawing.Point(10, 20)
$outputBox.Size = New-Object System.Drawing.Size(780, 290)
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$outputGroup.Controls.Add($outputBox)

# === MANAGEMENT TAB CONTROLS ===

# Management Connection Section
$mgmtConnectionGroup = New-Object System.Windows.Forms.GroupBox
$mgmtConnectionGroup.Text = "Connection & User Management"
$mgmtConnectionGroup.Location = New-Object System.Drawing.Point(20, 20)
$mgmtConnectionGroup.Size = New-Object System.Drawing.Size(800, 80)
$managementTab.Controls.Add($mgmtConnectionGroup)

$mgmtConnectButton = New-Object System.Windows.Forms.Button
$mgmtConnectButton.Text = "Connect to Microsoft Graph"
$mgmtConnectButton.Location = New-Object System.Drawing.Point(20, 25)
$mgmtConnectButton.Size = New-Object System.Drawing.Size(200, 25)
$mgmtConnectButton.Add_Click({ 
    if (Connect-ToMicrosoftGraph) {
        Load-UserList
    }
})
$mgmtConnectionGroup.Controls.Add($mgmtConnectButton)

$refreshUsersButton = New-Object System.Windows.Forms.Button
$refreshUsersButton.Text = "Refresh User List"
$refreshUsersButton.Location = New-Object System.Drawing.Point(240, 25)
$refreshUsersButton.Size = New-Object System.Drawing.Size(150, 25)
$refreshUsersButton.Enabled = $false
$refreshUsersButton.Add_Click({ Load-UserList })
$mgmtConnectionGroup.Controls.Add($refreshUsersButton)

$deprovisionSelectedButton = New-Object System.Windows.Forms.Button
$deprovisionSelectedButton.Text = "Deprovision Selected"
$deprovisionSelectedButton.Location = New-Object System.Drawing.Point(410, 25)
$deprovisionSelectedButton.Size = New-Object System.Drawing.Size(160, 25)
$deprovisionSelectedButton.BackColor = [System.Drawing.Color]::LightCoral
$deprovisionSelectedButton.Enabled = $false
$deprovisionSelectedButton.Add_Click({ Start-SelectedUserDeprovisioning })
$mgmtConnectionGroup.Controls.Add($deprovisionSelectedButton)

# User List Section
$userListGroup = New-Object System.Windows.Forms.GroupBox
$userListGroup.Text = "Users in Tenant"
$userListGroup.Location = New-Object System.Drawing.Point(20, 120)
$userListGroup.Size = New-Object System.Drawing.Size(800, 520)
$managementTab.Controls.Add($userListGroup)

# Create ListView for users
$userListView = New-Object System.Windows.Forms.ListView
$userListView.Location = New-Object System.Drawing.Point(10, 20)
$userListView.Size = New-Object System.Drawing.Size(780, 480)
$userListView.View = [System.Windows.Forms.View]::Details
$userListView.CheckBoxes = $true
$userListView.FullRowSelect = $true
$userListView.GridLines = $true

# Add columns
[void]$userListView.Columns.Add("Display Name", 250)
[void]$userListView.Columns.Add("User Principal Name", 300)
[void]$userListView.Columns.Add("Account Enabled", 120)
[void]$userListView.Columns.Add("Licenses", 200)

$userListGroup.Controls.Add($userListView)

# Initial log message
Write-GuiLog "M365 Copilot User Management GUI loaded successfully"
Write-GuiLog "Complete workflow: Create -> View -> Deprovision"
Write-GuiLog "Click Connect to Microsoft Graph on either tab to begin"

# Show the form
[System.Windows.Forms.Application]::Run($form)
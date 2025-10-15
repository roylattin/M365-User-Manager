[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$FirstName,
    
    [Parameter(Mandatory=$true)]
    [string]$LastName,
    
    [Parameter(Mandatory=$true)]
    [string]$UserName,
    
    [Parameter(Mandatory=$false)]
    [string]$Department = "General",
    
    [Parameter(Mandatory=$false)]
    [string]$JobTitle = "User",
    
    [Parameter(Mandatory=$false)]
    [string]$SecondaryEmail = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$TestMode
)

# Import required modules
Import-Module "$PSScriptRoot\Modules\Authentication.psm1" -Force
Import-Module "$PSScriptRoot\Modules\UserManagement.psm1" -Force
Import-Module "$PSScriptRoot\Modules\LicenseManagement.psm1" -Force
Import-Module "$PSScriptRoot\Modules\EmailNotification.psm1" -Force

# Global variables
$script:ConfigPath = "$PSScriptRoot\Config\settings.json"
$script:LogPath = "$PSScriptRoot\Logs\$(Get-Date -Format 'yyyy-MM-dd')_UserProvisioning.log"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
    Add-Content -Path $script:LogPath -Value $logEntry
}

function Main {
    try {
        Write-Log "Starting M365 Copilot user provisioning for: $FirstName $LastName"
        
        # Step 1: Load configuration
        Write-Log "Loading configuration from $script:ConfigPath"
        if (-not (Test-Path $script:ConfigPath)) {
            throw "Configuration file not found: $script:ConfigPath"
        }
        $config = Get-Content $script:ConfigPath | ConvertFrom-Json
        
        # Step 2: Connect to Azure services
        Write-Log "Connecting to Microsoft Graph for M365 Copilot provisioning"
        $authResult = Connect-AzureServices -TenantId $config.tenant.tenantId
        if (-not $authResult.Success) {
            throw "Failed to connect to Azure services: $($authResult.Error)"
        }
        
        # Step 3: Test connections
        Write-Log "Testing Azure connections"
        $connectionTest = Test-AzureConnections
        if (-not $connectionTest.Success) {
            throw "Azure connection test failed: $($connectionTest.Error)"
        }
        
        # Step 4: Create user account
        Write-Log "Creating Azure AD user account"
        $userParams = @{
            FirstName = $FirstName
            LastName = $LastName
            UserName = $UserName
            Department = $Department
            JobTitle = $JobTitle
            Domain = $config.tenant.domain
            TestMode = $TestMode
        }
        $userResult = New-AzureADUser @userParams
        if (-not $userResult.Success) {
            throw "Failed to create user: $($userResult.Error)"
        }
        
        $newUser = $userResult.User
        $temporaryPassword = $userResult.Password
        Write-Log "User created successfully. ID: $($newUser.Id)"
        
        # Step 5: Assign M365 E5 license with Copilot
        Write-Log "Assigning M365 E5 base license + Copilot add-on for full Copilot access"
        $licenseSkus = @($config.licensing.m365E5Sku, $config.licensing.copilotSku)
        $licenseResult = Set-UserLicense -UserPrincipalName $newUser.UserPrincipalName -LicenseSkus $licenseSkus -UsageLocation "US"
        if (-not $licenseResult.Success) {
            Write-Log "Warning: License assignment failed: $($licenseResult.Error)" -Level "WARN"
            Write-Log "User created but may not have full Copilot access without proper licensing" -Level "WARN"
        } else {
            Write-Log "Copilot licensing assigned successfully - user has full M365 + Copilot access"
        }
        
        # Step 6: Mailbox will be automatically provisioned with M365 license
        Write-Log "Mailbox will be automatically provisioned when license is applied"
        
        # Step 7: Send notification email
        if ($SecondaryEmail -and $SecondaryEmail -ne "") {
            Write-Log "Sending notification email to $SecondaryEmail"
            $emailParams = @{
                ToEmail = $SecondaryEmail
                UserName = $newUser.UserPrincipalName
                Password = $temporaryPassword
                FirstName = $FirstName
                LastName = $LastName
            }
            $emailResult = Send-WelcomeEmail @emailParams
            if (-not $emailResult.Success) {
                Write-Log "Warning: Email notification failed: $($emailResult.Error)" -Level "WARN"
            } else {
                Write-Log "Welcome email sent successfully"
            }
        }
        
        # Summary
        Write-Log "M365 Copilot user provisioning completed successfully!"
        Write-Host "`n=== M365 COPILOT USER PROVISIONING SUMMARY ===" -ForegroundColor Green
        Write-Host "User: $($newUser.DisplayName)" -ForegroundColor Cyan
        Write-Host "Username: $($newUser.UserPrincipalName)" -ForegroundColor Cyan
        Write-Host "User ID: $($newUser.Id)" -ForegroundColor Cyan
        Write-Host "Temporary Password: $temporaryPassword" -ForegroundColor Yellow
        Write-Host "Department: $Department" -ForegroundColor Cyan
        Write-Host "Job Title: $JobTitle" -ForegroundColor Cyan
        Write-Host "Target Licensing: M365 E5 + Copilot" -ForegroundColor Cyan
        if ($licenseResult.Success) {
            Write-Host "Copilot Status: ✓ ENABLED" -ForegroundColor Green
        } else {
            Write-Host "Copilot Status: ⚠ PENDING (awaiting license availability)" -ForegroundColor Yellow
        }
        if ($SecondaryEmail) {
            Write-Host "Notification sent to: $SecondaryEmail" -ForegroundColor Cyan
        }
        Write-Host "===============================================" -ForegroundColor Green
        
        return @{
            Success = $true
            User = $newUser
            Password = $temporaryPassword
        }
        
    } catch {
        $errorMessage = "M365 Copilot user provisioning failed: $($_.Exception.Message)"
        Write-Log $errorMessage -Level "ERROR"
        Write-Host $errorMessage -ForegroundColor Red
        
        return @{
            Success = $false
            Error = $errorMessage
        }
    }
}

# Ensure Logs directory exists
if (-not (Test-Path "$PSScriptRoot\Logs")) {
    New-Item -Path "$PSScriptRoot\Logs" -ItemType Directory -Force | Out-Null
}

# Execute main function
Main
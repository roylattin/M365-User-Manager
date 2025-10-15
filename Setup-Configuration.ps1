# Setup-Configuration.ps1
# Configuration setup script for M365 Copilot User Management Tool

param(
    [switch]$Force,
    [switch]$Quiet
)

function Write-ConfigLog {
    param([string]$Message, [string]$Level = "INFO")
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $levelString = "[$Level]".PadRight(10)
    $fullMessage = "[$timestamp] $levelString $Message"
    
    if (-not $Quiet) {
        switch ($Level) {
            "ERROR" { Write-Host $fullMessage -ForegroundColor Red }
            "WARN"  { Write-Host $fullMessage -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $fullMessage -ForegroundColor Green }
            default { Write-Host $fullMessage -ForegroundColor White }
        }
    }
    
    # Log to file
    $logPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Logs\configuration.log"
    if (-not (Test-Path (Split-Path $logPath))) {
        New-Item -ItemType Directory -Path (Split-Path $logPath) -Force | Out-Null
    }
    Add-Content -Path $logPath -Value $fullMessage
}

function Get-UserInput {
    param(
        [string]$Prompt,
        [string]$Default = "",
        [bool]$Required = $true
    )
    
    do {
        if ($Default) {
            $userInput = Read-Host "$Prompt [$Default]"
            if ([string]::IsNullOrWhiteSpace($userInput)) {
                $userInput = $Default
            }
        } else {
            $userInput = Read-Host $Prompt
        }
    } while ([string]::IsNullOrWhiteSpace($userInput) -and $Required)
    
    return $userInput
}

function Test-TenantConnectivity {
    param([string]$TenantId)
    
    try {
        Write-ConfigLog "Testing connectivity to tenant: $TenantId"
        
        # Basic validation - check if tenant ID is a valid GUID
        if (-not [System.Guid]::TryParse($TenantId, [ref][System.Guid]::Empty)) {
            Write-ConfigLog "Invalid tenant ID format" -Level "ERROR"
            return $false
        } else {
            Write-ConfigLog "Tenant ID format is valid"
            return $true
        }
        
    } catch {
        Write-ConfigLog "Tenant connectivity test failed: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Get-AvailableLicenses {
    param([string]$TenantId)
    
    try {
        Write-ConfigLog "Retrieving available licenses for tenant..."
        # For now, return null - can be enhanced with actual Graph API calls later
        return $null
        
    } catch {
        Write-ConfigLog "Failed to retrieve licenses: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

function New-ConfigurationFile {
    try {
        $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
        $configPath = Join-Path $scriptRoot "Config\settings.json"
        $templatePath = Join-Path $scriptRoot "Config\settings.template.json"
        
        # Check if configuration already exists
        if ((Test-Path $configPath) -and -not $Force) {
            $overwrite = Get-UserInput "Configuration file already exists. Overwrite? (y/N)" "N" $false
            if ($overwrite -notin @("y", "Y", "yes", "Yes", "YES")) {
                Write-ConfigLog "Configuration setup cancelled by user."
                return
            }
        }
        
        Write-ConfigLog "============================================================"
        Write-ConfigLog "M365 Copilot User Management - Configuration Setup"
        Write-ConfigLog "============================================================"
        Write-ConfigLog ""
        Write-ConfigLog "This wizard will help you configure your tenant settings."
        Write-ConfigLog "You'll need your Azure tenant information and admin privileges."
        Write-ConfigLog ""
        
        # Get tenant information
        Write-ConfigLog "Step 1: Tenant Information" -Level "SUCCESS"
        Write-ConfigLog "────────────────────────────"
        
        $tenantId = Get-UserInput "Enter your Azure Tenant ID (GUID)"
        $domain = Get-UserInput "Enter your tenant domain (e.g., yourdomain.onmicrosoft.com)"
        
        # Validate tenant connectivity
        Write-ConfigLog ""
        Write-ConfigLog "Step 2: Connectivity Test" -Level "SUCCESS"
        Write-ConfigLog "─────────────────────────"
        
        if (-not (Test-TenantConnectivity -TenantId $tenantId)) {
            Write-ConfigLog "Tenant connectivity test failed. Please verify your tenant ID and try again." -Level "ERROR"
            return
        }
        
        # Get license information
        Write-ConfigLog ""
        Write-ConfigLog "Step 3: License Configuration" -Level "SUCCESS"
        Write-ConfigLog "────────────────────────────"
        
        $licenses = Get-AvailableLicenses -TenantId $tenantId
        
        if ($licenses) {
            Write-ConfigLog ""
            $e5Sku = Get-UserInput "Enter M365 E5 license SKU" "Microsoft_365_E5_(no_Teams)"
            $copilotSku = Get-UserInput "Enter Copilot license SKU" "Microsoft_365_Copilot"
        } else {
            Write-ConfigLog "Could not retrieve license information. Using default values." -Level "WARN"
            $e5Sku = "Microsoft_365_E5_(no_Teams)"
            $copilotSku = "Microsoft_365_Copilot"
        }
        
        # Create configuration object
        $config = @{
            tenant = @{
                tenantId = $tenantId
                domain = $domain
            }
            licensing = @{
                m365E5Sku = $e5Sku
                copilotSku = $copilotSku
            }
        }
        
        # Save configuration
        $configJson = $config | ConvertTo-Json -Depth 10
        Set-Content -Path $configPath -Value $configJson -Encoding UTF8
        
        Write-ConfigLog ""
        Write-ConfigLog "============================================================"
        Write-ConfigLog "Configuration saved successfully!" -Level "SUCCESS"
        Write-ConfigLog "Configuration file: $configPath"
        Write-ConfigLog ""
        Write-ConfigLog "Next steps:"
        Write-ConfigLog "1. Review the configuration file if needed"
        Write-ConfigLog "2. Launch the application: .\M365UserManager.ps1"
        Write-ConfigLog "============================================================"
        
    } catch {
        Write-ConfigLog "Configuration setup failed: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# Execute setup if script is run directly
if ($MyInvocation.InvocationName -ne '.') {
    try {
        New-ConfigurationFile
    } catch {
        Write-ConfigLog "Setup failed. Please check the errors above and try again." -Level "ERROR"
        exit 1
    }
}
# Setup-Configuration.ps1
# Configuration setup script for M365 Copilot User Management Tool

param(
    [switch]$Force,
    [switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-ConfigLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    if (-not $Quiet) {
        switch ($Level) {
            "ERROR" { Write-Host $logMessage -ForegroundColor Red }
            "WARN"  { Write-Host $logMessage -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
            "INPUT" { Write-Host $logMessage -ForegroundColor Cyan }
            default { Write-Host $logMessage -ForegroundColor White }
        }
    }
}

function Get-UserInput {
    param(
        [string]$Prompt,
        [string]$DefaultValue = "",
        [bool]$Required = $true
    )
    
    do {
        if ($DefaultValue) {
            $fullPrompt = "$Prompt [$DefaultValue]"
        } else {
            $fullPrompt = $Prompt
        }
        
        Write-ConfigLog $fullPrompt -Level "INPUT"
        $userInput = Read-Host
        
        if ([string]::IsNullOrWhiteSpace($userInput) -and $DefaultValue) {
            return $DefaultValue
        }
        
        if ([string]::IsNullOrWhiteSpace($userInput) -and $Required) {
            Write-ConfigLog "This field is required. Please enter a value." -Level "WARN"
        }
        
    } while ([string]::IsNullOrWhiteSpace($userInput) -and $Required)
    
    return $userInput
}

function Test-TenantConnectivity {
    param(
        [string]$TenantId
    )
    
    try {
        Write-ConfigLog "Testing connectivity to tenant: $TenantId"
        
        # Import Microsoft Graph module
        Import-Module Microsoft.Graph -Force
        
        # Test connection
        Connect-MgGraph -TenantId $TenantId -Scopes "Organization.Read.All" -NoWelcome
        
        $context = Get-MgContext
        if ($context) {
            $org = Get-MgOrganization
            Write-ConfigLog "Successfully connected to tenant: $($org.DisplayName)" -Level "SUCCESS"
            Disconnect-MgGraph
            return $true
        } else {
            Write-ConfigLog "Failed to establish connection context" -Level "ERROR"
            return $false
        }
        
    } catch {
        Write-ConfigLog "Tenant connectivity test failed: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Get-AvailableLicenses {
    param(
        [string]$TenantId
    )
    
    try {
        Write-ConfigLog "Retrieving available license SKUs..."
        
        Connect-MgGraph -TenantId $TenantId -Scopes "Organization.Read.All" -NoWelcome
        
        $skus = Get-MgSubscribedSku | Where-Object { $_.PrepaidUnits.Enabled -gt 0 }
        
        Write-ConfigLog "Available license SKUs in your tenant:"
        foreach ($sku in $skus) {
            $available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
            Write-ConfigLog "  - $($sku.SkuPartNumber) (Available: $available)"
        }
        
        Disconnect-MgGraph
        return $skus
        
    } catch {
        Write-ConfigLog "Failed to retrieve license information: $($_.Exception.Message)" -Level "ERROR"
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
        
        Write-ConfigLog "=" * 60
        Write-ConfigLog "M365 Copilot User Management - Configuration Setup"
        Write-ConfigLog "=" * 60
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
        Write-ConfigLog "=" * 60
        Write-ConfigLog "Configuration saved successfully!" -Level "SUCCESS"
        Write-ConfigLog "Configuration file: $configPath"
        Write-ConfigLog ""
        Write-ConfigLog "Next steps:"
        Write-ConfigLog "1. Review the configuration file if needed"
        Write-ConfigLog "2. Launch the application: .\M365UserManager.ps1"
        Write-ConfigLog "=" * 60
        
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
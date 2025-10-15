# Simple Configuration Setup
# Creates basic configuration for M365 User Management

$configPath = ".\Config\settings.json"

# Check if configuration already exists with real values
if (Test-Path $configPath) {
    try {
        $existingConfig = Get-Content $configPath | ConvertFrom-Json
        if ($existingConfig.tenant.tenantId -ne "your-tenant-id-here" -and 
            $existingConfig.tenant.domain -ne "yourdomain.onmicrosoft.com") {
            Write-Host "Configuration file already exists with custom values." -ForegroundColor Green
            Write-Host "Skipping configuration setup to preserve your settings." -ForegroundColor Green
            exit 0
        }
    } catch {
        Write-Host "Existing configuration file is invalid. Creating new one..." -ForegroundColor Yellow
    }
}

# Create default configuration
$config = @{
    tenant = @{
        tenantId = "your-tenant-id-here"
        domain = "yourdomain.onmicrosoft.com"
    }
    licensing = @{
        m365E5Sku = "Microsoft_365_E5_(no_Teams)"
        copilotSku = "Microsoft_365_Copilot"
    }
}

# Save configuration
$configJson = $config | ConvertTo-Json -Depth 10
Set-Content -Path $configPath -Value $configJson -Encoding UTF8

Write-Host "Configuration template created at: $configPath" -ForegroundColor Green
Write-Host "Please edit the configuration file with your actual tenant details." -ForegroundColor Yellow
exit 0
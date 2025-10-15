# Quick-Start.ps1
# One-click setup script for M365 Copilot User Management Tool

param(
    [switch]$SkipModuleInstall,
    [switch]$Quiet
)

function Write-QuickLog {
    param([string]$Message, [string]$Level = "INFO")
    if (-not $Quiet) {
        switch ($Level) {
            "ERROR" { Write-Host $Message -ForegroundColor Red }
            "WARN"  { Write-Host $Message -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $Message -ForegroundColor Green }
            "TITLE" { Write-Host $Message -ForegroundColor Cyan }
            default { Write-Host $Message -ForegroundColor White }
        }
    }
}

try {
    Write-QuickLog "🚀 M365 Copilot User Management - Quick Start" -Level "TITLE"
    Write-QuickLog "=" * 60 -Level "TITLE"
    Write-QuickLog ""
    
    # Step 1: Environment Setup
    if (-not $SkipModuleInstall) {
        Write-QuickLog "Step 1: Setting up environment..." -Level "SUCCESS"
        & ".\Setup-Environment.ps1" -Quiet:$Quiet
        if ($LASTEXITCODE -ne 0) {
            throw "Environment setup failed"
        }
        Write-QuickLog "✓ Environment setup completed" -Level "SUCCESS"
        Write-QuickLog ""
    }
    
    # Step 2: Configuration
    Write-QuickLog "Step 2: Configuring tenant settings..." -Level "SUCCESS"
    & ".\Setup-Configuration.ps1" -Quiet:$Quiet
    if ($LASTEXITCODE -ne 0) {
        throw "Configuration setup failed"
    }
    Write-QuickLog "✓ Configuration completed" -Level "SUCCESS"
    Write-QuickLog ""
    
    # Step 3: Launch Application
    Write-QuickLog "Step 3: Launching application..." -Level "SUCCESS"
    Write-QuickLog ""
    Write-QuickLog "Starting M365 Copilot User Management GUI..." -Level "SUCCESS"
    
    & ".\M365UserManager.ps1"
    
} catch {
    Write-QuickLog "Quick start failed: $($_.Exception.Message)" -Level "ERROR"
    Write-QuickLog ""
    Write-QuickLog "Manual setup options:" -Level "WARN"
    Write-QuickLog "1. Run .\Setup-Environment.ps1" -Level "WARN"
    Write-QuickLog "2. Run .\Setup-Configuration.ps1" -Level "WARN"
    Write-QuickLog "3. Launch .\M365CopilotProvisioningGUI_Complete.ps1" -Level "WARN"
    exit 1
}
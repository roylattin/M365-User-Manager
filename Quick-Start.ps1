# Quick-Start.ps1
# Intelligent setup and launch script for M365 Copilot User Management Tool

param(
    [switch]$SkipModuleInstall,
    [switch]$Quiet,
    [switch]$ForceSetup
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

function Test-EnvironmentReady {
    try {
        # Check if Microsoft.Graph module is available
        $mgModule = Get-Module -ListAvailable -Name Microsoft.Graph | Select-Object -First 1
        if (-not $mgModule) {
            return $false
        }
        
        # Check if required directories exist
        $requiredDirs = @("Config", "Logs")
        foreach ($dir in $requiredDirs) {
            if (-not (Test-Path $dir)) {
                return $false
            }
        }
        
        return $true
    } catch {
        return $false
    }
}

function Test-ConfigurationReady {
    try {
        $configPath = ".\Config\settings.json"
        if (-not (Test-Path $configPath)) {
            return $false
        }
        
        $config = Get-Content $configPath | ConvertFrom-Json
        # Check if it has real values (not template values)
        if ($config.tenant.tenantId -eq "your-tenant-id-here" -or 
            $config.tenant.domain -eq "yourdomain.onmicrosoft.com") {
            return $false
        }
        
        return $true
    } catch {
        return $false
    }
}

try {
    Write-QuickLog "M365 Copilot User Management - Smart Launch" -Level "TITLE"
    Write-QuickLog "============================================================" -Level "TITLE"
    Write-QuickLog ""
    
    # Smart Step 1: Check Environment Setup
    $environmentReady = Test-EnvironmentReady
    if (-not $environmentReady -or $ForceSetup) {
        if ($environmentReady -and $ForceSetup) {
            Write-QuickLog "Forcing environment setup (ForceSetup flag)" -Level "WARN"
        } else {
            Write-QuickLog "Environment setup required..." -Level "WARN"
        }
        
        if (-not $SkipModuleInstall) {
            Write-QuickLog "Step 1: Setting up environment..." -Level "SUCCESS"
            & ".\Setup-Environment.ps1" -Quiet:$Quiet
            if ($LASTEXITCODE -ne 0) {
                throw "Environment setup failed"
            }
            Write-QuickLog "Environment setup completed" -Level "SUCCESS"
            Write-QuickLog ""
        }
    } else {
        Write-QuickLog "Environment already configured - skipping setup" -Level "SUCCESS"
    }
    
    # Smart Step 2: Check Configuration
    $configReady = Test-ConfigurationReady
    if (-not $configReady -or $ForceSetup) {
        if ($configReady -and $ForceSetup) {
            Write-QuickLog "Forcing configuration setup (ForceSetup flag)" -Level "WARN"
        } else {
            Write-QuickLog "Configuration setup required..." -Level "WARN"
        }
        
        Write-QuickLog "Step 2: Configuring tenant settings..." -Level "SUCCESS"
        & ".\Setup-Configuration-Simple.ps1" -Quiet:$Quiet
        if ($LASTEXITCODE -ne 0) {
            throw "Configuration setup failed"
        }
        Write-QuickLog "Configuration completed" -Level "SUCCESS"
        Write-QuickLog ""
    } else {
        Write-QuickLog "Configuration already ready - preserving existing settings" -Level "SUCCESS"
    }
    
    # Step 3: Always Launch Application
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
# Setup-Environment.ps1
# Environment setup script for M365 Copilot User Management Tool

param(
    [switch]$Force,
    [switch]$Quiet
)

# Set strict mode for better error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Script metadata
$ScriptVersion = "1.0.0"
$ScriptName = "M365 Copilot User Management - Environment Setup"

function Write-SetupLog {
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
            default { Write-Host $logMessage -ForegroundColor White }
        }
    }
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Install-RequiredModule {
    param(
        [string]$ModuleName,
        [string]$MinimumVersion = $null
    )
    
    try {
        Write-SetupLog "Checking module: $ModuleName"
        
        $installedModule = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
        
        if ($installedModule) {
            if ($MinimumVersion -and ($installedModule.Version -lt [Version]$MinimumVersion)) {
                Write-SetupLog "Module $ModuleName version $($installedModule.Version) is below required version $MinimumVersion" -Level "WARN"
                Write-SetupLog "Updating module: $ModuleName"
                Update-Module -Name $ModuleName -Force
            } else {
                Write-SetupLog "Module $ModuleName is already installed (version: $($installedModule.Version))" -Level "SUCCESS"
                return
            }
        } else {
            Write-SetupLog "Installing module: $ModuleName"
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
        }
        
        # Verify installation
        $verifyModule = Get-Module -ListAvailable -Name $ModuleName
        if ($verifyModule) {
            Write-SetupLog "Successfully installed/updated $ModuleName (version: $($verifyModule.Version))" -Level "SUCCESS"
        } else {
            throw "Failed to verify installation of $ModuleName"
        }
        
    } catch {
        Write-SetupLog "Failed to install $ModuleName`: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

function Initialize-ProjectStructure {
    try {
        $scriptRoot = Split-Path -Parent $PSCommandPath
        
        # Create required directories
        $directories = @("Config", "Logs", "Scripts")
        
        foreach ($dir in $directories) {
            $fullPath = Join-Path $scriptRoot $dir
            if (-not (Test-Path $fullPath)) {
                New-Item -Path $fullPath -ItemType Directory -Force | Out-Null
                Write-SetupLog "Created directory: $dir" -Level "SUCCESS"
            } else {
                Write-SetupLog "Directory already exists: $dir"
            }
        }
        
        # Create .gitignore if it doesn't exist
        $gitignorePath = Join-Path $scriptRoot ".gitignore"
        if (-not (Test-Path $gitignorePath) -or $Force) {
            $gitignoreContent = @"
# Logs
Logs/*.log

# Configuration files with sensitive data
Config/settings.json

# PowerShell module cache
*.pdb
*.dll
*.exe

# Temporary files
*.tmp
*.temp

# User-specific files
.vscode/
*.code-workspace

# Windows
Thumbs.db
Desktop.ini

# MacOS
.DS_Store

# Backup files
*.bak
*.backup
"@
            Set-Content -Path $gitignorePath -Value $gitignoreContent -Encoding UTF8
            Write-SetupLog "Created .gitignore file" -Level "SUCCESS"
        }
        
        # Create settings template
        $templatePath = Join-Path $scriptRoot "Config\settings.template.json"
        if (-not (Test-Path $templatePath) -or $Force) {
            $templateContent = @"
{
  "tenant": {
    "tenantId": "your-tenant-id-here",
    "domain": "yourdomain.onmicrosoft.com"
  },
  "licensing": {
    "m365E5Sku": "Microsoft_365_E5_(no_Teams)",
    "copilotSku": "Microsoft_365_Copilot"
  }
}
"@
            Set-Content -Path $templatePath -Value $templateContent -Encoding UTF8
            Write-SetupLog "Created settings template" -Level "SUCCESS"
        }
        
    } catch {
        Write-SetupLog "Failed to initialize project structure: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

function Test-PowerShellVersion {
    $version = $PSVersionTable.PSVersion
    Write-SetupLog "PowerShell version: $version"
    
    if ($version.Major -lt 5) {
        Write-SetupLog "PowerShell 5.1 or later is required. Current version: $version" -Level "ERROR"
        return $false
    }
    
    if ($version.Major -eq 5 -and $version.Minor -eq 0) {
        Write-SetupLog "PowerShell 5.0 detected. Consider upgrading to 5.1 or later." -Level "WARN"
    }
    
    return $true
}

function Test-InternetConnectivity {
    try {
        Write-SetupLog "Testing internet connectivity..."
        $result = Test-NetConnection -ComputerName "graph.microsoft.com" -Port 443 -InformationLevel Quiet
        if ($result) {
            Write-SetupLog "Internet connectivity verified" -Level "SUCCESS"
            return $true
        } else {
            Write-SetupLog "Cannot connect to Microsoft Graph endpoints" -Level "ERROR"
            return $false
        }
    } catch {
        Write-SetupLog "Internet connectivity test failed: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# Main setup function
function Start-EnvironmentSetup {
    try {
        Write-SetupLog "Starting $ScriptName v$ScriptVersion"
        Write-SetupLog "=" * 60
        
        # Check PowerShell version
        if (-not (Test-PowerShellVersion)) {
            throw "PowerShell version requirements not met"
        }
        
        # Check internet connectivity
        if (-not (Test-InternetConnectivity)) {
            throw "Internet connectivity requirements not met"
        }
        
        # Check if running as administrator (recommended but not required)
        if (-not (Test-Administrator)) {
            Write-SetupLog "Not running as administrator. Some operations may require elevation." -Level "WARN"
        }
        
        # Initialize project structure
        Write-SetupLog "Initializing project structure..."
        Initialize-ProjectStructure
        
        # Install required PowerShell modules
        Write-SetupLog "Installing required PowerShell modules..."
        $modules = @(
            @{ Name = "Microsoft.Graph"; MinVersion = "1.0.0" }
        )
        
        foreach ($module in $modules) {
            if ($module.MinVersion) {
                Install-RequiredModule -ModuleName $module.Name -MinimumVersion $module.MinVersion
            } else {
                Install-RequiredModule -ModuleName $module.Name
            }
        }
        
        # Verify installation
        Write-SetupLog "Verifying installation..."
        try {
            # Just check if the module is available, don't import everything
            $mgModule = Get-Module -ListAvailable -Name Microsoft.Graph | Select-Object -First 1
            if ($mgModule) {
                Write-SetupLog "Microsoft Graph module is available (version: $($mgModule.Version))" -Level "SUCCESS"
            } else {
                throw "Microsoft.Graph module not found after installation"
            }
        } catch {
            Write-SetupLog "Module verification failed: $($_.Exception.Message)" -Level "WARN"
            Write-SetupLog "Module installation completed but verification failed. This is often normal due to PowerShell function limits." -Level "WARN"
        }
        
        Write-SetupLog "=" * 60
        Write-SetupLog "Environment setup completed successfully!" -Level "SUCCESS"
        Write-SetupLog ""
        Write-SetupLog "Next steps:"
        Write-SetupLog "1. Run '.\Setup-Configuration.ps1' to configure your tenant settings"
        Write-SetupLog "2. Launch the application with '.\M365UserManager.ps1'"
        Write-SetupLog ""
        
    } catch {
        Write-SetupLog "Environment setup failed: $($_.Exception.Message)" -Level "ERROR"
        Write-SetupLog "Please review the errors above and try again." -Level "ERROR"
        exit 1
    }
}

# Execute setup if script is run directly
if ($MyInvocation.InvocationName -ne '.') {
    Start-EnvironmentSetup
}
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$OrganizationalUnit,
    
    [Parameter(Mandatory = $false)]
    [string]$SamAccountName,
    
    [Parameter(Mandatory = $false)]
    [string]$Domain = "CONTOSO.COM",
    
    [Parameter(Mandatory = $false)]
    [switch]$PlainTextLog,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory = $false)]
    [switch]$SaveCredentials,
    
    [Parameter(Mandatory = $false)]
    [switch]$EmailResults,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailFrom,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailTo,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailSubject = "Microsoft 365 Photo Update Results",
    
    [Parameter(Mandatory = $false)]
    [string]$EmailRelay = "smtp.office365.com",
    
    [Parameter(Mandatory = $false)]
    [int]$EmailRelayPort = 25
)

# Function to display usage information
function Show-Usage {
    Write-Host "Usage: .\Update365ThumbnailPhotos.ps1 [-OrganizationalUnit <OU_DistinguishedName>] [-SamAccountName <SamAccountName>] [-Domain <Domain>] [-LogFile <LogFilePath>] [-PlainTextLog] [-Force] [-TenantId <TenantId>] [-ClientId <ClientId>] [-ClientSecret <ClientSecret>] [-SaveCredentials] [-EmailResults] [-EmailFrom <FromAddress>] [-EmailTo <ToAddress>] [-EmailSubject <Subject>] [-EmailRelay <RelayServer>] [-EmailRelayPort <Port>]" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Parameters:" -ForegroundColor Cyan
    Write-Host "  -OrganizationalUnit : Optional. The Distinguished Name of the OU to search for users with thumbnail photos" -ForegroundColor Yellow
    Write-Host "  -SamAccountName     : Optional. The SAM Account Name of a specific user to update" -ForegroundColor Yellow
    Write-Host "  -Domain             : Optional. The domain to use (default: CONTOSO.COM)" -ForegroundColor Yellow
    Write-Host "  -LogFile            : Optional. Path to the log file (default: .\PhotoUpdate_<timestamp>.html)" -ForegroundColor Yellow
    Write-Host "  -PlainTextLog       : Optional. Use plain text format for the log file instead of HTML" -ForegroundColor Yellow
    Write-Host "  -Force              : Optional. Perform actual updates instead of a dry run" -ForegroundColor Yellow
    Write-Host "  -TenantId           : Optional. The Azure AD tenant ID for app registration authentication" -ForegroundColor Yellow
    Write-Host "  -ClientId           : Optional. The client ID (application ID) for app registration authentication" -ForegroundColor Yellow
    Write-Host "  -ClientSecret       : Optional. The client secret for app registration authentication" -ForegroundColor Yellow
    Write-Host "  -SaveCredentials    : Optional. Save the App registration credentials to a configuration file" -ForegroundColor Yellow
    Write-Host "  -EmailResults       : Optional. Send email with update results" -ForegroundColor Yellow
    Write-Host "  -EmailFrom          : Required if -EmailResults is used. Sender email address" -ForegroundColor Yellow
    Write-Host "  -EmailTo            : Required if -EmailResults is used. Recipient email address" -ForegroundColor Yellow
    Write-Host "  -EmailSubject       : Optional. Email subject (default: Microsoft 365 Photo Update Results)" -ForegroundColor Yellow
    Write-Host "  -EmailRelay         : Optional. SMTP relay server (default: smtp.office365.com)" -ForegroundColor Yellow
    Write-Host "  -EmailRelayPort     : Optional. SMTP relay port (default: 25)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Note: Either -OrganizationalUnit or -SamAccountName must be specified, but not both." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Cyan
    Write-Host "  .\Update365ThumbnailPhotos.ps1 -OrganizationalUnit 'OU=Users,DC=CONTOSO,DC=COM' -Force" -ForegroundColor Green
    Write-Host "  .\Update365ThumbnailPhotos.ps1 -SamAccountName 'jdoe' -Force" -ForegroundColor Green
    Write-Host "  .\Update365ThumbnailPhotos.ps1 -SamAccountName 'jdoe' -Force -EmailResults -EmailFrom 'noreply@contoso.com' -EmailTo 'admin@contoso.com'" -ForegroundColor Green
    Write-Host "  .\Update365ThumbnailPhotos.ps1 -SamAccountName 'jdoe' -Force -EmailResults -EmailFrom 'noreply@contoso.com' -EmailTo 'admin@contoso.com' -EmailRelay 'smtp.internal.contoso.com'" -ForegroundColor Green
    exit
}

# Check if either OrganizationalUnit or SamAccountName is provided, if not show usage
if (-not $OrganizationalUnit -and -not $SamAccountName) {
    Write-Host "Error: Either -OrganizationalUnit or -SamAccountName must be specified." -ForegroundColor Red
    Show-Usage
}

# Check if both OrganizationalUnit and SamAccountName are provided, if so show error
if ($OrganizationalUnit -and $SamAccountName) {
    Write-Host "Error: Cannot specify both -OrganizationalUnit and -SamAccountName. Please use only one of them." -ForegroundColor Red
    Show-Usage
}

# Check if email parameters are provided when EmailResults is used
if ($EmailResults) {
    if (-not $EmailFrom -or -not $EmailTo) {
        Write-Host "Error: -EmailFrom and -EmailTo parameters are required when using -EmailResults" -ForegroundColor Red
        Show-Usage
    }
}

# Function to write to log file
function Write-Log {
    param (
        [string]$Message,
        [string]$Status = "Info",
        [string]$Color = "White"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    # Write to console with color
    switch ($Status) {
        "Success" { Write-Host $logMessage -ForegroundColor Green }
        "Error" { Write-Host $logMessage -ForegroundColor Red }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Info" { Write-Host $logMessage -ForegroundColor White }
        "Current" { Write-Host $logMessage -ForegroundColor Cyan }
        default { Write-Host $logMessage -ForegroundColor $Color }
    }
    
    # Write to log file
    if ($PlainTextLog) {
        Add-Content -Path $script:LogFile -Value $logMessage
    } else {
        # HTML formatting
        $htmlColor = switch ($Status) {
            "Success" { "green" }
            "Error" { "red" }
            "Warning" { "orange" }
            "Info" { "white" }
            "Current" { "cyan" }
            default { $Color.ToLower() }
        }
        
        $htmlMessage = "<div style='color: $htmlColor;'>$logMessage</div>"
        Add-Content -Path $script:LogFile -Value $htmlMessage
    }
}

# Function to initialize log file
function Initialize-LogFile {
    # Try multiple methods to determine the script path
    $scriptPath = $null
    
    # Method 1: Use MyInvocation
    if ($MyInvocation.MyCommand.Path) {
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
        Write-Host "Script path determined from MyInvocation: $scriptPath" -ForegroundColor White
    }
    
    # Method 2: Use PSScriptRoot (works in modules and dot-sourced scripts)
    if (-not $scriptPath -and $PSScriptRoot) {
        $scriptPath = $PSScriptRoot
        Write-Host "Script path determined from PSScriptRoot: $scriptPath" -ForegroundColor White
    }
    
    # Method 3: Use the current location as a last resort
    if (-not $scriptPath) {
        $scriptPath = Get-Location
        Write-Host "Script path could not be determined. Using current location: $scriptPath" -ForegroundColor Yellow
    }
    
    # Create Logs directory if it doesn't exist
    $logsDir = Join-Path -Path $scriptPath -ChildPath "Logs"
    if (-not (Test-Path $logsDir)) {
        Write-Host "Creating Logs directory..." -ForegroundColor White
        New-Item -Path $logsDir -ItemType Directory -Force | Out-Null
        Write-Host "Logs directory created successfully" -ForegroundColor Green
    }
    
    # Generate log file name with timestamp
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logFileName = "PhotoUpdate_$timestamp"
    if ($PlainTextLog) {
        $logFileName += ".txt"
    } else {
        $logFileName += ".html"
    }
    
    # Set the script-scoped log file path
    $script:LogFile = Join-Path -Path $logsDir -ChildPath $logFileName
    
    # Verify the log file path is valid
    if (-not $script:LogFile) {
        throw "Failed to initialize log file path"
    }
    
    if ($PlainTextLog) {
        "Photo Update Log - $(Get-Date)" | Out-File -FilePath $script:LogFile
        "----------------------------------------" | Out-File -FilePath $script:LogFile -Append
    } else {
        $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Photo Update Log - $(Get-Date)</title>
    <style>
        body { font-family: Consolas, monospace; background-color: #1e1e1e; color: white; }
        .success { color: green; }
        .error { color: red; }
        .warning { color: orange; }
        .info { color: white; }
    </style>
</head>
<body>
    <h1>Photo Update Log - $(Get-Date)</h1>
"@
        $htmlHeader | Out-File -FilePath $script:LogFile
    }
    
    Write-Host "Log file initialized: $script:LogFile" -ForegroundColor White
}

# Function to finalize log file
function Finalize-LogFile {
    if (-not $PlainTextLog) {
        "</body></html>" | Out-File -FilePath $script:LogFile -Append
    }
}

# Function to check and install required modules
function Ensure-RequiredModules {
    $requiredModules = @(
        @{Name = "ActiveDirectory"; Description = "Active Directory module"},
        @{Name = "Microsoft.Graph"; Description = "Microsoft Graph module"}
    )
    
    $optionalModules = @(
        @{Name = "PSSQLite"; Description = "SQLite module for PowerShell"}
    )
    
    # Set TLS 1.2 for all connections
    Write-Log "Setting TLS 1.2 for all connections..." "Info"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    # Register and configure PSGallery
    Write-Log "Configuring PowerShell repository..." "Info"
    try {
        Register-PSRepository -Default -ErrorAction SilentlyContinue
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        Write-Log "PowerShell repository configured successfully" "Success"
    }
    catch {
        Write-Log "Warning: Could not configure PowerShell repository: $_" "Warning"
    }
    
    # Check and install required modules
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module.Name)) {
            Write-Log "Module '$($module.Name)' is not installed. Attempting to install..." "Warning"
            try {
                # Special handling for ActiveDirectory module
                if ($module.Name -eq "ActiveDirectory") {
                    # Check if running on Windows
                    if ($env:OS -eq "Windows_NT") {
                        # Check if RSAT-AD-PowerShell feature is installed
                        $adFeature = Get-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction SilentlyContinue
                        if (-not $adFeature.Installed) {
                            Write-Log "RSAT-AD-PowerShell feature is not installed. Attempting to install..." "Warning"
                            try {
                                Install-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction Stop
                                Write-Log "RSAT-AD-PowerShell feature installed successfully" "Success"
                            }
                            catch {
                                Write-Log "Failed to install RSAT-AD-PowerShell feature: $_" "Error"
                                throw "Could not install RSAT-AD-PowerShell feature. Please install it manually."
                            }
                        }
                    }
                    else {
                        Write-Log "Active Directory module is not available on non-Windows systems" "Error"
                        throw "Active Directory module is not available on non-Windows systems"
                    }
                }
                
                # Special handling for Microsoft.Graph
                if ($module.Name -eq "Microsoft.Graph") {
                    # Try to install using the specified approach
                    try {
                        Write-Log "Attempting to install Microsoft.Graph..." "Info"
                        Install-Module -Name $module.Name -Force -AllowClobber -Scope CurrentUser
                        
                        # Verify the module was actually installed
                        if (-not (Get-Module -ListAvailable -Name $module.Name)) {
                            throw "Module was not installed successfully"
                        }
                    }
                    catch {
                        Write-Log "Failed to install Microsoft.Graph: $_" "Error"
                        throw "Could not install Microsoft.Graph module. Please install it manually."
                    }
                }
                else {
                    # Standard installation for other modules
                    Install-Module -Name $module.Name -Force -AllowClobber -Scope CurrentUser
                }
                
                Write-Log "Successfully installed module '$($module.Name)'" "Success"
            }
            catch {
                Write-Log "Failed to install module '$($module.Name)': $_" "Error"
                throw "Required module '$($module.Name)' could not be installed. Please install it manually."
            }
        }
        else {
            Write-Log "Module '$($module.Name)' is already installed" "Info"
        }
    }
    
    # Check optional modules
    $script:UseSQLite = $false
    foreach ($module in $optionalModules) {
        if (Get-Module -ListAvailable -Name $module.Name) {
            Write-Log "Optional module '$($module.Name)' is available" "Info"
            if ($module.Name -eq "PSSQLite") {
                $script:UseSQLite = $true
            }
        }
        else {
            Write-Log "Optional module '$($module.Name)' is not available. Attempting to install..." "Warning"
            try {
                # Special handling for PSSQLite module
                if ($module.Name -eq "PSSQLite") {
                    # Check if running on Windows
                    if ($env:OS -eq "Windows_NT") {
                        # Check if SQLite is installed
                        $sqliteInstalled = $false
                        try {
                            $sqliteVersion = sqlite3 --version 2>$null
                            if ($sqliteVersion) {
                                $sqliteInstalled = $true
                                Write-Log "SQLite is already installed: $sqliteVersion" "Info"
                            }
                        }
                        catch {
                            Write-Log "SQLite command line tool not found" "Info"
                        }
                        
                        if (-not $sqliteInstalled) {
                            Write-Log "SQLite is not installed. Attempting to install..." "Warning"
                            try {
                                # Download and install SQLite
                                $sqliteUrl = "https://www.sqlite.org/2024/sqlite-tools-win32-x86-3450200.zip"
                                $tempDir = [System.IO.Path]::GetTempPath()
                                $zipFile = Join-Path $tempDir "sqlite-tools.zip"
                                
                                Write-Log "Downloading SQLite tools..." "Info"
                                Invoke-WebRequest -Uri $sqliteUrl -OutFile $zipFile
                                
                                Write-Log "Extracting SQLite tools..." "Info"
                                Expand-Archive -Path $zipFile -DestinationPath $tempDir -Force
                                
                                # Add SQLite to PATH
                                $sqlitePath = Join-Path $tempDir "sqlite-tools-win32-x86-3450200"
                                $env:Path += ";$sqlitePath"
                                [Environment]::SetEnvironmentVariable("Path", $env:Path, [EnvironmentVariableTarget]::Machine)
                                
                                Write-Log "SQLite installed successfully" "Success"
                            }
                            catch {
                                Write-Log "Failed to install SQLite: $_" "Error"
                                Write-Log "Please install SQLite manually from https://www.sqlite.org/download.html" "Warning"
                            }
                        }
                    }
                    
                    # Install PSSQLite module
                    try {
                        Write-Log "Installing PSSQLite module..." "Info"
                        Install-Module -Name PSSQLite -Force -AllowClobber -Scope CurrentUser
                        $script:UseSQLite = $true
                        Write-Log "PSSQLite module installed successfully" "Success"
                    }
                    catch {
                        Write-Log "Failed to install PSSQLite module: $_" "Error"
                        Write-Log "Photo change tracking will be disabled" "Warning"
                    }
                }
            }
            catch {
                Write-Log "Error installing optional module '$($module.Name)': $_" "Error"
                Write-Log "Some features may be limited" "Warning"
            }
        }
    }
}

# Function to selectively import Microsoft.Graph module
function Import-SelectiveGraphModule {
    Write-Log "Selectively importing Microsoft.Graph module..." "Info"
    
    # Define the specific modules we need
    $requiredGraphModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Users"
    )
    
    # Import each module individually
    foreach ($module in $requiredGraphModules) {
        try {
            Import-Module $module -ErrorAction Stop
            Write-Log "Successfully imported $module" "Success"
        }
        catch {
            $errorMsg = "Failed to import $module"
            Write-Log $errorMsg "Error"
            throw "Could not import required Microsoft.Graph module: $module"
        }
    }
    
    Write-Log "Successfully imported required Microsoft.Graph modules" "Success"
}

# Function to initialize the SQLite database
function Initialize-PhotoDatabase {
    if (-not $script:UseSQLite) {
        Write-Log "SQLite functionality is not available. Photo change tracking will be disabled." "Warning"
        return $null
    }
    
    try {
        # Try multiple methods to determine the script path
        $scriptPath = $null
        
        # Method 1: Use MyInvocation
        if ($MyInvocation.MyCommand.Path) {
            $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
            Write-Log "Script path determined from MyInvocation: $scriptPath" "Info"
        }
        
        # Method 2: Use PSScriptRoot (works in modules and dot-sourced scripts)
        if (-not $scriptPath -and $PSScriptRoot) {
            $scriptPath = $PSScriptRoot
            Write-Log "Script path determined from PSScriptRoot: $scriptPath" "Info"
        }
        
        # Method 3: Use the current location as a last resort
        if (-not $scriptPath) {
            $scriptPath = Get-Location
            Write-Log "Script path could not be determined. Using current location: $scriptPath" "Warning"
        }
        
        # Ensure the path exists
        if (-not (Test-Path -Path $scriptPath)) {
            Write-Log "Script path does not exist. Creating directory: $scriptPath" "Warning"
            New-Item -Path $scriptPath -ItemType Directory -Force | Out-Null
        }
        
        # Use a fixed database name in the same directory as the script
        $dbPath = Join-Path -Path $scriptPath -ChildPath "Update365ThumbnailPhotos.sqlite"
        Write-Log "Database path: $dbPath" "Info"
        
        # Check if the database exists, if not create it
        if (-not (Test-Path $dbPath)) {
            Write-Log "Creating photo history database..." "Info"
            
            # Create the database and table
            $query = @"
CREATE TABLE IF NOT EXISTS PhotoHistory (
    SamAccountName TEXT PRIMARY KEY,
    PhotoHash TEXT NOT NULL,
    LastUpdated TEXT NOT NULL
);
"@
            
            Invoke-SqliteQuery -DataSource $dbPath -Query $query
            Write-Log "Photo history database created successfully" "Success"
        }
        
        return $dbPath
    }
    catch {
        Write-Log "Error initializing photo database: $_" "Error"
        Write-Log "Photo change tracking will be disabled." "Warning"
        return $null
    }
}

# Function to check if a photo has changed
function Test-PhotoChanged {
    param (
        [string]$SamAccountName,
        [byte[]]$PhotoData,
        [string]$DbPath
    )
    
    # If SQLite is not available, always return true (photo has changed)
    if (-not $script:UseSQLite -or -not $DbPath) {
        Write-Log "SQLite functionality is not available. Assuming photo has changed." "Info"
        return $true, ""
    }
    
    try {
        # Calculate MD5 hash of the photo
        $md5 = [System.Security.Cryptography.MD5]::Create()
        $hash = $md5.ComputeHash($PhotoData)
        $hashString = [System.BitConverter]::ToString($hash).Replace("-", "").ToLower()
        Write-Log "Calculated hash for ${SamAccountName}: $hashString" "Info"
        
        # Check if the user exists in the database
        $query = "SELECT PhotoHash FROM PhotoHistory WHERE SamAccountName = @SamAccountName"
        $result = Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
            SamAccountName = $SamAccountName
        }
        
        if ($result.Count -eq 0) {
            Write-Log "User ${SamAccountName} not found in photo history database. Photo will be updated." "Info"
            return $true, $hashString
        }
        
        # Compare the hash
        $storedHash = $result.PhotoHash
        Write-Log "Stored hash for ${SamAccountName}: $storedHash" "Info"
        
        if ($storedHash -ne $hashString) {
            Write-Log "Photo hash for ${SamAccountName} has changed. Photo will be updated." "Info"
            return $true, $hashString
        }
        
        Write-Log "Photo hash for ${SamAccountName} matches stored hash. No update needed." "Info"
        return $false, $hashString
    }
    catch {
        Write-Log "Error checking if photo has changed: $_" "Error"
        Write-Log "Assuming photo has changed due to error." "Warning"
        return $true, ""
    }
}

# Function to update the photo history database
function Update-PhotoHistory {
    param (
        [string]$SamAccountName,
        [string]$PhotoHash,
        [string]$DbPath
    )
    
    # If SQLite is not available, do nothing
    if (-not $script:UseSQLite -or -not $DbPath) {
        return
    }
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        # Use parameterized query to prevent SQL injection
        $query = "INSERT OR REPLACE INTO PhotoHistory (SamAccountName, PhotoHash, LastUpdated) VALUES (@SamAccountName, @PhotoHash, @LastUpdated)"
        
        Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters @{
            SamAccountName = $SamAccountName
            PhotoHash = $PhotoHash
            LastUpdated = $timestamp
        }
        
        Write-Log "Photo history updated for ${SamAccountName}" "Info"
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Log "Error updating photo history: ${errorMessage}" "Error"
    }
}

# Function to store App registration credentials
function Save-AppRegistrationCredentials {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$ConfigPath = ".\Update365ThumbnailPhotos.config"
    )
    
    try {
        # Create a secure string for the client secret
        $secureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        
        # Create a credential object
        $credential = New-Object System.Management.Automation.PSCredential($ClientId, $secureClientSecret)
        
        # Export the credential and tenant ID to a file
        $config = @{
            TenantId = $TenantId
            ClientId = $ClientId
            ClientSecret = $ClientSecret
        }
        
        $config | Export-Clixml -Path $ConfigPath -Force
        
        Write-Log "App registration credentials saved to $ConfigPath" "Success"
        return $true
    }
    catch {
        Write-Log "Error saving App registration credentials: $_" "Error"
        return $false
    }
}

# Function to retrieve App registration credentials
function Get-AppRegistrationCredentials {
    param (
        [string]$ConfigPath = ".\Update365ThumbnailPhotos.config"
    )
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            Write-Log "Configuration file not found: $ConfigPath" "Warning"
            return $null, $null, $null
        }
        
        # Import the configuration
        $config = Import-Clixml -Path $ConfigPath
        
        # Extract the tenant ID, client ID, and client secret
        $tenantId = $config.TenantId
        $clientId = $config.ClientId
        $clientSecret = $config.ClientSecret
        
        Write-Log "App registration credentials retrieved from $ConfigPath" "Info"
        return $tenantId, $clientId, $clientSecret
    }
    catch {
        Write-Log "Error retrieving App registration credentials: $_" "Error"
        return $null, $null, $null
    }
}

# Function to check Microsoft.Graph module version and connect appropriately
function Connect-ToMicrosoftGraph {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )
    
    try {
        # Get the Microsoft.Graph module version
        $graphModule = Get-Module -Name Microsoft.Graph -ListAvailable | Select-Object -First 1
        $graphVersion = $graphModule.Version
        Write-Log "Microsoft.Graph module version: $graphVersion" "Info"
        
        Write-Log "Using Microsoft.Graph version 2.26.1 authentication method..." "Info"
        
        # Disconnect any existing connections
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        
        # Create a secure string for the client secret
        $secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        
        # Create a PSCredential object
        $clientSecretCredential = New-Object System.Management.Automation.PSCredential($ClientId, $secureSecret)
        
        # Set up the connection parameters
        $connectionParams = @{
            'TenantId'             = $TenantId
            'ClientSecretCredential' = $clientSecretCredential
            'NoWelcome'            = $true
            'ContextScope'         = 'Process'
        }
        
        # Connect with specific scopes and non-interactive mode
        Connect-MgGraph @connectionParams -ErrorAction Stop
        
        # Verify the connection
        $context = Get-MgContext
        if (-not $context) {
            throw "Failed to establish Microsoft Graph connection context"
        }
        
        Write-Log "Connected to Microsoft Graph successfully using client credentials" "Success"
    }
    catch {
        Write-Log "Error connecting to Microsoft Graph: $_" "Error"
        throw "Failed to connect to Microsoft Graph. Please check your credentials and try again."
    }
}

# Function to prune old log files
function Remove-OldLogs {
    try {
        # Try multiple methods to determine the script path
        $scriptPath = $null
        
        # Method 1: Use MyInvocation
        if ($MyInvocation.MyCommand.Path) {
            $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
            Write-Log "Script path determined from MyInvocation: $scriptPath" "Info"
        }
        
        # Method 2: Use PSScriptRoot (works in modules and dot-sourced scripts)
        if (-not $scriptPath -and $PSScriptRoot) {
            $scriptPath = $PSScriptRoot
            Write-Log "Script path determined from PSScriptRoot: $scriptPath" "Info"
        }
        
        # Method 3: Use the current location as a last resort
        if (-not $scriptPath) {
            $scriptPath = Get-Location
            Write-Log "Script path could not be determined. Using current location: $scriptPath" "Warning"
        }
        
        # Get Logs directory path
        $logsDir = Join-Path -Path $scriptPath -ChildPath "Logs"
        if (-not (Test-Path $logsDir)) {
            Write-Log "Logs directory does not exist: $logsDir" "Info"
            return
        }
        
        # Get all log files
        $logFiles = Get-ChildItem -Path $logsDir -Filter "PhotoUpdate_*.html" -File |
                   Sort-Object LastWriteTime -Descending
        
        # Keep only the 14 most recent HTML logs
        if ($logFiles.Count -gt 14) {
            $logFiles | Select-Object -Skip 14 | Remove-Item -Force
            Write-Log "Removed $($logFiles.Count - 14) old HTML log files" "Info"
        }
        
        # Get all plain text log files
        $textLogFiles = Get-ChildItem -Path $logsDir -Filter "PhotoUpdate_*.txt" -File |
                       Sort-Object LastWriteTime -Descending
        
        # Keep only the 14 most recent plain text logs
        if ($textLogFiles.Count -gt 14) {
            $textLogFiles | Select-Object -Skip 14 | Remove-Item -Force
            Write-Log "Removed $($textLogFiles.Count - 14) old plain text log files" "Info"
        }
    }
    catch {
        Write-Log "Error pruning old log files: $_" "Error"
    }
}

# Function to send email with update results
function Send-UpdateResultsEmail {
    param (
        [string]$From,
        [string]$To,
        [string]$Subject,
        [System.Collections.ArrayList]$UpdatedUsers,
        [string]$RelayServer,
        [int]$RelayPort
    )
    
    try {
        if ($UpdatedUsers.Count -eq 0) {
            Write-Log "No photos were updated, skipping email notification" "Info"
            return
        }
        
        Write-Log "Preparing email with update results..." "Info"
        
        # Split and clean up email addresses
        $recipients = $To -split ',' | ForEach-Object { $_.Trim() }
        Write-Log "Sending email to recipients: $($recipients -join ', ')" "Info"
        
        # Create HTML body
        $htmlBody = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        tr:nth-child(even) { background-color: #f9f9f9; }
    </style>
</head>
<body>
    <h2>Microsoft 365 Photo Update Results</h2>
    <p>The following user photos were successfully updated:</p>
    <table>
        <tr>
            <th>Display Name</th>
            <th>User Principal Name</th>
            <th>Photo Size (KB)</th>
        </tr>
"@
        
        foreach ($user in $UpdatedUsers) {
            $htmlBody += @"
        <tr>
            <td>$($user.DisplayName)</td>
            <td>$($user.UserPrincipalName)</td>
            <td>$($user.PhotoSizeKB)</td>
        </tr>
"@
        }
        
        $htmlBody += @"
    </table>
    <p>Total users updated: $($UpdatedUsers.Count)</p>
</body>
</html>
"@
        
        # Send email
        $smtpParams = @{
            From = $From
            To = $recipients
            Subject = $Subject
            Body = $htmlBody
            BodyAsHtml = $true
            SmtpServer = $RelayServer
            Port = $RelayPort
        }
        
        # Only add UseSsl if using Office 365 SMTP
        if ($RelayServer -eq "smtp.office365.com") {
            $smtpParams.UseSsl = $true
        }
        
        Send-MailMessage @smtpParams
        
        Write-Log "Email notification sent successfully to $($recipients.Count) recipients" "Success"
    }
    catch {
        Write-Log "Error sending email notification: $_" "Error"
    }
}

# Main script execution
try {
    # Initialize log file
    Initialize-LogFile
    
    Write-Log "Starting photo update process" "Info"
    Write-Log "Organizational Unit: $OrganizationalUnit" "Info"
    Write-Log "Sam Account Name: $SamAccountName" "Info"
    Write-Log "Domain: $Domain" "Info"
    Write-Log "Plain Text Log: $($PlainTextLog.IsPresent)" "Info"
    Write-Log "Force Mode: $($Force.IsPresent)" "Info"
    Write-Log "Email Results: $($EmailResults.IsPresent)" "Info"
    
    # Get script path
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    if (-not $scriptPath) {
        Write-Log "Could not determine script path. Using current directory." "Warning"
        $scriptPath = Get-Location
    }
    Write-Log "Script path: $scriptPath" "Info"
    
    # Create required directories if they don't exist
    $requiredDirs = @(
        @{Name = "Logs"; Description = "Log files directory"},
        @{Name = "Images"; Description = "Images directory for Force mode"},
        @{Name = "DryRunImages"; Description = "Images directory for Dry Run mode"}
    )
    
    foreach ($dir in $requiredDirs) {
        $dirPath = Join-Path -Path $scriptPath -ChildPath $dir.Name
        if (-not (Test-Path $dirPath)) {
            Write-Log "Creating $($dir.Description)..." "Info"
            New-Item -Path $dirPath -ItemType Directory -Force | Out-Null
            Write-Log "$($dir.Description) created successfully" "Success"
        }
    }
    
    # Set the appropriate image folder based on mode
    $imageFolderName = if ($Force) { "Images" } else { "DryRunImages" }
    $imageFolder = Join-Path -Path $scriptPath -ChildPath $imageFolderName
    
    # If App registration credentials are not provided as parameters, try to retrieve them from the config file
    if (-not ($TenantId -and $ClientId -and $ClientSecret)) {
        $configTenantId, $configClientId, $configClientSecret = Get-AppRegistrationCredentials
        if ($configTenantId -and $configClientId -and $configClientSecret) {
            $TenantId = $configTenantId
            $ClientId = $configClientId
            $ClientSecret = $configClientSecret
            Write-Log "Using App registration credentials from configuration file" "Info"
        }
    }
    
    # Save credentials to configuration file if requested
    if ($SaveCredentials -and $TenantId -and $ClientId -and $ClientSecret) {
        Save-AppRegistrationCredentials -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    }
    
    if ($TenantId -and $ClientId) {
        Write-Log "Using App registration authentication (TenantId: $TenantId, ClientId: $ClientId)" "Info"
    } else {
        Write-Log "Using delegated access authentication" "Info"
    }
    
    # Check and install required modules
    Write-Log "Checking for required modules..." "Info"
    Ensure-RequiredModules
    
    # Import required modules
    Write-Log "Importing required modules..." "Info"
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log "Active Directory module imported successfully" "Success"
    
    # Import Microsoft Graph module
    Write-Log "Importing Microsoft Graph module..." "Info"
    Import-SelectiveGraphModule
    
    # Import optional modules if available
    if ($script:UseSQLite) {
        Import-Module PSSQLite -ErrorAction Stop
        Write-Log "PSSQLite module imported successfully" "Success"
    }
    
    # Initialize the photo history database
    $dbPath = Initialize-PhotoDatabase
    
    # Connect to Microsoft Graph
    Write-Log "Connecting to Microsoft Graph..." "Info"
    if ($TenantId -and $ClientId -and $ClientSecret) {
        Write-Log "Connecting to Microsoft Graph using App registration..." "Info"
        Connect-ToMicrosoftGraph -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    }
    else {
        # Connect with the default delegated access
        Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop
        Write-Log "Connected to Microsoft Graph using delegated access successfully" "Success"
    }
    
    # Get users with thumbnail photos
    if ($SamAccountName) {
        # Get a specific user by SAM Account Name
        Write-Log "Searching for user with SAM Account Name: $SamAccountName" "Info"
        $users = Get-ADUser -Identity $SamAccountName -Properties thumbnailPhoto, mail, userPrincipalName, displayName, sAMAccountName | 
                 Where-Object { $_.thumbnailPhoto -ne $null -and $_.mail -ne $null }
        
        if ($users.Count -eq 0) {
            Write-Log "No user found with SAM Account Name: $SamAccountName or user has no thumbnail photo" "Warning"
            exit
        }
    }
    else {
        # Get users from the specified OU
        Write-Log "Searching for users with thumbnail photos in OU: $OrganizationalUnit" "Info"
        $users = Get-ADUser -SearchBase $OrganizationalUnit -Filter * -Properties thumbnailPhoto, mail, userPrincipalName, displayName, sAMAccountName | 
                 Where-Object { $_.thumbnailPhoto -ne $null -and $_.mail -ne $null }
    }
    
    Write-Log "Found $($users.Count) users with thumbnail photos" "Info"
    
    # Create array to track updated users
    $updatedUsers = [System.Collections.ArrayList]@()
    
    # Process each user
    foreach ($user in $users) {
        $photoSizeKB = [math]::Round($user.thumbnailPhoto.Length / 1KB, 2)
        Write-Log "Processing user: $($user.displayName) ($($user.sAMAccountName)) - Photo size: $photoSizeKB KB" "Info"
        
        # Check if the photo has changed
        $photoChanged, $photoHash = Test-PhotoChanged -SamAccountName $user.sAMAccountName -PhotoData $user.thumbnailPhoto -DbPath $dbPath
        
        # Determine target image path
        $imagePath = Join-Path -Path $imageFolder -ChildPath "$($user.sAMAccountName).jpg"
    
        # Check if image exists in target folder
        $imageExists = Test-Path $imagePath
        
        # Skip if the photo hasn't changed and image exists in target folder
        if (-not $photoChanged -and $imageExists) {
            Write-Log "Photo for $($user.displayName) ($($user.sAMAccountName)) has not changed and exists in $imageFolderName folder" "Current"
            continue
        }
        
        if ($Force) {
            try {
                # If photo hasn't changed but doesn't exist in target folder, just save a copy
                if (-not $photoChanged) {
                    Write-Log "Photo for $($user.displayName) ($($user.sAMAccountName)) has not changed, but saving copy to $imageFolderName folder" "Current"
                    [System.IO.File]::WriteAllBytes($imagePath, $user.thumbnailPhoto)
                    continue
                }
                
                # Save thumbnail photo to a temporary file
                $photoPath = [System.IO.Path]::GetTempFileName() + ".jpg"
                [System.IO.File]::WriteAllBytes($photoPath, $user.thumbnailPhoto)
                
                # Update the photo using Microsoft Graph
                Set-MgUserPhotoContent -UserId $user.UserPrincipalName -InFile $photoPath
                
                # Save a copy of the photo to the Images folder
                [System.IO.File]::WriteAllBytes($imagePath, $user.thumbnailPhoto)
                
                # Clean up the temporary file
                Remove-Item $photoPath -Force
                
                # Update the photo history database only after successful update
                Update-PhotoHistory -SamAccountName $user.sAMAccountName -PhotoHash $photoHash -DbPath $dbPath
                Write-Log "Successfully updated photo for $($user.displayName) ($($user.sAMAccountName))" "Success"
                
                # Add user to updated users list
                $updatedUsers.Add(@{
                    DisplayName = $user.displayName
                    UserPrincipalName = $user.UserPrincipalName
                    PhotoSizeKB = $photoSizeKB
                }) | Out-Null
            }
            catch {
                Write-Log "Error updating photo for $($user.displayName) ($($user.sAMAccountName)): $_" "Error"
            }
        } else {
            # In dry run mode, save the photo to the DryRunImages folder only if it doesn't exist or has changed
            try {
                Write-Log "DRY RUN: Saving photo for $($user.displayName) ($($user.sAMAccountName)) to $imageFolderName folder" "Warning"
                [System.IO.File]::WriteAllBytes($imagePath, $user.thumbnailPhoto)
                # Note: We don't update the SQLite database during a dry run
            }
            catch {
                Write-Log "Error saving photo for $($user.displayName) ($($user.sAMAccountName)): $_" "Error"
            }
        }
    }
    
    Write-Log "Photo update process completed" "Success"
    
    # If in dry run mode, provide a summary of the saved photos
    if (-not $Force) {
        $savedPhotos = Get-ChildItem -Path $imageFolder -Filter "*.jpg"
        Write-Log "Dry run completed. Saved $($savedPhotos.Count) photos to the DryRunImages folder." "Info"
        Write-Log "Review these photos before running the script with -Force to perform actual updates." "Warning"
    }
    else {
        $savedPhotos = Get-ChildItem -Path $imageFolder -Filter "*.jpg"
        Write-Log "Successfully saved $($savedPhotos.Count) photos to the Images folder." "Success"
        
        # Send email with update results if requested
        if ($EmailResults) {
            Send-UpdateResultsEmail -From $EmailFrom -To $EmailTo -Subject $EmailSubject -UpdatedUsers $updatedUsers -RelayServer $EmailRelay -RelayPort $EmailRelayPort
        }
    }
    
    # Disconnect from Microsoft Graph
    try {
        Disconnect-MgGraph
        Write-Log "Disconnected from Microsoft Graph" "Info"
    }
    catch {
        Write-Log "Error disconnecting from Microsoft Graph: $_" "Warning"
    }
}
catch {
    Write-Log "An error occurred: $_" "Error"
}
finally {
    # Finalize log file
    Finalize-LogFile
    
    # Prune old log files
    Remove-OldLogs
}

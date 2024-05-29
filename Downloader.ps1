# Read configuration from JSON file
$configPath = "$PSScriptRoot\config.json"
$config = Get-Content -Path $configPath | ConvertFrom-Json

# Get the date format from the configuration, or use the default format if not provided
$chromedateFormat = $config.chrome.logging.logDateFormat
if (-not $chromedateFormat) {
    $chromedateFormat = "dd/MM/yyyy HH:mm:ss"
}

# Function to log messages with the specified date format
$chromelogFileName = $config.chrome.logging.fileName
$chromelogFileFormat = $config.chrome.logging.fileFormat
if (-not $chromelogFileName) {
    $chromelogFileName = "chrome_downloader"
}
if (-not $chromelogFileFormat) {
    $chromelogFileFormat = "log"
}

$chromelogFileNameFormat = $chromelogFileName+"."+$chromelogFileFormat

# Get the date format from the configuration, or use the default format if not provided
$amazonworkspacedateFormat = $config.chrome.logging.logDateFormat
if (-not $amazonworkspacedateFormat) {
    $amazonworkspacedateFormat = "dd/MM/yyyy HH:mm:ss"
}

# Function to log messages with the specified date format
$amazonworkspacelogFileName = $config.amazonWorkspace.logging.fileName
$amazonworkspacelogFileFormat = $config.amazonWorkspace.logging.fileFormat
if (-not $amazonworkspacelogFileName) {
    $amazonworkspacelogFileName = "chrome_downloader"
}
if (-not $amazonworkspacelogFileFormat) {
    $amazonworkspacelogFileFormat = "log"
}

$amazonworkspacelogFileNameFormat = $amazonworkspacelogFileName+"."+$amazonworkspacelogFileFormat

function chrome-Log-Message {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format $chromedateFormat
    Write-Output "[$timestamp] - $message" | Out-File -Append -FilePath "$PSScriptRoot\$chromelogFileNameFormat" -Encoding utf8
}

function amazonworkspace-Log-Message {
    param (
        [string]$amazonworkspacemessage
    )
    $amazonworkspacetimestamp = Get-Date -Format $amazonworkspacedateFormat
    Write-Output "[$amazonworkspacetimestamp] - $amazonworkspacemessage" | Out-File -Append -FilePath "$PSScriptRoot\$amazonworkspacelogFileNameFormat" -Encoding utf8
}

# Log the start of the script

if ($config.chrome.options.enableRegularVersion) {
    chrome-Log-Message "Debug: Script started"
}
elseif ($config.chrome.options.enableForcedVersion) {
    chrome-Log-Message "Debug: Script started"
}
if ($config.amazonWorkspace.options.download) {
    amazonworkspace-Log-Message "Debug: Script started"
}


# Check if both options are disabled and log a message
if (-not $config.chrome.options.enableRegularVersion -and -not $config.chrome.options.enableForcedVersion -and -not $config.amazonWorkspace.options.download) {
    chrome-Log-Message "Warn: Neither Chrome or amazonWorkspace is selected. Please enable at least one option to proceed."
    exit
}

if ($config.chrome.options.checkExist) {
    $testPath = "$PSScriptRoot\Chrome - *"

    # Store subfolders before deletion
    $subfolders = @()
    
    if (Test-Path $testPath) {
        $subfolders = Get-ChildItem -Path $testPath -Directory | ForEach-Object { $_.FullName }
        
        # Remove the folders
        Remove-Item $testPath -Recurse -Force
    }
    
    # Output the subfolders
    foreach ($subfolder in $subfolders) {
        chrome-Log-Message "Info: The subfolder '$subfolder\' has been deleted."
    }
}
if ($config.amazonWorkspace.options.checkExist) {
    $amazonworkspacetestPath = "$PSScriptRoot\Amazon Workspace - *"

    # Store subfolders before deletion
    $amazonworkspacesubfolders = @()
    
    if (Test-Path $amazonworkspacetestPath) {
        $amazonworkspacesubfolders = Get-ChildItem -Path $amazonworkspacetestPath -Directory | ForEach-Object { $_.FullName }
        
        # Remove the folders
        Remove-Item $amazonworkspacetestPath -Recurse -Force
    }
    
    # Output the subfolders
    foreach ($amazonworkspacesubfolder in $amazonworkspacesubfolders) {
        try {
            amazonworkspace-Log-Message "Info: The subfolder '$amazonworkspacesubfolder\' has been deleted."
        } catch {
            Write-Host "Error logging message: $_"
        }
    }
}


# Define URLs
$chrome64BitUrl = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
$chrome32BitUrl = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise.msi"
$amazonworkspace64BitUrl = "https://d2td7dqidlhjx7.cloudfront.net/prod/global/windows/Amazon+WorkSpaces.msi"

# Define source and destination folders
$sourceFolderRegular = "$PSScriptRoot\Template\Chrome-Template"
$sourceFolderForced = "$PSScriptRoot\Template\Chrome-Template-Forced"
$amazonworkspacesourceFolderRegular = "$PSScriptRoot\Template\Amazon-Workspace-Template"

$destinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "Chrome - VERSION"
$forceUpdateFolder = Join-Path -Path $PSScriptRoot -ChildPath "Chrome - VERSION_force_update"
$amazonworkspacedestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "Amazon Workspace - VERSION"

# Conditional execution based on config
if ($config.chrome.options.enableRegularVersion) {
    # Create main folder and files folder if they don't exist
    $folderName = "Chrome - VERSION"
    $folderPath = Join-Path -Path $PSScriptRoot -ChildPath $folderName
    $filesFolder = Join-Path -Path $folderPath -ChildPath "Files"

    if (-not (Test-Path $filesFolder)) {
        try {
            New-Item -Path $filesFolder -ItemType Directory -ErrorAction Stop
            chrome-Log-Message "Info: Directory creation, 'Chrome - VERSION' and 'Files' folder successfully created in $PSScriptRoot"
        } catch {
            chrome-Log-Message "Error: Directory creation failed - $_"
        }
    }

    # Copy items from source folder to destination folder
    try {
        Copy-Item -Path $sourceFolderRegular\* -Destination $destinationFolder -Recurse -Force -ErrorAction Stop
        chrome-Log-Message "Info: Regular Template successfully copied to $destinationFolder"
    } catch {
        chrome-Log-Message "Error: Failed to copy Regular Template - $_"
    }
    
    # Download 64-bit Chrome installer
    $fileName1 = [System.IO.Path]::GetFileName($chrome64BitUrl)
    $filePath1 = Join-Path -Path $filesFolder -ChildPath $fileName1
    try {
        Invoke-RestMethod -Uri $chrome64BitUrl -OutFile $filePath1 -ErrorAction Stop
        chrome-Log-Message "Info: Download complete, 64-bit version of Chrome successfully downloaded to $filePath1"
    } catch {
        chrome-Log-Message "Error: 64-bit Chrome download failed - $_"
    }

    # Download 32-bit Chrome installer
    $fileName2 = [System.IO.Path]::GetFileName($chrome32BitUrl)
    $filePath2 = Join-Path -Path $filesFolder -ChildPath $fileName2
    try {
        Invoke-RestMethod -Uri $chrome32BitUrl -OutFile $filePath2 -ErrorAction Stop
        chrome-Log-Message "Info: Download complete, 32-bit version of Chrome successfully downloaded to $filePath2"
    } catch {
        chrome-Log-Message "Error: 32-bit Chrome download failed - $_"
    }
}

if ($config.chrome.options.enableForcedVersion) {
    # Create force update folder if it doesn't exist
    if (-not (Test-Path $forceUpdateFolder)) {
        try {
            New-Item -Path $forceUpdateFolder -ItemType Directory -ErrorAction Stop
            chrome-Log-Message "Info: Directory creation, 'Chrome - VERSION_force_update' successfully created in $PSScriptRoot"
        } catch {
            chrome-Log-Message "Error: Force update directory creation failed - $_"
        }
    }

    # Copy items from forced source folder to force update folder
    try {
        Copy-Item -Path "$sourceFolderForced\*" -Destination $forceUpdateFolder -Recurse -Force -ErrorAction Stop
        chrome-Log-Message "Info: Forced Template successfully copied to $forceUpdateFolder"
    } catch {
        chrome-Log-Message "Error: Failed to copy Forced Template - $_"
    }

    # If the regular version is not enabled, download 64-bit Chrome installer directly to the force update folder
    if (-not $config.chrome.options.enableRegularVersion) {
        $fileName1 = [System.IO.Path]::GetFileName($chrome64BitUrl)
        $filePath1 = Join-Path -Path $forceUpdateFolder -ChildPath $fileName1
        try {
            Invoke-RestMethod -Uri $chrome64BitUrl -OutFile $filePath1 -ErrorAction Stop
            chrome-Log-Message "Info: Download complete, 64-bit version of Chrome successfully downloaded to force update folder at $filePath1"
        } catch {
            chrome-Log-Message "Error: 64-bit Chrome download to force update folder failed - $_"
        }
    } else {
        # If the regular version is enabled, copy the downloaded 64-bit installer to the force update folder
        $fileName1 = [System.IO.Path]::GetFileName($chrome64BitUrl)
        $filePath1 = Join-Path -Path $filesFolder -ChildPath $fileName1
        if (Test-Path $filePath1) {
            try {
                Copy-Item -Path $filePath1 -Destination $forceUpdateFolder -Force -ErrorAction Stop
                chrome-Log-Message "Info: 64-bit version of Chrome copied to force update folder at $forceUpdateFolder"
            } catch {
                chrome-Log-Message "Error: Failed to copy 64-bit installer to force update folder - $_"
            }
        } else {
            chrome-Log-Message "Warn: 64-bit version of Chrome was not downloaded and could not be copied to force update folder."
        }
    }
}

if ($config.amazonWorkspace.options.download) {
# Create main folder and files folder if they don't exist
$amazonworkspacefolderName = "Amazon Workspace - VERSION"
$amazonworkspacefolderPath = Join-Path -Path $PSScriptRoot -ChildPath $amazonworkspacefolderName
$amazonworkspacefilesFolder = Join-Path -Path $amazonworkspacefolderPath -ChildPath "Files"

if (-not (Test-Path $amazonworkspacefilesFolder)) {
    try {
        New-Item -Path $amazonworkspacefilesFolder -ItemType Directory -ErrorAction Stop
        amazonworkspace-Log-Message "Info: Directory creation, 'Amazon Workspace - VERSION' and 'Files' folder successfully created in $PSScriptRoot"
    } catch {
        amazonworkspace-Log-Message "Error: Directory creation failed - $_"
    }
}

# Copy items from source folder to destination folder
try {
    Copy-Item -Path $amazonworkspacesourceFolderRegular\* -Destination $amazonworkspacedestinationFolder -Recurse -Force -ErrorAction Stop
    amazonworkspace-Log-Message "Info: Regular Template successfully copied to $amazonworkspacedestinationFolder"
} catch {
    amazonworkspace-Log-Message "Error: Failed to copy Regular Template - $_"
}

# Download 64-bit Amazon Workspace installer
$amazonworkspacefileName1 = [System.IO.Path]::GetFileName($amazonworkspace64BitUrl)
$amazonworkspacefilePath1 = Join-Path -Path $amazonworkspacefilesFolder -ChildPath $amazonworkspacefileName1
try {
    Invoke-RestMethod -Uri $amazonworkspace64BitUrl -OutFile $amazonworkspacefilePath1 -ErrorAction Stop
    amazonworkspace-Log-Message "Info: Download complete, 64-bit version of Amazon Workspace successfully downloaded to $filePath1"
} catch {
    amazonworkspace-Log-Message "Error: 64-bit Amazon Workspace download failed - $_"
}
}


if ($config.chrome.options.folderNumberedVersion -or $config.amazonWorkspace.options.folderNumberedVersion) {
	# Check if the script is running with administrative privileges
	if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
		chrome-Log-Message "Error: the config 'folderNumberedVersion' requires administrative privileges to run."
        amazonworkspace-Log-Message "Error: the config 'folderNumberedVersion' requires administrative privileges to run."
	}
	else {
		& $PSScriptRoot\Rename.ps1
	}
}
else {
    Write-Output "For additional logs, please refer to $PSScriptRoot\$logFileName."
}
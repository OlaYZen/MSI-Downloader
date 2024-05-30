# Read configuration from JSON file
$configPath = "$PSScriptRoot\config.json"
$config = Get-Content -Path $configPath | ConvertFrom-Json

$chromeNaming = $config.chrome.options.folderName
$workspacesNaming = $config.amazonWorkspace.options.folderName

# Get the date format from the configuration, or use the default format if not provided
$chromedateFormat = $config.chrome.logging.logDateFormat
if (-not $chromedateFormat) {
    $chromedateFormat = "dd'/'MM'/'yyyy HH:mm:ss"
}

# Function to log messages with the specified date format
$chromelogFileName = $config.chrome.logging.logName
$chromelogFileFormat = $config.chrome.logging.logFormat
if (-not $chromelogFileName) {
    $chromelogFileName = "google_chrome"
}
if (-not $chromelogFileFormat) {
    $chromelogFileFormat = "log"
}

$chromelogFileNameFormat = $chromelogFileName+"."+$chromelogFileFormat

# Get the date format from the configuration, or use the default format if not provided
$amazonworkspacedateFormat = $config.amazonWorkspace.logging.logDateFormat
if (-not $amazonworkspacedateFormat) {
    $amazonworkspacedateFormat = "dd'/'MM'/'yyyy HH:mm:ss"
}

# Function to log messages with the specified date format
$amazonworkspacelogFileName = $config.amazonWorkspace.logging.logName
$amazonworkspacelogFileFormat = $config.amazonWorkspace.logging.logFormat
if (-not $amazonworkspacelogFileName) {
    $amazonworkspacelogFileName = "amazon_workspaces"
}
if (-not $amazonworkspacelogFileFormat) {
    $amazonworkspacelogFileFormat = "log"
}

$amazonworkspacelogFileNameFormat = $amazonworkspacelogFileName+"."+$amazonworkspacelogFileFormat

if ($config.chrome.logging.clearLogs) {
    # Construct the full path to the log file
    $logFilePathchrome = Join-Path -Path $PSScriptRoot -ChildPath $chromelogFileNameFormat
    # Clear the contents of the log file
    Set-Content -Path $logFilePathchrome -Value $null
}

if ($config.amazonWorkspace.logging.clearLogs) {
    # Construct the full path to the log file
    $logFilePathworkspaces = Join-Path -Path $PSScriptRoot -ChildPath $amazonworkspacelogFileNameFormat
    # Clear the contents of the log file
    Set-Content -Path $logFilePathworkspaces -Value $null 
}

function chrome-Log-Message {
    param (
        [string]$message
    )
    if ($config.chrome.options.enableRegularVersion -or $config.chrome.options.enableForcedVersion) {
    $timestamp = Get-Date -Format $chromedateFormat
    Write-Output "[$timestamp] - $message" | Out-File -Append -FilePath "$PSScriptRoot\$chromelogFileNameFormat" -Encoding utf8
    }
}

function amazonworkspace-Log-Message {
        param (
        [string]$amazonworkspacemessage
    )
    if ($config.amazonWorkspace.options.download) {
    $amazonworkspacetimestamp = Get-Date -Format $amazonworkspacedateFormat
    Write-Output "[$amazonworkspacetimestamp] - $amazonworkspacemessage" | Out-File -Append -FilePath "$PSScriptRoot\$amazonworkspacelogFileNameFormat" -Encoding utf8
    }
}

# Check if both options are disabled and log a message
if (-not $config.chrome.options.enableRegularVersion -and -not $config.chrome.options.enableForcedVersion -and -not $config.amazonWorkspace.options.download) {
    chrome-Log-Message "Warn: Neither Chrome or Amazon Workspaces is selected. Please enable at least one option to proceed."
    exit
}

# Log the start of the script
if ($config.amazonWorkspace.logging.logName -eq $config.chrome.logging.logName) {
    chrome-Log-Message "Debug: Script started"
}
else {
    if ($config.chrome.options.enableRegularVersion) {
        chrome-Log-Message "Debug: Script started"
    }
    elseif ($config.chrome.options.enableForcedVersion) {
        chrome-Log-Message "Debug: Script started"
    }
    if ($config.amazonWorkspace.options.download) {
        amazonworkspace-Log-Message "Debug: Script started"
    }
}

if ($config.chrome.options.checkExist) {
    $testPath = "$PSScriptRoot\$chromeNaming *"

    # Store subfolders before deletion
    $subfolders = @()
    
    if (Test-Path $testPath) {
        $subfolders = Get-ChildItem -Path $testPath -Directory | ForEach-Object { $_.FullName }
        
        # Remove the folders
        Remove-Item $testPath -Recurse -Force
    }
    
    # Output the folders
    foreach ($subfolder in $subfolders) {
        chrome-Log-Message "Info: The folder '$subfolder\' has been deleted."
    }
}
if ($config.amazonWorkspace.options.checkExist) {
    $amazonworkspacetestPath = "$PSScriptRoot\$workspacesNaming *"

    # Store subfolders before deletion
    $amazonworkspacesubfolders = @()
    
    if (Test-Path $amazonworkspacetestPath) {
        $amazonworkspacesubfolders = Get-ChildItem -Path $amazonworkspacetestPath -Directory | ForEach-Object { $_.FullName }
        
        # Remove the folders
        Remove-Item $amazonworkspacetestPath -Recurse -Force
    }
    
    # Output the folders
    foreach ($amazonworkspacesubfolder in $amazonworkspacesubfolders) {
        try {
            amazonworkspace-Log-Message "Info: The folder '$amazonworkspacesubfolder\' has been deleted."
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

$destinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$chromeNaming VERSION"
$forceUpdateFolder = Join-Path -Path $PSScriptRoot -ChildPath "$chromeNaming VERSION_force_update"
$amazonworkspacedestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$workspacesNaming VERSION"

# Conditional execution based on config
if ($config.chrome.options.enableRegularVersion) {
    # Create main folder and files folder if they don't exist
    $folderName = "$chromeNaming VERSION"
    $folderPath = Join-Path -Path $PSScriptRoot -ChildPath $folderName
    $filesFolder = Join-Path -Path $folderPath -ChildPath "Files"

    if (-not (Test-Path $filesFolder)) {
        try {
            New-Item -Path $filesFolder -ItemType Directory -ErrorAction Stop
            chrome-Log-Message "Info: Directory creation, '$chromeNaming VERSION' and 'Files' folder successfully created in $PSScriptRoot"
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
            chrome-Log-Message "Info: Directory creation, '$chromeNaming VERSION_force_update' successfully created in $PSScriptRoot"
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
$amazonworkspacefolderName = "$workspacesNaming VERSION"
$amazonworkspacefolderPath = Join-Path -Path $PSScriptRoot -ChildPath $amazonworkspacefolderName
$amazonworkspacefilesFolder = Join-Path -Path $amazonworkspacefolderPath -ChildPath "Files"

if (-not (Test-Path $amazonworkspacefilesFolder)) {
    try {
        New-Item -Path $amazonworkspacefilesFolder -ItemType Directory -ErrorAction Stop
        amazonworkspace-Log-Message "Info: Directory creation, '$workspacesNaming VERSION' and 'Files' folder successfully created in $PSScriptRoot"
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
		if ($config.amazonWorkspace.logging.logName -eq $config.chrome.logging.logName) {
            chrome-Log-Message "Error: the config 'folderNumberedVersion' requires administrative privileges to run."
        }
        else {
            chrome-Log-Message "Error: the config 'folderNumberedVersion' requires administrative privileges to run."
            amazonworkspace-Log-Message "Error: the config 'folderNumberedVersion' requires administrative privileges to run."
        }
	}
	else {
		if ($config.chrome.options.enableRegularVersion -and -not $config.chrome.options.enableForcedVersion) {
            $msiPath = "$PSScriptRoot\$chromeNaming VERSION\Files\googlechromestandaloneenterprise64.msi"
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet" -Wait
            $chromeRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            $chromeVersion = Get-ChildItem -Path $chromeRegPath |
                                Get-ItemProperty |
                                Where-Object { $_.DisplayName -like "*Google Chrome*" } |
                                Select-Object -ExpandProperty DisplayVersion
            # Rename the folder if the version was retrieved
            if ($chromeVersion) {
                $newFolderName = "$chromeNaming $chromeVersion"
                try {
                    Rename-Item -Path $destinationFolder -NewName $newFolderName -ErrorAction Stop
                    chrome-Log-Message "Info: Folder renamed to $newFolderName"
                } catch {
                    chrome-Log-Message "Error: Failed to rename folder - $_"
                }
            } else {
                chrome-Log-Message "Warn: Chrome version could not be determined. Folder was not renamed."
            }
        }
        elseif ($config.chrome.options.enableForcedVersion -and -not $config.chrome.options.enableRegularVersion) {
            $msiPath = "$PSScriptRoot\$chromeNaming VERSION_force_update\googlechromestandaloneenterprise64.msi"
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet" -Wait
            $chromeRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            $chromeVersion = Get-ChildItem -Path $chromeRegPath |
                                Get-ItemProperty |
                                Where-Object { $_.DisplayName -like "*Google Chrome*" } |
                                Select-Object -ExpandProperty DisplayVersion
            # Rename the folder if the version was retrieved
            if ($chromeVersion) {
                $newFolderName = "$chromeNaming $chromeVersion" + "_force_update"
                try {
                    Rename-Item -Path $forceUpdateFolder -NewName $newFolderName -ErrorAction Stop
                    chrome-Log-Message "Info: Folder renamed to $newFolderName"
                } catch {
                    chrome-Log-Message "Error: Failed to rename folder - $_"
                }
            } else {
                chrome-Log-Message "Warn: Chrome version could not be determined. Folder was not renamed."
            }
        }
        elseif ($config.chrome.options.enableForcedVersion -and $config.chrome.options.enableRegularVersion) {
            $msiPath = "$PSScriptRoot\$chromeNaming VERSION_force_update\googlechromestandaloneenterprise64.msi"
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet" -Wait
        
            $chromeRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            $chromeVersion = Get-ChildItem -Path $chromeRegPath |
                                Get-ItemProperty |
                                Where-Object { $_.DisplayName -like "*Google Chrome*" } |
                                Select-Object -ExpandProperty DisplayVersion
        
            # Rename both folders if the version was retrieved
            if ($chromeVersion) {
                # Regular version folder
                $newRegularFolderName = "$chromeNaming $chromeVersion"
                try {
                    Rename-Item -Path $destinationFolder -NewName $newRegularFolderName -ErrorAction Stop
                    chrome-Log-Message "Info: Folder renamed to $newRegularFolderName"
                } catch {
                    chrome-Log-Message "Error: Failed to rename folder - $_"
                }
        
                # Forced version folder
                $newForcedFolderName = "$chromeNaming $chromeVersion" + "_force_update"
                try {
                    Rename-Item -Path $forceUpdateFolder -NewName $newForcedFolderName -ErrorAction Stop
                    chrome-Log-Message "Info: Folder renamed to $newForcedFolderName"
                } catch {
                    chrome-Log-Message "Error: Failed to rename folder - $_"
                }
            } else {
                chrome-Log-Message "Warn: Chrome version could not be determined. Folders were not renamed."
            }
        }
        if ($config.amazonWorkspace.options.download) {
        $msiPathamazonworkspace = "$PSScriptRoot\$workspacesNaming VERSION\Files\Amazon+WorkSpaces.msi"
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPathamazonworkspace`" /quiet" -Wait
        $workspacesRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        $workspacesVersion = Get-ChildItem -Path $workspacesRegPath |
                            Get-ItemProperty |
                            Where-Object { $_.DisplayName -like "*Amazon Workspaces*" } |
                            Select-Object -ExpandProperty DisplayVersion
        # Rename the folder if the version was retrieved
        if ($workspacesVersion) {
            $newFolderNameamazonworkspace = "$workspacesNaming $workspacesVersion"
            try {
                Rename-Item -Path $amazonworkspacedestinationFolder -NewName $newFolderNameamazonworkspace -ErrorAction Stop
                amazonworkspace-Log-Message "Info: Folder renamed to $newFolderNameamazonworkspace"
            } catch {
                amazonworkspace-Log-Message "Error: Failed to rename folder - $_"
            }
        } else {
            amazonworkspace-Log-Message "Warn: Amazon Workspaces version could not be determined. Folder was not renamed."
        }
        }
        if ($config.amazonWorkspace.logging.logName -eq $config.chrome.logging.logName) {
            Write-Output "For additional logs, please refer to $PSScriptRoot\$chromelogFileNameFormat."
        }
        else {
            if ($config.chrome.options.enableRegularVersion) {
                Write-Output "For additional logs, please refer to $PSScriptRoot\$chromelogFileNameFormat."
            }
            elseif ($config.chrome.options.enableForcedVersion) {
                Write-Output "For additional logs, please refer to $PSScriptRoot\$chromelogFileNameFormat."
            }
            if ($config.amazonWorkspace.options.download) {
                Write-Output "For additional logs, please refer to $PSScriptRoot\$amazonworkspacelogFileNameFormat."
            }
        }
	}
}
else {
    if ($config.amazonWorkspace.logging.logName -eq $config.chrome.logging.logName) {
        Write-Output "For additional logs, please refer to $PSScriptRoot\$chromelogFileNameFormat."
    }
    else {
        if ($config.chrome.options.enableRegularVersion) {
            Write-Output "For additional logs, please refer to $PSScriptRoot\$chromelogFileNameFormat."
        }
        elseif ($config.chrome.options.enableForcedVersion) {
            Write-Output "For additional logs, please refer to $PSScriptRoot\$chromelogFileNameFormat."
        }
        if ($config.amazonWorkspace.options.download) {
            Write-Output "For additional logs, please refer to $PSScriptRoot\$amazonworkspacelogFileNameFormat."
        }
    }
}
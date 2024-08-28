# Writes out OlaYZen's Name

write-host "|============================================|"
write-host "|=------------------------------------------=|"
write-host "|============================================|"
write-host "|                                            |"
write-host "|                                            |"
write-host "| YYYY             YYYYZZZZZZZZZZZZZZZZZZZZZ |"
write-host "|  Y::Y           Y::::::::::::::::::::::::Z |"
write-host "|   Y::Y         Y::YZZZZZZZZZZZZZZZZZZZ:::Z |"
write-host "|    Y::Y       Y::Y                 Z:::Z   |"
write-host "|     Y:::Y   Y:::Y                 Z:::Z    |"
write-host "|      Y:::Y Y:::Y                 Z:::Z     |"
write-host "|       Y:::Y:::Y                 Z:::Z      |"
write-host "|        Y:::::Y                Z:::Z        |"
write-host "|         Y:::Y                Z:::Z         |"
write-host "|         Y:::Y               Z:::Z          |"
write-host "|         Y:::Y             Z:::Z            |"
write-host "|         Y:::Y            Z:::Z             |"
write-host "|         Y:::Y           Z:::Z              |"
write-host "|         Y:::Y         Z:::ZZZZZZZZZZZZZZZZ |"
write-host "|         Y:::Y        Z:::::::::::::::::::Z |"
write-host "|         YYYYY        ZZZZZZZZZZZZZZZZZZZZZ |"
write-host "|                                            |"
write-host "|                                            |"
write-host "|============================================|"
write-host "|=------------------------------------------=|"
write-host "|=              made by OlaYZen             =|"
write-host "|=------------------------------------------=|"
write-host "|============================================|"
write-host "                                              "

# Read configuration from JSON file
$configPath = "$PSScriptRoot\config.json"
$config = Get-Content -Path $configPath | ConvertFrom-Json

if($config.license){
    # Writes out the license to the end user
    $copyrightUrl = "https://forgejo.olayzen.com/OlaYZen/MSI-Downloader/raw/branch/main/LICENSE"
    $copyrightResponse = Invoke-WebRequest -Uri $copyrightUrl
    $copyrightContent = $copyrightResponse.Content
    Write-Host $copyrightContent
}

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

# Define template name
$ctRegular = $config.chrome.template.templateFolderNameRegular
$ctForced = $config.chrome.template.templateFolderNameForced
$wt = $config.amazonWorkspace.template.templateFolderName

# Define prefix and suffix
$bp = $config.chrome.options.bothPrefix
$fs = $config.chrome.options.forcedSuffix
$ap = $config.amazonWorkspace.options.amazonPrefix

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
    if ($config.chrome.options.downloadRegular -or $config.chrome.options.downloadForced) {
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

if (-not $ctRegular) {
    $ctRegular = "Chrome-Template"
}

if (-not $ctForced) {
    $ctForced = "Chrome-Template-Forced"
}

if (-not $wt) {
    $wt = "Amazon-Workspace-Template"
}

if (-not $bp) {
    $bp = "VERSION"
}

if (-not $ap) {
    $ap = "VERSION"
}

if (-not $fs) {
    $fs = "_force_update"
}

if ($config.debug){
    if($config.chrome.logging.logName -eq $config.amazonWorkspace.logging.logName){
        if($config.chrome.logging.logFormat -eq $config.amazonWorkspace.logging.logFormat){
            chrome-Log-Message "Debug: |============================================|"
            chrome-Log-Message "Debug: |=------------------------------------------=|"
            chrome-Log-Message "Debug: |============================================|"
            chrome-Log-Message "Debug: |                                            |"
            chrome-Log-Message "Debug: |                                            |"
            chrome-Log-Message "Debug: | YYYY             YYYYZZZZZZZZZZZZZZZZZZZZZ |"
            chrome-Log-Message "Debug: |  Y::Y           Y::::::::::::::::::::::::Z |"
            chrome-Log-Message "Debug: |   Y::Y         Y::YZZZZZZZZZZZZZZZZZZZ:::Z |"
            chrome-Log-Message "Debug: |    Y::Y       Y::Y                 Z:::Z   |"
            chrome-Log-Message "Debug: |     Y:::Y   Y:::Y                 Z:::Z    |"
            chrome-Log-Message "Debug: |      Y:::Y Y:::Y                 Z:::Z     |"
            chrome-Log-Message "Debug: |       Y:::Y:::Y                 Z:::Z      |"
            chrome-Log-Message "Debug: |        Y:::::Y                Z:::Z        |"
            chrome-Log-Message "Debug: |         Y:::Y                Z:::Z         |"
            chrome-Log-Message "Debug: |         Y:::Y               Z:::Z          |"
            chrome-Log-Message "Debug: |         Y:::Y             Z:::Z            |"
            chrome-Log-Message "Debug: |         Y:::Y            Z:::Z             |"
            chrome-Log-Message "Debug: |         Y:::Y           Z:::Z              |"
            chrome-Log-Message "Debug: |         Y:::Y         Z:::ZZZZZZZZZZZZZZZZ |"
            chrome-Log-Message "Debug: |         Y:::Y        Z:::::::::::::::::::Z |"
            chrome-Log-Message "Debug: |         YYYYY        ZZZZZZZZZZZZZZZZZZZZZ |"
            chrome-Log-Message "Debug: |                                            |"
            chrome-Log-Message "Debug: |                                            |"
            chrome-Log-Message "Debug: |============================================|"
            chrome-Log-Message "Debug: |=------------------------------------------=|"
            chrome-Log-Message "Debug: |=              made by OlaYZen             =|"
            chrome-Log-Message "Debug: |=------------------------------------------=|"
            chrome-Log-Message "Debug: |============================================|"

            chrome-Log-Message "Debug: Log files share the same name '$chromelogFileNameFormat'"
        }
    } else {
        chrome-Log-Message "Debug: |============================================|"
        chrome-Log-Message "Debug: |=------------------------------------------=|"
        chrome-Log-Message "Debug: |============================================|"
        chrome-Log-Message "Debug: |                                            |"
        chrome-Log-Message "Debug: |                                            |"
        chrome-Log-Message "Debug: | YYYY             YYYYZZZZZZZZZZZZZZZZZZZZZ |"
        chrome-Log-Message "Debug: |  Y::Y           Y::::::::::::::::::::::::Z |"
        chrome-Log-Message "Debug: |   Y::Y         Y::YZZZZZZZZZZZZZZZZZZZ:::Z |"
        chrome-Log-Message "Debug: |    Y::Y       Y::Y                 Z:::Z   |"
        chrome-Log-Message "Debug: |     Y:::Y   Y:::Y                 Z:::Z    |"
        chrome-Log-Message "Debug: |      Y:::Y Y:::Y                 Z:::Z     |"
        chrome-Log-Message "Debug: |       Y:::Y:::Y                 Z:::Z      |"
        chrome-Log-Message "Debug: |        Y:::::Y                Z:::Z        |"
        chrome-Log-Message "Debug: |         Y:::Y                Z:::Z         |"
        chrome-Log-Message "Debug: |         Y:::Y               Z:::Z          |"
        chrome-Log-Message "Debug: |         Y:::Y             Z:::Z            |"
        chrome-Log-Message "Debug: |         Y:::Y            Z:::Z             |"
        chrome-Log-Message "Debug: |         Y:::Y           Z:::Z              |"
        chrome-Log-Message "Debug: |         Y:::Y         Z:::ZZZZZZZZZZZZZZZZ |"
        chrome-Log-Message "Debug: |         Y:::Y        Z:::::::::::::::::::Z |"
        chrome-Log-Message "Debug: |         YYYYY        ZZZZZZZZZZZZZZZZZZZZZ |"
        chrome-Log-Message "Debug: |                                            |"
        chrome-Log-Message "Debug: |                                            |"
        chrome-Log-Message "Debug: |============================================|"
        chrome-Log-Message "Debug: |=------------------------------------------=|"
        chrome-Log-Message "Debug: |=              made by OlaYZen             =|"
        chrome-Log-Message "Debug: |=------------------------------------------=|"
        chrome-Log-Message "Debug: |============================================|"
        amazonworkspace-Log-Message "Debug: |============================================|"
        amazonworkspace-Log-Message "Debug: |=------------------------------------------=|"
        amazonworkspace-Log-Message "Debug: |============================================|"
        amazonworkspace-Log-Message "Debug: |                                            |"
        amazonworkspace-Log-Message "Debug: |                                            |"
        amazonworkspace-Log-Message "Debug: | YYYY             YYYYZZZZZZZZZZZZZZZZZZZZZ |"
        amazonworkspace-Log-Message "Debug: |  Y::Y           Y::::::::::::::::::::::::Z |"
        amazonworkspace-Log-Message "Debug: |   Y::Y         Y::YZZZZZZZZZZZZZZZZZZZ:::Z |"
        amazonworkspace-Log-Message "Debug: |    Y::Y       Y::Y                 Z:::Z   |"
        amazonworkspace-Log-Message "Debug: |     Y:::Y   Y:::Y                 Z:::Z    |"
        amazonworkspace-Log-Message "Debug: |      Y:::Y Y:::Y                 Z:::Z     |"
        amazonworkspace-Log-Message "Debug: |       Y:::Y:::Y                 Z:::Z      |"
        amazonworkspace-Log-Message "Debug: |        Y:::::Y                Z:::Z        |"
        amazonworkspace-Log-Message "Debug: |         Y:::Y                Z:::Z         |"
        amazonworkspace-Log-Message "Debug: |         Y:::Y               Z:::Z          |"
        amazonworkspace-Log-Message "Debug: |         Y:::Y             Z:::Z            |"
        amazonworkspace-Log-Message "Debug: |         Y:::Y            Z:::Z             |"
        amazonworkspace-Log-Message "Debug: |         Y:::Y           Z:::Z              |"
        amazonworkspace-Log-Message "Debug: |         Y:::Y         Z:::ZZZZZZZZZZZZZZZZ |"
        amazonworkspace-Log-Message "Debug: |         Y:::Y        Z:::::::::::::::::::Z |"
        amazonworkspace-Log-Message "Debug: |         YYYYY        ZZZZZZZZZZZZZZZZZZZZZ |"
        amazonworkspace-Log-Message "Debug: |                                            |"
        amazonworkspace-Log-Message "Debug: |                                            |"
        amazonworkspace-Log-Message "Debug: |============================================|"
        amazonworkspace-Log-Message "Debug: |=------------------------------------------=|"
        amazonworkspace-Log-Message "Debug: |=              made by OlaYZen             =|"
        amazonworkspace-Log-Message "Debug: |=------------------------------------------=|"
        amazonworkspace-Log-Message "Debug: |============================================|"
    }


chrome-Log-Message "Debug: ctRegular set to '$ctRegular'"
chrome-Log-Message "Debug: ctForced set to '$ctForced'"
amazonworkspace-Log-Message "Debug: wt set to '$wt'"

chrome-Log-Message "Debug: folderName set to '$chromeNaming'"
amazonworkspace-Log-Message "Debug: folderName set to '$workspacesNaming'"


}

# Check if both options are disabled and log a message
if (-not $config.chrome.options.downloadRegular -and -not $config.chrome.options.downloadForced -and -not $config.amazonWorkspace.options.download) {
    chrome-Log-Message "Warn: Neither Chrome or Amazon Workspaces is selected. Please enable at least one option to proceed."
    exit
}

# Log the start of the script
if ($config.amazonWorkspace.logging.logName -eq $config.chrome.logging.logName) {
    chrome-Log-Message "Info: Script started"
}
else {
    if ($config.chrome.options.downloadRegular) {
        chrome-Log-Message "Info: Script started"
    }
    elseif ($config.chrome.options.downloadForced) {
        chrome-Log-Message "Info: Script started"
    }
    if ($config.amazonWorkspace.options.download) {
        amazonworkspace-Log-Message "Info: Script started"
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
# Chrome 64-bit URL
if ($config.chrome.options.spesificChromeURL64 -eq "") {
   $chrome64BitUrl = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
}
else {
    $chrome64BitUrl = $config.chrome.options.spesificChromeURL64
}

# Chrome 32-bit URL
if ($config.chrome.options.spesificChromeURL32 -eq "") {
    $chrome32BitUrl = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise.msi"
}
else {
    $chrome32BitUrl = $config.chrome.options.spesificChromeURL32
}

# AmazonWorkspaces URL
if ($config.amazonWorkspace.options.spesificAmazonURL -eq "") {
    $amazonworkspace64BitUrl = "https://d2td7dqidlhjx7.cloudfront.net/prod/global/windows/Amazon+WorkSpaces.msi"
}
else {
     $amazonworkspace64BitUrl = $config.amazonWorkspace.options.spesificAmazonURL
}
 


# Define source and destination folders
$sourceFolderRegular = "$PSScriptRoot\Template\$ctRegular"
$sourceFolderForced = "$PSScriptRoot\Template\$ctForced"
$amazonworkspacesourceFolderRegular = "$PSScriptRoot\Template\$wt"

$destinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$chromeNaming $bp"
$forceUpdateFolder = Join-Path -Path $PSScriptRoot -ChildPath "$chromeNaming $bp$fs"
$amazonworkspacedestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$workspacesNaming $ap"

# Conditional execution based on config
if ($config.chrome.options.downloadRegular) {
    # Create main folder and files folder if they don't exist
    $folderName = "$chromeNaming $bp"
    $folderPath = Join-Path -Path $PSScriptRoot -ChildPath $folderName
    $filesFolder = Join-Path -Path $folderPath -ChildPath "Files"

    if (-not (Test-Path $filesFolder)) {
        try {
            New-Item -Path $filesFolder -ItemType Directory -ErrorAction Stop
            chrome-Log-Message "Info: Directory creation, '$chromeNaming $bp' and 'Files' folder successfully created in $PSScriptRoot"
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

if ($config.chrome.options.downloadForced) {
    # Create force update folder if it doesn't exist
    if (-not (Test-Path $forceUpdateFolder)) {
        try {
            New-Item -Path $forceUpdateFolder -ItemType Directory -ErrorAction Stop
            chrome-Log-Message "Info: Directory creation, '$chromeNaming $bp $fs' successfully created in $PSScriptRoot"
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
    if (-not $config.chrome.options.downloadRegular) {
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
$amazonworkspacefolderName = "$workspacesNaming $ap"
$amazonworkspacefolderPath = Join-Path -Path $PSScriptRoot -ChildPath $amazonworkspacefolderName
$amazonworkspacefilesFolder = Join-Path -Path $amazonworkspacefolderPath -ChildPath "Files"

if (-not (Test-Path $amazonworkspacefilesFolder)) {
    try {
        New-Item -Path $amazonworkspacefilesFolder -ItemType Directory -ErrorAction Stop
        amazonworkspace-Log-Message "Info: Directory creation, '$workspacesNaming $ap' and 'Files' folder successfully created in $PSScriptRoot"
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

if ($config.chrome.options.folderNumber -or $config.amazonWorkspace.options.folderNumber) {
	# Check if the script is running with administrative privileges
	if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
		if ($config.amazonWorkspace.logging.logName -eq $config.chrome.logging.logName) {
            chrome-Log-Message "Error: the config 'folderNumber' requires administrative privileges to run."
        }
        else {
            chrome-Log-Message "Error: the config 'folderNumber' requires administrative privileges to run."
            amazonworkspace-Log-Message "Error: the config 'folderNumber' requires administrative privileges to run."
        }
	}
	else {
		if ($config.chrome.options.downloadRegular -and -not $config.chrome.options.downloadForced) {
            $msiPath = "$PSScriptRoot\$chromeNaming $bp\Files\googlechromestandaloneenterprise64.msi"
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
        elseif ($config.chrome.options.downloadForced -and -not $config.chrome.options.downloadRegular) {
            $msiPath = "$PSScriptRoot\$chromeNaming $bp $fs \googlechromestandaloneenterprise64.msi"
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet" -Wait
            $chromeRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            $chromeVersion = Get-ChildItem -Path $chromeRegPath |
                                Get-ItemProperty |
                                Where-Object { $_.DisplayName -like "*Google Chrome*" } |
                                Select-Object -ExpandProperty DisplayVersion
            # Rename the folder if the version was retrieved
            if ($chromeVersion) {
                $newFolderName = "$chromeNaming $chromeVersion" + "$fs"
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
        elseif ($config.chrome.options.downloadForced -and $config.chrome.options.downloadRegular) {
            $msiPath = "$PSScriptRoot\$chromeNaming $bp $fs \googlechromestandaloneenterprise64.msi"
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
                $newForcedFolderName = "$chromeNaming $chromeVersion" + "$fs"
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
        $msiPathamazonworkspace = "$PSScriptRoot\$workspacesNaming $ap\Files\Amazon+WorkSpaces.msi"
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
            if ($config.chrome.options.downloadRegular) {
                Write-Output "For additional logs, please refer to $PSScriptRoot\$chromelogFileNameFormat."
            }
            elseif ($config.chrome.options.downloadForced) {
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
        if ($config.chrome.options.downloadRegular) {
            Write-Output "For additional logs, please refer to $PSScriptRoot\$chromelogFileNameFormat."
        }
        elseif ($config.chrome.options.downloadForced) {
            Write-Output "For additional logs, please refer to $PSScriptRoot\$chromelogFileNameFormat."
        }
        if ($config.amazonWorkspace.options.download) {
            Write-Output "For additional logs, please refer to $PSScriptRoot\$amazonworkspacelogFileNameFormat."
        }
    }
}
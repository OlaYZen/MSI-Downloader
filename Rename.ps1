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
        [string]$message
    )
    $timestamp = Get-Date -Format $amazonworkspacedateFormat
    Write-Output "[$timestamp] - $message" | Out-File -Append -FilePath "$PSScriptRoot\$amazonworkspacelogFileNameFormat" -Encoding utf8
}

if ($config.chrome.options.enableRegularVersion -and -not $config.chrome.options.enableForcedVersion) {
    $msiPath = "$PSScriptRoot\Chrome - VERSION\Files\googlechromestandaloneenterprise64.msi"
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet" -Wait
    $chromeRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $chromeVersion = Get-ChildItem -Path $chromeRegPath |
                        Get-ItemProperty |
                        Where-Object { $_.DisplayName -like "*Google Chrome*" } |
                        Select-Object -ExpandProperty DisplayVersion
    # Rename the folder if the version was retrieved
    if ($chromeVersion) {
        $newFolderName = "Chrome - $chromeVersion"
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
    $msiPath = "$PSScriptRoot\Chrome - VERSION_force_update\googlechromestandaloneenterprise64.msi"
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet" -Wait
    $chromeRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $chromeVersion = Get-ChildItem -Path $chromeRegPath |
                        Get-ItemProperty |
                        Where-Object { $_.DisplayName -like "*Google Chrome*" } |
                        Select-Object -ExpandProperty DisplayVersion
    # Rename the folder if the version was retrieved
    if ($chromeVersion) {
        $newFolderName = "Chrome - $chromeVersion" + "_force_update"
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
    $msiPath = "$PSScriptRoot\Chrome - VERSION_force_update\googlechromestandaloneenterprise64.msi"
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet" -Wait

    $chromeRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $chromeVersion = Get-ChildItem -Path $chromeRegPath |
                        Get-ItemProperty |
                        Where-Object { $_.DisplayName -like "*Google Chrome*" } |
                        Select-Object -ExpandProperty DisplayVersion

    # Rename both folders if the version was retrieved
    if ($chromeVersion) {
        # Regular version folder
        $newRegularFolderName = "Chrome - $chromeVersion"
        try {
            Rename-Item -Path $destinationFolder -NewName $newRegularFolderName -ErrorAction Stop
            chrome-Log-Message "Info: Folder renamed to $newRegularFolderName"
        } catch {
            chrome-Log-Message "Error: Failed to rename folder - $_"
        }

        # Forced version folder
        $newForcedFolderName = "Chrome - $chromeVersion" + "_force_update"
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
$msiPathamazonworkspace = "$PSScriptRoot\Amazon Workspace - VERSION\Files\Amazon+WorkSpaces.msi"
Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPathamazonworkspace`" /quiet" -Wait
$workspacesRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$workspacesVersion = Get-ChildItem -Path $workspacesRegPath |
                    Get-ItemProperty |
                    Where-Object { $_.DisplayName -like "*Amazon Workspaces*" } |
                    Select-Object -ExpandProperty DisplayVersion
# Rename the folder if the version was retrieved
if ($workspacesVersion) {
    $newFolderNameamazonworkspace = "Amazon Workspace - $workspacesVersion"
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
Write-Output "For additional logs, please refer to $PSScriptRoot\$logFileName."
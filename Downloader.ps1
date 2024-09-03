# Read configuration from JSON file
$configPath = "$PSScriptRoot\config.json"
$config = Get-Content -Path $configPath | ConvertFrom-Json

# Define headers
$headers = @{
    "User-Agent"      = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    "Accept"          = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
    "Accept-Language" = "en-US,en;q=0.5"
    "Accept-Encoding" = "gzip, deflate, br"
    "Cache-Control"   = "no-cache"
}


if($config.license){
    # Writes out the license to the end user
    $copyrightUrl = "https://forgejo.olayzen.com/OlaYZen/MSI-Downloader/raw/branch/main/LICENSE"
    $copyrightResponse = Invoke-WebRequest -Uri $copyrightUrl -Headers $headers
    $copyrightContent = $copyrightResponse.Content
    Write-Host $copyrightContent
}

# Get the date format from the configuration, or use the default format if not provided
$dateFormat = $config.logging.logDateFormat
if ( $dateFormat -eq "") {
    $dateFormat = "dd'/'MM'/'yyyy HH:mm:ss"
}

# Function to log messages with the specified date format
$logFileName = $config.logging.logName
$logFileFormat = $config.logging.logFormat
if ($logFileName -eq "") {
    $logFileName = "Downloader"
}
if ($logFileFormat -eq "") {
    $logFileFormat = "log"
}

$logFileNameFormat = $logFileName+"."+$logFileFormat


# Define template name
$chromeREGULARtemplate = $config.chrome.template.templateFolderNameRegular
$chromeFORCEDtemplate = $config.chrome.template.templateFolderNameForced
$WORKSPACEStemplate = $config.amazonWorkspace.template.templateFolderName
$7ZIPtemplate = $config.SevenZip.template.templateFolderName
$VLCtemplate = $config.VLC.template.templateFolderName
$WinRARtemplate = $config.WinRAR.template.templateFolderName

# Define prefix and suffix
$CHROMEprefix = $config.chrome.options.Prefix
$ChromeFORCEDsuffix = $config.chrome.options.forcedSuffix
$WORKSPACESprefix = $config.amazonWorkspace.options.Prefix
$7ZIPprefix = $config.SevenZip.options.Prefix
$VLCprefix = $config.VLC.options.Prefix
$WinRARprefix = $config.WinRAR.options.Prefix

# Define folder names
$chromeNaming = $config.chrome.options.folderName
$workspacesNaming = $config.amazonWorkspace.options.folderName
$7ZipNaming = $config.SevenZip.options.folderName
$VLCNaming = $config.VLC.options.folderName
$WinRARNaming = $config.WinRAR.options.folderName



if ($config.logging.clearLogs) {
    # Construct the full path to the log file
    $logFilePath = Join-Path -Path $PSScriptRoot -ChildPath $logFileNameFormat
    # Clear the contents of the log file
    Set-Content -Path $logFilePath -Value $null
}

function Log_Message {
    param (
    [string]$message
)
    $timestamp = Get-Date -Format $dateFormat
    Write-Output "[$timestamp] - $message" | Out-File -Append -FilePath "$PSScriptRoot\$logFileNameFormat" -Encoding utf8

}

# Log the start of the script
Log_Message "Info: Script started"

if ($config.chrome.options.downloadRegular -or $config.chrome.options.downloadForced){
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
            Log_Message "Info: The folder '$subfolder\' has been deleted."
        }
    }
}
if ($config.amazonWorkspace.options.download){
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
                Log_Message "Info: The folder '$amazonworkspacesubfolder\' has been deleted."
            } catch {
                Write-Host "Error logging message: $_"
            }
        }
    }
}
if ($config.SevenZip.options.download){
    if ($config.SevenZip.options.checkExist) {
        $7ZiptestPath = "$PSScriptRoot\$7ZipNaming *"

        # Store subfolders before deletion
        $7Zipsubfolders = @()
        
        if (Test-Path $7ZiptestPath) {
            $7Zipsubfolders = Get-ChildItem -Path $7ZiptestPath -Directory | ForEach-Object { $_.FullName }
            
            # Remove the folders
            Remove-Item $7ZiptestPath -Recurse -Force
        }
        
        # Output the folders
        foreach ($7Zipsubfolder in $7Zipsubfolders) {
            try {
                Log_Message "Info: The folder '$7Zipsubfolder\' has been deleted."
            } catch {
                Write-Host "Error logging message: $_"
            }
        }
    }
}
if ($config.VLC.options.download){
    if ($config.VLC.options.checkExist) {
        $VLCtestPath = "$PSScriptRoot\$VLCNaming *"

        # Store subfolders before deletion
        $VLCsubfolders = @()
        
        if (Test-Path $VLCtestPath) {
            $VLCsubfolders = Get-ChildItem -Path $VLCtestPath -Directory | ForEach-Object { $_.FullName }
            
            # Remove the folders
            Remove-Item $VLCtestPath -Recurse -Force
        }
        
        # Output the folders
        foreach ($VLCsubfolder in $VLCsubfolders) {
            try {
                Log_Message "Info: The folder '$VLCsubfolder\' has been deleted."
            } catch {
                Write-Host "Error logging message: $_"
            }
        }
    }
}
if ($config.WinRAR.options.download){
    if ($config.WinRAR.options.checkExist) {
        $WinRARtestPath = "$PSScriptRoot\$WinRARNaming *"

        # Store subfolders before deletion
        $WinRARsubfolders = @()
        
        if (Test-Path $WinRARtestPath) {
            $WinRARsubfolders = Get-ChildItem -Path $WinRARtestPath -Directory | ForEach-Object { $_.FullName }
            
            # Remove the folders
            Remove-Item $WinRARtestPath -Recurse -Force
        }
        
        # Output the folders
        foreach ($WinRARsubfolder in $WinRARsubfolders) {
            try {
                Log_Message "Info: The folder '$WinRARsubfolder\' has been deleted."
            } catch {
                Write-Host "Error logging message: $_"
            }
        }
    }
}
# Define URLs
if ($config.chrome.options.downloadRegular -or $config.chrome.options.downloadForced){
    # Chrome 64-bit URL
    if ($config.chrome.options.specificURL64 -eq "") {
    $chrome64BitUrl = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
    }
    else {
        $chrome64BitUrl = $config.chrome.options.specificURL64
    }
}

if ($config.chrome.options.downloadRegular){
    # Chrome 32-bit URL
    if ($config.chrome.options.specificURL32 -eq "") {
        $chrome32BitUrl = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise.msi"
    }
    else {
        $chrome32BitUrl = $config.chrome.options.specificURL32
    }
}

if ($config.amazonWorkspace.options.download){
    # AmazonWorkspaces URL
    if ($config.amazonWorkspace.options.specificURL -eq "") {
        $amazonworkspace64BitUrl = "https://d2td7dqidlhjx7.cloudfront.net/prod/global/windows/Amazon+WorkSpaces.msi"
    }
    else {
        $amazonworkspace64BitUrl = $config.amazonWorkspace.options.specificURL
    }
}



if ($config.SevenZip.options.download){
    # 7-Zip URL
    if ($config.SevenZip.options.specificURL -eq "") {

        # Define the URL
        $7zipurl = "https://www.7-zip.org/download.html"

        # Download the HTML content of the page
        $7zipresponse = Invoke-WebRequest -Uri $7zipurl -UseBasicParsing  -Headers $headers

        # Use regex to find the first x64.msi link
        $7zipmsiLink = $7zipresponse.Content -match 'href="([^"]*x64\.msi)"' | Out-Null; $7zipmsiLink = $matches[1]

        # If a matching link was found, print it out
        if ($7zipmsiLink) {
            # Check if the link is a relative URL
            if ($7zipmsiLink -notmatch "^http") {
                # Get the base URL from the original request
                $7zipbaseUri = [uri]::new($7zipurl)
                $7zipmsiLink = [uri]::new($7zipbaseUri, $7zipmsiLink).AbsoluteUri
            }
            
            # Extract the file name from the URL
            $7ZipfileName = [System.IO.Path]::GetFileName($7zipmsiLink)
            
        } else {
            Log_Message "ERROR: 7-Zip URL not found."
        }


        $7Zip64BitUrl = $7zipmsiLink
        $7Zip64BitUrlClean = $7ZipfileName
    }
    else {
        $7Zip64BitUrl = $config.SevenZip.options.specificURL

        # Extract the file number from the URL
        $7Zip64BitUrlClean = $config.SevenZip.options.specificURL -replace '^https:\/\/www\.7-zip\.org\/a\/', ''

    }
}

if ($config.VLC.options.download){
    # VLC URL
    if ($config.VLC.options.specificURL -eq "") {
        # Define the URL
        $vlcurl = "https://download.videolan.org/pub/videolan/vlc/last/win64/"

        # Download the HTML content of the page
        $vlcresponse = Invoke-WebRequest -Uri $vlcurl -UseBasicParsing  -Headers $headers

        # Use regex to find the first win64.exe link
        $vlcexeLink = $vlcresponse.Content -match 'href="([^"]*win64\.exe)"' | Out-Null; $vlcexeLink = $matches[1]

        # If a matching link was found, print it out
        if ($vlcexeLink) {
        # Check if the link is a relative URL
        if ($vlcexeLink -notmatch "^http") {
        # Get the base URL from the original request
        $vlcbaseUri = [uri]::new($vlcurl)
        $vlcexeLink = [uri]::new($vlcbaseUri, $vlcexeLink).AbsoluteUri
        }

        # Extract the file name from the URL
        $vlcfileName = [System.IO.Path]::GetFileName($vlcexeLink)

        } else {
            Log_Message "ERROR: VLC URL not found."
        }

        $VLC64BitUrl = $vlcexeLink
        $VLC64BitUrlClean = $vlcfileName
    }
    else {
        $VLC64BitUrl = $config.VLC.options.specificURL

        # Extract the file number from the URL
        $VLC64BitUrlClean = $config.VLC.options.specificURL -replace '^https:\/\/www\.7-zip\.org\/a\/', ''

    }
}

if ($config.WinRAR.options.download){
    # WinRAR URL
    if ($config.WinRAR.options.specificURL -eq "") {


        # Define the URL
        $winrarurl = "https://www.win-rar.com/download.html?&L=0"

        # Download the HTML content of the page
        $winrarresponse = Invoke-WebRequest -Uri $winrarurl -UseBasicParsing  -Headers $headers

        # Use regex to find the first win64.exe link
        $winrarexeLink = $winrarresponse.Content -match 'href="([^"]*.exe)"' | Out-Null; $winrarexeLink = $matches[1]

        # If a matching link was found, print it out
        if ($winrarexeLink) {
        # Check if the link is a relative URL
        if ($winrarexeLink -notmatch "^http") {
        # Get the base URL from the original request
        $winrarbaseUri = [uri]::new($winrarurl)
        $winrarexeLink = [uri]::new($winrarbaseUri, $winrarexeLink).AbsoluteUri
        }

        # Extract the file name from the URL
        $winrarfileName = [System.IO.Path]::GetFileName($winrarexeLink)

        } else {
            Log_Message "ERROR: winrar URL not found."
        }

        $winrar64BitUrl = $winrarexeLink
        $winrar64BitUrlClean = $winrarfileName
    }
    else {
    $winrar64BitUrl = $config.winrar.options.specificURL

    # Extract the file number from the URL
    $winrar64BitUrlClean = $config.winrar.options.specificURL -replace '^https:\/\/www\.7-zip\.org\/a\/', ''

    }
}

# Define source and destination folders
$sourceFolderRegular = "$PSScriptRoot\Template\$chromeREGULARtemplate"
$sourceFolderForced = "$PSScriptRoot\Template\$chromeFORCEDtemplate"
$amazonworkspacesourceFolderRegular = "$PSScriptRoot\Template\$WORKSPACEStemplate"
$7ZipsourceFolderRegular = "$PSScriptRoot\Template\$7ZIPtemplate"
$VLCsourceFolderRegular = "$PSScriptRoot\Template\$VLCtemplate"
$WinRARsourceFolderRegular = "$PSScriptRoot\Template\$WinRARtemplate"

$destinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$chromeNaming $CHROMEprefix"
$forceUpdateFolder = Join-Path -Path $PSScriptRoot -ChildPath "$chromeNaming $CHROMEprefix$ChromeFORCEDsuffix"
$amazonworkspacedestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$workspacesNaming $WORKSPACESprefix"
$7ZipdestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$7ZipNaming $7ZIPprefix"
$VLCdestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$VLCNaming $VLCprefix"
$WinRARdestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$WinRARNaming $WinRARprefix"



# Conditional execution based on config
if ($config.chrome.options.downloadRegular) {
    # Create main folder and files folder if they don't exist
    $folderName = "$chromeNaming $CHROMEprefix"
    $folderPath = Join-Path -Path $PSScriptRoot -ChildPath $folderName
    $filesFolder = Join-Path -Path $folderPath -ChildPath "Files"

    if (-not (Test-Path $filesFolder)) {
        try {
            New-Item -Path $filesFolder -ItemType Directory -ErrorAction Stop
            Log_Message "Info: Directory creation, '$chromeNaming $CHROMEprefix' and 'Files' folder successfully created in $PSScriptRoot"
        } catch {
            Log_Message "Error: Directory creation failed - $_"
        }
    }

    # Copy items from source folder to destination folder
    try {
        Copy-Item -Path $sourceFolderRegular\* -Destination $destinationFolder -Recurse -Force -ErrorAction Stop
        Log_Message "Info: Regular Template successfully copied to $destinationFolder"
    } catch {
        Log_Message "Error: Failed to copy Regular Template - $_"
    }
    
    # Download 64-bit Chrome installer
    $fileName1 = [System.IO.Path]::GetFileName($chrome64BitUrl)
    $filePath1 = Join-Path -Path $filesFolder -ChildPath $fileName1
    try {
        Invoke-RestMethod -Uri $chrome64BitUrl -OutFile $filePath1 -ErrorAction Stop
        Log_Message "Info: Download complete, 64-bit version of Chrome successfully downloaded to $filePath1"
    } catch {
        Log_Message "Error: 64-bit Chrome download failed - $_"
    }

    # Download 32-bit Chrome installer
    $fileName2 = [System.IO.Path]::GetFileName($chrome32BitUrl)
    $filePath2 = Join-Path -Path $filesFolder -ChildPath $fileName2
    try {
        Invoke-RestMethod -Uri $chrome32BitUrl -OutFile $filePath2 -ErrorAction Stop
        Log_Message "Info: Download complete, 32-bit version of Chrome successfully downloaded to $filePath2"
    } catch {
        Log_Message "Error: 32-bit Chrome download failed - $_"
    }
}

if ($config.chrome.options.downloadForced) {
    # Create force update folder if it doesn't exist
    if (-not (Test-Path $forceUpdateFolder)) {
        try {
            New-Item -Path $forceUpdateFolder -ItemType Directory -ErrorAction Stop
            Log_Message "Info: Directory creation, '$chromeNaming $CHROMEprefix $ChromeFORCEDsuffix' successfully created in $PSScriptRoot"
        } catch {
            Log_Message "Error: Force update directory creation failed - $_"
        }
    }

    # Copy items from forced source folder to force update folder
    try {
        Copy-Item -Path "$sourceFolderForced\*" -Destination $forceUpdateFolder -Recurse -Force -ErrorAction Stop
        Log_Message "Info: Forced Template successfully copied to $forceUpdateFolder"
    } catch {
        Log_Message "Error: Failed to copy Forced Template - $_"
    }

    # If the regular version is not enabled, download 64-bit Chrome installer directly to the force update folder
    if (-not $config.chrome.options.downloadRegular) {
        $fileName1 = [System.IO.Path]::GetFileName($chrome64BitUrl)
        $filePath1 = Join-Path -Path $forceUpdateFolder -ChildPath $fileName1
        try {
            Invoke-RestMethod -Uri $chrome64BitUrl -OutFile $filePath1 -ErrorAction Stop
            Log_Message "Info: Download complete, 64-bit version of Chrome successfully downloaded to force update folder at $filePath1"
        } catch {
            Log_Message "Error: 64-bit Chrome download to force update folder failed - $_"
        }
    } else {
        # If the regular version is enabled, copy the downloaded 64-bit installer to the force update folder
        $fileName1 = [System.IO.Path]::GetFileName($chrome64BitUrl)
        $filePath1 = Join-Path -Path $filesFolder -ChildPath $fileName1
        if (Test-Path $filePath1) {
            try {
                Copy-Item -Path $filePath1 -Destination $forceUpdateFolder -Force -ErrorAction Stop
                Log_Message "Info: 64-bit version of Chrome copied to force update folder at $forceUpdateFolder"
            } catch {
                Log_Message "Error: Failed to copy 64-bit installer to force update folder - $_"
            }
        } else {
            Log_Message "Warn: 64-bit version of Chrome was not downloaded and could not be copied to force update folder."
        }
    }
}

if ($config.amazonWorkspace.options.download) {
# Create main folder and files folder if they don't exist
$amazonworkspacefolderName = "$workspacesNaming $WORKSPACESprefix"
$amazonworkspacefolderPath = Join-Path -Path $PSScriptRoot -ChildPath $amazonworkspacefolderName
$amazonworkspacefilesFolder = Join-Path -Path $amazonworkspacefolderPath -ChildPath "Files"

if (-not (Test-Path $amazonworkspacefilesFolder)) {
    try {
        New-Item -Path $amazonworkspacefilesFolder -ItemType Directory -ErrorAction Stop
        Log_Message "Info: Directory creation, '$workspacesNaming $WORKSPACESprefix' and 'Files' folder successfully created in $PSScriptRoot"
    } catch {
        Log_Message "Error: Directory creation failed - $_"
    }
}

# Copy items from source folder to destination folder
try {
    Copy-Item -Path $amazonworkspacesourceFolderRegular\* -Destination $amazonworkspacedestinationFolder -Recurse -Force -ErrorAction Stop
    Log_Message "Info: Regular Template successfully copied to $amazonworkspacedestinationFolder"
} catch {
    Log_Message "Error: Failed to copy Regular Template - $_"
}

# Download 64-bit Amazon Workspace installer
$amazonworkspacefileName1 = [System.IO.Path]::GetFileName($amazonworkspace64BitUrl)
$amazonworkspacefilePath1 = Join-Path -Path $amazonworkspacefilesFolder -ChildPath $amazonworkspacefileName1
try {
    Invoke-RestMethod -Uri $amazonworkspace64BitUrl -OutFile $amazonworkspacefilePath1 -ErrorAction Stop
    Log_Message "Info: Download complete, 64-bit version of Amazon Workspace successfully downloaded to $filePath1"
} catch {
    Log_Message "Error: 64-bit Amazon Workspace download failed - $_"
}
}


if ($config.SevenZip.options.download) {
    # Create force update folder if it doesn't exist
    if (-not (Test-Path $7ZipdestinationFolder)) {
        try {
            New-Item -Path $7ZipdestinationFolder -ItemType Directory -ErrorAction Stop
            Log_Message "Info: Directory creation, '$7ZipNaming $7ZIPprefix' successfully created in $PSScriptRoot"
        } catch {
            Log_Message "Error: Force update directory creation failed - $_"
        }
    }

    # Copy items from forced source folder to force update folder
    try {
        Copy-Item -Path "$7ZipsourceFolderRegular\*" -Destination $7ZipdestinationFolder -Recurse -Force -ErrorAction Stop
        Set-Content -Path "$7ZipdestinationFolder\install.cmd" -Value "`"%~dp0$7Zip64BitUrlClean`" /q ALLUSERS=1 REBOOT=ReallySuppress"
        Log_Message "Info: Forced Template successfully copied to $7ZipdestinationFolder"
    } catch {
        Log_Message "Error: Failed to copy Forced Template - $_"
    }

    $7ZipfileName1 = [System.IO.Path]::GetFileName($7Zip64BitUrl)
    $7ZipfilePath1 = Join-Path -Path $7ZipdestinationFolder -ChildPath $7ZipfileName1
    try {
        Invoke-RestMethod -Uri $7Zip64BitUrl -OutFile $7ZipfilePath1 -ErrorAction Stop
        Log_Message "Info: Download complete, 64-bit version of 7-Zip successfully downloaded to force update folder at $7ZipfilePath1"
    } catch {
        Log_Message "Error: 64-bit 7-Zip download to force update folder failed - $_"
    }
}

if ($config.VLC.options.download) {
    # Create force update folder if it doesn't exist
    if (-not (Test-Path $VLCdestinationFolder)) {
        try {
            New-Item -Path $VLCdestinationFolder -ItemType Directory -ErrorAction Stop
            Log_Message "Info: Directory creation, '$VLCNaming $VLCprefix' successfully created in $PSScriptRoot"
        } catch {
            Log_Message "Error: Force update directory creation failed - $_"
        }
    }

    # Copy items from forced source folder to force update folder
    try {
        Copy-Item -Path "$VLCsourceFolderRegular\*" -Destination $VLCdestinationFolder -Recurse -Force -ErrorAction Stop
        Set-Content -Path "$VLCdestinationFolder\install.cmd" -Value "`"%~dp0$VLC64BitUrlClean`" /q ALLUSERS=1 REBOOT=ReallySuppress"
        Log_Message "Info: Forced Template successfully copied to $VLCdestinationFolder"
    } catch {
        Log_Message "Error: Failed to copy Forced Template - $_"
    }

    $VLCfileName1 = [System.IO.Path]::GetFileName($VLC64BitUrl)
    $VLCfilePath1 = Join-Path -Path $VLCdestinationFolder -ChildPath $VLCfileName1
    try {
        Invoke-RestMethod -Uri $VLC64BitUrl -OutFile $VLCfilePath1 -ErrorAction Stop
        Log_Message "Info: Download complete, 64-bit version of VLC successfully downloaded to force update folder at $VLCfilePath1"
    } catch {
        Log_Message "Error: 64-bit VLC download to force update folder failed - $_"
    }
}

if ($config.WinRAR.options.download) {
    # Create force update folder if it doesn't exist
    if (-not (Test-Path $WinRARdestinationFolder)) {
        try {
            New-Item -Path $WinRARdestinationFolder -ItemType Directory -ErrorAction Stop
            Log_Message "Info: Directory creation, '$WinRARNaming $WinRARprefix' successfully created in $PSScriptRoot"
        } catch {
            Log_Message "Error: Force update directory creation failed - $_"
        }
    }

    # Copy items from forced source folder to force update folder
    try {
        Copy-Item -Path "$WinRARsourceFolderRegular\*" -Destination $WinRARdestinationFolder -Recurse -Force -ErrorAction Stop
        Set-Content -Path "$WinRARdestinationFolder\install.cmd" -Value "`"%~dp0$WinRAR64BitUrlClean`" /q ALLUSERS=1 REBOOT=ReallySuppress"
        Log_Message "Info: Forced Template successfully copied to $WinRARdestinationFolder"
    } catch {
        Log_Message "Error: Failed to copy Forced Template - $_"
    }

    $WinRARfileName1 = [System.IO.Path]::GetFileName($WinRAR64BitUrl)
    $WinRARfilePath1 = Join-Path -Path $WinRARdestinationFolder -ChildPath $WinRARfileName1
    try {
        Invoke-RestMethod -Uri $WinRAR64BitUrl -OutFile $WinRARfilePath1 -ErrorAction Stop
        Log_Message "Info: Download complete, 64-bit version of WinRAR successfully downloaded to force update folder at $WinRARfilePath1"
    } catch {
        Log_Message "Error: 64-bit WinRAR download to force update folder failed - $_"
    }
}

if ($config.chrome.options.folderNumber -or $config.amazonWorkspace.options.folderNumber -or $config.SevenZip.options.folderNumber -or $config.VLC.options.folderNumber -or $config.WinRAR.options.folderNumber) {
	# Check if the script is running with administrative privileges
	if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Log_Message "Error: the config 'folderNumber' requires administrative privileges to run."
	}
	else {
		if ($config.chrome.options.downloadRegular -and -not $config.chrome.options.downloadForced) {
            $msiPath = "$PSScriptRoot\$chromeNaming $CHROMEprefix\Files\googlechromestandaloneenterprise64.msi"
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
                    Log_Message "Info: Folder renamed to $newFolderName"
                } catch {
                    Log_Message "Error: Failed to rename folder - $_"
                }
            } else {
                Log_Message "Warn: Chrome version could not be determined. Folder was not renamed."
            }
        }
        elseif ($config.chrome.options.downloadForced -and -not $config.chrome.options.downloadRegular) {
            $msiPath = "$PSScriptRoot\$chromeNaming $CHROMEprefix$ChromeFORCEDsuffix\googlechromestandaloneenterprise64.msi"
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet" -Wait
            $chromeRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            $chromeVersion = Get-ChildItem -Path $chromeRegPath |
                                Get-ItemProperty |
                                Where-Object { $_.DisplayName -like "*Google Chrome*" } |
                                Select-Object -ExpandProperty DisplayVersion
            # Rename the folder if the version was retrieved
            if ($chromeVersion) {
                $newFolderName = "$chromeNaming $chromeVersion" + "$ChromeFORCEDsuffix"
                try {
                    Rename-Item -Path $forceUpdateFolder -NewName $newFolderName -ErrorAction Stop
                    Log_Message "Info: Folder renamed to $newFolderName"
                } catch {
                    Log_Message "Error: Failed to rename folder - $_"
                }
            } else {
                Log_Message "Warn: Chrome version could not be determined. Folder was not renamed."
            }
        }
        elseif ($config.chrome.options.downloadForced -and $config.chrome.options.downloadRegular) {
            $msiPath = "$PSScriptRoot\$chromeNaming $CHROMEprefix$ChromeFORCEDsuffix\googlechromestandaloneenterprise64.msi"
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
                    Log_Message "Info: Folder renamed to $newRegularFolderName"
                } catch {
                    Log_Message "Error: Failed to rename folder - $_"
                }
        
                # Forced version folder
                $newForcedFolderName = "$chromeNaming $chromeVersion" + "$ChromeFORCEDsuffix"
                try {
                    Rename-Item -Path $forceUpdateFolder -NewName $newForcedFolderName -ErrorAction Stop
                    Log_Message "Info: Folder renamed to $newForcedFolderName"
                } catch {
                    Log_Message "Error: Failed to rename folder - $_"
                }
            } else {
                Log_Message "Warn: Chrome version could not be determined. Folders were not renamed."
            }
        }
        if ($config.amazonWorkspace.options.download) {
        $msiPathamazonworkspace = "$PSScriptRoot\$workspacesNaming $WORKSPACESprefix\Files\Amazon+WorkSpaces.msi"
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
                Log_Message "Info: Folder renamed to $newFolderNameamazonworkspace"
            } catch {
                Log_Message "Error: Failed to rename folder - $_"
            }
        } else {
            Log_Message "Warn: Amazon Workspaces version could not be determined. Folder was not renamed."
        }
        }
        if ($config.SevenZip.options.download) {
            $msiPath7Zip = "$PSScriptRoot\$VLCNaming $7ZIPprefix\$7Zip64BitUrlClean"
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath7Zip`" /quiet" -Wait
            $7ZipRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            $7ZipVersion = Get-ChildItem -Path $7ZipRegPath |
                                Get-ItemProperty |
                                Where-Object { $_.DisplayName -like "*7-Zip*" } |
                                Select-Object -ExpandProperty DisplayVersion
            # Rename the folder if the version was retrieved
            if ($7ZipVersion) {
                $newFolderName7Zip = "$7ZipNaming $7ZipVersion"
                try {
                    Rename-Item -Path $7ZipdestinationFolder -NewName $newFolderName7Zip -ErrorAction Stop
                    Log_Message "Info: Folder renamed to $newFolderName7Zip"
                } catch {
                    Log_Message "Error: Failed to rename folder - $_"
                }
            } else {
                Log_Message "Warn: 7-Zip version could not be determined. Folder was not renamed."
            }
            }

            if ($config.VLC.options.download) {
                $msiPathVLC = "$PSScriptRoot\$VLCNaming $VLCprefix\$VLC64BitUrlClean"
                Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPathVLC`" /quiet" -Wait
                $VLCRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
                $VLCVersion = Get-ChildItem -Path $VLCRegPath |
                                    Get-ItemProperty |
                                    Where-Object { $_.DisplayName -like "*VLC*" } |
                                    Select-Object -ExpandProperty DisplayVersion
                # Rename the folder if the version was retrieved
                if ($VLCVersion) {
                    $newFolderNameVLC = "$VLCNaming $VLCVersion"
                    try {
                        Rename-Item -Path $VLCdestinationFolder -NewName $newFolderNameVLC -ErrorAction Stop
                        Log_Message "Info: Folder renamed to $newFolderNameVLC"
                    } catch {
                        Log_Message "Error: Failed to rename folder - $_"
                    }
                } else {
                    Log_Message "Warn: 7-Zip version could not be determined. Folder was not renamed."
                }
            }
            if ($config.WinRAR.options.download) {
                $msiPathWinRAR = "$PSScriptRoot\$WinRARNaming $WinRARprefix\$WinRAR64BitUrlClean"
                Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPathWinRAR`" /quiet" -Wait
                $WinRARRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
                $WinRARVersion = Get-ChildItem -Path $WinRARRegPath |
                                    Get-ItemProperty |
                                    Where-Object { $_.DisplayName -like "*WinRAR*" } |
                                    Select-Object -ExpandProperty DisplayVersion
                # Rename the folder if the version was retrieved
                if ($WinRARVersion) {
                    $newFolderNameWinRAR = "$WinRARNaming $WinRARVersion"
                    try {
                        Rename-Item -Path $WinRARdestinationFolder -NewName $newFolderNameWinRAR -ErrorAction Stop
                        Log_Message "Info: Folder renamed to $newFolderNameWinRAR"
                    } catch {
                        Log_Message "Error: Failed to rename folder - $_"
                    }
                } else {
                    Log_Message "Warn: 7-Zip version could not be determined. Folder was not renamed."
                }
            }

            Write-Output "For additional logs, please refer to $PSScriptRoot\$logFileNameFormat."
	}
}
else {
    Write-Output "For additional logs, please refer to $PSScriptRoot\$logFileNameFormat."
}
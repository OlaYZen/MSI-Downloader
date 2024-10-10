param (
    [switch]$y,
    [switch]$h,
    [switch]$u
)

Clear-Host
$currentVersion = "v1.0.3"

if ($h) {
    Write-Host "Usage: .\Downloader.ps1 [-y] [-u] [-h]"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -y    Automatically starts the script without requiring a Y/n response if the script is outdated."
    Write-Host "  -u    Updates the script to the latest version and restarts the script."
    Write-Host "  -h    Displays this help message."
    exit
}

if ($u) {
    $url = "https://api.github.com/repos/OlaYZen/MSI-Downloader/releases/latest"
    $response = Invoke-WebRequest -Uri $url -Headers $headers -UseBasicParsing
    $latestRelease = $response.Content | ConvertFrom-Json
    $latestVersion = $latestRelease.tag_name
    
    if ($latestVersion -ne $currentVersion) {
        $latestVersionUrl = $latestRelease.assets[0].browser_download_url

        $tempFile = "$env:TEMP\Downloader.ps1"
        
        # Delete any old Downloader.ps1 in the temp folder
        if (Test-Path -Path $tempFile) {
            Remove-Item -Path $tempFile -Force
        }

        Invoke-WebRequest -Uri $latestVersionUrl -OutFile $tempFile

        Write-Host "Downloading the latest version of the script..."
        Start-Sleep -Seconds 2

        Write-Host "Replacing the current script with the latest version..."
        Copy-Item -Path $tempFile -Destination $MyInvocation.MyCommand.Definition -Force

        Write-Host "The script has been updated. Running the latest version..."
        & $MyInvocation.MyCommand.Definition
        exit
    } else {
        Write-Host "You are already using the latest version of the script."
    }
    exit
}

function Check-NewVersion {
    param (
        [string]$repoUrl,
        [string]$currentVersion,
        [switch]$autoYes
    )

    try {
        $apiUrl = "$repoUrl/releases/latest"
        $response = Invoke-WebRequest -Uri $apiUrl -Headers $headers -UseBasicParsing
        $latestRelease = $response.Content | ConvertFrom-Json
        $latestVersion = $latestRelease.tag_name

        if ($latestVersion -ne $currentVersion) {
            Write-Host "The version $latestVersion exists. Please update from https://github.com/OlaYZen/MSI-Downloader."
            if ($autoYes) {
                $userInput = "Y"
            } else {
                $userInput = Read-Host "Do you want to start the script? (Y/n)"
            }
            if ($userInput -eq "Y" -or $userInput -eq "y" -or $userInput -eq "") {
            } else {
                exit
            }
        } else {
            Write-Host "You are using the latest version of the script."
        }
    } catch {
        Write-Host "Failed to check for a new version of the script. Please check your internet connection or the repository URL."
    }
}

$repoUrl = "https://api.github.com/repos/OlaYZen/MSI-Downloader"

Check-NewVersion -repoUrl $repoUrl -currentVersion $currentVersion -autoYes:$y




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
    $copyrightUrl = "https://raw.githubusercontent.com/OlaYZen/MSI-Downloader/refs/heads/main/LICENSE"
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
$Firefoxtemplate = $config.Firefox.template.templateFolderName
$WORKSPACEStemplate = $config.amazonWorkspace.template.templateFolderName
$7ZIPtemplate = $config.SevenZip.template.templateFolderName
$WinRARtemplate = $config.WinRAR.template.templateFolderName
$NotepadPlusPlusTemplate = $config.NotepadPlusPlus.template.templateFolderName
$VLCtemplate = $config.VLC.template.templateFolderName
$LenovoSystemUpdateTemplate = $config.LenovoSystemUpdate.template.templateFolderName
$DellCommandUpdateTemplate = $config.DellCommandUpdate.template.templateFolderName
$JabraDirectTemplate = $config.JabraDirect.template.templateFolderName

# Define prefix and suffix
$CHROMEprefix = $config.chrome.options.Prefix
$ChromeFORCEDsuffix = $config.chrome.options.forcedSuffix
$Firefoxprefix = $config.Firefox.options.Prefix
$WORKSPACESprefix = $config.amazonWorkspace.options.Prefix
$7ZIPprefix = $config.SevenZip.options.Prefix
$WinRARprefix = $config.WinRAR.options.Prefix
$NotepadPlusPlusPrefix = $config.NotepadPlusPlus.options.Prefix
$VLCprefix = $config.VLC.options.Prefix
$LenovoSystemUpdatePrefix = $config.LenovoSystemUpdate.options.Prefix
$DellCommandUpdatePrefix = $config.DellCommandUpdate.options.Prefix
$JabraDirectPrefix = $config.JabraDirect.options.Prefix

# Define folder names
$chromeNaming = $config.chrome.options.folderName
$FirefoxNaming = $config.Firefox.options.folderName
$workspacesNaming = $config.amazonWorkspace.options.folderName
$7ZipNaming = $config.SevenZip.options.folderName
$WinRARNaming = $config.WinRAR.options.folderName
$NotepadPlusPlusNaming = $config.NotepadPlusPlus.options.folderName
$VLCNaming = $config.VLC.options.folderName
$LenovoSystemUpdateNaming = $config.LenovoSystemUpdate.options.folderName
$DellCommandUpdateNaming = $config.DellCommandUpdate.options.folderName
$JabraDirectNaming = $config.JabraDirect.options.folderName

# Define source and destination folders
$sourceFolderRegular = "$PSScriptRoot\Template\$chromeREGULARtemplate"
$sourceFolderForced = "$PSScriptRoot\Template\$chromeFORCEDtemplate"
$FirefoxsourceFolderRegular = "$PSScriptRoot\Template\$Firefoxtemplate"
$amazonworkspacesourceFolderRegular = "$PSScriptRoot\Template\$WORKSPACEStemplate"
$7ZipsourceFolderRegular = "$PSScriptRoot\Template\$7ZIPtemplate"
$WinRARsourceFolderRegular = "$PSScriptRoot\Template\$WinRARtemplate"
$NotepadPlusPlusSourceFolderRegular = "$PSScriptRoot\Template\$NotepadPlusPlusTemplate"
$VLCsourceFolderRegular = "$PSScriptRoot\Template\$VLCtemplate"
$LenovoSystemUpdateSourceFolderRegular = "$PSScriptRoot\Template\$LenovoSystemUpdateTemplate"
$DellCommandUpdateSourceFolderRegular = "$PSScriptRoot\Template\$DellCommandUpdateTemplate\Template"
$DellCommandUpdateSource = "$PSScriptRoot\Template\$DellCommandUpdateTemplate"
$JabraDirectSourceFolderRegular = "$PSScriptRoot\Template\$JabraDirectTemplate"

$destinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$chromeNaming $CHROMEprefix"
$forceUpdateFolder = Join-Path -Path $PSScriptRoot -ChildPath "$chromeNaming $CHROMEprefix$ChromeFORCEDsuffix"
$FirefoxdestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$FirefoxNaming $Firefoxprefix"
$amazonworkspacedestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$workspacesNaming $WORKSPACESprefix"
$7ZipdestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$7ZipNaming $7ZIPprefix"
$WinRARdestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$WinRARNaming $WinRARprefix"
$NotepadPlusPlusDestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$NotepadPlusPlusNaming $NotepadPlusPlusPrefix"
$VLCdestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$VLCNaming $VLCprefix"
$LenovoSystemUpdateDestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$LenovoSystemUpdateNaming $LenovoSystemUpdatePrefix"
$DellCommandUpdateDestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$DellCommandUpdateNaming $DellCommandUpdatePrefix"
$JabraDirectDestinationFolder = Join-Path -Path $PSScriptRoot -ChildPath "$JabraDirectNaming $JabraDirectPrefix"


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
$apps = @(
    @{ name = "chrome"; download = $config.chrome.options.downloadRegular -or $config.chrome.options.downloadForced; deleteExist = $config.chrome.options.deleteExist; naming = $chromeNaming },
    @{ name = "Firefox"; download = $config.Firefox.options.download; deleteExist = $config.Firefox.options.deleteExist; naming = $FirefoxNaming },
    @{ name = "amazonWorkspace"; download = $config.amazonWorkspace.options.download; deleteExist = $config.amazonWorkspace.options.deleteExist; naming = $workspacesNaming },
    @{ name = "SevenZip"; download = $config.SevenZip.options.download; deleteExist = $config.SevenZip.options.deleteExist; naming = $7ZipNaming },
    @{ name = "WinRAR"; download = $config.WinRAR.options.download; deleteExist = $config.WinRAR.options.deleteExist; naming = $WinRARNaming },
    @{ name = "NotepadPlusPlus"; download = $config.NotepadPlusPlus.options.download; deleteExist = $config.NotepadPlusPlus.options.deleteExist; naming = $NotepadPlusPlusNaming },
    @{ name = "VLC"; download = $config.VLC.options.download; deleteExist = $config.VLC.options.deleteExist; naming = $VLCNaming },
    @{ name = "LenovoSystemUpdate"; download = $config.LenovoSystemUpdate.options.download; deleteExist = $config.LenovoSystemUpdate.options.deleteExist; naming = $LenovoSystemUpdateNaming },
    @{ name = "DellCommandUpdate"; download = $config.DellCommandUpdate.options.download; deleteExist = $config.DellCommandUpdate.options.deleteExist; naming = $DellCommandUpdateNaming },
    @{ name = "JabraDirect"; download = $config.JabraDirect.options.download; deleteExist = $config.JabraDirect.options.deleteExist; naming = $JabraDirectNaming }
)

foreach ($app in $apps) {
    if ($app.download -and $app.deleteExist) {
        $testPath = "$PSScriptRoot\$($app.naming) *"
        $subfolders = @()

        if (Test-Path $testPath) {
            $subfolders = Get-ChildItem -Path $testPath -Directory | ForEach-Object { $_.FullName }
            Remove-Item $testPath -Recurse -Force
        }

        foreach ($subfolder in $subfolders) {
            try {
                Log_Message "Info: The Folder `"$subfolder\`" has been deleted."
            } catch {
                Write-Host "Error: logging message: $_"
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
    if ($config.debug -eq $true) {Log_Message "Debug: `"Chrome 64-bit`" URL set to `"$chrome64BitUrl`""}
}

if ($config.chrome.options.downloadRegular){
    # Chrome 32-bit URL
    if ($config.chrome.options.specificURL32 -eq "") {
        $chrome32BitUrl = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise.msi"
    }
    else {
        $chrome32BitUrl = $config.chrome.options.specificURL32
    }
    if ($config.debug -eq $true) {Log_Message "Debug: `"Chrome 32-bit`" URL set to `"$chrome32BitUrl`""}
}

if ($config.Firefox.options.download){
    # Firefox URL
    if ($config.Firefox.options.specificURL -eq "") {
        # Define the initial URL
        $FirefoxInitialUrl = "https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=en-US"

        # Create a new HttpWebRequest object for the initial URL
        $FirefoxHttpRequest = [System.Net.HttpWebRequest]::Create($FirefoxInitialUrl)
        $FirefoxHttpRequest.Method = "HEAD"
        $FirefoxHttpRequest.AllowAutoRedirect = $false

        # Get the response to capture the redirect location
        $FirefoxResponse = $FirefoxHttpRequest.GetResponse()
        $FirefoxRedirectUrl = $FirefoxResponse.GetResponseHeader("Location")

        # Follow the redirect to get the final URL
        $FirefoxHttpRequest = [System.Net.HttpWebRequest]::Create($FirefoxRedirectUrl)
        $FirefoxHttpRequest.Method = "HEAD"
        $FirefoxResponse = $FirefoxHttpRequest.GetResponse()
        $FirefoxFinalUrl = $FirefoxResponse.ResponseUri.AbsoluteUri

        # Extract the final file name from the URL and replace %20 with a space
        $FirefoxFileName = [System.IO.Path]::GetFileName($FirefoxFinalUrl) -replace '%20', ' '

        $Firefox64BitUrl = $FirefoxFinalUrl
        $Firefox64BitUrlClean = $FirefoxFileName

    }
    else {
        $Firefox64BitUrl = $config.Firefox.options.specificURL
    }
    if ($config.debug -eq $true) {Log_Message "Debug: `"Firefox`" URL set to `"$Firefox64BitUrl`""}
}

if ($config.amazonWorkspace.options.download){
    # AmazonWorkspaces URL
    if ($config.amazonWorkspace.options.specificURL -eq "") {
        $amazonworkspace64BitUrl = "https://d2td7dqidlhjx7.cloudfront.net/prod/global/windows/Amazon+WorkSpaces.msi"
        $amazonworkspace64BitUrlClean = "Amazon+WorkSpaces.msi"
    }
    else {
        $amazonworkspace64BitUrl = $config.amazonWorkspace.options.specificURL
    }
    if ($config.debug -eq $true) {Log_Message "Debug: `"Amazon Workspaces`" URL set to `"$amazonworkspace64BitUrl`""}
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
            Log_Message "Error: 7-Zip URL not found."
        }


        $7Zip64BitUrl = $7zipmsiLink
        $7Zip64BitUrlClean = $7ZipfileName
    }
    else {
        $7Zip64BitUrl = $config.SevenZip.options.specificURL

        # Extract the file number from the URL
        $7Zip64BitUrlClean = $config.SevenZip.options.specificURL -replace '^https:\/\/www\.7-zip\.org\/a\/', ''

    }
    if ($config.debug -eq $true) {Log_Message "Debug: `"7-Zip`" URL set to `"$7Zip64BitUrl`""}
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
            Log_Message "Error: WinRAR URL not found."
        }

        $winrar64BitUrl = $winrarexeLink
        $winrar64BitUrlClean = $winrarfileName
    }
    else {
    $winrar64BitUrl = $config.winrar.options.specificURL

    # Extract the file number from the URL
    $winrar64BitUrlClean = $config.winrar.options.specificURL -replace '^https:\/\/www\.win-rar\.com\/fileadmin\/winrar-versions\/winrar\/', ''

    }
    if ($config.debug -eq $true) {Log_Message "Debug: `"WinRAR`" URL set to `"$winrar64BitUrl`""}
}

if ($config.NotepadPlusPlus.options.download){
    # NotepadPlusPlus URL
    if ($config.NotepadPlusPlus.options.specificURL -eq "") {
        # URL of the webpage
        $url = "https://notepad-plus-plus.org/downloads/"

        # Get the web page content
        $response = Invoke-RestMethod -Uri $url

        # Extract the HTML content as a string
        $htmlContent = $response

        # Define a regex to find the first link that matches the version pattern "/v[number]"
        # This looks for something like '/downloads/v8.7/'
        $versionRegex = '/downloads/v([\d\.]+)/'

        # Use regex to match the first occurrence
        $matches = [regex]::Matches($htmlContent, $versionRegex)
        
        if ($matches.Count -gt 0) {
            # Extract the version number
            $version = $matches[0].Groups[1].Value
            $NotepadPlusPlus64BitUrl = "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v$version/npp.$version.Installer.x64.exe"
            $NotepadPlusPlus64BitUrlClean = "npp.$version.Installer.x64.exe"
        }
        else {
            Log_Message "Warn: Version number not found in the HTML content."
        }
    }
    else {
        $NotepadPlusPlus64BitUrl = $config.NotepadPlusPlus.options.specificURL
    }
    if ($config.debug -eq $true) {Log_Message "Debug: `"Notepad++`" URL set to `"$NotepadPlusPlus64BitUrl`""}
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
            Log_Message "Error: VLC URL not found."
        }

        $VLC64BitUrl = $vlcexeLink
        $VLC64BitUrlClean = $vlcfileName
    }
    else {
        $VLC64BitUrl = $config.VLC.options.specificURL

        # Extract the file number from the URL
        $VLC64BitUrlClean = $config.VLC.options.specificURL -replace '^https:\/\/www\.7-zip\.org\/a\/', ''

    }
    if ($config.debug -eq $true) {Log_Message "Debug: `"VLC`" URL set to `"$VLC64BitUrl`""}
}

if ($config.LenovoSystemUpdate.options.download){
    # LenovoSystemUpdate URL
    if ($config.LenovoSystemUpdate.options.specificURL -eq "") {
    # Fetch the HTML content from the website
    $LenovoSystemUpdateurl = "https://pcsupport.lenovo.com/no/nb/products/laptops-and-netbooks/thinkpad-l-series-laptops/thinkpad-l14-gen-2-type-20x1-20x2/20x1/20x100glmx/pf3pdw07/downloads/ds012808-lenovo-system-update-for-windows-10-7-32-bit-64-bit-desktop-notebook-workstation?Products=LAPTOPS-AND-NETBOOKS%2FTHINKPAD-L-SERIES-LAPTOPS%2FTHINKPAD-L14-GEN-2-TYPE-20X1-20X2%2F20X1%2F20X100GLMX%2FPF3PDW07"
    $LenovoSystemUpdateresponse = Invoke-WebRequest -Uri $LenovoSystemUpdateurl

    # Convert the response content (HTML) into a string that can be searched
    $LenovoSystemUpdatehtmlContent = $LenovoSystemUpdateresponse.Content

    # Define a regex pattern to search for the system_update_*.exe files
    $LenovoSystemUpdateexePattern = "system_update_.*?\.exe"

    # Use regex to find all matches for system_update_*.exe in the HTML content
    $LenovoSystemUpdateMatches = [regex]::Matches($LenovoSystemUpdatehtmlContent, $LenovoSystemUpdateexePattern)

    # Check if any matches were found and display them
    if ($LenovoSystemUpdatematches.Count -gt 0) {
        foreach ($LenovoSystemUpdatematch in $LenovoSystemUpdatematches) {
            $LenovoSystemUpdate64BitUrl = "https://download.lenovo.com/pccbbs/thinkvantage_en/$($LenovoSystemUpdatematch.Value)"
            $LenovoSystemUpdate64BitUrlClean = $LenovoSystemUpdatematch.Value
        }
    } else {
            Write-Host "No system_update_*.exe file found in the HTML content."
        }


    }
    else {
        $LenovoSystemUpdate64BitUrl = $config.LenovoSystemUpdate.options.specificURL
    }
    if ($config.debug -eq $true) {Log_Message "Debug: `"Lenovo System Update`" URL set to `"$LenovoSystemUpdate64BitUrl`""}
}

if ($config.DellCommandUpdate.options.download){


    # Remove the dell.txt file if it exists
    $dellTxtPath = "$DellCommandUpdateSource\dell.txt"
    if (Test-Path $dellTxtPath) {
        Remove-Item $dellTxtPath -Force
    }
    
    # Check if Python is installed
    $pythonInstalled = Get-Command python -ErrorAction SilentlyContinue

    if (-not $pythonInstalled) {
        Log_Message "Error: Python is not installed. The script will be stopped."
        exit
    }

    # Path to the requirements.txt file
    $requirementsPath = "$DellCommandUpdateSource\requirements.txt"

    # Install the requirements using pip
    $requirementsInstalled = & python -m pip list | Select-String -Pattern "requests|selenium|beautifulsoup4|webdriver-manager"

    if ($requirementsInstalled) {
    } else {
        try {
            & python -m pip install -r $requirementsPath
        } catch {
            Log_Message "Error: Failed to install requirements from requirements.txt. $_"
            exit
        }
    }

    # Path to the Python script
    $pythonScriptPath = "$DellCommandUpdateSource\dell.py"

    # Start the Python script
    try {
        $process = Start-Process -FilePath "python" -ArgumentList $pythonScriptPath -NoNewWindow -Wait -PassThru
        if ($process.ExitCode -eq 0) {
        } else {
            Log_Message "Error: Python script dell.py failed with exit code $($process.ExitCode)"
            exit
        }
    } catch {
        Log_Message "Error: Failed to start the Python script dell.py. $_"
        exit
    }

    # Dell Command Update URL
    if ($config.DellCommandUpdate.options.specificURL -eq "") {
        $DellCommandUpdate64BitUrl = Get-Content -Path "$DellCommandUpdateSource\dell.txt" -Raw
        $DellCommandUpdate64BitUrlClean = [System.IO.Path]::GetFileName($DellCommandUpdate64BitUrl)
    }
    else {
        $DellCommandUpdate64BitUrl = $config.DellCommandUpdate.options.specificURL
    }
    if ($config.debug -eq $true) {Log_Message "Debug: `"Dell Command Update`" URL set to `"$DellCommandUpdate64BitUrl`""}
}

if ($config.JabraDirect.options.download){
    # Jabra Direct URL
    if ($config.JabraDirect.options.specificURL -eq "") {
        $JabraDirect64BitUrl = "https://jabraxpressonlineprdstor.blob.core.windows.net/jdo/JabraDirectSetup.exe"
        $JabraDirect64BitUrlClean = "JabraDirectSetup.exe"
    }
    else {
        $JabraDirect64BitUrl = $config.JabraDirect.options.specificURL
    }
    if ($config.debug -eq $true) {Log_Message "Debug: `"Jabra Direct`" URL set to `"$JabraDirect64BitUrl`""}
}

Clear-Host

# Conditional execution based on config
if ($config.chrome.options.downloadRegular) {
    # Create main folder and files folder if they don't exist
    $folderName = "$chromeNaming $CHROMEprefix"
    $folderPath = Join-Path -Path $PSScriptRoot -ChildPath $folderName
    $filesFolder = Join-Path -Path $folderPath -ChildPath "Files"

    if (-not (Test-Path $filesFolder)) {
        try {
            New-Item -Path $filesFolder -ItemType Directory -ErrorAction Stop
            Log_Message "Info: Directory creation, `"$chromeNaming $CHROMEprefix`" and `"Files`" folder successfully created in `"$PSScriptRoot`""
        } catch {
            Log_Message "Error: Directory creation failed - $_"
        }
    }

    # Copy items from source folder to destination folder
    try {
        Copy-Item -Path $sourceFolderRegular\* -Destination $destinationFolder -Recurse -Force -ErrorAction Stop
        Log_Message "Info: Regular Template successfully copied to `"$destinationFolder`""
    } catch {
        Log_Message "Error: Failed to copy Regular Template - $_"
    }
    
    # Download 64-bit Chrome installer
    $fileName1 = [System.IO.Path]::GetFileName($chrome64BitUrl)
    $filePath1 = Join-Path -Path $filesFolder -ChildPath $fileName1
    try {
        if ($config.debug -eq $true) {Log_Message "Debug: Started downloading `"$fileName1`" from `"$chrome64BitUrl`""}
        else {Log_Message "Info: Started downloading `"$fileName1`""}
        Invoke-RestMethod -Uri $chrome64BitUrl -OutFile $filePath1 -ErrorAction Stop
        Log_Message "Info: Download complete, `"64-bit`" version of Chrome successfully downloaded to $filePath1"
    } catch {
        Log_Message "Error: `"64-bit`" Chrome download failed - $_"
    }

    # Download 32-bit Chrome installer
    $fileName2 = [System.IO.Path]::GetFileName($chrome32BitUrl)
    $filePath2 = Join-Path -Path $filesFolder -ChildPath $fileName2
    try {
        if ($config.debug -eq $true) {Log_Message "Debug: Started downloading `"$fileName2`" from `"$chrome32BitUrl`""}
        else {Log_Message "Info: Started downloading `"$fileName2`""}
        Invoke-RestMethod -Uri $chrome32BitUrl -OutFile $filePath2 -ErrorAction Stop
        Log_Message "Info: Download complete,  `"32-bit`"  version of Chrome successfully downloaded to $filePath2"
    } catch {
        Log_Message "Error:  `"32-bit`"  Chrome download failed - $_"
    }
}

if ($config.chrome.options.downloadForced) {
    # Create force update folder if it doesn't exist
    if (-not (Test-Path $forceUpdateFolder)) {
        try {
            New-Item -Path $forceUpdateFolder -ItemType Directory -ErrorAction Stop
            Log_Message "Info: Directory creation, `"$chromeNaming $CHROMEprefix $ChromeFORCEDsuffix`" successfully created in `"$PSScriptRoot`""
        } catch {
            Log_Message "Error: Force update directory creation failed - $_"
        }
    }

    # Copy items from forced source folder to force update folder
    try {
        Copy-Item -Path "$sourceFolderForced\*" -Destination $forceUpdateFolder -Recurse -Force -ErrorAction Stop
        Log_Message "Info: Forced Template successfully copied to `"$forceUpdateFolder`""
    } catch {
        Log_Message "Error: Failed to copy Forced Template - $_"
    }

    # If the regular version is not enabled, download 64-bit Chrome installer directly to the force update folder
    if (-not $config.chrome.options.downloadRegular) {
        $fileName1 = [System.IO.Path]::GetFileName($chrome64BitUrl)
        $filePath1 = Join-Path -Path $forceUpdateFolder -ChildPath $fileName1
        try {
            if ($config.debug -eq $true) {Log_Message "Debug: Started downloading `"$fileName1`" from `"$chrome64BitUrl`""}
            else {Log_Message "Info: Started downloading `"$fileName1`""}
            Invoke-RestMethod -Uri $chrome64BitUrl -OutFile $filePath1 -ErrorAction Stop
            Log_Message "Info: Download complete, `"64-bit`" version of Chrome successfully downloaded to force update folder at `"$filePath1`""
        } catch {
            Log_Message "Error: `"64-bit`" Chrome download to force update folder failed - $_"
        }
    } else {
        # If the regular version is enabled, copy the downloaded 64-bit installer to the force update folder
        $fileName1 = [System.IO.Path]::GetFileName($chrome64BitUrl)
        $filePath1 = Join-Path -Path $filesFolder -ChildPath $fileName1
        if (Test-Path $filePath1) {
            try {
                Copy-Item -Path $filePath1 -Destination $forceUpdateFolder -Force -ErrorAction Stop
                Log_Message "Info: `"64-bit`" version of Chrome copied to force update folder at $forceUpdateFolder"
            } catch {
                Log_Message "Error: Failed to copy `"64-bit`" installer to force update folder - $_"
            }
        } else {
            Log_Message "Warn: `"64-bit`" version of Chrome was not downloaded and could not be copied to force update folder."
        }
    }
}

function CreateFolder {
    param (
        [string]$folderPath,
        [string]$logMessage
    )
    if (-not (Test-Path $folderPath)) {
        try {
            New-Item -Path $folderPath -ItemType Directory -ErrorAction Stop
            Log_Message "Info: Directory creation, `"$logMessage`" successfully created in `"$PSScriptRoot`""
        } catch {
            Log_Message "Error: Directory creation failed - $_"
        }
    }
}

function CopyTemplate {
    param (
        [string]$sourceFolder,
        [string]$destinationFolder
    )
    try {
        Copy-Item -Path "$sourceFolder\*" -Destination $destinationFolder -Recurse -Force -ErrorAction Stop
        Log_Message "Info: Template successfully copied to `"$destinationFolder`""
    } catch {
        Log_Message "Error: Failed to copy Template - $_"
    }
}

function DownloadInstaller {
    param (
        [string]$url,
        [string]$destinationFolder
    )
    $fileName = [System.IO.Path]::GetFileName($url)
    $filePath = Join-Path -Path $destinationFolder -ChildPath $fileName
    try {
        if ($config.debug -eq $true) {Log_Message "Debug: Started downloading `"$fileName`" from `"$url`""}
        else {Log_Message "Info: Started downloading `"$fileName`""}
        Invoke-RestMethod -Uri $url -OutFile $filePath -ErrorAction Stop -Headers $headers
        Log_Message "Info: Download complete, `"$fileName`" successfully downloaded to `"$filePath`""
    } catch {
        Log_Message "Error: Download failed - $_"
    }
}

function CreateCmd {
    param (
        [string]$destinationFolder,
        [string]$content,
        [string]$fileName
    )
    try {
        Set-Content -Path "$destinationFolder\$fileName" -Value $content
        Log_Message "Info: $fileName successfully created in `"$destinationFolder`""
    } catch {
        Log_Message "Error: Failed to create $fileName - $_"
    }
}

function MoveFolder {
    param (
        [string]$sourceFolder,
        [string]$file
    )
    $filesFolder = Join-Path -Path $sourceFolder -ChildPath "Files"
    $filePath = Join-Path -Path $sourceFolder -ChildPath $file

    if (-not (Test-Path $filesFolder)) {
        try {
            New-Item -Path $filesFolder -ItemType Directory -ErrorAction Stop | Out-Null
            Log_Message "Info: Directory creation, `"$filesFolder`" successfully created."
        } catch {
            Log_Message "Error: Directory creation failed - $_"
            return
        }
    }

    try {
        Move-Item -Path $filePath -Destination $filesFolder -ErrorAction Stop | Out-Null
        Log_Message "Info: File `"$filePath`" successfully moved to `"$filesFolder`""
    } catch {
        Log_Message "Error: File move failed - $_"
    }
}

if ($config.Firefox.options.download) {
    CreateFolder -folderPath $FirefoxdestinationFolder -logMessage "$FirefoxNaming $Firefoxprefix"
    CopyTemplate -sourceFolder $FirefoxsourceFolderRegular -destinationFolder $FirefoxdestinationFolder
    CreateCmd -destinationFolder $FirefoxdestinationFolder -content "`"%~dp0$Firefox64BitUrlClean`" /q /norestart `nif not exist `"C:\Program Files\Mozilla Firefox\distribution\`" mkdir `"C:\Program Files\Mozilla Firefox\distribution\`" `nxcopy `"%~dp0policies.json`" `"C:\Program Files\Mozilla Firefox\distribution\`" /y /s" -fileName "install.cmd"
    DownloadInstaller -url $Firefox64BitUrl -destinationFolder $FirefoxdestinationFolder
    # Replace %20 with spaces in the file name
    $firefoxFileName = [System.IO.Path]::GetFileName($Firefox64BitUrl).Replace('%20', ' ')
    Rename-Item -Path (Join-Path -Path $FirefoxdestinationFolder -ChildPath ([System.IO.Path]::GetFileName($Firefox64BitUrl))) -NewName $firefoxFileName
    Log_Message "Info: Firefox file renamed to `"$firefoxFileName`""
}

if ($config.amazonWorkspace.options.download) {
    CreateFolder -folderPath $amazonworkspacedestinationFolder -logMessage "$workspacesNaming $WORKSPACESprefix"
    CopyTemplate -sourceFolder $amazonworkspacesourceFolderRegular -destinationFolder $amazonworkspacedestinationFolder
    DownloadInstaller -url $amazonworkspace64BitUrl -destinationFolder $amazonworkspacedestinationFolder
    MoveFolder -sourceFolder $amazonworkspacedestinationFolder -file $amazonworkspace64BitUrlClean
}

if ($config.SevenZip.options.download) {
    CreateFolder -folderPath $7ZipdestinationFolder -logMessage "$7ZipNaming $7ZIPprefix"
    CopyTemplate -sourceFolder $7ZipsourceFolderRegular -destinationFolder $7ZipdestinationFolder
    CreateCmd -destinationFolder $7ZipdestinationFolder -content "`"%~dp0$7Zip64BitUrlClean`" /q ALLUSERS=1 REBOOT=ReallySuppress" -fileName "install.cmd"
    DownloadInstaller -url $7Zip64BitUrl -destinationFolder $7ZipdestinationFolder
}

if ($config.WinRAR.options.download) {
    CreateFolder -folderPath $WinRARdestinationFolder -logMessage "$WinRARNaming $WinRARprefix"
    CopyTemplate -sourceFolder $WinRARsourceFolderRegular -destinationFolder $WinRARdestinationFolder
    CreateCmd -destinationFolder $WinRARdestinationFolder -content "`"%~dp0$WinRAR64BitUrlClean`" /q ALLUSERS=1 REBOOT=ReallySuppress" -fileName "install.cmd"
    DownloadInstaller -url $WinRAR64BitUrl -destinationFolder $WinRARdestinationFolder
}

if ($config.NotepadPlusPlus.options.download) {
    CreateFolder -folderPath $NotepadPlusPlusDestinationFolder -logMessage "$NotepadPlusPlusNaming $NotepadPlusPlusPrefix"
    CopyTemplate -sourceFolder $NotepadPlusPlussourceFolderRegular -destinationFolder $NotepadPlusPlusDestinationFolder
    CreateCmd -destinationFolder $NotepadPlusPlusDestinationFolder -content "taskkill /F /IM notepad++.exe`n`"%~dp0$NotepadPlusPlus64BitUrlClean`" /S" -fileName "install.cmd"
    DownloadInstaller -url $NotepadPlusPlus64BitUrl -destinationFolder $NotepadPlusPlusDestinationFolder
}

if ($config.VLC.options.download) {
    CreateFolder -folderPath $VLCdestinationFolder -logMessage "$VLCNaming $VLCprefix"
    CopyTemplate -sourceFolder $VLCsourceFolderRegular -destinationFolder $VLCdestinationFolder
    CreateCmd -destinationFolder $VLCdestinationFolder -content "`"%~dp0$VLC64BitUrlClean`" /q ALLUSERS=1 REBOOT=ReallySuppress" -fileName "install.cmd"
    DownloadInstaller -url $VLC64BitUrl -destinationFolder $VLCdestinationFolder
}

if ($config.LenovoSystemUpdate.options.download) {
    CreateFolder -folderPath $LenovoSystemUpdateDestinationFolder -logMessage "$LenovoSystemUpdateNaming $LenovoSystemUpdateprefix"
    CopyTemplate -sourceFolder $LenovoSystemUpdatesourceFolderRegular -destinationFolder $LenovoSystemUpdateDestinationFolder
    CreateCmd -destinationFolder $LenovoSystemUpdateDestinationFolder -content "`"%~dp0$LenovoSystemUpdate64BitUrlClean`" /verysilent /norestart /suppressmsgboxes /sp-" -fileName "install.cmd"
    DownloadInstaller -url $LenovoSystemUpdate64BitUrl -destinationFolder $LenovoSystemUpdateDestinationFolder
}

if ($config.DellCommandUpdate.options.download) {
    CreateFolder -folderPath $DellCommandUpdateDestinationFolder -logMessage "$DellCommandUpdateNaming $DellCommandUpdatePrefix"
    CopyTemplate -sourceFolder $DellCommandUpdatesourceFolderRegular -destinationFolder $DellCommandUpdateDestinationFolder
    CreateCmd -destinationFolder $DellCommandUpdateDestinationFolder -content "`"%~dp0$DellCommandUpdate64BitUrlClean`" /S`nping localhost -n 5 >NUL`nexit" -fileName "install.cmd"
    CreateCmd -destinationFolder $DellCommandUpdateDestinationFolder -content "`"C:\Program Files\Dell\CommandUpdate\dcu-cli.exe`" /applyUpdates -silent" -fileName "run.cmd"
    DownloadInstaller -url $DellCommandUpdate64BitUrl -destinationFolder $DellCommandUpdateDestinationFolder
}

if ($config.JabraDirect.options.download) {
    CreateFolder -folderPath $JabraDirectDestinationFolder -logMessage "$JabraDirectNaming $JabraDirectPrefix"
    CopyTemplate -sourceFolder $JabraDirectsourceFolderRegular -destinationFolder $JabraDirectdestinationFolder
    DownloadInstaller -url $JabraDirect64BitUrl -destinationFolder $JabraDirectDestinationFolder
    MoveFolder -sourceFolder $JabraDirectDestinationFolder -file $JabraDirect64BitUrlClean
}

$CheckerFolder = $config.chrome.options.folderNumber -or $config.Firefox.options.folderNumber -or $config.amazonWorkspace.options.folderNumber -or $config.SevenZip.options.folderNumber -or $config.WinRAR.options.folderNumber -or $config.NotepadPlusPlus.options.folderNumber -or $config.VLC.options.folderNumber -or $config.LenovoSystemUpdate.options.folderNumber -or $config.DellCommandUpdate.options.folderNumber -or $config.JabraDirect.options.folderNumber 

if ($CheckerFolder) {
	# Check if the script is running with administrative privileges
	if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Log_Message "Error: the config 'folderNumber' requires administrative privileges to run."
	}
	else {
		if ($config.chrome.options.downloadRegular -and -not $config.chrome.options.downloadForced) {
            $msiPath = "$PSScriptRoot\$chromeNaming $CHROMEprefix\Files\googlechromestandaloneenterprise64.msi"
            if ($config.debug -eq $true) {Log_Message "Debug: msiPath: `"$msiPath`""}
            Log_Message "Info: Starting Chrome installation"
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet" -Wait
            Log_Message "Info: Chrome installation completed"
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
                    Log_Message "Info: Folder renamed to `"$newFolderName`""
                } catch {
                    Log_Message "Error: Failed to rename folder - $_"
                }
            } else {
                Log_Message "Warn: Chrome version could not be determined. Folder was not renamed."
            }
        }
        elseif ($config.chrome.options.downloadForced -and -not $config.chrome.options.downloadRegular) {
            $msiPath = "$PSScriptRoot\$chromeNaming $CHROMEprefix$ChromeFORCEDsuffix\googlechromestandaloneenterprise64.msi"
            if ($config.debug -eq $true) {Log_Message "Debug: msiPath: `"$msiPath`""}
            Log_Message "Info: Starting Chrome installation"
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet" -Wait
            Log_Message "Info: Chrome installation completed"
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
                    Log_Message "Info: Folder renamed to `"$newFolderName`""
                } catch {
                    Log_Message "Error: Failed to rename folder - $_"
                }
            } else {
                Log_Message "Warn: Chrome version could not be determined. Folder was not renamed."
            }
        }
        elseif ($config.chrome.options.downloadForced -and $config.chrome.options.downloadRegular) {
            $msiPath = "$PSScriptRoot\$chromeNaming $CHROMEprefix$ChromeFORCEDsuffix\googlechromestandaloneenterprise64.msi"
            if ($config.debug -eq $true) {Log_Message "Debug: msiPath: `"$msiPath`""}
            Log_Message "Info: Starting Chrome installation"
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet" -Wait
            Log_Message "Info: Chrome installation completed"
        
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
                    Log_Message "Info: Folder renamed to `"$newRegularFolderName`""
                } catch {
                    Log_Message "Error: Failed to rename folder - $_"
                }
        
                # Forced version folder
                $newForcedFolderName = "$chromeNaming $chromeVersion" + "$ChromeFORCEDsuffix"
                try {
                    Rename-Item -Path $forceUpdateFolder -NewName $newForcedFolderName -ErrorAction Stop
                    Log_Message "Info: Folder renamed to `"$newForcedFolderName"
                } catch {
                    Log_Message "Error: Failed to rename folder - $_"
                }
            } else {
                Log_Message "Warn: Chrome version could not be determined. Folders were not renamed."
            }
        }
        $appsToInstall = @(
            @{ name = "Firefox"; config = $config.Firefox; msiPath = "$PSScriptRoot\$FirefoxNaming $Firefoxprefix\$Firefox64BitUrlClean"; installArgs = "/i `"$PSScriptRoot\$FirefoxNaming $Firefoxprefix\$Firefox64BitUrlClean`" /quiet /norestart" },
            @{ name = "Amazon Workspaces"; config = $config.amazonWorkspace; msiPath = "$PSScriptRoot\$workspacesNaming $WORKSPACESprefix\Files\Amazon+WorkSpaces.msi"; installArgs = "/i `"$PSScriptRoot\$workspacesNaming $WORKSPACESprefix\Files\Amazon+WorkSpaces.msi`" /quiet /norestart" },
            @{ name = "7-Zip"; config = $config.SevenZip; msiPath = "$PSScriptRoot\$7ZipNaming $7ZIPprefix\$7Zip64BitUrlClean"; installArgs = "/i `"$PSScriptRoot\$7ZipNaming $7ZIPprefix\$7Zip64BitUrlClean`" /quiet /norestart" },
            @{ name = "WinRAR"; config = $config.WinRAR; msiPath = "$PSScriptRoot\$WinRARNaming $WinRARprefix\$WinRAR64BitUrlClean"; installArgs = "/S" },
            @{ name = "Notepad++"; config = $config.NotepadPlusPlus; msiPath = "$PSScriptRoot\$NotepadPlusPlusNaming $NotepadPlusPlusprefix\$NotepadPlusPlus64BitUrlClean"; installArgs = "/S" },
            @{ name = "VLC"; config = $config.VLC; msiPath = "$PSScriptRoot\$VLCNaming $VLCprefix\$VLC64BitUrlClean"; installArgs = "/S" },
            @{ name = "Lenovo System Update"; config = $config.LenovoSystemUpdate; msiPath = "$PSScriptRoot\$LenovoSystemUpdateNaming $LenovoSystemUpdateprefix\$LenovoSystemUpdate64BitUrlClean"; installArgs = "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART" },
            @{ name = "Dell Command | Update"; config = $config.DellCommandUpdate; msiPath = "$PSScriptRoot\$DellCommandUpdateNaming $DellCommandUpdatePrefix\$DellCommandUpdate64BitUrlClean"; installArgs = "/S" },
            @{ name = "Jabra Direct"; config = $config.JabraDirect; msiPath = "$PSScriptRoot\$JabraDirectNaming $JabraDirectprefix\Files\JabraDirectSetup.exe"; installArgs = "/install /quiet /norestart" }
        )

        foreach ($app in $appsToInstall) {
            if ($app.config.options.download) {
                try {
                    if ($app.msiPath -like "*.msi") {
                        if ($config.debug -eq $true) {Log_Message "Debug: msiPath: `"$($app.msiPath)`""}
                        Log_Message "Info: Starting $($app.name) installation"
                        Start-Process -FilePath "msiexec.exe" -ArgumentList $app.installArgs -Wait
                        Log_Message "Info: $($app.name) installation completed"
                    } elseif ($app.msiPath -like "*.exe") {
                        if ($config.debug  -eq $true) {Log_Message "Debug: msiPath: `"$($app.msiPath)`""}
                        Log_Message "Info: Starting $($app.name) installation"
                        Start-Process -FilePath $app.msiPath -ArgumentList $app.installArgs -Wait
                        Log_Message "Info: $($app.name) installation completed"
                    } else {
                        Log_Message "Error: Unsupported file type for $($app.name) installation."
                    }
                } catch {
                    Log_Message "Error: Failed to start $($app.name) installation - $_"
                    continue
                }

                $appVersion = $null
                $regPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
                foreach ($regPath in $regPaths) {
                    $appInfo = Get-ChildItem -Path $regPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "*$($app.name)*" }
                    if ($appInfo) {
                        $appVersion = $appInfo | Sort-Object -Property InstallDate -Descending | Select-Object -First 1 | Select-Object -ExpandProperty DisplayVersion
                        break
                    }
                }

                if ($appVersion) {
                    $newFolderName = "$($app.config.options.folderName) $appVersion"
                    $oldFolderPath = "$PSScriptRoot\$($app.config.options.folderName) $($app.config.options.Prefix)"
                    if (Test-Path -Path $oldFolderPath) {
                        try {
                            Rename-Item -Path $oldFolderPath -NewName $newFolderName -ErrorAction Stop
                            Log_Message "Info: Folder renamed to `"$newFolderName`""
                        } catch {
                            Log_Message "Error: Failed to rename folder - $_"
                        }
                    } else {
                        Log_Message "Error: Folder path '$oldFolderPath' does not exist. Cannot rename folder."
                    }
                } else {
                    Log_Message "Warn: $($app.name) version could not be determined. Folder was not renamed."
                }
            }
        }
	}
}
else {
    Write-Output "For additional logs, please refer to $PSScriptRoot\$logFileNameFormat"
}
$Checker = $config.chrome.options.downloadRegular -or $config.chrome.options.downloadForced -or $config.Firefox.options.download -or $config.amazonWorkspace.options.download -or $config.SevenZip.options.download -or $config.WinRAR.options.download -or $config.NotepadPlusPlus.options.download -or $config.VLC.options.download -or $config.LenovoSystemUpdate.options.download -or $config.DellCommandUpdate.options.download -or $config.JabraDirect.options.download 



if ($config.old -eq $true -and $checker -eq $true) {
    $oldFolderPath = "$PSScriptRoot\.Old"
    if (Test-Path -Path $oldFolderPath) {
        if ($config.debug  -eq $true) {
            Log_Message "Debug: .Old folder exists at '$oldFolderPath'."
            Log_Message "Debug: Starting old folder check."
        }
        else {Log_Message "Info: Starting old folder check."}
    $downloadedApps = $apps | Where-Object { $_.download }
    foreach ($app in $downloadedApps) {
        $folderNamePattern = "$($app.naming) *"
        $downloadedFolders = Get-ChildItem -Path $PSScriptRoot -Directory | Where-Object { $_.Name -like $folderNamePattern }
        foreach ($folder in $downloadedFolders) {
            $oldFolder = Join-Path -Path $oldFolderPath -ChildPath $folder.Name
            if (Test-Path -Path $oldFolder) {
                if ($config.debug -eq $true) {Log_Message "Debug: The folder `"$($folder.Name)`" exists in the .Old folder."}
                Remove-Item -Path $folder.FullName -Recurse -Force
                if ($config.debug -eq $true) {Log_Message "Debug: The folder `"$($folder.Name)`" has been deleted."}
            } else {
                Log_Message "Info: New file detected `"$($folder.Name)`""
                Write-Output "Info: New file detected `"$($folder.Name)`""
            }
        }
    }
    } else {
        Log_Message "Warn: .Old folder does not exist at '$oldFolderPath'."
    }
}

if (-not $Checker) {
    Log_Message "Warn: No download version selected in the configuration."
}
else {
    Write-Output "For additional logs, please refer to $PSScriptRoot\$logFileNameFormat"
    Log_Message "Info: Script ended"
}
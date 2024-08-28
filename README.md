# Google Chrome & Amazon Workspaces Downloader MSI Script

This PowerShell script automates the process of downloading and organizing Google Chrome and Amazon Workspaces installers based on specified configurations. It supports downloading both 64-bit and 32-bit versions of Chrome and organizing them into appropriate folders. (64-bit only on Amazon Workspaces)

## Script Overview

## Configuration

The script reads its configuration from a JSON file named config.json located in the same directory as the script. The configuration options include:

### Configuration File
The `config.json` file should be structured as follows:
```json
{
  "chrome":{
    "options": {
      "downloadRegular": true,
      "downloadForced": false,
      "folderNumber": false,
      "checkExist": false,
      "folderName": "Chrome -",
      "bothPrefix": "VERSION",
      "forcedSuffix": "_force_update",
      "spesificChromeURL64": "",
      "spesificChromeURL32": ""
      },
      "logging": {
        "logName": "google_chrome",
        "logFormat": "log",
        "logDateFormat": "dd'/'MM'/'yyyy HH:mm:ss",
        "clearLogs": false
      },
      "template": {
        "templateFolderNameRegular": "Chrome-Template",
        "templateFolderNameForced": "Chrome-Template-Forced"
      }
  },
  "amazonWorkspace":{
    "options": {
      "download": false,
      "folderNumber": false,
      "checkExist": false,
      "folderName": "WorkSpaces -",
      "AmazonPrefix": "VERSION",
      "spesificAmazonURL": ""
      },
      "logging": {
        "logName": "amazon_workspaces",
        "logFormat": "log",
        "logDateFormat": "dd'/'MM'/'yyyy HH:mm:ss",
        "clearLogs": false
      },
      "template": {
        "templateFolderName": "Amazon-Workspace-Template"
      }
  },
  "license": true,
  "debug": false
}
```
#### Options:
---
##### Chrome Specific:
- `downloadRegular`: A boolean flag to enable downloading and installing the regular version of Chrome.
- `downloadForced`: A boolean flag to enable downloading and installing the forced update version of Chrome.
- `bothPrefix`: A String defining the prefix for both the regular and forced update versions of Chrome. The default prefix is `VERSION`.
- `forcedSuffix`: A string defining the suffix for the forced update version of Chrome. The default suffix is `_force_update`.
- `spesificChromeURL64`: A string defining custom URL for the 64-bit version of Chrome. Leave empty to use the default URL.
- `spesificChromeURL32`: A string defining custom URL for the 32-bit version of Chrome. Leave empty to use the default URL.

##### WorkSpaces Specific:

- `download`: A boolean flag to enable downloading and installing Amazon Workspaces.
- `AmazonPrefix`: A String defining the prefix for the Amazon Workspaces version. The default prefix is `VERSION`.
- `spesificAmazonURL`: A string defining custom URL for the Amazon Workspaces version. Leave empty to use the default URL.

##### Universal:

- `folderNumber`: A boolean flag to enable the automatic renaming of the folder to the newest version of Chrome. ⚠️ **This option requires administrative privileges when executing the script!** ⚠️
- `checkExist`: A boolean flag to delete old Chrome folders when the script is executed. ⚠️ **This action will delete your Chrome folders, so ensure you have backups if you wish to retain them.** ⚠️
- `folderName`: A string defining the name of the folders. The default name is `Chrome -` and  `WorkSpaces -`.
- `license`: A boolean flag to enable/disable the MIT license showing on script start.
- `debug`: A boolean flag to enable/disable debugging in the logs.

#### Logging options:
---
- `logName`: A string defining the name of the log(s) file(s). The default name is `google_chrome` and  `amazon_workspaces`.
- `logFormat`: A string defining the format of the log(s) file(s). The default format is `log`.
- `logDateFormat`: A string defining the format of timestamps in logs. The default format is `dd/MM/yyyy HH:mm:ss`.
- `clearLogs`: A boolean flag to enable clearing of the log(s) file(s). This will clear the content inside of the log file(s).

#### Template options:
---
- `templateFolderNameRegular`: A string defining the name of the regular GoogleChrome template folder name.
- `templateFolderNameForced`: A string defining the name of the forced Google Chrome template folder name.
- `templateFolderName`: A string defining the name of Amazon WorkSpaces template folder name.

### Date Configuration
---

##### `yyyy`: This specifier represents the year portion of the date. It uses four digits to represent the year. For example, 2024.

##### `MM`: This specifier represents the month portion of the date. It uses two digits to represent the month, with leading zeros if necessary. For example, 05 represents May.

##### `dd`: This specifier represents the day portion of the date. It uses two digits to represent the day of the month, with leading zeros if necessary. For example, 23.

##### `HH`: This specifier represents the hour portion of the time in 24-hour format. It uses two digits to represent the hour, ranging from 00 to 23. For example, 14 represents 2 PM in 24-hour format.

<details>
<summary><b>More info on HH format</b></summary>

##### `HH` (24-hour format): When HH is used, it represents the hour portion of the time in a 24-hour format, where the hour is represented with two digits from 00 to 23. The HH specifier does not use AM/PM designators since it covers the full 24-hour range. Example: HH:mm:ss might output 14:30:00, representing 2:30 PM in 24-hour format.

##### `hh` (12-hour format): When hh is used, it represents the hour portion of the time in a 12-hour format, where the hour is represented with one or two digits from 1 to 12. The hh specifier is typically used alongside the tt specifier (AM/PM designator) to indicate whether the time is in the AM or PM. Example: hh:mm:ss tt might output 02:30:00 PM, representing 2:30 PM.
</details>



##### `mm`: This specifier represents the minute portion of the time. It uses two digits to represent the minutes, with leading zeros if necessary. For example, 30.

##### `ss`: This specifier represents the second portion of the time. It uses two digits to represent the seconds, with leading zeros if necessary. For example, 45.

##### `fff`: This specifier represents the millisecond portion of the time. It uses three digits to represent the milliseconds, with leading zeros if necessary. For example, 123.

##### `tt`: This specifier represents the AM/PM designator in a 12-hour time format. It is typically used alongside the hh specifier to indicate whether the time is in the AM or PM. For example, AM or PM.

### Examples

```json
"logDateFormat": "yyyy'/'MM'/'dd hh:mm:ss tt"
```
Output: <code>2024/06/29 03:19:30 p.m.</code>

```json
"logDateFormat": "MM/dd/yyyy HH:mm:ss"
```
Output: <code>06.29.2024 15:19:30</code>

```json
"logDateFormat": "dd-MM-yyyy HH:mm:ss"
```
Output: <code>29-06-2024 15:19:30</code>

### Numbered Version
---
`folderNumber`: Set this to `true` to enable automatic renaming of the folder based on the downloaded Chrome version. This action requires administrative privileges.

For example, if this option is enabled, the folders will be named as follows:

`false`:
```css
Chrome - VERSION_force_update
```

`true`:
```css
Chrome - 125.0.6422.113_force_update
```

The `folderNumber` configuration requires administrative privileges because the only way to obtain the Chrome version number is by installing the MSI file and retrieving the version from the Windows registry.

## Script Usage
### 1. Prepare the Environment:

Ensure that `config.json` is present in the same directory as the script.
Create the following template folders and populate them with necessary files:
- `Template\Chrome-Template`
- `Template\Chrome-Template-Forced`
- `Template\Amazon-Workspace-Template`

### 2. Downloading the Script:

You can download the script using `git clone` command. Follow these steps:

1. Open your terminal or command prompt.
2. Navigate to the directory where you want to download the script.
3. Run the following command:
```
git clone https://github.com/OlaYZen/MSI-Downloader.git
```

This command will clone the repository into your current directory.

### 3. Run the Script:

- Open PowerShell and navigate to the directory containing the script and config.json.
- Execute the script:
```css
& '.\Downloader.ps1' 
```

### 4. Monitor the Logs:

- Check `google_chrome.log` or `amazon_workspaces.log` in the script directory for detailed logs of the execution process, including any errors encountered.

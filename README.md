# MSI-Downloader

![MSI-Downloader Banner](./banner.png)

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Supported Programs](#supported-programs)
- [Configuration](#configuration)
  - [Configuration File Structure](#configuration-file-structure)
  - [Options](#options)
  - [Logging Options](#logging-options)
  - [Optional Extras](#optional-extras)
- [Date Configuration](#date-configuration)
  - [Format Specifiers](#format-specifiers)
  - [Examples](#examples)
- [Numbered Versioning](#numbered-versioning)
- [Dell Command | Update](#dell-command--update)
- [NTFY Integration](#ntfy-integration)
  - [Configuration](#configuration-1)
  - [Notification Triggers](#notification-triggers)
  - [Example Notification](#example-notification)
- [Usage](#usage)
  - [Installation](#installation)
  - [Running the Script](#running-the-script)
  - [Monitoring Logs](#monitoring-logs)
- [Arguments](#arguments)
- [Integration with Other Programs](#integration-with-other-programs)
- [Script Update](#script-update)
- [Recommendations](#recommendations)
- [License](#license)
- [Contributing](#contributing)

## Overview

**MSI-Downloader** is a PowerShell script designed to automate the downloading and organization of various software installers based on customizable configurations. Unlike tools like `winget`, MSI-Downloader focuses on downloading both 64-bit versions of specific applications and organizing them into designated folders. The script is user-friendly, easily configurable, and supports EXE files despite its name suggesting MSI files.

## Features

- **Automated Downloads**: Downloads and installs specified programs automatically.
- **Organized Storage**: Sorts downloaded programs into appropriate folders.
- **Version Management**: Renames folders to reflect the newest version of each program and removes outdated versions.
- **Customizable Intervals**: Supports timer intervals for scheduled operations.
- **Custom URLs**: Allows specifying custom download URLs for programs.
- **Logging**: Maintains detailed logs of all operations.
- **User-Friendly Configuration**: Easily configurable via a JSON file.

## Supported Programs

Currently, MSI-Downloader supports the following applications:

- Google Chrome
- Firefox
- Amazon Workspaces
- 7-Zip
- WinRAR
- Notepad++
- VLC Media Player
- Lenovo System Update
- Dell Command | Update
- Jabra Direct

*Additional programs may be supported in future releases.*

## Configuration

MSI-Downloader utilizes a `config.json` file located in the same directory as the script to manage its settings. This file allows users to customize download options, folder structures, logging preferences, and more.

### Configuration File Structure

Below is an example of the `config.json` structure:

```json
{
  "chrome": {
    "options": {
      "downloadRegular": true,
      "downloadForced": false,
      "folderNumber": false,
      "deleteExist": false,
      "folderName": "Chrome -",
      "Prefix": "VERSION",
      "forcedSuffix": "_force_update",
      "specificURL64": "",
      "specificURL32": ""
    },
    "template": {
      "templateFolderNameRegular": "Chrome-Template",
      "templateFolderNameForced": "Chrome-Template-Forced"
    }
  },
  "Firefox": {
    "options": {
      "download": false,
      "folderNumber": false,
      "deleteExist": false,
      "folderName": "Firefox -",
      "Prefix": "VERSION",
      "specificURL": ""
    },
    "template": {
      "templateFolderName": "Firefox-Template"
    }
  },
  // ... Additional program configurations ...
  "logging": {
    "logName": "Downloader",
    "logFormat": "log",
    "logDateFormat": "dd'/'MM'/'yyyy HH:mm:ss",
    "clearLogs": false
  },
  "ntfy": {
    "URL": "https://ntfy.yourdomain.com/your-topic",
    "Title": "",
    "Priority": "default",
    "Enabled": false
  },
  "license": true,
  "debug": false,
  "old": false
}
```

### Options

#### Chrome Specific

- `downloadRegular` (bool): Enable downloading and installing the regular version of Chrome.
- `downloadForced` (bool): Enable downloading and installing the forced update version of Chrome.
- `forcedSuffix` (string): Suffix for the forced update version (default: `_force_update`).
- `specificURL64` (string): Custom URL for the 64-bit version of Chrome. Leave empty for default.
- `specificURL32` (string): Custom URL for the 32-bit version of Chrome. Leave empty for default.
- `templateFolderNameRegular` (string): Name of the regular Chrome template folder.
- `templateFolderNameForced` (string): Name of the forced Chrome template folder.

#### Universal

- `download` (bool): Enable downloading and installing the application.
- `folderNumber` (bool): Enable automatic renaming of the folder to the newest version. **Requires administrative privileges**.
- `deleteExist` (bool): Delete existing folders upon script execution. **Caution: This will delete folders like Chrome, AWS, 7-Zip, etc. Ensure backups are in place if needed.**
- `folderName` (string): Base name for folders (e.g., `Chrome -`, `WorkSpaces -`).
- `Prefix` (string): Prefix for the application version (default: `VERSION`).
- `specificURL` (string): Custom URL for the application version. Leave empty for default.
- `templateFolderName` (string): Name of the application's template folder.

### Logging Options

- `logName` (string): Name of the log file(s) (default: `Downloader`).
- `logFormat` (string): Format of the log file(s) (default: `log`).
- `logDateFormat` (string): Timestamp format in logs (default: `dd'/'MM'/'yyyy HH:mm:ss`).
- `clearLogs` (bool): Enable clearing of log files' content on execution.

### Optional Extras

- `license` (bool): Display the MIT license upon script start.
- `debug` (bool): Enable debug mode for additional troubleshooting information.
- `old` (bool): Moves downloaded applications into the .Old folder. This can then be used to find out if a newer version of a program exists

## Date Configuration

Customize timestamp formats using the following specifiers:

| Specifier | Description |
|-----------|-------------|
| `yyyy`    | Four-digit year (e.g., 2024) |
| `MM`      | Two-digit month with leading zeros (e.g., 05 for May) |
| `dd`      | Two-digit day with leading zeros (e.g., 23) |
| `HH`      | Two-digit 24-hour format hour (00-23) |
| `hh`      | Two-digit 12-hour format hour (01-12) |
| `mm`      | Two-digit minutes with leading zeros (e.g., 30) |
| `ss`      | Two-digit seconds with leading zeros (e.g., 45) |
| `fff`     | Three-digit milliseconds with leading zeros (e.g., 123) |
| `tt`      | AM/PM designator for 12-hour format |

<details>
  <summary><strong>More on Hour Formats</strong></summary>
  
- **`HH` (24-hour format)**: Represents hours from 00 to 23 without AM/PM.
  
  *Example:* `14:30:00` represents 2:30 PM.
  
- **`hh` (12-hour format)**: Represents hours from 01 to 12 with AM/PM.
  
  *Example:* `02:30:00 PM` represents 2:30 PM.
</details>

### Examples

```json
"logDateFormat": "yyyy'/'MM'/'dd hh:mm:ss tt"
```
*Output:* 2024/06/29 03:19:30 PM

```json
"logDateFormat": "MM/dd/yyyy HH:mm:ss"
```
*Output:* 06.29.2024 15:19:30

```json
"logDateFormat": "dd-MM-yyyy HH:mm:ss"
```
*Output:* 29-06-2024 15:19:30


## Numbered Versioning

Enable automatic folder renaming based on the downloaded application's version by setting `folderNumber` to `true`. **Note:** This requires administrative privileges.

**Examples:**

- **Disabled (`folderNumber`: `false`):**
  ```plaintext
  Chrome - VERSION_force_update
  ```

- **Enabled (`folderNumber`: `true`):**
  ```plaintext
  Chrome - 129.0.6668.101_force_update
  ```

*Reason:* Retrieving the application's version number necessitates administrative access to install the MSI file and access the Windows registry.

## Dell Command | Update

The Dell Command | Update configuration leverages Python to bypass restrictions that prevent scripts from fetching their HTML files directly. The Python script mimics a browser to obtain the latest download URL. 

**Requirements:**

- **Python Installation:** Python must be installed manually (recommended version: 3.12.4).
- **Automatic Dependency Installation:** The script will automatically install required Python packages via the `requirements.txt` file. Manual installation of Python is required.

## NTFY Integration

MSI-Downloader integrates with [NTFY](https://ntfy.sh), a notification service that sends updates about script execution and detected new versions.

### Configuration

To enable NTFY notifications, update the `ntfy` section in `config.json`:

```json
"ntfy": {
    "URL": "https://ntfy.yourdomain.com/your-topic",
    "Title": "",
    "Priority": "default",
    "Enabled": false
}
```

- **URL**: Endpoint for NTFY notifications. Replace with your actual NTFY URL.
- **Title**: Customize the notification title. Leave empty for automatic titles.
- **Priority**: Set notification priority (`max`, `high`, `default`, `low`, `min`).
- **Enabled**: Set to `true` to activate notifications.

### Notification Triggers

Notifications are sent in the following scenarios:

- Detection of a new program version.
- Detection of a new script version.
- Execution errors (e.g., failed downloads or installations).

### Example Notification

When a new version is detected:

```
Title: NotepadPlusPlus - 8.7 | MSI-Downloader
Message: New version of NotepadPlusPlus - 8.7 detected
```

This ensures you stay informed about the script's activities and any critical updates or issues.

## Usage

### Installation

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/OlaYZen/MSI-Downloader.git
   ```

### Running the Script

1. **Open PowerShell:**

2. **Navigate to the downloaded folder:**

3. **Execute the Script:**

   ```powershell
   .\Downloader.ps1
   ```

### Monitoring Logs

- **Log File Location:** `Downloader.log` in the script directory.
- **Purpose:** Provides detailed logs of the execution process, including errors.

## Arguments

The script supports several command-line arguments to enhance its functionality and provide flexibility in its execution. Here are the available arguments:

![Arguments Image](./image.png)

These arguments allow you to control the script's behavior directly from the command line, making it easier to automate and integrate into other workflows.

## Integration with Other Programs

MSI-Downloader can be integrated with various other programs and scripts to enhance its functionality and automate workflows. Here are some ideas and examples of how you can integrate MSI-Downloader with other tools:

### PowerShell Integration

You can call the MSI-Downloader script from another PowerShell script to automate the installation of multiple applications.
```powershell
.\Downloader.ps1 -p "SevenZip" -c "deleteExist folderNumber"
```

the names is the same as they are in `config.json`
NotepadPlusPlus, amazonWorkspace etc.

### Python Integration

You can also integrate MSI-Downloader with Python scripts. This allows you to leverage Python's capabilities for data processing, web scraping, or even creating a GUI for user interaction.

Here’s an example of a Python script that triggers the MSI-Downloader to install applications based on a configuration file:

```python
import subprocess
import json

# Load configuration from a JSON file
with open('config.json') as config_file:
    config = json.load(config_file)

# List of applications to install based on the config
apps_to_install = [app for app, details in config.items() if details['options']['download']]

for app in apps_to_install:
    print(f"Installing {app}...")
    subprocess.run(["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", "Downloader.ps1", "-p", app, "-c", "deleteExist folderNumber"])
```
<sub>Note: The script above was AI generated and has not been tested. Proceed with caution.</sub>

### Scheduled Tasks

You can set up a scheduled task in Windows to run the MSI-Downloader script at specific intervals. This is useful for keeping applications up to date automatically.

1. Open Task Scheduler.
2. Create a new task and set the trigger to your desired schedule (e.g., daily, weekly).
3. In the "Actions" tab, set the action to start a program and point it to PowerShell with the arguments to run your script:

```plaintext
Program/script: powershell.exe
Add arguments: -ExecutionPolicy Bypass -File "C:\Path\To\Downloader.ps1"
```
<sub>Note: The script above was AI generated and has not been tested. Proceed with caution.</sub>

### Web Integration

If you have a web application, you can create an API endpoint that triggers the MSI-Downloader script. This can be done using a web framework like Flask in Python. 

```python
from flask import Flask
import subprocess

app = Flask(__name__)

@app.route('/install/<app_name>', methods=['POST'])
def install_app(app_name):
    subprocess.run(["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", "Downloader.ps1", "-p", app_name])
    return f"Triggered installation for {app_name}"

if __name__ == '__main__':
    app.run(port=5000)
```
<sub>Note: The script above was AI generated and has not been tested. Proceed with caution.</sub>


### Conclusion

Integrating MSI-Downloader with other programs can significantly enhance your automation capabilities. Whether through PowerShell, Python, or web applications, you can create a seamless experience for managing software installations. Feel free to customize the examples above to fit your specific needs and workflows.

## Script Update

To update the script to the latest version from GitHub:

```powershell
.\Downloader.ps1 -Update
```

**Caution:** This command currently only updates the `Downloader.ps1` script, which may potentially break functionality. For optimal results, manually upgrade to ensure the latest `config.json` and template folders are included.

## Recommendations

- **PowerShell Version:** It is recommended to use **PowerShell 7** or **Windows PowerShell ISE** for optimal performance.
  
  - **Reason:** PowerShell 5 has significantly slower download speeds compared to PowerShell 7 or ISE. This improvement is due to the more efficient `Invoke-RestMethod` cmdlet in newer versions, which facilitates faster HTTP and HTTPS requests essential for downloading content from the web.

## License

This project is licensed under the [MIT License](LICENSE).

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request for any enhancements or bug fixes.

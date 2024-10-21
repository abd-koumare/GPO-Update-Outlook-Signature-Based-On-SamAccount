# Outlook Signature Automation Script

## Description
This PowerShell script automates the process of setting up email signatures in Microsoft Outlook by converting user-specific signature images into embedded base64 format and configuring the appropriate registry settings. The script supports Outlook 2010, 2016, and 2021.

## Features
- Automatically retrieves current username
- Converts signature images to base64 for reliable email embedding
- Creates necessary Outlook signature directories
- Configures registry settings for default signatures
- Tracks signature deployment progress
- Supports multiple Outlook versions (2010, 2016, 2021)

## Prerequisites
- Windows operating system
- PowerShell execution permissions
- Network access to shared signature directory
- Appropriate read/write permissions
- Microsoft Outlook installation (2010, 2016, or 2021)

## Configuration
Update the following variables at the beginning of the script:
```powershell
$publicAvailableUNC = "\\<ip_address>\netlogon\<signature_img_directory>\"
$signatureName = "Octobre-2024"
$imageFileExtension = ".png"
```

## Directory Structure
The script expects/creates the following directory structure:
- Network Share: `\\<ip_address>\netlogon\<signature_img_directory>\`
  - Individual signature images: `username.png`
  - Progress tracking: `Progression\`
- Local: `%APPDATA%\Microsoft\Signatures\`

## Usage
1. Ensure your signature image exists in the network share with your username as the filename
2. Run the script with appropriate permissions
```powershell
.\Set-OutlookSignature.ps1
```

## Error Handling
The script includes error handling for common scenarios:
- Missing signature directories
- File access issues
- Registry modification failures

## Version History
- 2.0 (15/10/2024)
  - Fixed signature image embedding using base64 conversion
  - Improved error handling
  - Added progress tracking

## Author
- Author: abd.koumare@gmail.com

## Notes
- Signature images should be in PNG format
- Image height is set to 353px by default
- The script automatically tracks successful deployments in the "Progression" folder

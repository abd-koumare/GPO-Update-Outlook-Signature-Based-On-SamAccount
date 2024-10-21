
<#
  Version:       2.0
  Author:        abd.koumare@gmail.com
  Updated:       15/10/2024

  Issue Fixed: Signature is not include when email is sent fix by using embed image in img src tag value 
#>


$publicAvailableUNC = "\\<ip_address>\netlogon\<signature_img_directory>\"


$username = ((Get-CimInstance Win32_ComputerSystem).username -split '\\')[1]
$signatureName = "Octobre-2024"
$imageFileExtension = ".png"


$outlookSignaturePath = Join-Path -Path $Env:appdata -ChildPath 'Microsoft\Signatures\'

$userSignatureUNC = $publicAvailableUNC + $username + $imageFileExtension


# Progress folder 
$signatureHasBeenSetFolderName = Join-Path -Path $publicAvailableUNC -ChildPath "Progression" 

$signatureFilePath = Join-Path -Path $outlookSignaturePath -ChildPath $signatureName




if( Test-Path -Path $userSignatureUNC ) {

    $bytes = [System.IO.File]::ReadAllBytes($userSignatureUNC)
    $base64String = [Convert]::ToBase64String($bytes)

    $dataUrl = "data:image/png;base64,$base64String"

    

    if (-not (Test-Path -Path $outlookSignaturePath)) {
        try {
            New-Item -ItemType directory -Path $outlookSignaturePath
        }
        catch {
            Write-Host "Error: Unable to create the signatures folder. Details: $($_.Exception.Message)"
            exit
        }
       
    }


    Copy-Item -Path $userSignatureUNC -Destination $outlookSignaturePath -Force



# Building HTML
$signature = 
@"

    <html xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:w="urn:schemas-microsoft-com:office:word"
    xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
    xmlns="http://www.w3.org/TR/REC-html40">

    <head>
    <meta http-equiv=Content-Type content="text/html; charset=windows-1252">
    <meta name=ProgId content=Word.Document>
    <meta name=Generator content="Microsoft Word 15">
    <meta name=Originator content="Microsoft Word 15">
    </head>

    <body lang=FR>

    <p class="signature-img-wrapper">
    <img src="$dataUrl" alt="$($username)" width="auto" height="353px"/>
    </p>

    </body>
    </html>
"@
  
    # Save the HTML to the signature file
    try {
        $signature | Out-File -FilePath "$signatureFilePath.htm" -Encoding ascii
    }
    catch {
        Write-Host "Error: Unable to save the HTML signature file. Details: $($_.Exception.Message)"
        exit
    }


    # Setting the regkeys for Outlook 2016 and 2021
    if (test-path "HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General") {
        get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name disablesignatures -value 0 -propertytype DWord -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value $signatureName -propertytype string -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value $signatureName -propertytype string -force
        Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook\\Setup -Name "First-Run" -ErrorAction silentlycontinue
    }

    # Setting the regkeys for Outlook 2010
    if (test-path "HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General") {
        get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name disablesignatures -value 0 -propertytype DWord -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ MailSettings | new-Itemproperty -name NewSignature -value $signatureName -propertytype string -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ MailSettings | new-Itemproperty -name ReplySignature -value $signatureName -propertytype string -force
        Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\14.0\\Outlook\\Setup -Name "First-Run" -ErrorAction silentlycontinue
    }


     # If the folder does not exist create it
     if (-not (Test-Path -Path $signatureHasBeenSetFolderName)) {
        try {
            New-Item -ItemType directory -Path $signatureHasBeenSetFolderName
        }
        catch {
            Write-Host "Error: Unable to create the $signatureHasBeenSetFolderName folder. Details: $($_.Exception.Message)"
            exit
        }
    }

    Copy-Item -Path $userSignatureUNC -Destination $signatureHasBeenSetFolderName -Force

} else {
    echo "Signature image file has not been found !"
}

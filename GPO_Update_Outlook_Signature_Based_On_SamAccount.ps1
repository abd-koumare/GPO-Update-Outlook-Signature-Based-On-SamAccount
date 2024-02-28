
<#
  Version: 1.0
  Author : abdloulayek@embross.com or abd.koumare@gmail.com
  Created: 25/02/2024

  Roaming signature RegKey: DisableRoamingSignaturesTemporaryToggle Path HKEY_CURRENT_USER\Software\Microsoft\Office.<16.0 or 8.0>\Outlook\Setup\
#>

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[System.Windows.Forms.MessageBox]::Show('Signature Update Is Running ...', 'WARNING')

try {
    $user = (([adsisearcher]"(&(objectCategory=User)(samaccountname=$env:username))").FindOne().Properties)
}
catch {
    Write-Host "Error: Unable to query Active Directory for user information. Details: $($_.Exception.Message)"
    exit
}

  
# Create the signatures folder and sets the name of the signature file
$msOutlookSignatureLocation = Join-Path -Path $Env:appdata -ChildPath 'Microsoft\Signatures\'
$signatureFileName = 'ORGANISATION-SIGNATURE'
$signatureFilePath = Join-Path -Path $msOutlookSignatureLocation -ChildPath $signatureFileName

$publicAvailableUNC = ""
$imageFileExtension = ".png"
$userSignatureImageRemoteFile = $publicAvailableUNC + $user.samaccountname + $imageFileExtension


if( Test-Path -Path $userSignatureImageRemoteFile) {


    # If the folder does not exist create it
    if (-not (Test-Path -Path $msOutlookSignatureLocation)) {
        try {
            New-Item -ItemType directory -Path $msOutlookSignatureLocation
        }
        catch {
            Write-Host "Error: Unable to create the signatures folder. Details: $($_.Exception.Message)"
            exit
        }
    }

    Copy-Item -Path $userSignatureImageRemoteFile -Destination $msOutlookSignatureLocation -Force
    $imgFileName = $user.samaccountname + $imageFileExtension

# Building Style Sheet
$style = 
@"
<style>
.signature-img-wrapper img {
 width: auto;
 height: 353px;
}
.signature-img-wrapper {
    text-align: left;
}
</style>
"@


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
 <img src="$($user.samaccountname)$($imageFileExtension)" alt="$($user.samaccountname)"/>
</p>

</body>
"@

    # Save the HTML to the signature file
    try {
        $style + $signature | Out-File -FilePath "$signatureFilePath.htm" -Encoding ascii
    }
    catch {
        Write-Host "Error: Unable to save the HTML signature file. Details: $($_.Exception.Message)"
        exit
    }


    # Setting the regkeys for Outlook 2016
    if (test-path "HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General") {
        get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name disablesignatures -value 0 -propertytype DWord -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value $signatureFileName -propertytype string -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value $signatureFileName -propertytype string -force
        Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook\\Setup -Name "First-Run" -ErrorAction silentlycontinue
    }

    # Setting the regkeys for Outlook 2010 - Thank you AJWhite1970 for the 2010 registry keys
    if (test-path "HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General") {
        get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name disablesignatures -value 0 -propertytype DWord -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ MailSettings | new-Itemproperty -name NewSignature -value $signatureFileName -propertytype string -force
        get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ MailSettings | new-Itemproperty -name ReplySignature -value $signatureFileName -propertytype string -force
        Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\14.0\\Outlook\\Setup -Name "First-Run" -ErrorAction silentlycontinue
    }

}

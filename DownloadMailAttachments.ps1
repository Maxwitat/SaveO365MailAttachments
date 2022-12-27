﻿# Frank Maxwitat
# Version 1.0, Dec 2022

#------------------Begin Function------------------------------------------------------
function LogWrite {
    Param(
          [parameter(Mandatory=$true)]
          [String]$LogFilePath,

          [parameter(Mandatory=$true)]
          [String]$LogTxt,

          [Parameter(Mandatory=$true)]
          [String]$Severity,

          [Parameter(Mandatory=$false)]
          [String]$WithLog = $true
    )
#--------------End Function---------------------------------------------------------------------------

#--------------Initialization-------------------------------------------------------------------------
$LogFilePath = $psscriptroot + "\DownloadMailAttachments.log"
LogWrite -LogFilePath $LogFilePath -LogTxt "Starting DownloadMailAttachment version 1.0" -Severity 'I'
LogWrite -LogFilePath $LogFilePath -LogTxt  ("LogFile: " + $LogFilePath) -Severity 'I' -WithLog $false

$DownloadPath = $psscriptroot + "\Downloads"
if(!(Test-Path $DownloadPath)) 
{
    Try{
        New-Item -ItemType Directory -Path $DownloadPath
        LogWrite -LogFilePath $LogFilePath -LogTxt ("Create Folder: Downloads go to " + $DownloadPath) -Severity I
    }
    Catch
    {
        Write-Error ($_ | Out-String)
        LogWrite -LogFilePath $LogFilePath -LogTxt ($_ | Out-String) -Component  -Severity E    }
    }
else 
{
    LogWrite -LogFilePath $LogFilePath -LogTxt ("Downloads go to " + $DownloadPath) -Severity I
}

$TenantId    = '' #ENTER YOUR TENANT ID HERE "12345678-1234-1234-1234-123456789012"
$ClientId    = '' #ENTER YOUR CLIENT ID HERE "12345678-1234-1234-1234-123456789012"

$mailUser = '' #username@happyadmin.com


$thumbPrint = "" #ENTER THE THUMBPRINT OF YOUR CERTIFICATE HERE

#--------------End Initialization----------------------------------------------------------------------

# SaveMailAttachementsTo
Import-Module -Name MSAL.PS -Force

# Use TLS 1.2 connection (Server OS don't use TLS1.2 by default)
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($thumbPrint)"
Try{
    $token = Get-MsalToken -ClientId $clientID -TenantId $tenantID -ClientCertificate $ClientCertificate
}
Catch{
    LogWrite -LogFilePath $LogFilePath -LogTxt $_ -Severity E
    LogWrite -LogFilePath $LogFilePath -LogTxt 'Are you running the script with admin rights?' -Severity E 
}

# Inspect the Access Token using JWTDetails PowerShell Module
$accessToken = $token.AccessToken

$uri = "https://graph.microsoft.com/v1.0/users/$mailUser/mailFolders/deleteditems"
$del = Invoke-RestMethod -Uri $uri -Headers @{Authorization=("bearer {0}" -f $accessToken)}
$deleteditemsfolderid = $del.id

$url = "https://graph.microsoft.com/v1.0/users/$mailUser/mailFolders/Inbox/messages"
$messagequery = $url + "?' $select-Id&'$filter=HasAttachments eq true"
$messages = Invoke-RestMethod $messagequery -Headers @{Authorization=("bearer {0}" -f $accessToken)}

foreach ($message in $messages.value)
{
$query = $url + "/" + $message.id + "/attachments"
$attachments = Invoke-RestMethod $query -Headers @{Authorization=("bearer {0}" -f $accessToken)}

foreach($attachment in $attachments.value)
{    
    $attachment.Name
        
    $path = $DownloadPath + "\"+ $attachment.Name
    LogWrite -LogFilePath $LogFilePath -LogTxt ("Downloading " + $attachment.Name) -Severity I

    $content = [System.Convert]::FromBase64String($attachment.ContentBytes)
    Set-Content -Path $path -Value $content -Encoding Byte
    }

    $query = $url + "/" + $message.id + "/move"

    $body = "{""DestinationId"": ""$deleteditemsfolderid""}"

    Invoke-RestMethod $query -Body $body -ContentType "Application/json" -Method Post -Headers @{Authorization=("bearer {0}" -f $accessToken)}    
}

LogWrite -LogFilePath $LogFilePath -LogTxt "Ending DownloadMailAttachment version 1.0" -Severity I
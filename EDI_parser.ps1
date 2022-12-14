#read settings
$settings = (get-content (Join-Path -Path $PSScriptRoot -ChildPath "settings.json") | ConvertFrom-json)

add-type -path ((join-path -path (join-path -Path $PSScriptRoot -ChildPath "lib") -ChildPath "OpenPop.dll"))
Import-Module Transferetto -Force
Import-Module ((join-path -path (join-path -Path $PSScriptRoot -ChildPath "lib") -ChildPath "source.psm1"))

# remove everything from folders
Get-ChildItem $settings.inboxFolder | Remove-Item -Force -Recurse
Get-ChildItem $settings.outboxFolder | Remove-Item -Force -Recurse

# first we connect to POP3 and store it to variable
$pop3Client = Connect-Mail -Server $settings.mailServer -Port $settings.mailPort -Username $settings.mailUsername -Password $settings.mailPassword

# now we check to see if there is any mail for us
$mails = Check-Mail $pop3Client -From $settings.mailTargetfrom

# if none - exit
if ( ($mails | Where-Object {$_.target -eq $true}).count -eq 0 ) { write-host "No matching emails found. Goodbye!" ; write-host "Will exit in 10 seconds..." ; start-sleep 10; exit 0 }

# proceed mails, if $_.target -eq $true - download and mark for deletion, else mark for deletion
foreach ($mail in $mails) { 
  if ($mail.target -eq $true) {
    FetchAndSave-Attachment -pop3Client $pop3Client -Folder $settings.inboxFolder -messageIndex $mail.index  
    $pop3Client.DeleteMessage($mail.index)
  } else {
    $pop3Client.DeleteMessage($mail.index)
  }
}
# done with mailbox
$pop3Client.dispose()

$incomingFiles = (Get-ChildItem $settings.inboxFolder)

# proceed files - move, parse, download and shit
foreach ($incomingFile in $incomingFiles) {
  $incomingFileBaseName = ($incomingFile.BaseName -split "-_-")[1]
  #$incomingFileTimeStamp = ($incomingFile.BaseName -split "-_-")[0]
  $source = (Parse-HTML $incomingFile.FullName)
  $HTMLheader = $source[0]
  $HTMLdate = $source[1]
  $sourceData = $source[2]
  $folder = ( (Join-Path -Path $settings.outboxFolder -ChildPath "$incomingFileBasename-$HTMLdate")  )
  mkdir $folder
  Download-Pages -sourceData $sourceData -Folder $folder -Prefix $incomingFileBaseName -HTMLdate $HTMLdate -HTMLheader $HTMLheader
  Copy-Item $settings.cssFile $folder
} 

# connect to SFTP
$sftpClient = Connect-SFTP -Server $settings.sftpServer -Port $settings.sftpPort -Verbose -Username $settings.sftpUsername -Password $settings.sftpPassword

# Create folders and upload files
foreach ($folder in (Get-ChildItem $settings.outboxFolder)) {
  $sftpFolder = $settings.sftpParentFolder + $folder.BaseName
  $sftpClient.CreateDirectory("$sftpFolder")
  foreach ($file in (Get-ChildItem $folder.FullName)){
    Send-SFTPFile -SftpClient $sftpClient -LocalPath $file.FullName -RemotePath "$sftpFolder/$($file.name)" -AllowOverride
  }
}

Disconnect-SFTP $sftpClient

Remove-Module source 

write-host "will exit in 10 seconds..."
start-sleep 10

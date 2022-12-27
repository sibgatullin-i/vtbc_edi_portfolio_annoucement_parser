add-type -path ((join-path -path (join-path -Path $PSScriptRoot -ChildPath "lib") -ChildPath "OpenPop.dll"))
Import-Module Transferetto -Force
Import-Module (join-path -Path $PSScriptRoot -ChildPath "source.psm1")
#Import-Module (join-path -Path $PSScriptRoot -ChildPath "telegram.psm1")

$inboxFolder = (Join-Path -Path $PSScriptRoot -ChildPath "inbox")
$outboxFolder = (Join-Path -Path $PSScriptRoot -ChildPath "outbox")
$cssFile = (Join-Path -path (Join-Path -Path $PSScriptRoot -ChildPath ./lib) -ChildPath "style.css")
$sftpServer = (read-host "sftpServer")
$sftpPort = 8022
$sftpUsername = (read-host "sftpusername")
$sftpPassword = (read-host "sftppassword")
$sftpParentFolder = "/sftp/EDI/"
$mailServer = "outlook.office365.com"
$mailPort = "995"
$mailUsername = (read-host "mailusername")
$mailPassword = (read-host "mailpassword")

# remove everything from folders
Get-ChildItem $inboxFolder | Remove-Item -Force -Recurse
Get-ChildItem $outboxFolder | Remove-Item -Force -Recurse


# first we connect to POP3 and store it to variable
$pop3Client = Connect-Mail -Server $mailServer -Port $mailPort -Username $mailUsername -Password $mailPassword

# now we check to see if there is any mail for us
$mails = Check-Mail $pop3Client

# if none - exit
if ( ($mails | Where-Object {$_.target -eq $true}).count -eq 0 ) { write-host "No matching emails found. Goodbye!" ; exit 0 } #else {Send-TelegramMessage "got some mail"}

# proceed mails, if $_.target -eq $true - download and mark for deletion, else mark for deletion
foreach ($mail in $mails) { 
  if ($mail.target -eq $true) {
    FetchAndSave-Attachment -pop3Client $pop3Client -Folder $inboxFolder -messageIndex $mail.index  
    $pop3Client.DeleteMessage($mail.index)
  } else {
    $pop3Client.DeleteMessage($mail.index)
  }
}
# done with mailbox
$pop3Client.dispose()

$incomingFiles = (Get-ChildItem $inboxFolder)

# proceed files - move, parse, download and shit
foreach ($incomingFile in $incomingFiles) {
  $incomingFileBaseName = ($incomingFile.BaseName -split "-_-")[1]
  $incomingFileTimeStamp = ($incomingFile.BaseName -split "-_-")[0]
  $folder = ( (Join-Path -Path $outboxFolder -ChildPath "$incomingFileBasename-$incomingFileTimeStamp")  )
  mkdir $folder
  $sourceData = (Parse-HTML $incomingFile)
  Download-Pages -sourceData $sourceData -Folder $folder -Prefix $incomingFileBaseName
  Copy-Item $cssFile $folder
} 

# connect to SFTP
$sftpClient = Connect-SFTP -Server $sftpServer -Port $sftpPort -Verbose -Username $sftpUsername -Password $sftpPassword

# Create folders and upload files
foreach ($folder in $(Get-ChildItem $outboxFolder)) {
  $sftpFolder = $sftpParentFolder + $folder.BaseName
  $sftpClient.CreateDirectory("$sftpFolder")
  foreach ($file in (Get-ChildItem $folder)){
    Send-SFTPFile -SftpClient $sftpClient -LocalPath $file.FullName -RemotePath "$sftpFolder/$($file.name)" -AllowOverride
  }
}

Disconnect-SFTP $sftpClient

Remove-Module source 


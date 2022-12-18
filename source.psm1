
function Download-Pages {
  param(
    [Parameter(Mandatory)][string]$Filepath,
    [Parameter(Mandatory)][string]$Folder
  )
  $incomingFile = get-childitem $Filepath
  $array = (get-content $incomingFile | ConvertFrom-Json)
  foreach ($item in $array) {
    $page = (Invoke-WebRequest -UseBasicParsing $item.Url).Content
    $newName = $incomingFile.BaseName + '_' + (get-date -Format 'yyyyMMdd_hhMMssffff') + '.html'
    $newPath = Join-Path -Path $Folder -ChildPath $newName
    $page = $page -replace "(?s)<script.+?</script>", ""
    $page = $page -replace "(?s)<style.+?</style>", "<style type='text/css'>@import url('./style.css');</style>"
    Set-Content -Path $newPath -Value $page
    $item.EventID = "<a href='./" + $newName + "'>" + $item.EventID + "</a>" 
    $item.Url = "<a href='./" + $newName + "'></a>"
  }
  $array
  $array | select-object -ExcludeProperty Url| convertto-html -CssUri "./style.css" | % {$_ -replace "&#39;","'" -replace '&lt;','<' -replace '&gt;','>'} | Out-File (Join-Path -Path $folder -ChildPath ("index_" + $incomingFile.BaseName + (get-date -Format 'yyyyMMdd_hhMMssffff') + '.html'))
  #$array | ConvertTo-Json | Out-File $incomingFile
}

function Parse-HTML {
  param(
    [Parameter(Mandatory)][string]$Path
  )
  if (!(Test-Path $path)) {write-warning "File $Path not found"; return $false}
  $HTML = (Get-Content $Path | ConvertFrom-Html)
  $HTML = $HTML.SelectNodes('//table') | where-object {$_.InnerText -like "event*"}
  $HTML = $HTML.SelectNodes('tr')
  $tableHeader = $HTML[0].SelectNodes('td').InnerText
  $HTML = $HTML | where-object {$_.InnerText -notlike "event*"}
  $array = @()
  foreach ($line in $HTML) {
    $array += [pscustomobject]@{
      $tableHeader[0] = $line.SelectNodes('td')[0].InnerText
      $tableHeader[1] = $line.SelectNodes('td')[1].InnerText
      $tableHeader[2] = $line.SelectNodes('td')[2].InnerText
      $tableHeader[3] = $line.SelectNodes('td')[3].InnerText
      $tableHeader[4] = $line.SelectNodes('td')[4].InnerText
      $tableHeader[5] = $line.SelectNodes('td')[5].InnerText
      $tableHeader[6] = $line.SelectNodes('td')[6].InnerText
      $tableHeader[7] = $line.SelectNodes('td')[7].InnerText
      $tableHeader[8] = $line.SelectNodes('td')[8].InnerText
      Url = ($line.InnerHtml -split "'" | Where-Object {$_ -like "http*"})
    }
  }
  Return $array

}

function Get-Mail {
  param(
    [Parameter(Mandatory)][string]$Server,
    [Parameter(Mandatory)][string]$Port,
    [Parameter(Mandatory)][string]$Username,
    [Parameter(Mandatory)][string]$Password,
    [Parameter(Mandatory)][string]$enableSSL = $true
  )
  $pop3Client = New-Object OpenPop.Pop3.Pop3Client
  $pop3Client.connect( $server, $port, $enableSSL )

  if ( !$pop3Client.connected )
    {
      throw "Unable to create POP3 client. Connection failed with server $server"
    }
  $pop3Client.authenticate( $username, $password )

  
}
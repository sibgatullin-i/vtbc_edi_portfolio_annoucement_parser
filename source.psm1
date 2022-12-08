

function Download-Pages2 {
  param(
    [Parameter(Mandatory)][string]$filePath,
    [Parameter(Mandatory)][string]$folder
    )
  $incomingFile = get-childitem $filePath
  $URLs = @()
  $rawURLs = get-content $incomingFile
  $rawURLs = $rawURLs -split "\'"
  $rawURLs = $rawURLs | where-object {$_ -match "https://*"}
  foreach ($rawURL in $rawURLs) {
    $URLDate = (get-date -format "yyyyMMdd_hhMMssffff")
    $URLId = ($rawURL -split "&" | where-object {$_ -match "eventid=*"} | % {$_ -replace "eventid=", ""})
    $URLExchgcd = ($rawURL -split "&" | where-object {$_ -match "exchgcd=*"} | % {$_ -replace "exchgcd=", ""})
    $URLSedol = ($rawURL -split "&" | where-object {$_ -match "sedol=*"} | % {$_ -replace "sedol=", ""})
    $URLs += [pscustomobject]@{
      Date = $URLDate
      ID = $URLId
      URL = $rawURL
      Filename = ($incomingFile.Basename + "_" + $URLExchgcd + "_" + $URLSedol + "_" + $URLId + "_" + $URLDate + ".html")
    }
  }
  foreach ($URL in $URLs) {
    $file = join-path -path $folder -childpath $URL.Filename
    if (test-path $file) {write-warning "Overwriting $file"}
    $html = (invoke-webrequest -UseBasicParsing $URL.URL).content
    $html = $html -replace "(?s)<script.+?</script>", ""
    $html = $html -replace "(?s)<style.+?</style>", ""
    set-content -path $file -value $html
  }
}

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
    $page = $page -replace "(?s)<style.+?</style>", ""
    Set-Content -Path $newPath -Value $page
    $item.Url = "./" + $newName
  }
  $array
  $array | convertto-html | Out-File (Join-Path -Path $folder -ChildPath ("index_" + $incomingFile.BaseName + (get-date -Format 'yyyyMMdd_hhMMssffff') + '.html'))
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
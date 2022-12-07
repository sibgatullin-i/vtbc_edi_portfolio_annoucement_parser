

function Download-Pages {
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
    $URLDate = (get-date -format yyyyMMdd)
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
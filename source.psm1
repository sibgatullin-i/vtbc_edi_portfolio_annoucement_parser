

function Download-Pages {
  param(
    [Parameter(Mandatory)][string]$filePath
    )
  $URLs = @()
  $rawURLs = get-content $filePath
  $rawURLs = $rawURLs -split "\'"
  $rawURLs = $rawURLs | where-object {$_ -match "https://*"}
  foreach ($rawURL in $rawURLs) {
    $URLDate = (get-date -format yyMMdd)
    $URLId = ($rawURL -split "&" | where-object {$_ -match "eventid=*"} | % {$_ -replace "eventid=", ""})
    $URLs += [pscustomobject]@{
      Date = $URLDate
      ID = $URLId
      URL = $rawURL
      Filename = ("$URLId_$URLDate.html")
    }
  }
  echo $URLs
}
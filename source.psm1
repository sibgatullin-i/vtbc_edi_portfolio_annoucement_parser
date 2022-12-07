

function Download-Pages {
  param(
    [Parameter(Mandatory)][string]$filePath
    )
  $URLs = @()
  $rawURLs = get-content $filePath
  $rawURLs = $rawURLs -split "\'"
  $rawURLs = $rawURLs | where-object {$_ -match "https://*"}
  foreach ($rawURL in $rawURLs) {
    $URLs += [pscustomobject]@{
      Date = (get-date -format yyMMdd)
      ID = ($rawURL -split "&" | where-object {$_ -match "eventid=*"} | % {$_ -replace "eventid=", ""})
    }
  }
  echo URLs
}
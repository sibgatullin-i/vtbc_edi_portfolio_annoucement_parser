

function Download-Pages {
  param(
    [Parameter(Mandatory)][string]$filePath,
    [Parameter(Mandatory)][string]$folder
    )
  $URLs = @()
  $rawURLs = get-content $filePath
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
      Filename = ($URLExchgcd + "_" + $URLSedol + "_" + $URLId + "_" + $URLDate + ".html")
    }
  }
  foreach ($URL in $URLs) {
    $file = join-path -path $folder -childpath $URL.Filename
    $html = (invoke-webrequest -UseBasicParsing $URL.URL).content
    $html = $html -replace "(?s)<script.+?</script>",""
    set-content -path $file -value $html
  }
}
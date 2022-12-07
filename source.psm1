

function Download-Pages {
  ([Parameter(Mandatory)][string]$fileName)
  $URLs = get-content $fileName
  $URLs = $URLs -split "\'"
  $URLs = $URLs | where-object {$_ -match "https://*"}
  echo $URLs
}
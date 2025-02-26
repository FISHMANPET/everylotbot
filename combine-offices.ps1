$type = "offices"
$output = "$type.csv"
$sum = 0
$locations = @()
foreach ($i in 0..9) {
  $locations_temp = Import-Csv -Path "$type$i.csv"
  $locations += $locations_temp
  $sum += $locations_temp.Length
  Write-Output $locations.Length
}

$locations | Export-Csv -Path $output
Write-Output $sum

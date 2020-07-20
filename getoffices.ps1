$offices = [System.Collections.ArrayList]@()
$log = ".\zip.log"
$csv = ".\offices.csv"
$null | set-content -path $log
$null | set-content -path $csv
foreach ($zip in (16040..99950)) {
    $zipcode = $zip.ToString("00000")
    $zipcode | Add-Content -Path $log
    $locations = Invoke-RestMethod -Uri "https://tools.usps.com/UspsToolsRestServices/rest/POLocator/findLocations" `
    -Method "POST" `
    -Headers @{
    "method"           = "POST"
    "authority"        = "tools.usps.com"
    "scheme"           = "https"
    "path"             = "/UspsToolsRestServices/rest/POLocator/findLocations"
    "accept"           = "application/json, text/javascript, */*; q=0.01"
    "x-requested-with" = "XMLHttpRequest"
    "user-agent"       = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36"
    "origin"           = "https://tools.usps.com"
    "sec-fetch-site"   = "same-origin"
    "sec-fetch-mode"   = "cors"
    "sec-fetch-dest"   = "empty"
    "referer"          = "https://tools.usps.com/find-location.htm"
    "accept-encoding"  = "gzip, deflate, br"
    "accept-language"  = "en-US,en;q=0.9"
} `
    -ContentType "application/json;charset=UTF-8" `
    -Body "{`"maxDistance`":`"100`",`"lbro`":`"`",`"requestType`":`"po`",`"requestServices`":`"`",`"requestRefineTypes`":`"`",`"requestRefineHours`":`"`",`"requestZipCode`":`"$zipcode`",`"requestZipPlusFour`":`"`"}"
    $location = $locations.locations | Where-Object zip5 -eq $zipcode
    foreach ($po in $location) {
        $null = $offices.Add($po)
        Write-Host "$($po.zip5) in $($po.city), $($po.state)"
        $po | Export-Csv -NoTypeInformation -Append -Path $csv -Force
    }
}

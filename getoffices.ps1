[CmdletBinding()]
param (
    [Parameter()]
    [int[]]
    $range = 0..9
)
$ErrorActionPreference = 'Stop'
foreach ($i in $range) {
    $offices = [System.Collections.ArrayList]@()
    $log = ".\zip$i.log"
    $csvtemp = ".\boxes$i-temp.csv"
    $csv = ".\boxes$i.csv"
    $xml = ".\boxes$i.xml"
    $null | set-content -path $log
    $null | set-content -path $csv
    $null | set-content -path $xml
    $null | set-content -path $csvtemp
    switch ($i) {
        0 { $zips = (00501..09999) }
        9 { $zips = (90000..99950)}
        Default { $zips = (($i * 10000)..($i * 10000 + 9999)) }
    }
    foreach ($zip in $zips) {
        do {
            $found = $false
            $zipcode = $zip.ToString("00000")
            $zipcode | Add-Content -Path $log
            try {
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
                    -Body "{`"maxDistance`":`"100`",`"lbro`":`"`",`"requestType`":`"collectionbox`",`"requestServices`":`"`",`"requestRefineTypes`":`"`",`"requestRefineHours`":`"`",`"requestZipCode`":`"$zipcode`",`"requestZipPlusFour`":`"`"}"
                $locations | Add-Content -Path $log
                if ($locations.PSObject.properties.name -notcontains 'locations') {
                    if ($locations.errorcode -in ('800412df', '800412fd')) {
                        $found = $true
                    } elseif ($locations.psobject.Properties.name -notcontains 'errorcode') {
                        $found = $false
                    } else {
                        throw "uncaught error code $locations"
                    }
                } else {
                    $location = $locations.locations | Where-Object zip5 -eq $zipcode
                    foreach ($po in $location) {
                        $null = $offices.Add($po)
                        Write-Host "$($po.zip5) in $($po.city), $($po.state)"
                        $po | Export-Csv -NoTypeInformation -Append -Path $csvtemp -Force
                    }
                    $found = $true
                }
            } catch {
                start-sleep -seconds 10
                $found = $false
            }
        } until ($found)
    }
    $offices | export-csv -NoTypeInformation -Path $csv
    $offices | Export-Clixml -Path $xml
}

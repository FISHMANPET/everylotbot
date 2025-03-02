[CmdletBinding()]
param (
    [Parameter()]
    [int[]]
    $range = 0..9
)
$ErrorActionPreference = 'Stop'
foreach ($i in $range) {
    $offices = [System.Collections.ArrayList]@()
    $type = "boxes"
    $log = ".\zip$i.log"
    $csvtemp = ".\$type$i-temp.csv"
    $csv = ".\$type$i.csv"
    $xml = ".\$type$i.xml"
    $null | set-content -path $log
    $null | set-content -path $csv
    $null | set-content -path $xml
    $null | set-content -path $csvtemp
    switch ($i) {
        0 { $zips = (00501..09999) }
        9 { $zips = (90000..99950)}
        Default { $zips = (($i * 10000)..($i * 10000 + 9999)) }
    }
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $session.UserAgent = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36"
    $session.Cookies.Add((New-Object System.Net.Cookie("ak_bmsc", "19D767FEBFBDDE771070FAAB3A494A0A~000000000000000000000000000000~YAAQGSkeuFfMk/+UAQAAYNunKRriv+HxcIKmSmouzI2v8lzs7cZ5hMuECXtBKkec1bPKhPoUzDiIHbxo5enHICJ9qiS9p5cqD5XL5LNRXxzCCf1syNsu9NgA6RcI8efzr9CrEy6mv6Cs+V2WdhdrYz2sqqvIAn1eSDCvhm0hH15EwdWezmiGqfg9j8NfFVXIDZX3rOZ1BMl472grDBMFURBqEilbQW6b7D3B9vHMA65FW1h1EhgjbiRuJRk/7UNmYcMDW/+5eoseJoT6zGogukzdnMs0rkzxyY/PbZnBH1QvbhOXDwVQPun3EgoylG2+xdorDYUSgEGZvKN/yHsZ2bYnDoYsXipoJDLQikPN1OR5BXrGlXhrebBYf7Ne1C7VmrFVBKOUQXypdUih7Gg=", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_gcl_au", "1.1.95102927.1740160624", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("mab_usps", "48", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("tmab_usps", "64", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_ga", "GA1.1.511488693.1740160624", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_rdt_uuid", "1740160623861.46a6cce0-f936-4b10-8e84-cae8ba46e74e", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("sm_uuid", "1740161349136", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_scid", "PtGlErZtZN1CRYE8bPvRiCBmYZ7flj02J-Puyg", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_scid_r", "PtGlErZtZN1CRYE8bPvRiCBmYZ7flj02J-Puyg", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_fbp", "fb.1.1740160623904.487878365208427964", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_uetsid", "41eea950f07d11efa35e4580373a11a6", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_uetvid", "1eead280db3c11eebe1245238a9c7042", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_pin_unauth", "dWlkPU1USXlObVJsWVRVdFpqSXhZaTAwWmpBMExXRTVOR1l0TlRZM05HVmlPR0ZoWlRsbA", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_ScCbts", "%5B%5D", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_sctr", "1%7C1740117600000", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_ga_QM3XHZ2B95", "GS1.1.1740160623.1.0.1740160629.0.0.0", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("TLTSID", "03aabc17d85d16638a0b00e0ed96a2ca", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("NSC_uppmt_lvcfsofuft", "ffffffff3b225a4f45525d5f4f58455e445a4a42378b", "/", "tools.usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_ga_4WMWRRBML7", "GS1.1.1740160630.1.0.1740160630.0.0.0", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_ga_CSLL4ZEK4L", "GS1.1.1740160623.1.1.1740160630.0.0.0", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("mdLogger", "false", "/", "tools.usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("kampyleUserSession", "1740160630569", "/", "tools.usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("kampyleUserSessionsCount", "4", "/", "tools.usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("kampyleSessionPageCounter", "1", "/", "tools.usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("bm_sv", "79DDDF23A036F440DC5EC88A161581C1~YAAQzgfUF/mHKBWVAQAAnRGoKRr9zG83k/fOY53sf9As6QEp2SVrbSGYloR45/mC43M1IhksYFvgZ53soYJzU/DjBA8HpaZDVDFuW+ontFDs4nD7r7wXF9Q3zVdvD+6p5nE9xKB7JYsJvE19ExsZExm5KN0dOdessfNx+gliYZJzsbQd8C7uuTwQ9pLIG8iDJtVm2Q2x3MDr6Bmwi+vOSiq5AXsZFPdJhhQR8bxu4NmGtiRl3pzGnQERpGpVcw==~1", "/", ".usps.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("_ga_3NXP3C8S9V", "GS1.1.1740160623.1.1.1740160687.0.0.0", "/", ".usps.com")))
    foreach ($zip in $zips) {
        do {
            $found = $false
            $zipcode = $zip.ToString("00000")
            $zipcode | Add-Content -Path $log

            try {
                # $locations = Invoke-RestMethod -Uri "https://tools.usps.com/UspsToolsRestServices/rest/POLocator/findLocations" `
                #     -Method "POST" `
                #     -Headers @{
                #     "method"           = "POST"
                #     "authority"        = "tools.usps.com"
                #     "scheme"           = "https"
                #     "path"             = "/UspsToolsRestServices/rest/POLocator/findLocations"
                #     "accept"           = "application/json, text/javascript, */*; q=0.01"
                #     "x-requested-with" = "XMLHttpRequest"
                #     "user-agent"       = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36"
                #     "origin"           = "https://tools.usps.com"
                #     "sec-fetch-site"   = "same-origin"
                #     "sec-fetch-mode"   = "cors"
                #     "sec-fetch-dest"   = "empty"
                #     "referer"          = "https://tools.usps.com/find-location.htm"
                #     "accept-encoding"  = "gzip, deflate, br"
                #     "accept-language"  = "en-US,en;q=0.9"
                # } `
                #     -ContentType "application/json;charset=UTF-8" `
                #     -Body "{`"maxDistance`":`"100`",`"lbro`":`"`",`"requestType`":`"collectionbox`",`"requestServices`":`"`",`"requestRefineTypes`":`"`",`"requestRefineHours`":`"`",`"requestZipCode`":`"$zipcode`",`"requestZipPlusFour`":`"`"}"
                # $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
                # $session.UserAgent = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36"
                # $session.Cookies.Add((New-Object System.Net.Cookie("ak_bmsc", "19D767FEBFBDDE771070FAAB3A494A0A~000000000000000000000000000000~YAAQGSkeuFfMk/+UAQAAYNunKRriv+HxcIKmSmouzI2v8lzs7cZ5hMuECXtBKkec1bPKhPoUzDiIHbxo5enHICJ9qiS9p5cqD5XL5LNRXxzCCf1syNsu9NgA6RcI8efzr9CrEy6mv6Cs+V2WdhdrYz2sqqvIAn1eSDCvhm0hH15EwdWezmiGqfg9j8NfFVXIDZX3rOZ1BMl472grDBMFURBqEilbQW6b7D3B9vHMA65FW1h1EhgjbiRuJRk/7UNmYcMDW/+5eoseJoT6zGogukzdnMs0rkzxyY/PbZnBH1QvbhOXDwVQPun3EgoylG2+xdorDYUSgEGZvKN/yHsZ2bYnDoYsXipoJDLQikPN1OR5BXrGlXhrebBYf7Ne1C7VmrFVBKOUQXypdUih7Gg=", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_gcl_au", "1.1.95102927.1740160624", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("mab_usps", "48", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("tmab_usps", "64", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_ga", "GA1.1.511488693.1740160624", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_rdt_uuid", "1740160623861.46a6cce0-f936-4b10-8e84-cae8ba46e74e", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("sm_uuid", "1740161349136", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_scid", "PtGlErZtZN1CRYE8bPvRiCBmYZ7flj02J-Puyg", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_scid_r", "PtGlErZtZN1CRYE8bPvRiCBmYZ7flj02J-Puyg", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_fbp", "fb.1.1740160623904.487878365208427964", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_uetsid", "41eea950f07d11efa35e4580373a11a6", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_uetvid", "1eead280db3c11eebe1245238a9c7042", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_pin_unauth", "dWlkPU1USXlObVJsWVRVdFpqSXhZaTAwWmpBMExXRTVOR1l0TlRZM05HVmlPR0ZoWlRsbA", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_ScCbts", "%5B%5D", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_sctr", "1%7C1740117600000", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_ga_QM3XHZ2B95", "GS1.1.1740160623.1.0.1740160629.0.0.0", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("TLTSID", "03aabc17d85d16638a0b00e0ed96a2ca", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("NSC_uppmt_lvcfsofuft", "ffffffff3b225a4f45525d5f4f58455e445a4a42378b", "/", "tools.usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_ga_4WMWRRBML7", "GS1.1.1740160630.1.0.1740160630.0.0.0", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_ga_CSLL4ZEK4L", "GS1.1.1740160623.1.1.1740160630.0.0.0", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("mdLogger", "false", "/", "tools.usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("kampyleUserSession", "1740160630569", "/", "tools.usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("kampyleUserSessionsCount", "4", "/", "tools.usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("kampyleSessionPageCounter", "1", "/", "tools.usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("bm_sv", "79DDDF23A036F440DC5EC88A161581C1~YAAQzgfUF/mHKBWVAQAAnRGoKRr9zG83k/fOY53sf9As6QEp2SVrbSGYloR45/mC43M1IhksYFvgZ53soYJzU/DjBA8HpaZDVDFuW+ontFDs4nD7r7wXF9Q3zVdvD+6p5nE9xKB7JYsJvE19ExsZExm5KN0dOdessfNx+gliYZJzsbQd8C7uuTwQ9pLIG8iDJtVm2Q2x3MDr6Bmwi+vOSiq5AXsZFPdJhhQR8bxu4NmGtiRl3pzGnQERpGpVcw==~1", "/", ".usps.com")))
                # $session.Cookies.Add((New-Object System.Net.Cookie("_ga_3NXP3C8S9V", "GS1.1.1740160623.1.1.1740160687.0.0.0", "/", ".usps.com")))
                # Invoke-RestMethod -Uri "https://tools.usps.com/locations/getLocations" -Method Post `
                #     -Body @{
                #         "maxDistance" = "10"
                #         "requestHours" = ""
                #         "requestServices" = ""
                #         "reqeustType" = "PO,CPU,VPO"
                #         "requestZipCode" = $zipcode
                #     }
                $locations = Invoke-RestMethod -Uri "https://tools.usps.com/locations/getLocations" -ErrorVariable 'locError' `
                    -Method "POST" `
                    -WebSession $session `
                    -Headers @{
                        "authority"="tools.usps.com"
                        "method"="POST"
                        "path"="/locations/getLocations"
                        "scheme"="https"
                        "accept"="application/json, text/javascript, */*; q=0.01"
                        "accept-encoding"="gzip, deflate, br, zstd"
                        "accept-language"="en-US,en;q=0.9"
                        "cache-control"="no-cache"
                        "origin"="https://tools.usps.com"
                        "pragma"="no-cache"
                        "priority"="u=1, i"
                        "referer"="https://tools.usps.com/locations/?_gl=1*ejtlgj*_gcl_au*OTUxMDI5MjcuMTc0MDE2MDYyNA..*_ga*NTExNDg4NjkzLjE3NDAxNjA2MjQ.*_ga_3NXP3C8S9V*MTc0MDE2MDYyMy4xLjAuMTc0MDE2MDYyMy4wLjAuMA..*_ga_QM3XHZ2B95*MTc0MDE2MDYyMy4xLjAuMTc0MDE2MDYyMy4wLjAuMA.."
                        "sec-ch-ua"="`"Not(A:Brand`";v=`"99`", `"Google Chrome`";v=`"133`", `"Chromium`";v=`"133`""
                        "sec-ch-ua-mobile"="?0"
                        "sec-ch-ua-platform"="`"macOS`""
                        "sec-fetch-dest"="empty"
                        "sec-fetch-mode"="cors"
                        "sec-fetch-site"="same-origin"
                        "x-requested-with"="XMLHttpRequest"
                    } `
                    -ContentType "application/json;charset=UTF-8" `
                    -Body "{`"requestZipCode`":`"$zipcode`",`"requestType`":`"COLLECTIONBOX`",`"maxDistance`":`"100`",`"requestServices`":`"`",`"requestHours`":`"`"}"
                $locations | Add-Content -Path $log
                if ($locations.PSObject.properties.name -notcontains 'locations') {
                    Write-Host $zipcode
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
            } catch [System.Net.WebException] {
                start-sleep -seconds 10
                $found = $false
            }
        } until ($found)
    }
    $offices | export-csv -NoTypeInformation -Path $csv
    $offices | Export-Clixml -Path $xml
}

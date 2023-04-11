##   KUIDAS SAADA KÄTTE MEILILISTID ##

1) Vaata kas sul Windows Powershell on arvutis olemas

2) Tee Powershell admin õigustega lahti

3) Jooksuta see kood: ```Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.1.0```  <-- Seda pead tegema ainult esimene kord

4) Seejärel jooksuta järgmine kood:

------------------------------------------------------

```
Connect-ExchangeOnline -ShowBanner:$False

Clear
$NF = ""

$DL = Read-Host "Sisesta soovitud Distribution Listi nimi"

Write-Host $DL

# Kontrollime, kas selline list on olemas
try {
    Get-DistributionGroup -Identity $DL
}
catch {
    $NF = "Distribution Listi ei leitud"
}

if ($NF -eq "") {
    # Specify the export file path
    $exportfile = 'C:\Meililisti_liikmed.csv'
    
    # Kontrollime kas fail on juba olemas
    if (Test-Path $exportfile) {
        $overwrite = $false
        while (-not $overwrite) {
            $newfilename = Read-Host "Selle nimega fail on juba olemas. Palun vali uus nimi koos faililaiendiga (.csv)"
            $newfile = Join-Path (Split-Path $exportfile) $newfilename
            if (Test-Path $newfile) {
                Write-Host "Selle nimega fail on juba olemas. Palun vali uus nimi koos faililaiendiga (.csv)"
            }
            else {
                $exportfile = $newfile
                $overwrite = $true
            }
        }
    }
    
    # Expordime listi CSV failina
    Get-DistributionGroupMember -Identity $DL | Select Name, PrimarySMTPAddress | Export-Csv -Path $exportfile -Force
    
    Write-Host "Export tehtud. Korras."
}
else {
    Write-Host "Distribution Listi ei leitud. Kas kirjutasid nime õigesti?"
}
```

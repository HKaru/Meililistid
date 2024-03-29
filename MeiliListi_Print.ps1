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
    $exportfile = Join-Path (Split-Path $MyInvocation.MyCommand.Path) "Meililisti_liikmed.csv"

    # Kontrollime kas fail on juba olemas
    if (Test-Path $exportfile) {
        $overwrite = $false
        while (-not $overwrite) {
            $newfilename = Read-Host "Selle nimega fail on juba olemas. Palun vali uus nimi koos faililaiendiga (.csv)"
            $newfile = Join-Path $PSScriptRoot $newfilename
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

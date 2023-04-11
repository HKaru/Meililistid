Connect-ExchangeOnline -ShowBanner:$False

clear
$NF =""


$DL = Read-Host "Pane Distribution Listi nimi"

write-host $DL
# Kood kontrollib kas see list on olemas
try{Get-DistributionGroup -Identity $DL} catch{$NF="Distribution List Not Found"} 
  
if($NF -eq "")  
 {
 # Salvesta list 
 $exportfile =  'C:\Meililisti_liikmed.csv' 

 # Siin valid mis tulbad sa expordid ning salvestad CSV failina 
 Get-DistributionGroupMember -Identity $DL | Select Name, PrimarySMTPAddress | Export-Csv -Path $exportfile  
}

Else 
{
Write-host "Distribution listi ei leitud, kontrolli õigekirja"
}
##   KUIDAS SAADA KÄTTE MEILILISTID ##

1) Vaata kas sul Windows Powershell on arvutis olemas

2) Tee Powershell admin õigustega lahti

3) Jooksuta see kood: Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.1.0  ### Seda pead tegema ainult esimene kord ###

4) Seejärel jooksuta järgmine kood:
------------------------------------------------------

Connect-ExchangeOnline -ShowBanner:$False

clear
$NF =""


$DL = Read-Host "Pane Distribution Listi nimi"

write-host $DL
# Kood kontrollib kas see list on olemas
try{Get-DistributionGroup -Identity $DL} catch{$NF="Distribution List Not Found"} 
  
if($NF -eq "")  
 {
 # Salvesta list - vajadusel saad muuta kohta
 $exportfile =  'C:\Meililisti_liikmed.csv' 

 # Siin valid mis tulbad sa expordid ning salvestad CSV failina 
 Get-DistributionGroupMember -Identity $DL | Select Name, PrimarySMTPAddress | Export-Csv -Path $exportfile  
}

Else 
{
Write-host "Distribution listi ei leitud, kontrolli õigekirja"
}

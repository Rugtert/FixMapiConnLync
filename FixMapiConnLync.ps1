#Check to see if Outlook is running and wait for it to stop running. End if Outlook has not stopped in 60 seconds.
While (Get-Process outlook -ErrorAction SilentlyContinue)
{
Write-host "Waiting for Outlook to stop running." -ForegroundColor Green
Wait-Process -Name Outlook -Timeout 60 -ErrorAction Stop
}

#Start Outlook
add-Type -assembly "Microsoft.Office.Interop.Outlook"
$outlook = new-object -ComObject outlook.application 
#Wait for MAPI Connection to establish.
While (!($outlook.GetNamespace("MAPI")).autodiscoverxml)
{
    Write-Host "Establishing MAPI Connection" -ForegroundColor Green
}

#Get necessary variables from the MAPI Namespace property AutodiscoverXML
$Autodiscover = $Outlook.GetNameSpace("MAPI")
[xml]$AutodiscoverXML = [xml]$Autodiscover.AutoDiscoverXml
$legacyDN = $autodiscoverxml.Autodiscover.Response.User.LegacyDN
$SMTPAddress = $AutodiscoverXML.Autodiscover.Response.User.AutoDiscoverSMTPAddress
[string]$outlookprofile = $outlook.DefaultProfileName

#Set Registry path to the Outlook profile
$HKCUPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\$outlookprofile"

#Get serviceUID Value in Bytes
$HKCUProfile = Get-ChildItem HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\$outlookprofile\9375CFF0413111d3B88A00104B2A6676 | Get-ItemProperty | Where {$_."Account Name" -eq "$SMTPAddress"}

#Convert ServiceUID from Bytes to String and remove Dashes (-)
$ServiceUID = ([System.BitConverter]::ToString($HKCUProfile.'Service UID')).Replace("-","")

#Get the destination path from property 01023d0d in $HKCUPath\$ServiceUID and remove the dashes.
$DestPath = ([system.bitconverter]::ToString((Get-ItemProperty -Path "$HKCUPath\$ServiceUID")."01023d0d")).replace("-","")
if (!(Get-ItemProperty -Path "$HKCUPath\$DestPath" -name 001e6603 -ErrorAction SilentlyContinue))
{
    New-ItemProperty -Path "$HKCUPath\$DestPath" -PropertyType String -name 001e6603 -Value $legacyDN
    Write-host "Adding $legacyDN Value to Reg_SZ 001e6603 in $HKCUPath\$DestPath to fix the Lync MAPI Connection." -ForegroundColor Green
}
Else
{
    Write-host "Reg_SZ 001e6603 is already configured with the value $legacyDN. No change necessary" -ForegroundColor Green
}
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | out-null
[gc]::collect() 
[gc]::WaitForPendingFinalizers()
Get-Process Outlook -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue
Remove-Variable -Name outlook
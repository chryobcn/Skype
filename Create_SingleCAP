#Create an Common Area Phone:
##Define Common Area phone

$SLineuri = "tel:+12037989988;ext=52425"
$sRegistrarPool = "XXXpblync01sba.eu.boehringer.com"
$sOU = "OU=UC-Devices,OU=Users,OU=BI,DC=eu,DC=boehringer,DC=com"
$sDisplayName = ""
$ssDescription = ""
$sDisplayNumber = ""
$sSipAddress = "sip:zpRDGCA52797@boehringer-ingelheim.com"

New-CsCommonAreaPhone -LineUri $SLineuri -RegistrarPool $sRegistrarPool -OU $sOU -Description $sDescription -DisplayName $sDisplayName -DisplayNumber $sDisplayNumber -SipAddress $sSipAddress

$sVoicePolicy = ""
$sDialPlan = ""

#Grant Lync phone a Voice Policy
Grant-CsVoicePolicy -Identity $sDisplayName  -Policyname $sVoicePolicy
#Grant Lync phone a Dial Plan
Grant-CsDialPlan -Identity $sDisplayName  -Policyname $sDialPlan
#Set Lync phone log in Pin
Set-CsClientPin -Identity $sDisplayName  -Pin 123456
#Set Lync phone a Conferencing Policy
Grant-CsClientPolicy -Identity $sDisplayName  -PolicyName PhoneEditionPolicy

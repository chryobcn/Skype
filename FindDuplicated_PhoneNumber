$match ="*+902123674185*"
Get-CsUser -Filter {PrivateLine -like $match} | ft DisplayName,lineuri,PrivateLineUser
Get-CsUser -Filter {LineURI -like $match} | ft DisplayName,lineuri


Get-CsAnalogDevice -Filter { lineURI -like $match} | ft DisplayName,lineuri
Get-CsCommonAreaPhone -Filter {LineURI -like $match} | ft DisplayName,lineuri
Get-CsExUmContact -Filter {LineURI -like $match} | ft DisplayName,lineuri
Get-CsDialInConferencingAccessNumber -Filter {LineURI -like $match} | ft DisplayName,lineuri


Get-CsTrustedApplicationEndpoint -Filter {LineURI -like $match}  | ft DisplayName,lineuri
Get-CsRgsWorkflow | Where-Object {$_.LineURI -like $match} | ft Name,lineuri

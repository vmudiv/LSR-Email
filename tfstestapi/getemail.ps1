function getEmail($teamname)
{
$myICMCertificate = Get-ChildItem Cert:\CurrentUser\my | Where-Object {$_.subject -match "mobile-center-operations-triagelooper"}
$sono = "https://icm.ad.msft.net/api/cert/oncall/teams?`$filter=PublicId eq '"+$teamname +"'"
$Sonooncal = Invoke-RestMethod -Method Get -Uri $sono -Certificate $myICMCertificate
$email=$Sonooncal.value.Email 
return $email
}

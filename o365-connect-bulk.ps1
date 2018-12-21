## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script allow login to multiple MS Cloud services by calling individual scripts

## Source - https://github.com/directorcia/Office365/blob/master/o365-connect-bulk.ps1

## Prerequisites = 1
## 1. All required scripts MUST be in the same directory as this script, so make sure you are in that directory before running this

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$warningmessagecolor = "yellow"

$nomfa = $true      ## MFA not required for login
$std = $false       ## MS Online connect
$aad = $false       ## Azure AD connect
$exo = $true        ## Exchange Online connect
$s4b = $false       ## Skype for Business Online connect
$sac = $false       ## Security and Compliance Center connect
$spo = $false       ## SharePoint Online connect
$tms = $false       ## Teams connect
$aadrm = $false     ## Azure AD Rights Management connect

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Start Script`n"

if ($nomfa) {
    write-host -foregroundcolor $processmessagecolor "Start - Non MFA login"
    if ($std) {.\o365-connect.ps1}
    if ($aad) {.\o365-connect-aad.ps1}                                                                                                                       
    if ($exo) {.\o365-connect-Exo.ps1}                                                                                                                                                                                                                                      
    if ($s4b) {.\o365-connect-s4b.ps1}                                                                                                                       
    if ($sac) {.\o365-connect-sac.ps1}                                                                                                                       
    if ($spo) {.\o365-connect-spo.ps1}                                                                                                                       
    if ($tms) {.\o365-connect-tms.ps1}
    if ($aadrm) {.\o365-connect-aadrm.ps1}
    write-host -foregroundcolor $processmessagecolor "Finish - Non MFA login`n"                                                                                                                       
}                                                                                                                           
else {
    write-host -foregroundcolor $processmessagecolor "Start - MFA login"
    if ($std) {.\o365-connect-mfa.ps1}
    if ($aad) {.\o365-connect-aad-mfa.ps1}                                                                                                                       
    if ($exo) {.\o365-connect-Exo-mfa.ps1}                                                                                                                                                                                                                                      
    if ($s4b) {.\o365-connect-s4b-mfa.ps1}                                                                                                                       
    if ($sac) {.\o365-connect-sac-mfa.ps1}                                                                                                                       
    if ($spo) {.\o365-connect-spo-mfa.ps1}                                                                                                                       
    if ($tms) {.\o365-connect-tms-mfa.ps1}
    if ($aadrm) {.\o365-connect-aadrm-mfa.ps1}                                                                                                                   
    write-host -foregroundcolor $processmessagecolor "Finish - MFA login`n"                             
}

write-host -foregroundcolor $systemmessagecolor "Finish Script`n"
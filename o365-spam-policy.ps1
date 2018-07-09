## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.
## Crowd source settings for filtering - http://blog.ciaops.com/2018/06/spam-filtering-in-office-365best.html

## Source - https://github.com/directorcia/Office365/blob/master/o365-spam-policy.ps1

## Description
## Script designed to configure a new Exchange Online spam filtering policy

## Prerequisites = 1
## 1. Ensure Exchange online PowerShell module installed or updated

## Variables
## Separate multiple domains with comma (,) e.g."domain1.com", "domain2.com", "domain3.com"
$domains = "M365B555418.onmicrosoft.com" 
$policyname = "Configured Policy"
$rulename = "Configured Recipients"

Clear-Host

write-host -foregroundcolor Cyan "Script started"

write-host -foregroundcolor Cyan "Set new spam policy"

$policyparams = @{
    "name" = $policyname;
    'Bulkspamaction' =  'movetojmf';
    'bulkthreshold' =  '7';
    'highconfidencespamaction' =  'movetojmf';
    'inlinesafetytipsenabled' = $true;
    'markasspambulkmail' = 'on';
    'increasescorewithimagelinks' = 'off'
    'increasescorewithnumericips' = 'on'
    'increasescorewithredirecttootherport' = 'on'
    'increasescorewithbizorinfourls' = 'on';
    'markasspamemptymessages' ='on';
    'markasspamjavascriptinhtml' = 'on';
    'markasspamframesinhtml' = 'on';
    'markasspamobjecttagsinhtml' = 'on';
    'markasspamembedtagsinhtml' ='on';
    'markasspamformtagsinhtml' = 'on';
    'markasspamwebbugsinhtml' = 'on';
    'markasspamsensitivewordlist' = 'on';
    'markasspamspfrecordhardfail' = 'on';
    'markasspamfromaddressauthfail' = 'on';
    'markasspamndrbackscatter' = 'on';
    'phishspamaction' = 'movetojmf';
    'spamaction' = 'movetojmf';
    'zapenabled' = $true
    }
    
new-hostedcontentfilterpolicy @policyparams

write-host -foregroundcolor Cyan "Set new filter rule"

$ruleparams = @{
    'name' = $rulename;
    'hostedcontentfilterpolicy' = $policyname;     ## this needs to match the above policy name
    'recipientdomainis' = $domains;     
    'Enabled' = $true
    }
    
New-hostedcontentfilterrule @ruleparams

write-host -foregroundcolor Cyan "Script complete"

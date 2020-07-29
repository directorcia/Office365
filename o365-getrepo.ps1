<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Download all contents of /directorcia/office365 repository from GitHub

Source - https://github.com/directorcia/Office365/blob/master/o365-getrepo.ps1

Original concept - https://gist.github.com/chrisbrownie/f20cb4508975fb7fb5da145d3d38024a

Prerequisites = 0

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$Owner = "directorcia"                      ## Repository owner
$Repository = "Office365"                   ## Repository name
$Path = ""                                  ## Subfolder within Repository if required
$DestinationPath = "C:\downloads\repo\"     ## Location for local copy of repository, does not need to exist prior

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Start Script`n"

[Net.ServicePointManager]::SecurityProtocol = "tls12,tls11,tls,ssl3"  ## Avoid failure to create secure channel

$baseUri = "https://api.github.com/"
$args = "repos/$Owner/$Repository/contents/$Path"
[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls,ssl3"  ## Avoid failure to create secure channel
$wr = Invoke-WebRequest -Uri $($baseuri+$args) -UseBasicParsing
$objects = $wr.Content | ConvertFrom-Json
$files = $objects | where {$_.type -eq "file"} | Select -exp download_url
$directories = $objects | where {$_.type -eq "dir"}
    
$directories | ForEach-Object { 
   DownloadFilesFromRepo -Owner $Owner -Repository $Repository -Path $_.path -DestinationPath $($DestinationPath+$_.name)
}
    
if (-not (Test-Path $DestinationPath)) {
   # Destination path does not exist, let's create it
   try {
       New-Item -Path $DestinationPath -ItemType Directory -ErrorAction Stop
   } catch {
            throw "Could not create path '$DestinationPath'!"
   }
}

foreach ($file in $files) {
    $fileDestination = Join-Path $DestinationPath (Split-Path $file -Leaf)
   try {
       Invoke-WebRequest -Uri $file -OutFile $fileDestination -ErrorAction Stop -Verbose -UseBasicParsing
       "Grabbed '$($file)' to '$fileDestination'"
  } catch {
       throw "Unable to download '$($file.path)'"
   }
}

write-host -foregroundcolor $systemmessagecolor "Script completed`n"
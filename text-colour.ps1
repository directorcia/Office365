<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Display all PowerShell copnsole colour combinations

Source - https://github.com/directorcia/Office365/blob/master/text-colour.ps1

Prerequisites = 0

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables

$Fg = @(        ## Foregrund colours
    "Black",
    "Blue",
    "Cyan",
    "DarkBlue",
    "DarkCyan",
    "DarkGray",
    "DarkGreen",
    "DarkMagenta",
    "DarkRed",
    "DarkYellow",
    "Gray",
    "Green",
    "Magenta",
    "Red",
    "White",
    "Yellow"
)

$bg = @(        ## Background colours
    "Black",
    "Blue",
    "Cyan",
    "DarkBlue",
    "DarkCyan",
    "DarkGray",
    "DarkGreen",
    "DarkMagenta",
    "DarkRed",
    "DarkYellow",
    "Gray",
    "Green",
    "Magenta",
    "Red",
    "White",
    "Yellow"
)

$systemmessagecolor = "cyan"

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

foreach ($bgcount in $bg) {
    Foreach ($fgcount in $fg) {
        Write-host "Foreground =", $fgcount, "/ Background =", $bgcount
        Write-host -foregroundcolor $fgcount -backgroundcolor $bgcount "The quick brown fox jumped over the lazy dog 1234567890!@#$%^&*()-+="
        write-host
    }
    write-host
}

write-host -foregroundcolor $systemmessagecolor "`nScript Completed`n"
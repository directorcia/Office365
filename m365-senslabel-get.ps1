# Get all sensitivity labels
$labels = Get-Label -IncludeDetailedLabelActions
#set-label
#https://learn.microsoft.com/en-us/graph/api/security-sensitivitylabel-get?view=graph-rest-beta
#https://practical365.com/sensitivity-label-settings-report/#:~:text=The%20Get-Label%20cmdlet%20returns%20sensitivity%20label%20settings%20in,like%3A%20%24LabelActions%20%3D%20%28Get-Label%20-Identity%20%24Label.ImmutableId%29.LabelActions%20%7C%20Convertfrom-JSON

# Create an array to store label information
$labelInfo = @()

# Loop through each label and collect relevant information
foreach ($label in $labels) {
    $labelDetails = [PSCustomObject]@{
        Name               = $label.DisplayName
        Description        = $label.Comment
        Tooltip            = $label.Tooltip
        ContentMarking     = $label.LabelActions -match "applycontentmarking"
        Watermarking       = $label.LabelActions -match "applywatermarking"
        Encryption         = $label.LabelActions -match "encrypt"
        EndpointProtection = $label.LabelActions -match "endpoint"
        ProtectGroup       = $label.LabelActions -match "protectgroup"
        ProtectSite        = $label.LabelActions -match "protectsite"
        ProtectTeams       = $label.LabelActions -match "protectteams"
    }
    $labelInfo += $labelDetails
    write-host $label.EncryptionRightsDefinitions 
    pause
}

# Display a summary of the labels
$labelInfo | Format-Table Name, Description, Tooltip


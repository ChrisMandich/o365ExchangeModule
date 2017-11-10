# Remove an email from Collection Array based on Inbox name
function Remove-FromCollection{

<#
    .Synopsis
    This removes an email from the inbox array. TBH, I don't like this module... 

    .Description
    Remove-FromCollection will search through the array for the inbox specified and remove all emails that match that inbox.
      
    .PARAMETER inbox
    Specifies inbox to remove from collection array.     

    .Example 
    Remove-FromCollection -inbox "example@email.com"
#>

    Param(
        [string]$inbox = $(read-host "Set inbox to search")
    )

    [System.Collections.ArrayList]$remove = @()

    foreach ($email in $global:collection){
        if ($email.Inbox -eq $inbox){
            $remove += $email
        }
    }

    foreach ($email in $remove){
        $global:collection.Remove($email)
    }
}
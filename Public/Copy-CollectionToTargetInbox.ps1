# Copy Collection to a specified target inbox. 
function Copy-CollectionToTargetInbox{
<#
    .SYNOPSIS
    This will copy the collection to a target inbox. 

    .DESCRIPTION
    It iterates through the collection, creates a query then opens each inbox to search for that query. Messages found will be copied to a target inbox. 

    .PARAMETER TargetMailbox
    Target inbox to copy messages to. 

    .PARAMETER TargetFolder
    Folder to be created within target inbox for copying to. Default: MM-d-yy HH:mm:ss

    .PARAMETER TranscriptPath
    Full path to save session transcript. Default: %USERPROFILE%\log\MM-d-yy\HHmmss_copy.log

    .EXAMPLE 
    Copy-CollectionToTargetInbox -target "InboxName"

    .EXAMPLE
    Copy-CollectionToTargetInbox -target "InboxName" -folder "This folder instead" -TranscriptPath "C:\I_WANT_MY_LOGS_HERE\copy.log"
#>

    Param(
        [String]$TargetMailbox = $(read-host "Target Inbox"),
        [String]$TargetFolder = $(get-date -format 'MM-%d-yy HH:mm:ss'),
        [String]$TranscriptPath = $($env:USERPROFILE + "\log\" + $(get-date -format 'MM-%d-yy') + "\" + $(Get-Date -format 'HHmmss') + "_copy.log")
    )

    # Array is used to store remove queue
    [System.Collections.ArrayList]$remove = @()

    #Start Transcript logging
    Start-Transcript -Path $TranscriptPath -Append -Force

    #Get Size of Collection
    $totIndex = $Global:collection.count

    foreach ($email in $global:collection){
        [String]$inbox = $email.Inbox
        
        # This is to keep track of index (# of emails processed)
        $index=([array]::IndexOf($global:collection, $email) + 1)

        #OUTPUT
        write-host "Index: $($index)/$($totIndex)"
        Write-Host "`tSearching Inbox: `"$($Email.inbox)`"" 
        Write-Host "`tQuery: `'$($Email.QueryBuilder())`'"
        Write-Host "`tTarget Inbox: `"$($TargetMailbox)`""
        Write-Host "`tFolder Name: `"$($TargetFolder)`""
        
        #Copy Message to Target Inbox
        $results = $Email.CopyMessageToTargetInbox($TargetMailbox,$TargetFolder)
        
        if($results -ge 0){
            #OUTPUT
            Write-Host "`tSearch Results: `"$results`"" -ForegroundColor Yellow
            write-host
        }
        else{
            #OUTPUT
            $remove += $email
            Write-Host "`tEmail Does Not Exist, Adding to Remove Queue" -ForegroundColor Red
        }

    }

    # Remove emails that do no exist. 
    foreach ($Email in $remove){
        Write-Host "`Removing Inbox from Collection: `"$($Email.inbox)`"" -ForegroundColor Yellow
        Write-Host "`tQuery: `'$($Email.QueryBuilder())`'" -ForegroundColor Yellow
        $global:collection.Remove($Email)
    }

    # Stop Transcript
    Stop-Transcript
}
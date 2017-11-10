# Remove all emails in the Collection Array from Target Inboxes. 
function Remove-CollectionFromInbox{
<#
    .Synopsis
    Removes Emails specified in collection array from inboxes. 

    .Description
    Iterates through collection array to build a query. The query is searched through the specified inbox. All matches are deleted on match. 

    .PARAMETER Force
    Force validates that you want to delete messages from the collection

    .PARAMETER TranscriptPath
    Full path to save session transcript. Default: %USERPROFILE%\log\MM-d-yy\HHmmss_copy.log

    .Example 
    Remove-CollectionFromInbox -Force
#>
    Param(
        [Parameter(Mandatory=$true)][Switch]$Force=$false,
        [String]$TranscriptPath = $($env:USERPROFILE + "\log\" + $(get-date -format 'MM-%d-yy') + "\" + $(Get-Date -format 'HHmmss') + "_remove.log")
    )

    if($Force -eq $true){
    
    #Get Size of Collection
    $totIndex = $Global:collection.count
        
    #Start Transcript logging
    Start-Transcript -Path $TranscriptPath -Append -Force    
    

        foreach ($email in $global:collection){
            [String]$inbox = $email.Inbox
        
            # This is to keep track of index (# of emails processed)
            $index=([array]::IndexOf($global:collection, $email) + 1)

            #OUTPUT
            write-host "Index: $($index)/$($totIndex)"
            Write-Host "`tSearching Inbox: `"$($Email.inbox)`"" 
            Write-Host "`tQuery: `'$($Email.QueryBuilder())`'"
        
            #Copy Message to Target Inbox
            $results = $Email.DeleteMessageFromInbox()
        
            if($results -ge 0){
                #OUTPUT
                Write-Host "`tSearch Results: `"$results`"" -ForegroundColor Yellow
                write-host
            }
            else{
                #OUTPUT
                Write-Host "`tEmail Does Not Exist, Adding to Remove Queue" -ForegroundColor Red
            }

        }
        #Clear Collection
        $global:collection.clear()
    }
    else{
        write-host "Use Force Switch to Verify You want to Delete the Messages in the collection" -ForegroundColor Red
    }


    # Stop Transcript
    Stop-Transcript
}
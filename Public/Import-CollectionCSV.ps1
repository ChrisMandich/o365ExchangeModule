
# Imports a CSV to the Collection Array. CSV headers must match email object fields. 
function Import-CollectionCSV{
<#
    .SYNOPSIS
    This allows you to batch import emails from a csv.

    .DESCRIPTION
    This allows you to batch import emails to a collection from a CSV. Make sure your headers include an inbox and at least one other key. Available keys include: to, from, subject, kind, cc, bcc, body, attachment, received and sent.
    
    Examples of keys can be found in "Add-ToCollection -detailed" 

    .PARAMETER path
    Full path to CSV to be imported.

    .EXAMPLE
    Import-CollectionCSV -FilePath .\path\filename.csv
#>

    Param(
        [string]$FilePath = $(read-host "CSV Location")
    )

    # Import CSV as from specified path. 
    $tmpCSV = Import-Csv $FilePath


    # Verify that inbox key exists in csv.
    if ($tmpCSV | Get-member -MemberType 'NoteProperty' | where name -eq "inbox"){
        # Add each email to collection
        foreach ($email in $tmpCSV){
                # Verify that inbox is not empty.
                if($email.inbox -ne ""){
                    write-host "Added -- Inbox: $($email.inbox)" -ForegroundColor Yellow
                    Add-ToCollection -participants $email.participants -inbox $email.inbox -to $email.to -from $email.from -subject $email.subject -kind $email.kind -cc $email.cc -bcc $email.bcc -body $email.body -attachment $email.attachment -received $email.received -sent $email.sent
                }
                else{
                    Write-Host "`tThere is no inbox on this line (0-0)" -ForegroundColor Red
                }
        }
    }
    else{
        Write-Error "Collection CSV does not include an inbox." -Category InvalidData
    }
}
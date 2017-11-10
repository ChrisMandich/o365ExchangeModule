# Add an email to the collection.
function Add-ToCollection{

<#
    .SYNOPSIS

    This allows you to add an email to the collection array.

    .DESCRIPTION

    This allows you to add an email to the collection array. Please specify an inbox and additional search terms. Try to be as specific as possible. 

    Additional information about searchable properties in exchange: https://technet.microsoft.com/en-us/library/dn774955(v=exchg.150).aspx

    .PARAMETER inbox

    Sets the inbox for email added to collection. Exact phrase. Ex: test@mail.com

    .PARAMETER participants

    Sets the participants field in an email message. This includes the following fields: (From, To, CC, and BCC)

    .PARAMETER to

    Sets the recipient of an email message. Will search for exact phrase or keyword

    .PARAMETER from

    Sets the sender of an email message. Will search for exact phrase or keyword

    .PARAMETER subject

    Sets the subject of an email message. Will search for exact phrase or keyword. EX: TEST EX: TEST EMAIL

    .PARAMETER kind

    Sets the kind field for query. Defauls to email. Available (contacts, docs, email, faxes, im, journals, meetings, notes, posts, rssfeeds, tasks, voicemail)

    .PARAMETER cc

    Sets the CC field of an email message.

    .PARAMETER bcc

    Sets the BCC field of an email message.

    .PARAMETER body

    Sets the body field of an email message.

    .PARAMETER attachment

    Sets the attachment field of an email message. EX: report.pdf EX: report.*

    .PARAMETER received

    Sets the date of an email message was received. EX: 5/2/2010

    .PARAMETER sent

    Sets the date of an email message was sent. EX: 5/2/2010

    .EXAMPLE 

    Add-ToCollection -inbox "email@example.com" -from "attacker@example.com"

    .EXAMPLE

    Add-ToCollection -inbox "email@example.com" -from "attacker@example.com" -subject "This is no joke"

    .EXAMPLE

    Add-ToCollection -inbox "email@example.com" -from "attacker@example.com" -attachment "tempura.doc"
#>

    Param(
        [Parameter(Mandatory=$true)][string]$inbox = "",
        [string]$to = "",
        [string]$from = "",
        [string]$participants = "",
        [string]$subject = "",
        [string]$kind = $("email"),
        [string]$cc = "",
        [string]$bcc = "",
        [string]$body = "",
        [string]$attachment = "",
        [string]$received = "",
        [string]$sent = ""
    )
    
    try{
        $result = Get-Mailbox -Identity $inbox -ErrorAction Stop
        }
    catch{
        write-host "Inbox does not exist"
        }

    #Create Email Object
    $email = New-EmailObject -inbox $inbox -to $to -from $from -participants $participants -subject $subject -kind $kind -cc $cc -bcc $bcc -body $body -attachment $attachment -received $received -sent $sent 
    
    #Add email object to collection array
    $global:collection += $email    
    
}
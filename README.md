# o365ExchangeModule
PowerShell module to collect and remove emails from Office 365

#### Add-ToCollection 
    Description: Adds the Email to the Collection Array
    Parameters: inbox(required), to, from, participants, subject, kind, cc, bcc, body, attachment, received, and sent
    Example: Add-ToCollection -inbox "email@example.com" -from "attacker@example.com" -subject "This is no joke"

#### Clear-Collection
    Description: Empties the collection Array
    Example: Clear-Collection 

#### Copy-CollectionToTargetInbox
    Description: This will attempt to copy the emails within the collection Array to the target inbox
    Parameters: TargetMailbox(required), TargetFolder, TranscriptPath
    Example: Copy-CollectionToTargetInbox -target "MyInbox@example.com"

#### Enter-ExchangeSession 
    Description: Creates a remote session with ps.outlook.com/powershell. Imports the session globally allowing users to interact with exchange online in the current PS window. Imports, Get-Mailbox, Search-Mailbox, and Get-MessageTrace. 
    Parameters: Credential(Credential Object), proxy(Switch)
    Example: 
        $creds = get-credential 
        Enter-ExchangeSession -proxy -Credential $creds

#### Get-MessageTraceCustom 
    Description: Creates a Custom Message Trace search, iterates through all pages and sets the default pages size to 5000. 
    Parameters: SenderAddress, RecipientAddress, FromIp, ToIp, Status, StartDate, EndDate
    Example: Get-MessageTraceCustom -StartDate $(get-date).AddDays(-3) -SenderAddress "test@email.com"

#### Import-CollectionCSV 
    Description: Allow for batch import emails from a CSV. Make sure your headers include an inbox and at least one other key. Available keys include: to, from, subject, kind, cc, bcc, body, attachment, received and sent.
    Parameters: FilePath
    Example: Import-CollectionCSV -FilePath .\path\filename.csv

#### New-EmailObject 
    Description: Creates a custom email object
    Parameters: inbox(required), to, from, participants, subject, kind, cc, bcc, body, attachment, received, and sent
    Example: New-EmailObject -inbox "email@example.com" -from "attacker@example.com" -subject "This is no joke"

#### Remove-CollectionFromInbox
    Description: Remove E-Mails specified in the collection array from the associated inbox. 
    Parameters: TranscriptPath, Force(Switch)
    Example: Remove-CollectionFromInbox -Force

#### Remove-ExchangeSession 
    Description: Remove current exchange session and module
    Example: Remove-ExchangeSession

#### Remove-FromCollection
    Description: Removes and email from the inbox array. 
    Parameters: inbox
    example: Remove-FromCollection -inbox "example@email.com"

#### Reset-ExchangeSessionState:
    Description: Resets the outlook exchange session. 
    Parameters: Proxy, Credential
    Example: Reset-ExchangeSessionState -proxy -Credential $creds

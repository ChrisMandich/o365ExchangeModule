function New-EmailObject{
    Param(
        [string]$inbox = "",
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

    #create email object and add members
    $email = New-Object -TypeName PSObject

    #Add Fields to Object
    if($inbox -ne ""){
        $email | Add-Member -Type NoteProperty -Name Inbox -value $inbox
        }
    else{
        #If an inbox was not specified, do not complete
        break
    }
    if($to -ne ""){
        $email | Add-Member -Type NoteProperty -Name To -value $to
        }
    if($from -ne ""){
        $email | Add-Member -Type NoteProperty -Name From -value $from
        }
    if($participants -ne ""){
        $email | Add-Member -Type NoteProperty -Name Participants -value $participants
        }
    if($subject -ne ""){
        $email | Add-Member -Type NoteProperty -Name Subject -value $subject
        }
    if($kind -ne ""){
        $email | Add-Member -Type NoteProperty -Name kind -value $kind
        }
    if($cc -ne ""){
        $email | Add-Member -Type NoteProperty -Name CC -value $cc
        }
    if($bcc -ne ""){
        $email | Add-Member -Type NoteProperty -Name BCC -value $bcc
        }
    if($body -ne ""){
        $email | Add-Member -Type NoteProperty -Name Body -value $body
        }
    if($attachment -ne ""){
        $email | Add-Member -Type NoteProperty -Name Attachment -value $attachment
        }
    if($received -ne ""){
        $email | Add-Member -Type NoteProperty -Name Received -value $received
        }
    if($sent -ne ""){
        $email | Add-Member -Type NoteProperty -Name Sent -value $sent
        }

    #Estimate Results
    $email | Add-Member -Type ScriptMethod -Name GetResultCount -Force -Value {

            # TODO Test exchange session state, Reset if not open. Need to implement Proxy Feature
            #Reset-ExchangeSessionState -Credential $Credential > $null

            [String]$SearchQuery = $this.QueryBuilder()

            $identity = $this.Inbox.toString()
    
            [System.Object]$Results = Search-Mailbox -Identity $identity -SearchQuery $SearchQuery -EstimateResultOnly -ErrorAction Stop -WarningAction silentlyContinue

            #Return results
            return $Results.ResultItemsCount
        }

    #Copy Message to Target Inbox
        $email | Add-Member -Type ScriptMethod -Name CopyMessageToTargetInbox -Force -Value {
            Param(
                [Parameter(Mandatory=$true)][string]$TargetMailbox,
                [Parameter(Mandatory=$true)][string]$TargetFolder
            )
            # TODO Test exchange session state, Reset if not open. Need to implement Proxy Feature
            #Reset-ExchangeSessionState -Credential $Credential > $null

            [String]$SearchQuery = $this.QueryBuilder()

            $identity = $this.Inbox.toString()
    
            [System.Object]$Results = Search-Mailbox -Identity $identity -SearchQuery $SearchQuery -TargetMailbox $TargetMailbox -TargetFolder $TargetFolder -ErrorAction Stop -WarningAction silentlyContinue -LogLevel Full

            #Return results
            return $Results.ResultItemsCount
        }

        #Delete Message from Inbox
        $email | Add-Member -Type ScriptMethod -Name DeleteMessageFromInbox -Force -Value {
            # TODO Test exchange session state, Reset if not open. Need to implement Proxy Feature
            #Reset-ExchangeSessionState -Credential $Credential > $null

            [String]$SearchQuery = $this.QueryBuilder()

            $identity = $this.Inbox.toString()
    
            [System.Object]$Results = Search-Mailbox -Identity $identity -SearchQuery $SearchQuery -DeleteContent -ErrorAction Stop -WarningAction silentlyContinue -Force

            #Return results
            return $Results.ResultItemsCount
        }

        #Create Search Query 
        $email | Add-Member -Type ScriptMethod -Name QueryBuilder -Force -Value {
            
            $email = $this

            [string]$query = ""

            foreach ($keyobj in ($email | Get-Member -MemberType NoteProperty  | Select-Object Name)){
               
                #get key and value from current email object
                $key = $keyobj.Name
                $value = $email.($key)
                 #create query
                if(($key -notmatch "inbox") -and ($query -eq "") -and ($value -ne "")){
                        $query = "" + $key + ":" + "`"" + $value + "`""
                    }
                elseif(($key -notmatch "inbox") -and ($value -ne "")){
                        $query =$query + " " + $key + ":" + "`"" + $value + "`""   
                    }
                else{#none
                }
                   
            }
            return $query
        }

    #return object
    return $email
}
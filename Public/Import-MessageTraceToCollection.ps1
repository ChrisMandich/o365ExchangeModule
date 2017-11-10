#Import message trace results to our collection
Function Import-MessageTraceToCollection {
<#
    .Synopsis
    Receives Message Trace Objects from Message Trace and adds them to the collection

    .Description
    Receives Message Trace Objects from Message Trace and adds them to the collection. 

    .PARAMETER tempCollection
    Receives Message Trace Objects from Message Trace

    .Example 
    Get-MessageTrace | ImportMessageTraceToCollection
#>
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$True,
		ValueFromPipeline=$True)]
		[Object[]]$tempCollection
	)
	BEGIN {
        Write-Host "Converting Message Trace"
    }
	PROCESS {
        $tempCollection | Select RecipientAddress, SenderAddress, Received, Subject | ForEach-Object{
            #Add Inbox Check
            $RecipientAddress = $_.RecipientAddress
            $SenderAddress = $_.SenderAddress
            $Subject = $_.Subject
            $Received = $_.Received

            if(Get-Mailbox -Identity $RecipientAddress){
                Write-Host "Inbox:`"$($RecipientAddress)`" From:`"$($SenderAddress)`" Subject:`"$($Subject)`" Received:`"$($Received)`""
                Add-ToCollection -from $SenderAddress -subject $subject -inbox $RecipientAddress -received $received -to "" -participants "" -kind "email" -cc "" -bcc "" -body "" -attachment "" -sent ""
            }
            elseif(Get-Mailbox -Identity $SenderAddress){
                Write-Host "Inbox:`"$($SenderAddress)`" From:`"$($RecipientAddress)`" Subject:`"$($Subject)`" Received:`"$($Received)`""
                Add-ToCollection -to $RecipientAddress -subject $_.subject -inbox $SenderAddress -received $_.received -to "" -participants "" -kind "email" -cc "" -bcc "" -body "" -attachment "" -sent ""
            }
            else{
                Write-Host "Unable to find Inbox.`n`tRecipient: $RecipientAddress `n`tSender: $SenderAddress" -ForegroundColor Red    
            }
        }
    }#End Process
	END {
        Write-Host "Import Complete"
    }
}
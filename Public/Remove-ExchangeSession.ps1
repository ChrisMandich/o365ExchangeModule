# Close a session with exchange
function Remove-ExchangeSession{
<#
    .SYNOPSIS
    Remove current Exchange Session based on Session Name: PSSecComp and PSOutlook.

    .DESCRIPTION  
    This function removes the remote session specified in $session variable.
 
    .EXAMPLE
    Remove-ExchangeSession
       
#>
    #Remove TMP_* Modules connected to Outlook 
    Get-Module tmp_* |% {if($_.Description -match "outlook.com" ) {Write-Output "Remove Module $($_.Name)" ; remove-module $_.name}}
    #Remove PSSession's that Match Outlook.com
    Get-PSSession | % { if( $_.ComputerName -match "outlook.com" ){Write-Output "Remove PSSession $($_.Name)"; Remove-PSSession -Id $_.id}  } 
}


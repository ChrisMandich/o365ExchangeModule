# Open a session with exchange
function Enter-ExchangeSession{
<#
    .SYNOPSIS

    Enter an Exchange Session. 

    .DESCRIPTION
    
    This function creates a remote session with "ps.outlook.com/powershell". Imports it globally allowing the user to interact with Exchange Online in the current PS Window. Imports Get-Mailbox, Search-Mailbox, and Get-MessageTrace.  

    .PARAMETER proxy

    Uses $proxysettings variable to establish session.

    
    .EXAMPLE

    Enter-ExchangeSession
       
    Create a session with Exchange

    .EXAMPLE

    Enter-ExchangeSession -proxy

    Create a session using proxy settings defined in $proxysettings variable.

#>

    Param(
        [switch]$proxy,
        [System.Object]$Credential = $(Get-Credential -Message "Enter your Office 365 Credential")
    )
    
    Try{
        #Set Session Options. Proxy and 10 Minute TimeOut
        #$soProxySettings = $($global:proxysettings)
        $soProxySettings = @($($global:proxysettings),$(New-PSSessionOption -IdleTimeout 600000))

        $so = New-PSSessionOption -IdleTimeout 600000

        #PowerShell Session URL's
        $PSOutlookURL = "https://ps.outlook.com/powershell"
        #$PSSecCompURL = "https://ps.compliance.protection.outlook.com/powershell-liveid"

        #Check to see if System Proxy Flag is Set.    
        if ($proxy.IsPresent){
            #Create new remote powershell session with exchange
            $global:PSOutlookSession = New-PSSession -Name "PSOutlook" -ConfigurationName Microsoft.Exchange -ConnectionUri $PSOutlookURL -Credential $Credential -WarningAction SilentlyContinue -Authentication Basic -AllowRedirection -SessionOption $soProxySettings -ErrorAction Stop
            #$global:PSSecComp = New-PSSession -Name "PSSecComp" -ConfigurationName Microsoft.Exchange -ConnectionUri $PSSecCompURL -Credential $Credential -WarningAction SilentlyContinue -Authentication Basic -AllowRedirection -SessionOption $soProxySettings -ErrorAction Stop 
        }
        else{
            #Create new remote powershell session with exchange
            $global:PSOutlookSession = New-PSSession -Name "PSOutlook" -ConfigurationName Microsoft.Exchange -ConnectionUri $PSOutlookURL -Credential $Credential -WarningAction SilentlyContinue -Authentication Basic -AllowRedirection -SessionOption $so -ErrorAction Stop
            #$global:PSSecComp = New-PSSession -Name "PSSecComp" -ConfigurationName Microsoft.Exchange -ConnectionUri $PSSecCompURL -Credential $Credential -WarningAction SilentlyContinue -Authentication Basic -AllowRedirection -SessionOption $so -ErrorAction Stop
        }
        
        #Specify import command names
        $tmpPSOutlookCommandNames = @("Get-MessageTrace", "Get-Mailbox", "Search-Mailbox" , "*-MailboxSearch")
        #$tmpSecCompImpCommandName = @("*compliancesearch*")

        #import PSSessions
        Import-Module ($global:importPSOutlookSession = Import-PSSession $global:PSOutlookSession -CommandName $tmpPSOutlookCommandNames -AllowClobber -DisableNameChecking -ErrorAction Stop) -Global -DisableNameChecking -Force 
        #import-module ($global:importPSSecCompSession = Import-PSSession $global:PSSecComp -CommandName $tmpSecCompImpCommandName -AllowClobber -DisableNameChecking -ErrorAction Stop) -Global -DisableNameChecking -Force 

        #output results
        $global:FormatEnumerationLimit = -1
        Write-Output "`nCreated session with Office 365`n"
        Write-Output "`tImported PowerShell Session Id: $($global:PSOutlookSession.Id)" 
        #Write-Output "`tImported Security and Compliance Center ID: $($global:PSSecComp.Id)"  
        
    }
    Catch{
        #Output Error information
        Write-Output "Session Failed to open`nCheck Creds or Proxy`n"
        Write-Host -ForegroundColor Red $_.Exception
    }
}
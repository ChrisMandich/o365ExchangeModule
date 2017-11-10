<#
Chris Mandich
version 0.0.4

X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X 
X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X 

CHANGELOG

20160506 CM -- Made get-help more relevant. Added additional examples and parameter explanations.
20160506 CM -- Added Distribution group dump function. This output all members of a distribution group.
20160506 CM -- Added participants flag for collection. This flag searches the following fields: "from, to, bcc, cc"
20160506 CM -- Added QueryBuilder function, this just prevents code reuse. Not useful to anyone.
20160506 CM -- Added test if mailbox exists before search. One less error to deal with. 
20160505 CM -- Fixed scoping issue in enter-exchangesession. This remote 365 commands to be run in your current session.
20160505 CM -- Seriously cleaned up collection array making it easier for everyone to use.
20170315 CM -- Added Convert-Message Trace Function
20170320 CM -- Updated Enter-ExchangeSession to incorporate the Security and Compliance commands and limit Exchange to Get-Mailbox, Search-Mailbox, and Get-Message Tracking
20170320 CM -- Updated Remove-ExchangeSession to removes sessions based on names that are specified when enter-exchange is used. 
20170424 CM -- Added Get-MessageTraceCustom - This iterates through all 200 pages of MessageTrace Events
20170424 CM -- Updated Copy-* and Remove-* to use Custom New-EmailObject
20170424 CM -- Cleaned up Convert-MessageTrace Function
20171109 CM -- Removed Compliance Related PS Session, Converted Module Structure to individual functions/commandlets. 

X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X 
X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X X 

#>

$global:proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
$global:PSOutlookSession;
$global:importPSOutlookSession;

<#
Enabled for PS Compliance Sessions
$global:importPSSecCompSession;
$global:PSSecCompSession;
#>
[System.Collections.ArrayList]$global:collection = @();

#Get public and private function definition files.
    $Public  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
    $Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )
    $ModuleRoot = $PSScriptRoot

#Dot source the files
    Foreach($import in @($Public + $Private))
    {
        Try
        {
            . $import.fullname
        }
        Catch
        {
            Write-Error -Message "Failed to import function $($import.fullname): $_"
        }
    }

#Specify allowed functions for users.
Export-ModuleMember -Function $Public.Basename



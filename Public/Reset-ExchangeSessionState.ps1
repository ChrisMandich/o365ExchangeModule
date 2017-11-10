function Reset-ExchangeSessionState{
    Param(
        [switch]$proxy,
        [System.Object]$Credential = $(Get-Credential -Message "Enter your Office 365 Credential")
    )    

    #Get sessions that have outlook.com in the computer name
    $tmpSession = Get-PSSession | where ComputerName -Match "outlook.com"

    if($tmpSession ){
    #end session with Microsoft Exchange
    $tmpSession | ForEach-Object{
            if($_.State -notcontains "Opened"){
                #Remove-ExchangeSession if they exist. 
                Remove-ExchangeSession 

                #Check Proxy Flag 
                if ($proxy.IsPresent){
                    Enter-ExchangeSession -Credential $Credential -proxy 
                    }
                else{
                    Enter-ExchangeSession -Credential $Credential
                }
            
                break;
            }
            else{
                write-output "`"$($_.Name)`" is Open." 
            }
        }
    }
    #If a session has not been established, ask for it now. 
    else{
        #Check Proxy Flag 
        if ($proxy.IsPresent){
            Enter-ExchangeSession -Credential $Credential -proxy 
            }
        else{
            Enter-ExchangeSession -Credential $Credential
        }
    }
}

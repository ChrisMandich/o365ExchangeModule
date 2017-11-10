Function Get-MessageTraceCustom {
<#
    .Synopsis
    Custom Message Trace Search. 

    .Description
    Creates a Custom Message Trace Search, Iterates through all 200 Pages and set default PageSize to 5000.

    .PARAMETER SenderAddress
    Message Trace Sender

    .PARAMETER RecipientAddress
    Message Trace Recipient

    .PARAMETER FromIp
    IP of Source MG

    .PARAMETER ToIp
    IP Of Destination MG


    .PARAMETER Status
    Message Trace Status, Defaults to Delivered

    .PARAMETER StartDate
    Set Start Date

    .PARAMETER EndDate
    Set End Date

    .Example 
    Get-MessageTraceCustom -StartDate $(get-date).AddDays(-3) -SenderAddress "test@email.com"
#>
	param(
	    $SenderAddress="",
        $RecipientAddress="", 
        $FromIp="", 
        $ToIp="",
        $Status = "Delivered",
        $StartDate = $(get-date).AddDays(-7), 
        $EndDate = $(get-date),
        $StartPage = 1,
        $EndPage = 200
	)
    
    #Set Acceptable Arguments
    $Arguments = @{
        "SenderAddress"=$SenderAddress;
        "RecipientAddress"=$RecipientAddress;
        "FromIp"=$FromIp;
        "ToIp"=$ToIp;
        "Status"=$Status;
        "StartDate"=$StartDate;
        "EndDate"=$EndDate;
        "PageSize"="5000";
    }

    #Create For Loop that navigates through the first 200 pages. 
    $newMessageTraceCustom = "foreach(`$number in `$StartPage..`$EndPage) {get-messagetrace -page `$number "

    $Arguments.Keys |  ForEach-Object {
        if($Arguments[$($_)]){
            $newMessageTraceCustom += "-" + $_ + " `'$($Arguments[$($_)])`'"
        }
    }

    $newMessageTraceCustom += "}"

    invoke-expression $newMessageTraceCustom
}
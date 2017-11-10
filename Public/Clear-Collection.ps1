# Clear Inbox Array
function Clear-Collection{
<#
    .Synopsis
    This empties the email collection.

    .Description
    Seriously though. It clears the array. 

    .Example 
    Clear-Collection
#>

    $global:collection.Clear()
}
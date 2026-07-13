function Get-UserGroupMemberships {
    <#
    .SYNOPSIS
    Retrieves Microsoft Entra ID group memberships for a user.

    .DESCRIPTION
    Retrieves all direct Microsoft Entra ID group memberships for the specified
    user and returns the group display name, type and description.

    .PARAMETER UPN
    The User Principal Name (UPN) of the user whose group memberships are to be
    retrieved.

    .EXAMPLE
    Get-UserGroupMemberships -UPN "john.smith@contoso.com"

    Returns all direct group memberships for the specified user.

    .OUTPUTS
    PSCustomObject containing:

    - DisplayName
    - Type
    - Description

    .NOTES
    Requires Microsoft Graph permissions:
    - Group.Read.All
    or
    - Group.ReadWrite.All

    #>
    param (
        [Parameter(Mandatory)][string]$UPN
    )

    begin { }
    process {

        Get-MgUserMemberOf -UserId $UPN |

        ForEach-Object {
            $group = Get-MgGroup -GroupId $_.Id
            [PSCustomObject]@{
                DisplayName = $group.DisplayName
                Type        = if ($group.GroupTypes -contains 'Unified') { 'Microsoft 365' }
                elseif ($group.SecurityEnabled) { 'Security' }
                else { 'Other' }
                Description = $group.Description
            }
        }
    }
    end { return $groups }
}


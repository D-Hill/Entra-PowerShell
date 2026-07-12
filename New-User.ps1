
<#
.SYNOPSIS
Creates a new Microsoft Entra ID user.

.DESCRIPTION
Creates a new user account in Microsoft Entra ID using Microsoft Graph PowerShell.
Optional user attributes such as department, job title and office location can be supplied.

.PARAMETER DisplayName
The display name of the new user.

.PARAMETER UserPrincipalName
The User Principal Name (UPN) for the new user.

.PARAMETER MailNickname
The mail nickname (alias) for the new user.

.PARAMETER GivenName
The user's first name.

.PARAMETER Surname
The user's surname.

.PARAMETER Password
The initial password for the user.

.PARAMETER ForceChangePasswordNextSignIn
Determines whether the user must change their password at first sign-in.

.PARAMETER AccountEnabled
Determines whether the account is enabled after creation.

.PARAMETER UsageLocation
The user's usage location. Defaults to GB.

.PARAMETER Department
The user's department.

.PARAMETER JobTitle
The user's job title.

.PARAMETER Office
The user's office location.

.PARAMETER GroupIds
Optional array of Entra group IDs to add the user to.

.EXAMPLE
New-User -DisplayName "AN Other" `
         -UserPrincipalName "AN.Other@contoso.com" `
         -MailNickname "AN.Other"

Creates a new Entra ID user.

.NOTES
Requires Microsoft Graph PowerShell permissions:
- User.ReadWrite.All

Author: D-Hill
Version: 1.0
#>
function New-User {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [string]$DisplayName,
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory)]
        [string]$MailNickname,
        [string]$GivenName,
        [string]$Surname,
        [string]$Password,
        [bool]$ForceChangePasswordNextSignIn = $true,
        [bool]$AccountEnabled = $true,
        [string]$UsageLocation = "GB",
        [string]$Department,
        [string]$JobTitle,
        [string]$Office,
        [string[]]$GroupIds
    )

    begin {
        # TODO: verify Connect-MgGraph session / required scopes
    }

    process {
        $passwordProfile = @{
            Password                      = $Password
            ForceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn
        }

        $userParams = @{
            DisplayName       = $DisplayName
            UserPrincipalName = $UserPrincipalName
            MailNickname      = $MailNickname
            AccountEnabled    = $AccountEnabled
            UsageLocation     = $UsageLocation
            PasswordProfile   = $passwordProfile
        }

        if ($GivenName)  { $userParams.GivenName      = $GivenName }
        if ($Surname)    { $userParams.Surname         = $Surname }
        if ($Department) { $userParams.Department      = $Department }
        if ($JobTitle)   { $userParams.JobTitle        = $JobTitle }
        if ($Office)     { $userParams.OfficeLocation  = $Office }

        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Create new Entra ID user")) {
             New-MgUser @userParams
        }
    }

    end {
    }
}

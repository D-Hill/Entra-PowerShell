

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

$NewUserParams = @{
    DisplayName                  = 'Aidan Whitfield'
    GivenName                    = 'Aidan'
    Surname                      = 'Whitfield'
    MailNickname                 = 'aidan.whitfield'
    JobTitle                     = 'Chief Executive Officer'
    Password                     = 'ChangeMe1!'
    UserPrincipalName            = 'aidan.whitfield@duncanjameshillgmailcom.onmicrosoft.com'
    Office                       = 'London'
    ForceChangePasswordNextSignIn = $true
}

New-User @NewUserParams
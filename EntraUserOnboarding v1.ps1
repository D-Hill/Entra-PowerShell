
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
        # TODO: generate password if not supplied

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

        if ($GivenName) { $userParams.GivenName = $GivenName }
        if ($Surname) { $userParams.Surname = $Surname }
        if ($Department) { $userParams.Department = $Department }
        if ($JobTitle) { $userParams.JobTitle = $JobTitle }
        if ($Office) { $userParams.OfficeLocation = $Office }

        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Create new Entra ID user")) {
            New-MgUser @userParams

            # TODO: optionally add to groups via $GroupIds (New-MgGroupMember)

            # TODO: output created user object
        }
    }

    end {
    }
}

function Confirm-Headers {

    param (
        [Parameter(Mandatory)]
        [string]$file
    )


    $expectedHeaders = @(
        'DisplayName',
        'GivenName',
        'Surname',
        'UserPrincipalName',
        'MailNickname',
        'JobTitle',
        'Department',
        'Office',
        'ManagerUserPrincipalName',
        'Password'
    )
    $csvHeaders = (Get-Content $file -First 1) -split ','
    $missing = $expectedHeaders | Where-Object { $_ -notin $csvHeaders }
    $extra = $csvHeaders | Where-Object { $_ -notin $expectedHeaders }

    if ($missing -or $extra) {
        Write-Warning "Header mismatch in $($file.Name):"
        if ($missing) { Write-Warning "  Missing: $($missing -join ', ') - Please correct input file and try again" }
        if ($extra) { Write-Warning "  Unexpected: $($extra -join ', ') - Please remove or correct this header as it is not configured in script" }
        $headersValidated = $false
        break

    }
    else {
        #Write-Host "Importing $($file.Name): headers OK - Moving to next step" -ForegroundColor Green
        $headersValidated = $true
    }

    $headersValidated
}

$dir = "C:\NewUsers"

$logFile = "$dir\Log\UsersCreated.csv"
if (!(test-path $logFile)) { write-warning 'No log file found - Please ensure the log file is in the directory - $($logfile)' ; break }

# checks import files
Write-host "Checking import file and verifying headers"
$importFile = Get-ChildItem -Path "$dir\Import" -Filter *.csv 

If ($importFile.count -gt 1) { write-warning "More than one file in the Import folder - Please ensure only one file is present and try again" }

# validate headers
if (Confirm-Headers -file $dir\import\$importFile) { Write-host "Checked headers of input file - All OK" -ForegroundColor green }
# import users
$users = Import-Csv $dir\import\$importFile | select -first 5

#trim all entries
foreach ($user in $users) {
    foreach ($property in $user.PSObject.Properties) {
        if ($property.Value -is [string]) {
            $property.Value = $property.Value.Trim()
        }
    }
}

read-host "$(($users | measure-object).count) users found in the input file. Any key to continue"

# main execution
foreach ($user in $users) {
    read-host "`Any key to continue to next user - $($user.userprincipalname)"

    $existingUser = $null
    $existingUser = Get-MgUser -UserId $user.UserPrincipalName -ErrorAction SilentlyContinue
    if ($existingUser) { write-warning "User $($existingUser.DisplayName) already exists in the tenant. Skipping to next user in input file" ; continue }

    if (-not $existingUser) {

        write-host "Creating $($user.UserPrincipalName)"
        try {
            $mailNickname = $null
            $mailNickname = $user.UserPrincipalName.Split("@")[0]
            $NewUserParams = @{

                DisplayName       = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                MailNickname      = $mailNickname
                GivenName         = $user.GivenName
                Surname           = $user.Surname
                Password          = $user.Password
                Department        = $user.Department
                JobTitle          = $user.JobTitle
                Office            = $user.Office
            }

            New-User @NewUserParams -ErrorAction Stop
            Write-Host "Successfully created $($user.UserPrincipalName)" -ForegroundColor Green

            $LogEntry = [PSCustomObject]@{
        
                DisplayName              = $user.DisplayName
                UserPrincipalName        = $user.UserPrincipalName
                JobTitle                 = $user.JobTitle
                Department               = $user.Department
                Office                   = $user.Office
                ManagerUserPrincipalName = $user.ManagerUserPrincipalName
                CreatedBy                = $env:USERNAME
                CreatedDate              = Get-Date -Format dd-MM-yyyy-HH:mm
                Status                   = "Created"
            }

            $LogEntry | Export-Csv -Path $LogFile -NoTypeInformation -Append

        }
        catch {
            Write-Warning "Failed to create $($user.UserPrincipalName)"
            Write-Warning $_.Exception.Message
        }

    }

  
}


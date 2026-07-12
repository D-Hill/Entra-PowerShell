
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

function Get-GroupAssignments {


    param (
        [Parameter(Mandatory)]
        [object]$user
    )

    $groupsToAdd = @()

    # All Users
    $groupsToAdd += Get-MgGroup -Filter "displayname eq 'SG-All Users'"

    # Offices
    If ($user.office -like 'London') { $groupsToAdd += Get-MgGroup -Filter "displayname eq 'SG-London Office'" }
    If ($user.office -like 'Manchester') { $groupsToAdd += Get-MgGroup -Filter "displayname eq 'SG-Manchester Office'" }
    If ($user.office -like 'Edinburgh') { $groupsToAdd += Get-MgGroup -Filter "displayname eq 'SG-Edinburgh Office'" }
    If ($user.office -like 'Birmingham') { $groupsToAdd += Get-MgGroup -Filter "displayname eq 'SG-Birmingham Office'" }

    # Departments
    If ($user.Department -like 'Finance') { $groupsToAdd += Get-MgGroup -Filter "displayname eq 'SG-Finance'" }
    If ($user.Department -like 'Information Technology') { $groupsToAdd += Get-MgGroup -Filter "displayname eq 'SG-Information Technology'" }
    If ($user.Department -like 'HR') { $groupsToAdd += Get-MgGroup -Filter "displayname eq 'SG-Human Resources'" }
    If ($user.Department -like 'Executive') { $groupsToAdd += Get-MgGroup -Filter "displayname eq 'SG-Executive'" }

    # titles
    If ($user.jobtitle -like 'IT Support Analyst') { $groupsToAdd += Get-MgGroup -Filter "displayname eq 'IT Support Hub'" }


    return $groupsToAdd

}

$dir = "C:\NewUsers"

$logFile = "$dir\Log\UsersCreated.csv"
if (!(test-path $logFile)) { write-warning 'No log file found - Please ensure the log file is in the directory - $($logfile)' ; break }

# checks import files
Write-host "Checking import file and verifying headers"
$importFile = Get-ChildItem -Path "$dir\Import" -Filter *.csv 

If ($importFile.count -gt 1) { write-warning "More than one file in the Import folder - Please ensure only one file is present and try again" }

write-host "`n *** Starting User Provisioninfg Script ***" -ForegroundColor cyan

# validate headers
if (Confirm-Headers -file $dir\import\$importFile) { Write-host "`n - Checked headers of input file - All OK" -ForegroundColor green }
# import users
$users = Import-Csv $dir\import\$importFile | Select-Object -first 18

#trim all entries
foreach ($user in $users) {
    foreach ($property in $user.PSObject.Properties) {
        if ($property.Value -is [string]) {
            $property.Value = $property.Value.Trim()
        }
    }
}

write-host " -  $(($users | measure-object).count) users found in the input file" -ForegroundColor green

# main execution

$i = 0
foreach ($user in $users) {

    $i++

    Write-host "`nProcessing row $i of $($users.count) - $($user.UserPrincipalName)" -foregroundcolor cyan
  #  read-host "`Any key to continue"

    $existingUser = $null
    $existingUser = Get-MgUser -UserId $user.UserPrincipalName -ErrorAction SilentlyContinue
    if ($existingUser) { Write-host "User $($existingUser.DisplayName) already exists in the tenant with the same UPN. This user will be skipped" -foregroundcolor yellow; continue }

    if (-not $existingUser) {

        write-host "Creating new user $($user.UserPrincipalName)" -ForegroundColor green
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

            $newUser = $null
            $newUser = New-User @NewUserParams -ErrorAction Stop 
          
        }
        catch {
            Write-Warning "Failed to create $($user.UserPrincipalName)"
            Write-Warning $_.Exception.Message
        }

        if ($newUser) { 
            write-host "`n - User $($newUser.DisplayName) created successfully" -ForegroundColor green

            Write-host "`n - Adding security groups" -ForegroundColor Cyan
            $groupsToAdd = Get-GroupAssignments -user $user

            foreach ($group in $groupsToAdd) {

                write-host " -- Adding $($newUser.DisplayName) to group $($group.DisplayName)" -ForegroundColor cyan
                New-MgGroupMember -GroupId $group.Id  -DirectoryObjectId $newUser.Id 
            
            }

            write-host "`n - Adding log file entry for $($user.DisplayName)" -ForegroundColor blue
       
            $LogEntry = [PSCustomObject]@{
        
                DisplayName              = $user.DisplayName
                UserPrincipalName        = $user.UserPrincipalName
                JobTitle                 = $user.JobTitle
                Department               = $user.Department
                Office                   = $user.Office
                ManagerUserPrincipalName = $user.ManagerUserPrincipalName
                GroupsAdded              = $groupsToAdd.displayname -join ';'
                CreatedBy                = (Get-MgContext).Account
                CreatedDate              = Get-Date -Format dd-MM-yyyy-HH:mm
                Status                   = "Created"
            }

            $LogEntry | Export-Csv -Path $LogFile -NoTypeInformation -Append
       
        }

    }
    
}




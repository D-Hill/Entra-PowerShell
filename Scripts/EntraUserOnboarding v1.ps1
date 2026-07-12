<#
.SYNOPSIS
Automates Microsoft Entra ID user onboarding.

.DESCRIPTION
This script provisions new Microsoft Entra ID users from a CSV input file.

The script performs the following tasks:

- Validates Microsoft Graph connectivity and required permissions
- Validates CSV input file structure and user data
- Checks whether users already exist in the tenant
- Creates new Entra ID users using Microsoft Graph PowerShell
- Assigns users to security groups based on attributes such as:
    - Office location
    - Department
    - Job title
- Creates an audit log entry for each successfully provisioned user

.INPUTS
CSV file containing user provisioning data.

Required CSV fields:
- DisplayName
- GivenName
- Surname
- UserPrincipalName
- JobTitle
- Department
- Office
- Password

.OUTPUTS
CSV audit log containing:

- Created user details
- Groups assigned
- Account that executed the script
- Creation timestamp
- Provisioning status

.REQUIREMENTS
Microsoft Graph PowerShell SDK

Required Microsoft Graph permissions:

- User.ReadWrite.All
- Group.ReadWrite.All

.EXAMPLE
.\Invoke-EntraUserOnboarding.ps1

Runs the onboarding process using the CSV file located in the configured import directory.

.NOTES
Author: D-Hill
Version: 1.0

This script is intended for automating user lifecycle onboarding
within Microsoft Entra ID environments.

#>
function New-User {


    <#
.SYNOPSIS
Creates a new Microsoft Entra ID user.

.DESCRIPTION
Creates a new user account in Microsoft Entra ID using Microsoft Graph PowerShell.
Optional attributes such as department, job title and office location can be supplied.

.PARAMETER DisplayName
Display name of the new user.

.PARAMETER UserPrincipalName
User Principal Name (UPN) for the new user.

.PARAMETER MailNickname
Mail nickname/alias for the user.

.PARAMETER GroupIds
Optional array of Entra group IDs to assign.

.EXAMPLE
New-User @NewUserParams

Creates a new Entra ID user using supplied parameters.

.NOTES
Requires Microsoft Graph permissions:
- User.ReadWrite.All

#>


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

        if ($GivenName) { $userParams.GivenName = $GivenName }
        if ($Surname) { $userParams.Surname = $Surname }
        if ($Department) { $userParams.Department = $Department }
        if ($JobTitle) { $userParams.JobTitle = $JobTitle }
        if ($Office) { $userParams.OfficeLocation = $Office }

        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Create new Entra ID user")) {
            New-MgUser @userParams
        }
    }

    end {
    }
}

function Confirm-Headers {
    <#
.SYNOPSIS
Validates CSV file headers.

.DESCRIPTION
Checks that the input CSV contains all expected headers required by the
user provisioning process. Also identifies unexpected headers that are not
configured for use by the script.

.PARAMETER File
Path to the CSV input file.

.EXAMPLE
Confirm-Headers -File ".\Users.csv"

Validates the structure of the input CSV file.

.NOTES
Used as a pre-validation step before user provisioning begins.

#>
    param (
        [Parameter(Mandatory)]
        [string]$file
    )


    $expectedHeaders = @(
        'DisplayName',
        'GivenName',
        'Surname',
        'UserPrincipalName',
        #  'MailNickname',
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

        if ($missing) { Write-Warning " - Missing header in input file: $($missing -join ', ') - Please correct input file and try again" }
        if ($extra) { Write-Warning " - Unexpected header in input file: $($extra -join ', ') - Please remove or correct this header as it is not configured in script" }
        $headersValidated = $false
        exit

    }
    else {
        #Write-Host "Importing $($file.Name): headers OK - Moving to next step" -ForegroundColor Green
        $headersValidated = $true
    }

    $headersValidated
}
function Confirm-CsvEntries {
    <#
.SYNOPSIS
Validates CSV user data.

.DESCRIPTION
Checks that mandatory fields contain values and validates that User Principal
Names are in the correct format before attempting user creation.

.PARAMETER File
Path to the CSV input file.

.EXAMPLE
Confirm-CsvEntries -File ".\Users.csv"

Validates user records before provisioning.

.NOTES
Returns:
$true  - Validation successful
$false - Validation failed

#>
    param (
        [Parameter(Mandatory)]
        [string]$File
    )

    $requiredFields = @(
        'DisplayName',
        'GivenName',
        'Surname',
        'UserPrincipalName',
        #  'MailNickname',
        'JobTitle',
        'Department',
        'Office',
        #     'ManagerUserPrincipalName',
        'Password'
    )

    $users = Import-Csv $File
    $errors = @()
    $row = 1
    foreach ($user in $users) {
        $row++
        foreach ($field in $requiredFields) {

            if ([string]::IsNullOrWhiteSpace($user.$field)) {
                $errors += " -- Checked content of input file.  $field is missing in row $row. $field is a madatory field. Please update and retry." 
            }

        }
        # Validate UPN format
        if ($user.UserPrincipalName -and 
            $user.UserPrincipalName -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {

            $errors += " -- Invalid UPN format in row $row - $($user.UserPrincipalName). Please update and retry."
        }
    }

    if ($errors.Count -gt 0) {

        #  Write-Warning "- CSV validation failed:"

        foreach ($error in $errors) {
            Write-host "$error" -foregroundcolor red
        }

        return $false
     
    }
    else {

        #   Write-Host " --- CSV entries validated successfully" -ForegroundColor Green
        return $true

    }
}
function Get-GroupAssignments {

    <#
.SYNOPSIS
Determines Entra ID groups for a user.

.DESCRIPTION
Returns a collection of Microsoft Entra ID groups that a user should be
assigned to based on their office location, department and job title.

The function evaluates user attributes and retrieves matching groups from
Microsoft Graph.

.PARAMETER User
User object containing attributes used for group assignment.

.EXAMPLE
$Groups = Get-GroupAssignments -User $User

Returns groups applicable to the user.

.NOTES
Requires Microsoft Graph permissions:
- Group.ReadWrite.All

#>
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
write-host "`n *** Starting Onboarding Script ***" -ForegroundColor green

# sets directory
$dir = "C:\NewUsers"

# Check Microsoft Graph connection
$GraphContext = Get-MgContext

write-host "`n - Checking connection to Microsoft Graph" -ForegroundColor Cyan

# connect to Grpah is no open session
if (-not $GraphContext) {
    Write-Host " -- Connecting to Microsoft Graph with User write permissions" -ForegroundColor Yellow

    # connects to graph and discards message for console readability
    $graphConnectionMessage = Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All" 
    if ( Get-MgContext ) { Write-Host " --- Connection to Microsoft Graph successful" -ForegroundColor Green } else { Write-warning " --- Connection to Microsoft Graph failed" ; exit }
}
else {
    Write-Host " -- Existing Microsoft Graph session found for $($GraphContext.Account)" -ForegroundColor Green
}

# checks import files
write-host "`n - Checking input and log file" -ForegroundColor Cyan
$importFile = Get-ChildItem -Path "$dir\Import" -Filter *.csv 
If ($importFile.count -gt 1) { write-warning " -- More than one file in the Import folder - Please ensure only one file is present and try again" ; exit }

# validate headers
if (Confirm-Headers -file $dir\import\$importFile) { Write-host " -- Checked headers of input file - All OK" -ForegroundColor green }
# import users
$users = Import-Csv $dir\import\$importFile | Select-Object -first 15
# validate contents
if (Confirm-CSVEntries -file $dir\import\$importFile) { Write-host " -- Checked contents of input file - All appears OK" -ForegroundColor green } else { exit }

write-host " -- $(($users | measure-object).count) users found in the input file" -ForegroundColor green

#trim all entries
write-host ' -- trimming CSV entries' -ForegroundColor Green
foreach ($user in $users) {
    foreach ($property in $user.PSObject.Properties) {
        if ($property.Value -is [string]) {
            $property.Value = $property.Value.Trim()
        }
    }
}
# checks logfile exists
$logFile = "$dir\Log\UsersCreated.csv"
if ((test-path $logFile)) { write-host ' -- Log file found' -ForegroundColor Green }
if (!(test-path $logFile)) { write-warning ' -- No log file found - Please ensure the log file is in the directory - $($logfile)' ; exit }

read-host "`nRun onboarding script? any key to continue"

# main execution - runs through all rows in input file
$i = 0
foreach ($user in $users) {

    # counter goes up by 1
    $i++

    Write-host "`nProcessing row $i of $($users.count) - $($user.UserPrincipalName)" -foregroundcolor cyan

    # checks for existing user by UPN
    $existingUser = $null
    $existingUser = Get-MgUser -UserId $user.UserPrincipalName -ErrorAction SilentlyContinue
    if ($existingUser) { Write-host " - User $($existingUser.UserPrincipalName) already exists in the tenant. This user will be skipped" -foregroundcolor yellow; continue }

    # if no user exists with that UPN create user
    if (-not $existingUser) {

        write-host " - Creating new user $($user.UserPrincipalName)" -ForegroundColor yello
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
            # calls New-User function
            $newUser = $null
            $newUser = New-User @NewUserParams -ErrorAction Stop 
        }
        catch {
            Write-Warning " - Failed to create $($user.UserPrincipalName)"
            #Write-Warning $_.Exception.Message
        }
        # if a new user has been created, add baseline security groups based on Office and Department
        if ($newUser) { 
            write-host " -- User $($newUser.UserPrincipalName) created successfully" -ForegroundColor green

            Write-host "`n - Adding security groups" -ForegroundColor Cyan
            
            # gets groups assignments based on office and department values
            $groupsToAdd = Get-GroupAssignments -user $user

            # adds groups
            foreach ($group in $groupsToAdd) {
                write-host " -- Adding $($newUser.DisplayName) to group $($group.DisplayName)" -ForegroundColor green
                New-MgGroupMember -GroupId $group.Id  -DirectoryObjectId $newUser.Id 
            }

            # appends log of users created to csv file
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

write-host "`n *** Script complete ***" -ForegroundColor green


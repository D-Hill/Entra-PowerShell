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

   $groups =       Get-MgUserMemberOf -UserId $UPN -All |     Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' } |

       ForEach-Object {
            $group = Get-MgGroup -GroupId $_.Id
            [PSCustomObject]@{
                DisplayName = $group.DisplayName
                Type        = if ($group.GroupTypes -contains 'Unified') { 'Microsoft 365' }
                elseif ($group.SecurityEnabled) { 'Security' }
                else { 'Other' }
                Description = $group.Description
                Decision    = ''
            }
        }
    }
    end { return $groups | Where-Object {$_.displayname} | Sort-Object DisplayName }
}

# import the importExcel module
try { import-module ImportExcel } catch { write-warning "please instal the Importexcel PowerShell module and rerun script" ; exit}

# sets repoprt directory
$dir = 'C:\UserReports'

write-host "`n*** Starting User Access Report Script ***" -ForegroundColor Cyan
# connect to Grpah is no open session

if (-not (Get-MgContext)) {
    Write-Host "`n - Connecting to Microsoft Graph with Read permissions"
    # connects to graph and discards message for console readability
    $null = Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" 
    if ( Get-MgContext ) { Write-Host " --- Connection to Microsoft Graph with Read permissions successful" -ForegroundColor Green } else { Write-warning " --- Connection to Microsoft Graph failed" ; exit }
}
else {
    Write-Host "`n - Existing Microsoft Graph session found for $($GraphContext.Account)" -ForegroundColor Cyan
}
if (-not $users) {
    $users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, Manager, OfficeLocation, Department, Jobtitle
}
# create hash table
write-host "`n - Getting user information" -ForegroundColor Cyan 
$userHashtable = @{}
foreach ($user in $users) {
    $userHashtable[$user.UserPrincipalName] = $user
}
# get all users with direct reports
write-host " - Getting managers and reports" -ForegroundColor Cyan 
$managersAndReports = $users |  Where-Object { Get-MgUserDirectReport -UserId $_.Id -ErrorAction SilentlyContinue } | 
ForEach-Object {
    $mgr = $_
    $directReports = (Get-MgUserDirectReport -UserId $_.Id -Property Id).id
        
    $directReports | ForEach-Object {

        $report = $null
        $report = Get-MgUser -UserId $_
        [PSCustomObject]@{
      
            ManagerUPN      = $mgr.UserPrincipalName
            DirectReportUPN = $report.UserPrincipalName
        }
    }
} 
# Get unique list of manager UPNs and retrieve their user objects from the hashtable
$managerList = $managersAndReports.ManagerUPN | Select-Object -Unique 
$managers = foreach ($managerUPN in $managerList) { $userHashtable[$managerUPN] }

# Loop through each manager creating reports
foreach ($manager in $managers) {

    # Get all direct reports for this manager
    $mgrReports = $null
    $mgrReports = $managersAndReports | Where-Object { $_.ManagerUPN -like $manager.UserPrincipalName } | Sort-Object DirectReportUPN

    $mgrReportsCount = $($mgrReports | Measure-Object).count
    # Define the Excel file path using manager's display name
    $excelPath = "$dir\$($manager.DisplayName)_Team Access Report.xlsx"

    # Create summary sheet with manager info and report generation date
    $initialData = [PSCustomObject]@{
        Manager            = $manager.DisplayName
        ManagerUPN         = $manager.UserPrincipalName
        "Report Generated" = Get-Date -Format dd-MM-yyyy-HH:mm
        'Number of Reports' =  $mgrReportsCount
    }

    $initialData | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize

    # Loop through each direct report for this manager
    foreach ($rpt in $mgrReports) {

        # Get the direct report user object and their group memberships
        $rptUser = $userHashtable[$rpt.DirectReportUPN]
        $rptGroups = Get-UserGroupMemberships -UPN $rptUser.UserPrincipalName

        # create worksheet name - max 31 char
        $worksheetName = $rptUser.DisplayName.Substring(0, [Math]::Min(31, $rptUser.DisplayName.Length))
        # Export group memberships to a new worksheet named after the direct report
        $rptGroups | Export-Excel -Path $excelPath -WorksheetName $worksheetName -Append -AutoSize -TableStyle Light1

        # Open the Excel package to add data validation and conditional formatting
        $pkg = Open-ExcelPackage -Path $excelPath
        $ws = $pkg.Workbook.Worksheets[$worksheetName]
        
        # Add dropdown validation in column D for Approve/Reject
        Add-ExcelDataValidationRule -Worksheet $ws -Range "D:D" -ValidationType List -Formula '"Approve,Reject"'

        # Add green conditional formatting for "Approve" entries
        Add-ConditionalFormatting -Worksheet $ws -Range "A:D" -RuleType Expression `
            -ConditionValue '=$D1="Approve"' -BackgroundColor Green -ForegroundColor White
    
        # Add red conditional formatting for "Reject" entries
        Add-ConditionalFormatting -Worksheet $ws -Range "A:D" -RuleType Expression `
            -ConditionValue '=$D1="Reject"' -BackgroundColor Red -ForegroundColor White
    
        # Save and close the Excel package
        Close-ExcelPackage $pkg
    }
write-host " - Created report for Manager $($manager.DisplayName) - $($mgrReportsCount) reports" -ForegroundColor Green
}
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
                Decision    = ''
            }
        }
    }
    end { return $groups }
}

import-module ImportExcel

$dir = 'C:\UserReports'

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
if (-not $users) {
    $users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, Manager, OfficeLocation, Department, Jobtitle
}
# create hash table
$userHashtable = @{}

foreach ($user in $users) {
    $userHashtable[$user.UserPrincipalName] = $user
}
# get all users with direct reports
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

# Loop through each manager
foreach ($manager in $managers) {

    # Get all direct reports for this manager
    $mgrReports = $managersAndReports | Where-Object { $managersAndReports.ManagerUPN -like $manager.UserPrincipalName }

    # Define the Excel file path using manager's display name
    $excelPath = "$dir\$($manager.DisplayName)_Direct Report Access Report.xlsx"

    # Create summary sheet with manager info and report generation date
    $initialData = [PSCustomObject]@{
        Manager            = $manager.DisplayName
        ManagerUPN         = $manager.UserPrincipalName
        "Report Generated" = Get-Date -Format 'yyyy-MM-dd HH:mm'
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
        $ws = $pkg.Workbook.Worksheets[$report.DisplayName]
        
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

}
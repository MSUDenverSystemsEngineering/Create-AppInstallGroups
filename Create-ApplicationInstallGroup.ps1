$AppList = Import-Csv -Path "C:\users\jhardy6\Documents\Powershell\SCCM Scripts\AppList.csv"

$AppList | ForEach-Object {

    $version = ($_.version -replace '[ ,_,.]')
    $name = ($_.Name -replace '[ ]')
    $ADInstall = "app_"+$name+"_"+$version+"_install"
    $ADUninstall = "app_"+$name+"_"+$version+"_uninstall"
    $PackageName = $_.Name + " " + $_.Version

    ##Creates Active Directory Security Groups
    New-ADGroup -Name "$ADInstall" -GroupCategory Security -GroupScope Global -Path "CN=$ADInstall,OU=Application Groups,OU=Groups,DC=winad,DC=msudenver,DC=edu" -Description "Members of this group have Visio 2019 automatically uninstalled"

    New-ADGroup -Name "$ADUninstall" -GroupCategory Security -GroupScope Global -Path "CN=$ADUninstall,OU=Application Groups,OU=Groups,DC=winad,DC=msudenver,DC=edu" -Description "Members of this group have Visio 2019 automatically uninstalled"

    ##Creates SCCM Device Collections
    New-CMDeviceCollection -Name "$ADInstall" -LimitingCollectionName "All Systems"

    New-CMDeviceCollection -Name "$ADUninstall" -LimitingCollectionName "All Systems"

    ##Adds the Device Collection Query Membership  Rule "Import CN from AD" to the install collection
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName "$ADInstall" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SecurityGroupName = 'WINAD\\$ADInstall'" -RuleName "Import CN from AD Install Group"

    ##Adds the Device Collection Exclude Membership Rule to the install collection
    Add-CMDeviceCollectionExcludeMembershipRule -CollectionName "$ADInstall" -ExcludeCollectionName "$ADUninstall"

    ##Adds the Device Collection Query Membership Rule to the uninstall collection
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName "$ADUninstall" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SecurityGroupName = 'WINAD\\App_Visio_19_Uninstall'" -RuleName "Import CN from AD Uninstall Group"

    ##Creates Required Installation Deployment
    New-CMApplicationDeployment -Name "$PackageName" -CollectionName "$ADInstall" -AvailableDateTime (Get-Date) -TimeBaseOn LocalTime -DeployAction Install -DeployPurpose Required -UserNotification DisplayAll -PreDeploy $true -OverrideServiceWindow $true

    ##Creates Required Uninstallation Deployment
    New-CMApplicationDeployment -Name "$PackageName" -CollectionName "$ADUninstall" -AvailableDateTime (Get-Date) -TimeBaseOn LocalTime -DeployAction Uninstall -DeployPurpose Required -UserNotification DisplayAll -PreDeploy $true -OverrideServiceWindow $true

    ##Find a way to adjust security permissions of AD Groups
}
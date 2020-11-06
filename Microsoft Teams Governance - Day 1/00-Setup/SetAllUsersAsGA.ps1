# Variables
$tenant = "TeamsWSBloomberg" # tenant name
$userCount = 24 # number of users
$roleName="Company Administrator" # Global Admin role name

# connect to services
write-host "Connecting to services, keep an eye out for sign-in dialogs"
Connect-AzureAD

# loop throiugh all users
For ($i=1; $i -le $userCount; $i++)
{
    $userName="user$($i)@$($tenant).onmicrosoft.com"
    $role = Get-AzureADDirectoryRole | Where {$_.displayName -eq $roleName}
    if ($role -eq $null) {
        $roleTemplate = Get-AzureADDirectoryRoleTemplate | Where {$_.displayName -eq $roleName}
        Enable-AzureADDirectoryRole -RoleTemplateId $roleTemplate.ObjectId
        $role = Get-AzureADDirectoryRole | Where {$_.displayName -eq $roleName}
    }
    Add-AzureADDirectoryRoleMember -ObjectId $role.ObjectId -RefObjectId (Get-AzureADUser | Where {$_.UserPrincipalName -eq $userName}).ObjectID
    # display
    Write-Host "User$($i) set to $($roleName)"
}

Write-Host -BackgroundColor Blue -ForegroundColor Yellow "Connecting to Tenant"
$orgName="teamswsbloomberg"
Connect-SPOService -Url https://$orgName-admin.sharepoint.com

$siteScriptRaw = '
{
    "$schema": "schema.json",
        "actions": [
            {
                "verb": "setSiteExternalSharingCapability",
                "capability": "Disabled"
            }
        ],
        "bindata": { },
    "version": 1
}
'

Add-SPOSiteScript -Title "Sharing Disabled Site Script" -Content $siteScriptRaw

$siteScript = Get-SPOSiteScript | where {$_.Title -eq "Sharing Disabled Site Script"} | select Id

Add-SPOSiteDesign -Title "Sharing Disabled Design" -WebTemplate 64 -SiteScripts $siteScript.Id -Description "Configures SharingCapability for Modern Team Sites that back a Microsoft Team"

Get-SPOSiteDesign
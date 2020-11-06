Write-Host -BackgroundColor Blue -ForegroundColor Yellow "Connecting to Tenant"
$orgName="teamsws4"
Connect-SPOService -Url https://$orgName-admin.sharepoint.com

Get-SPOSiteDesign
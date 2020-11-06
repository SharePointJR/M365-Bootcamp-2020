# vars
$configFileName = $PSScriptRoot+"/Elements.xml"

# Grab Elements.xml field configuration file
Write-Host "Obtaining Configuration file" -foregroundcolor black -backgroundcolor yellow  
[xml]$xmlinput = (Get-Content (Resolve-Path $configFileName))

# Credentials to connect to office 365 site collection url 
$url = $xmlinput.Elements.SPOnlineDetails.Url
$username= $xmlinput.Elements.SPOnlineDetails.UserName
$password= $xmlinput.Elements.SPOnlineDetails.Password

# Check to make sure information valued
Write-Host "Peform check to determine connection information in the configuration file is correct or not" -foregroundcolor black -backgroundcolor yellow
if (($url.length -gt 0) -And ($username.length -gt 0) -And ($password.length -gt 0))
{
    Write-Host "Connection information exists in the configuration file" -foregroundcolor black -backgroundcolor green
    Write-Host "URL=$($url); UserName=$($username)" -foregroundcolor black -backgroundcolor green
}
else
{
    Write-Host "Connection information does not exist in the configuration file...exiting process" -foregroundcolor black -backgroundcolor red
    $isFileValid = $false
    return
}
# Secure password
$Password = $password |ConvertTo-SecureString -AsPlainText -force

# Grab list info
$listTitle = $xmlinput.Elements.List.Title
$listUrl = $xmlinput.Elements.List.Url
$listDescription = $xmlinput.Elements.List.Description
$listTemplateType = $xmlinput.Elements.List.TemplateType

# Check to make sure information valued
Write-Host "Peform check to determine if the list information in the configuration file is correct or not" -foregroundcolor black -backgroundcolor yellow
if (($listTitle.length -gt 0) -And ($listUrl.length -gt 0) -And ($listDescription.length -gt 0) -And ($listTemplateType.length -gt 0))
{
    Write-Host "List information exists in the configuration file" -foregroundcolor black -backgroundcolor green
    Write-Host "List Title=$($listTitle); List Url=$($listUrl); Parent Content ID=$($listDescription); List Template Type=$($listTemplateType)" -foregroundcolor black -backgroundcolor green
}
else
{
    Write-Host "List information does not exist in the configuration file...exiting process" -foregroundcolor black -backgroundcolor red
    $isFileValid = $false
    return
}

# Grab the local SharePoint DLLs to perform our operations
Write-Host "Load CSOM libraries from local area" -foregroundcolor black -backgroundcolor yellow
Set-Location $PSScriptRoot
Add-Type -Path (Resolve-Path "../dlls/Microsoft.SharePoint.Client.dll")
Add-Type -Path (Resolve-Path "../dlls/Microsoft.SharePoint.Client.Runtime.dll")
Write-Host "CSOM libraries loaded successfully" -foregroundcolor black -backgroundcolor Green 

# Using the provided credentials, connect to the SPOnline/O365 site, and grab the Site, Web, and Fields objects
Write-Host "Authenticate using provided credentials to SharePoint Online site collection $url and get ClientContext object" -foregroundcolor black -backgroundcolor yellow  
$context = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password) 
$context.Credentials = $credentials 
#Delete List
$context.Web.Lists.GetByTitle($listTitle).DeleteObject()
try
{
    Write-Host "Deleting list $($listTitle)" -foregroundcolor black -backgroundcolor yellow  
    $context.ExecuteQuery()
    Write-Host "List $($listTitle) deleted" -foregroundcolor black -backgroundcolor green  
}
catch
{
    Write-Host "Exception deleting list $($listTitle) $_.Exception.Message" -foregroundcolor black -backgroundcolor red  
}

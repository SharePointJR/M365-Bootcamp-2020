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
$web = $context.Web
$context.Load($web)
try
{
    Write-Host "Loading SharePoint context" -foregroundcolor black -backgroundcolor yellow  
    $context.ExecuteQuery()
    Write-Host "SharePoint context loaded" -foregroundcolor black -backgroundcolor green  
}
catch
{
    Write-Host "Exception loading SharePoint context $_.Exception.Message" -foregroundcolor black -backgroundcolor red  
}

# Grab list info
$listTitle = $xmlinput.Elements.List.Title
$listUrl = $xmlinput.Elements.List.Url
$listDescription = $xmlinput.Elements.List.Description
$listTemplateType = $xmlinput.Elements.List.TemplateType
$listAddContentType = $xmlinput.Elements.List.AddContentType

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
    return
}

#Create list with "custom" list template
$listInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
$listInfo.Title = $listTitle
$listInfo.Url = $listUrl
$listInfo.TemplateType = $listTemplateType
$list = $context.Web.Lists.Add($listInfo)
$list.Description = $ListDescription
$list.Update()
try
{
    Write-Host "Creating list $($listTitle)" -foregroundcolor black -backgroundcolor yellow  
    $context.ExecuteQuery()
    Write-Host "List $($listTitle) created" -foregroundcolor black -backgroundcolor green  
}
catch
{
    Write-Host "Exception creating list $($listTitle) $_.Exception.Message" -foregroundcolor black -backgroundcolor red  
}
# grab content type
$contentTypeId = $xmlinput.Elements.ContentType.id
$contentType = $context.Web.ContentTypes.GetById($contentTypeID)
$context.Load($contentType)
try
{
    Write-Host "Getting content type" -foregroundcolor black -backgroundcolor yellow  
    $context.ExecuteQuery()
    Write-Host "Content type obtained" -foregroundcolor black -backgroundcolor green  
}
catch
{
    Write-Host "Exception getting content type $_.Exception.Message" -foregroundcolor black -backgroundcolor red  
}
# add content type to list
$listCTs = $list.ContentTypes # gets CTs for removal process later
$context.Load($listCTs)
$list.EnableAttachments = $false
$addedContentType = $list.ContentTypes.AddExistingContentType($contentType)
$list.Update()
try
{
    Write-Host "Adding content type to list" -foregroundcolor black -backgroundcolor yellow  
    $context.ExecuteQuery()
    Write-Host "Content type added to list" -foregroundcolor black -backgroundcolor green  
}
catch
{
    Write-Host "Exception adding content type to list $_.Exception.Message" -foregroundcolor black -backgroundcolor red  
}
#
# set multiple content types back to false
$list.ContentTypesEnabled = $false
$list.Update()
try
{
    Write-Host "Setting multiple content types on this list back to false" -foregroundcolor black -backgroundcolor yellow  
    $context.ExecuteQuery()
    Write-Host "Set multiple content types to false in list" -foregroundcolor black -backgroundcolor green  
}
catch
{
    Write-Host "Exception setting multiple content types to falst on this list $_.Exception.Message" -foregroundcolor black -backgroundcolor red  
}
#
# add fields to default view
$defaultView = $list.DefaultView
$idx = 1
foreach($field in $xmlinput.Elements.Fields.Field)
{
    if ($idx -gt 9)
    {
        break
    }
    $defaultView.ViewFields.Add($field.Name)
    $idx++
}
$defaultView.Update()
try
{
    Write-Host "Adding fields to default view" -foregroundcolor black -backgroundcolor yellow  
    $context.ExecuteQuery()
    Write-Host "Fields added to default view" -foregroundcolor black -backgroundcolor green  
}
catch
{
    Write-Host "Exception adding fields to default view $_.Exception.Message" -foregroundcolor black -backgroundcolor red  
}
# Delete Item Content Type
$itemCT = $listCTs | Where {$_.Name -eq "Item"}
If($itemCT -ne $Null)
{
    #Remove content type from list
    $itemCT.DeleteObject()
    try
    {
        Write-Host "Removing Item Content Type from List" -foregroundcolor black -backgroundcolor yellow  
        $context.ExecuteQuery()
        Write-Host "Item Content Type removed from List" -foregroundcolor black -backgroundcolor green  
    }
    catch
    {
        Write-Host "Exception removing Item Content Type from List $_.Exception.Message" -foregroundcolor black -backgroundcolor red  
    }
}
Write-Host "List creation operation complete." -foregroundcolor black -backgroundcolor Green 

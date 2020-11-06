# vars
$configFileName = $PSScriptRoot+"/Elements.xml"
$isFileValid = $true

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
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password) 
$Context.Credentials = $credentials 
$site = $context.Site 
$web = $context.Web
$webFields = $web.Fields
$contentTypes = $context.web.contenttypes      
$context.Load($web)
$context.Load($webFields)
$context.Load($site)
$context.Load($contentTypes)
# try/catch
try
{
    # Execute
    $context.ExecuteQuery()
    Write-Host "Authenticated to SharePoint Online site collection $url and ClientContext object obtained successfully" -foregroundcolor black -backgroundcolor Green
}
catch
{
    # Fail
    Write-Host "Unable to authenticate to SharePoint Online site collection $url $_.Exception.Message" -foregroundcolor black -backgroundcolor Red
    # we're done as we can't authenticate
    return
}

# Grab Content Type Information
$contentTypeName = $xmlinput.Elements.ContentType.Name
$contentTypeGroup = $xmlinput.Elements.ContentType.Group
$parentContentTypeId = $xmlinput.Elements.ContentType.ParentContentTypeId

# Check to make sure information valued
Write-Host "Peform check to determine if the content type information in the configuration file is correct or not" -foregroundcolor black -backgroundcolor yellow
if (($contentTypeName.length -gt 0) -And ($contentTypeGroup.length -gt 0) -And ($parentContentTypeId.length -gt 0))
{
    Write-Host "Content type information exists in the configuration file" -foregroundcolor black -backgroundcolor green
    Write-Host "Content Type Name=$($contentTypeName); Content Type Group=$($contentTypeGroup); Parent Content ID=$($parentContentTypeId)" -foregroundcolor black -backgroundcolor green
}
else
{
    Write-Host "Content type information does not exist in the configuration file...exiting process" -foregroundcolor black -backgroundcolor red
    $isFileValid = $false
    return
}

# Check to make sure the existing fields in the web contain the fields in the configuration file
Write-Host "Peform check to determine if the site columns in the configuration file exist in the site or not" -foregroundcolor black -backgroundcolor yellow
#  Loop through all fields in the configuration file
$fieldExists = $false
foreach($field in $xmlinput.Elements.Fields.Field)
{
    # start if field doesn't exist
    $fieldExists = $false
    foreach ($webField in $webFields)
    {
        if ($field.Name -eq $webField.InternalName)
        {
            # found it don't need to loop anymore
            Write-Host "Field $($field.Name) found, moving to next" -foregroundcolor black -backgroundcolor green
            $fieldExists = $true
            break
        }
    }
    if (!($fieldExists))
    {
        Write-Host "Field $($field.Name) does not exist in the web.  Field must exist in order to create the content type...exiting the routine" -foregroundcolor black -backgroundcolor red
        return
    }
}
if ($fieldExists)
{
    Write-Host "Field check valid and complete, moving forward" -foregroundcolor black -backgroundcolor yellow
}

# Check to make sure content type doesn't already exist
Write-Host "Peform check to determine if the content type in the configuration file exist in the site or not" -foregroundcolor black -backgroundcolor yellow
foreach ($contentType in $contentTypes){
    if ($contentType.name -eq $contentTypeName){
        write-host "Content type $($contentTypeName) already exists...exiting the routine" -foregroundcolor black -backgroundcolor red
        return
    }
}
Write-Host "Content type does not exist in the web, moving forward" -foregroundcolor black -backgroundcolor yellow

# Perform if we're good
if ($isFileValid)
{
    Write-Host "Configuration check complete, moving forward..." -foregroundcolor black -backgroundcolor yellow
    # create content type
    # load parent content type
    Write-Host "Loading parent content type" -foregroundcolor black -backgroundcolor yellow
    $parentContentType = $contentTypes.GetByID($parentContentTypeID)
    $context.load($parentContentType)
    # send the request containing all operations to the server
    try
    {
        $context.executeQuery()
        write-host "Parent Content Type loaded" -foregroundcolor black -backgroundcolor green
    }
    catch
    {
        write-host "Exception: $($_.Exception.Message); exiting the routine" -foregroundcolor black -backgroundcolor red
        return
    }
         
    # create Content Type using ContentTypeCreationInformation object (ctci)
    Write-Host "Loading ContentTypeCreationInformation object to create content type" -foregroundcolor black -backgroundcolor yellow
    $ctci = new-object Microsoft.SharePoint.Client.ContentTypeCreationInformation
    $ctci.name = $contentTypeName
    $ctci.ParentContentType = $parentContentType
    $ctci.group = $contentTypeGroup
    $ctci = $contentTypes.add($ctci)
    $context.load($ctci)
    # send the request containing all operations to the server
    try
    {
        $context.executeQuery()
        write-host "Content type created" -foregroundcolor black -backgroundcolor green
    }
    catch
    {
        write-host "Exception: $($_.Exception.Message); exiting routine" -foregroundcolor red
    }
    # get the new content type object
    Write-Host "Loading newly created content type to add fields" -foregroundcolor black -backgroundcolor yellow
    $newContentType = $context.web.contenttypes.getbyid($ctci.id)    
    # loop through all the columns that needs to be added
    foreach ($field in $xmlinput.Elements.Fields.Field)
    {
        $spField = $webFields.GetByInternalNameOrTitle($field.Name)
        #create FieldLinkCreationInformation object (flci)
        $flci = new-object Microsoft.SharePoint.Client.FieldLinkCreationInformation
        $flci.Field = $spField
        $addContentType = $newContentType.FieldLinks.Add($flci)
        write-host "Added $($field.Name) to content type" -foregroundcolor black -backgroundcolor yellow
    }
    Write-Host "Updating content type with new fields and updating configuration file with ID" -foregroundcolor black -backgroundcolor yellow
    $newContentType.Update($true)
    # send the request containing all operations to the server
    try
    {
        $context.executeQuery()
        write-host "Content type updated" -foregroundcolor black -backgroundcolor green
        write-host "Content type ID: $($ctci.id) " -foregroundcolor black -backgroundcolor green
        $contentTypeNode = $xmlinput.Elements.ContentType
        $contentTypeNode.SetAttribute("id", $ctci.id)
        $xmlinput.Save((Resolve-Path "Elements.xml"))
    }
    catch
    {
        write-host "Exception: $($_.Exception.Message)" -foregroundcolor black -backgroundcolor red
    }
}

# Write-Host "** Manual note: Update the 'Title' column of the content type to be hidden." -foregroundcolor black -backgroundcolor Green 
Write-Host "Content Type operation complete." -foregroundcolor black -backgroundcolor Green 

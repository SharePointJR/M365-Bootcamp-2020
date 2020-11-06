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
Add-Type -Path (Resolve-Path "../DLLs/Microsoft.SharePoint.Client.dll")
Add-Type -Path (Resolve-Path "../DLLs/Microsoft.SharePoint.Client.Runtime.dll")
Write-Host "CSOM libraries loaded successfully" -foregroundcolor black -backgroundcolor Green 

# Using the provided credentials, connect to the SPOnline/O365 site, and grab the Site, Web, and Fields objects
Write-Host "Authenticate using provided credentials to SharePoint Online site collection $url and get ClientContext object" -foregroundcolor black -backgroundcolor yellow  
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password) 
$Context.Credentials = $credentials 
$web = $context.Web
$webFields = $web.Fields;         
$site = $context.Site 
$context.Load($web)
$context.Load($webFields)
$context.Load($site)
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

# Check to make sure the existing fields in the web aren't in the configuration file
Write-Host "Peform check to determine if the site columns in the configuration file exist in the site or not" -foregroundcolor black -backgroundcolor yellow
#  Loop through all fields in the web
foreach ($webField in $webFields)
{
    # Loop through all the field elements in the configuration file
    foreach($field in $xmlinput.Elements.Fields.Field)
    {
        # $field.OuterXml
        if ($webField.InternalName -eq $field.Name)
        {
            Write-Host "Site column $($field.Name) / $($webField.InternalName) already exists...exiting..." -foregroundcolor black -backgroundcolor Red
            $isFileValid = $false
            return
        }
    }
}
Write-Host "Site column validation passed successfully" -foregroundcolor black -backgroundcolor Green

if ($isFileValid)
{
    Write-Host "Creating site columns from Configuration File" -foregroundcolor black -backgroundcolor yellow 
    foreach($field in $xmlinput.Elements.Fields.Field)
    {
        $fieldAsXML = $field.OuterXml
        $fieldOption = [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue
        $spField = $null
        $spField = $webFields.AddFieldAsXml($fieldAsXML, $true, $fieldOption)
        $context.Load($spField)
        try
        {
          $context.ExecuteQuery()
          Write-Host "Site column $($field.DisplayName) / $($field.Name) created successfully" -foregroundcolor black -backgroundcolor Green 
        }
        catch
        {
          Write-Host "Error while creating site column $($field.Name) $($_.Exception.Message)" -foregroundcolor black -backgroundcolor Red 
        }
    }
}
else
{
      Write-Host "Configuration File error" -foregroundcolor black -backgroundcolor Red 
}
Write-Host "Site columns created successfully...operation complete." -foregroundcolor black -backgroundcolor Green 
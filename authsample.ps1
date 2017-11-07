<# 
.SYNOPSIS  
    none working code sniplet to authenticate to office with a refresh token
.DESCRIPTION  
    you can use this code sample to login to office 365 using a refresh token obtained with createrestauthtoken.ps1
.NOTES  
    File Name  : authsample.ps1
    Author     : Kevin Miller (kemi@microsoft.com) - if you can call me an author, more like the guy who borrowed from other people to hack a thing
      
.LINK 
    base site used to build this script https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/ 
    the other side used to trouble shoot the script https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-protocols-oauth-code
#>
# variables
# tokenxml is the path and name to where you want to store the login credentials
$tokenxml = "c:\bob\authtoken.xml"
# logfile is the path and name to the file used for logging
$logfile = "C:\bob\authlog.txt"


# Create the REST authorization token
# Load our configuration data
$configData = Import-Clixml $tokenxml

Add-Type -AssemblyName System.Web
$client_id = $configData.client_id
$client_secret = $configData.client_secret
$redirectUrl = $configData.redirecturl
$refreshToken = $configData.refresh_token

# Now, refresh the authorization token
$loginUrl = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=" + 
[System.Web.HttpUtility]::UrlEncode($redirectUrl) + 
"&client_id=$client_id" + 
"&prompt=login"

$ReAuthorizationPostRequest = 
"grant_type=refresh_token" + "&" +
"redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($redirectUrl) + "&" +
"client_id=$client_id" + "&" +
"client_secret=" + [System.Web.HttpUtility]::UrlEncode("$client_secret") + "&" +
"refresh_token=" + $refreshToken + "&" +
"resource=" + [System.Web.HttpUtility]::UrlEncode("https://outlook.office365.com/")

# Actually make the request
$Authorization = invoke-restmethod -uri "https://login.microsoftonline.com/common/oauth2/token" -ContentType "application/x-www-form-urlencoded" -Method Post -Body $ReAuthorizationPostRequest

# Check to make sure we have an access token
if (!$Authorization.access_token) {
    "$((get-date).ToUniversalTime()) - Access token is not populated, exiting." | Out-File $logfile -Append
    "$((get-date).ToUniversalTime()) - $($errVar | fl)" | Out-File $logfile -Append
    "==================" | Out-File $logfile -Append
    Remove-Item $marker
    exit
}
# We do, let's proceed

# Write our refresh_token and other values back out to the XML
$configData.refresh_token = $Authorization.refresh_token
$configData | Export-Clixml $tokenxml

# check your mail to make sure everything is working.
$mail = Invoke-RestMethod -Headers @{Authorization =("Bearer "+ $Authorization.access_token)} -Uri "https://outlook.office365.com/api/v2.0/me/messages" -Method Get
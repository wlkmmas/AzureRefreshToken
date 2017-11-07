<# 
.SYNOPSIS  
    Captures an OAUTH REST login token, logs in with the token, then captures data and a refresh token in an XML file
.DESCRIPTION  
    This script requires a bunch of things to be in place - 
    -register an application in Azure - http://dev.office.com/app-registration
      -generate and capture a key for that application on the key page
      -generate and capture a reply url for that application
      -authorize the application to have access to your mailbox
    -see link here for how to do all of this once I write it.   
.NOTES  
    File Name  : CreateRestAuthToken.ps1  
    Author     : Kevin Miller (kemi@microsoft.com) - if you can call me an author, more like the guy who borrowed from other people to hack a thing
      
.LINK 
    base site used to build this script https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/ 
    the other side used to trouble shoot the script https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-protocols-oauth-code
#>

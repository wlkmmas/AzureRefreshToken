<# 
.SYNOPSIS  
    Captures an OAUTH REST login token, logs in with the token, then captures data and a refresh token in an XML file
.DESCRIPTION  
    This script requires a bunch of things to be in place - 
    -register an application in Azure - http://dev.office.com/app-registration
    -generate and capture a key for that application on the key page
    -generate and capture a reply url for that application
    -authorize the application to have access to your mailbox
    -see link here for how to do all of this -   
.NOTES  
    File Name  : CreateRestAuthToken.ps1  
    Author     : Kevin Miller (kemi@microsoft.com) - if you can call me an author, more like the guy who borrowed from other people to hack a thing
      
.LINK 
    base site used to build this script https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/ 
    the other side used to trouble shoot the script https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-protocols-oauth-code
#>
#function to create a login window to validate your Oauth creds
Function Show-OAuthWindow
{
    param(
        [System.Uri]$Url
    )


    Add-Type -AssemblyName System.Windows.Forms
 
    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url ) }
    $DocComp  = {
        $Global:uri = $web.Url.AbsoluteUri
        if ($Global:Uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null

    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
    }
    
    $output
}
# Create the first authorization request
# you need to fill application_ID, key_secret, and reply_url

Add-Type -AssemblyName System.Web
$Application_id = "5f0267bc-5073-4f68-869e-ff259053ee52"
$Key_secret = "3uGH5km9jFH50x0cvu6PUSg5owoIW1PDzxrPXtb9cu0="
$Reply_Url = "https://localhost/"
$loginUrl = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=" + 
            [System.Web.HttpUtility]::UrlEncode($reply_Url) + 
            "&client_id=$Application_id" + 
            "&prompt=login"

#Open the login window for user input
$queryOutput = Show-OAuthWindow -Url $loginUrl

#generate the authorization request URL 
$AuthorizationPostRequest = 
"grant_type=authorization_code" + "&" +
"redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($reply_Url) + "&" +
"client_id=$Application_id" + "&" +
"client_secret=" + [System.Web.HttpUtility]::UrlEncode("$Key_secret") + "&" +
"code=" + $queryOutput["code"] + "&" +
"resource=" + [System.Web.HttpUtility]::UrlEncode("https://outlook.office365.com")

# generate the authorization request and obtain the refresh token
$Authorization = Invoke-RestMethod   -Method Post `
                    -ContentType application/x-www-form-urlencoded `
                    -Uri https://login.microsoftonline.com/common/oauth2/token `
                    -Body $AuthorizationPostRequest

# construct the xml file output to capture the refresh token to use later
Add-Type -AssemblyName System.Web
$configData.client_id = $Application_id
$configData.client_secret = $Key_secret
$configData.redirecturl = $reply_Url
$configData.refresh_token= $Authorization.refresh_token

# save the XML output with all of the tokens
$configData | Export-Clixml $remailerxml                   

# check your mail to make sure everything is working.
$mail = Invoke-RestMethod -Headers @{Authorization =("Bearer "+ $Authorization.access_token)} -Uri "https://outlook.office365.com/api/v2.0/me/messages" -Method Get
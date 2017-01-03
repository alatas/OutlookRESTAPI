
<#PSScriptInfo

.VERSION 1.0.0.0

.GUID e2dfffdf-d767-4fce-9dca-238cbc44e22b

.AUTHOR Sukru Alatas

.COMPANYNAME 

.COPYRIGHT 

.TAGS outlook.com REST OAuth2

.LICENSEURI 

.PROJECTURI https://github.com/alatas/OutlookRESTAPI/

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#>

<# 

.DESCRIPTION 
 Deletes all the outlook.com events in the calendar that the user selects from the listed calendar. 

#> 
Param()

function Show-OAuthWindow {
  param(
    [System.Uri]$Url
  )

  #---------------#
  #based on https://blogs.technet.microsoft.com/heyscriptingguy/2013/07/01/use-powershell-3-0-to-get-more-out-of-windows-live/
  #---------------#
  
  Add-Type -AssemblyName System.Windows.Forms

  $form = New-Object -TypeName System.Windows.Forms.Form -Property @{ Width = 440; Height = 640 }
  $web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{ Width = 420; Height = 600; Url = ($url) }
  $formSuccess = $false

  $DocComp = {
    $Global:uri = $web.Url.AbsoluteUri
    if ($Global:Uri -match "error=[^&]*|access_token=[^&]*") {
      $script:formSuccess = $true
      $form.Close()
    }

  }
  $web.ScriptErrorsSuppressed = $true
  $web.Add_DocumentCompleted($DocComp)
  $form.Controls.Add($web)
  $form.Add_Shown({ $form.Activate() })
  $form.ShowDialog() | Out-Null
  if ($script:formSuccess) {

    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Fragment.Substring(1))
    $output = @{}
    foreach ($key in $queryOutput.Keys) {
      $output["$key"] = $queryOutput[$key]
    }

    return $output
  } else {
    return $null
  }
}

function Get-RESTApi {
  param(
    $Func,
    $Method = "GET"
  )

  Invoke-RestMethod -Headers @{ Authorization = ("Bearer " + $Authorization["access_token"]) } `
     -Uri https://outlook.office.com/api/v2.0/$Func `
     -Method $Method
}

function Get-Authenticate {
  Add-Type -AssemblyName System.Web
  $client_id = "d2a2b164-c156-4f9b-8dc4-c6ebe9f97177"
  $redirectUrl = "https://raw.githubusercontent.com/alatas/OutlookRESTAPI/master/README.md"

  $state = [guid]::NewGuid()

  $loginUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize" +
  "?response_type=token" +
  "&redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($redirectUrl) +
  "&client_id=$client_id" +
  "&scope=" + [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/calendars.readwrite") + 
  "&promt=login" +
  "&state=" + $state

  while ($true) {
    $script:Authorization = Show-OAuthWindow -Url $loginUrl
    if ($script:Authorization -ne $null) {
      break
    } else {
      if (1 -eq (Get-Menu -Title "Outlook Events Cleaner" `
             -Desc "Script cannot get token for authorization. Do you want to try again ?" `
             -Default 0 -Options "&Try Again","&Exit") `
        ) { exit }

    }
  }
}

function Get-Menu {
  param(
    $Title,
    $Desc,
    [string[]]$Options,
    $Default = 0
  )

  return $host.UI.PromptForChoice($Title,$Desc,$Options,$Default)
}

if (1 -eq (Get-Menu -Title "Outlook Events Cleaner" `
       -Desc "This script will authorize with your live.com / outlook.com account, search and delete all of calendar events." `
       -Default 0 -Options "&Continue","&Exit") `
  ) { exit }

echo "Now script is opening an authentication web page for authorize to read your calendar information"

Get-Authenticate

echo "Authentication is successful, fetching calendar information"

$rootFolders = Get-RESTApi -Func me/calendars | select -ExpandProperty value

$selection = Get-Menu -Title "Select Calendar" -Desc "Please select calendar you want to clean. All of the events in that calendar will be deleted" -Options $($rootFolders | select -ExpandProperty Name)

$selectedFolder = $rootFolders[$selection]

if (1 -eq (Get-Menu -Title "Confirmation" -Desc "All of the events in $($selectedFolder.Name) will be DELETED. THIS OPERATION ISNOT REVERSIBLE. ARE YOU SURE?" -Options "&No","&YES" -Default 0)) {

  while ($true) {
    $events = (Get-RESTApi -Func me/calendars/$($selectedFolder.Id)/events).value

    if (@( $events).Count -gt 0) {

      foreach ($event in $events) {
        echo "Deleting $($event.Subject)..."
        Get-RESTApi -Method DELETE -Func me/events/$($event.Id)
      }

    } else {
      break
    }
  }
}

echo ""
echo "Finished!"
echo "(press any key to exit)"
Read-Host
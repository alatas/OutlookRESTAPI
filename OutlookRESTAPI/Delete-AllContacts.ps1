
<#PSScriptInfo

.VERSION 1.0.0.0

.GUID b798c22f-a96e-4d13-b034-48c1faa211b3

.AUTHOR Sukru Alatas

.COMPANYNAME 

.COPYRIGHT 

.TAGS outlook.com REST OAuth2

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#>

<# 

.DESCRIPTION 
 Deletes all the outlook.com contacts in the folder that the user selects from the listed contacts folders. 

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
  "&scope=" + [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/contacts.readwrite") + 
  "&promt=login" +
  "&state=" + $state

  while ($true) {
    $script:Authorization = Show-OAuthWindow -Url $loginUrl
    if ($script:Authorization -ne $null) {
      break
    } else {
      if (1 -eq (Get-Menu -Title "Outlook Contacts Cleaner" `
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

if (1 -eq (Get-Menu -Title "Outlook Contacts Cleaner" `
       -Desc "This script will authorize with your live.com / outlook.com account, search and delete all of your orphaned contacts." `
       -Default 0 -Options "&Continue","&Exit") `
  ) { exit }

echo "Now script is opening an authentication web page for authorize to read your contacts information"

Get-Authenticate

echo "Authentication is successful, fetching contacts informations"

$contactsFolder = Get-RESTApi -Func me/contactfolders/Contacts

$rootFolders = Get-RESTApi -Func me/contactfolders/$($contactsFolder.ParentFolderId)/childfolders | select -ExpandProperty value

foreach ($folder in $rootFolders) {
  [int]$count = ((Get-RESTApi -Func me/contactfolders/$($folder.Id)/contacts/$('$count')) -ireplace '[^0-9]','')

  Add-Member -InputObject $folder -NotePropertyName Count -NotePropertyValue $count
  Add-Member -InputObject $folder -NotePropertyName Name -NotePropertyValue $('&' + $folder.DisplayName + ' (' + $folder.Count + ' contacts)')

}

$selection = Get-Menu -Title "Select Contacts Folder" -Desc "Please select contacts folder you want to clean. All of the contacts in that folder will be deleted" -Options $($rootFolders | select -ExpandProperty Name)

$selectedFolder = $rootFolders[$selection]

if (1 -eq (Get-Menu -Title "Confirmation" -Desc "All of the contacts in $($selectedFolder.DisplayName) Folder will be DELETED. THIS OPERATION IS IRREVERSIBLE. ARE YOU SURE?" -Options "&No","&YES" -Default 0)) {

  while ($true) {
    $contacts = (Get-RESTApi -Func me/contactfolders/$($selectedFolder.Id)/contacts).value

    if (@( $contacts).Count -gt 0) {

      foreach ($contact in $contacts) {
        echo "Deleting $($contact.FileAs)..."
        Get-RESTApi -Method DELETE -Func me/contacts/$($contact.Id)
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
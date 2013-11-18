Set-StrictMode -Version 2.0
$ErrorActionPreference="Stop"

# More info @ https://github.com/idvorkin/psOutlook

function global:Get-Outlook
{
    Add-type -assembly "Microsoft.Office.Interop.Outlook" 
    $olApp = New-Object -comObject Outlook.Application     $mapi = $olApp.GetNameSpace("MAPI")

    $ol = [PSCustomObject] @{
        MAPI = $mapi
        Application = $olApp
        Folders = @{}
    }

    # Add the Folders
    $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]

    $realFolderCount = 10 # manually set by looking at OlDefaultFolders struct.
    [enum]::GetValues($olFolders) | Select -First 10 | %  {
        # extract string from folderName e.g olFolderInbox
        $folderName = (([string]$_) -Split "olFolder")[1]
        $folderValue = ([int]$_)
        $ol.Folders[$folderName] = $mapi.GetDefaultFolder($folderValue)
    }

    # Useful member if you use delayed send.
    $ol | Add-Member -MemberType ScriptMethod -Name SendAllInOutbox -Value {
        $this.Folders.Outbox.Items | % {
            # Move Deferred Time To Past
            $_.DeferredDeliveryTime = [DateTime]::Now- [TimeSpan]::FromMinutes(10)
            $_.Send()
        }
    } 

    return $ol
}

Get-Outlook
psOutlook
=========

Outlook Powershell Module - Outlook automation from powershell. This will contain methods to automate my common outlook tasks.



Summary:

    # Return an object wrapping outlook functionality
    $ol = Get-Outlook()

    # Return Folders
    $ol.Folders

    # Access specific folders
    $ol.Folders.Inbox
    $ol.Folders.Outbox

    # Enumerate mails in inbox
    $ol.Folders.Inbox.items | -Property Subject 

    #Send all mail in outbox
    $ol.SendAllInOutbook()


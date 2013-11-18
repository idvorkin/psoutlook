psOutlook
=========

Outlook Powershell Module - Outlook automation from powershell. This will contain methods to automate my common outlook tasks.



Usage:
-------

```powershell

    # Return an object wrapping outlook functionality
    $ol = . .\outlook.ps1

    # To get another outlook instance load outlook directly.
    $ol = Get-Outlook()

    # Enumerate mails in inbox
    $ol.Folders.Inbox.Items | Select -Property SenderName, Subject, ReceivedTime

    # Enumerate mails in outbox
    $ol.Folders.Outbox.Items | Select -Property To, Subject, DeferredDeliveryTime

    #Send all mail in outbox
    $ol.SendAllInOutlook()

    # Enumerate folders 
    $ol.Folders
```
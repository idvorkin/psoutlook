psOutlook
=========

Outlook Powershell Module - Outlook automation from powershell. This will contain methods to automate my common outlook tasks.



Usage:
-------

```powershell
     # Load plugin
     . .\outlook.ps1

    # Return an object wrapping outlook functionality
    $ol = Get-Outlook()

    # Enumerate mails in inbox
    $ol.Folders.Inbox.Items | Select -Property SenderName, Subject, ReceivedTime

    #Send all mail in outbox
    $ol.SendAllInOutbook()

    # Enumerate folders 
    $ol.Folders
```
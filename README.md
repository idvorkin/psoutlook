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

    # Enumerate last 100 mails
    $ol.Folders.SentMail.Items | Select -First 100  | Select -Property To, Subject, SentOn

    #Send all mail in outbox
    $ol.SendAllInOutbox()

    # Enumerate Items in Calendar
    $ol.Folders.Calendar.Items | Select -First 200 | Select -Property Categories, Subject, Start, Duration
```
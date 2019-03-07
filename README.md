# setOutlookAging
Hi!

As you might knew already there is some problems in managing Outlook aging properties centralized way. Microsoft Office GPO templates cannot set all ther properties cause it is not stored in registry. 

This is a Powershell script that can silently change aging settings on every Outlook folder.
It should be run under the user account. It find user Outlook profile and on every email account in it runs recursive function that sets aging parameters.

I've used trial version of a wonderful tool OutlookSpy to analyze MAPI properties names and debug the script.

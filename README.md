# Outlook-to-Zendesk
Outlook Add-in Forward to Zendesk
For those that have tried using the Outlook to Zendesk case creation add-in know it’s limitations:
1.	You cannot send inline images
2.	Web outlook add-ins will not work for a shared inbox, (like a support inbox).

To work around this Zendesk does provide the ability to Forward an email and place some mail api code in it, which allows you to set the original sender as the ticket requester as noted here:
https://support.zendesk.com/hc/en-us/articles/203663316-Passing-an-email-to-your-support-address#topic_jfs_cwr_2k
(you can also forward if the subject has the fwd or fwd, but using the mail api eliminates the need for this)

This add-in builds on prior VB macro tips and code out there like this one from Andrew Bray:
https://support.zendesk.com/hc/en-us/community/posts/203458956-Forward-email-as-Ticket-Internal-Users

I’ve taken it a step further and created an installable program you can distribue to all users vs having to copy and paste a macro into each users outlook.

Before you start you'll need Visual Studio 2017.
Make sure to also shut down your outlook client.

Open the .sln program and make sure to do a "clean" of the code first.
You should be able to start the code and run it.

You will need to install your own signed certificate to distriute the code to your team.

Included in this code is a user install document as well.  Feel free to edit as needed.

Please ping me or add an issue to the list to fix for any problems. Thanks.


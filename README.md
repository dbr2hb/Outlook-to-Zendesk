# Outlook-to-Zendesk
Outlook Add-in Forward to Zendesk

Note: currently this is windows only.

For those that have tried using the Outlook to Zendesk case creation add-in know it’s limitations:

1. You cannot send inline images 
2. Web outlook add-ins will not work for a shared inbox, (like a support inbox).

To work around this Zendesk does provide the ability to Forward an email and place some mail api code in it, which allows you to set the original sender as the ticket requester as noted here:

https://support.zendesk.com/hc/en-us/articles/203663316-Passing-an-email-to-your-support-address#topic_jfs_cwr_2k (you can also forward if the subject has the fwd or fwd, but using the mail api eliminates the need for this)

This add-in builds on prior VB macro tips and code out there like this one from Andrew Bray:

https://support.zendesk.com/hc/en-us/community/posts/203458956-Forward-email-as-Ticket-Internal-Users

Basically it does a one click forward email to your support email address with the original sender.  Assuming you properly have zendesk setup to assign forwarded tickets it will create a ticket with the original sender as the requester.

<B>Enable Email forwarding under Admin > Settings > Agents</b>

I’ve taken it a step further and created an installable program you can distribute to all users vs having to copy and paste a macro into each users outlook.

Before you start you'll need Visual Studio 2017. Make sure to also shut down your outlook client.

Open the .sln program.  In the visual basic code make sure to put in your correct zendesk email address.  Make sure to do a "clean" of the code first. You should be able to start the code in debug mode and run it. It'll open up outlook and the add-in will show the on your ribbon.

You will need to install your own signed certificate to distribute the code to your team.

Included in this package is a user install document as well that you can edit and give you a better idea of what the add-in does.

Please ping me or add an issue to the list to fix for any problems or leave a comment here. Also if you'd like to contribute please put in a pull request. Thanks.


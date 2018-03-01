Public Class ThisAddIn


    Sub HelpdeskNewTicket()
        Dim helpdeskaddress As String
        Dim objMail As Outlook.MailItem
        'Dim olNS As Outlook.NameSpace
        Dim objItem As Outlook.MailItem
        Dim senderaddress As String
        Dim result As MsgBoxResult

        result = MsgBox("Is this a new Case?", vbYesNo)

        If result = vbYes Then

            ' Set this variable as your helpdesk e-mail address
            helpdeskaddress = "support@trackercorp.zendesk.com"


            objItem = GetCurrentItem()
            objMail = objItem.Forward

            ' Sender E=mail Address
            If objItem.SenderEmailType = "EX" Then

                senderaddress = objItem.Sender.GetExchangeUser().PrimarySmtpAddress
            Else

                senderaddress = objItem.SenderEmailAddress
            End If

            objMail.To = helpdeskaddress
            objMail.Subject = objItem.Subject
            objMail.HTMLBody = "#requester " & senderaddress & "<BR>" & objItem.HTMLBody
            objMail.Sender = Application.Session.CurrentUser.AddressEntry
            objItem.Categories = "Added to Zendesk"
            objItem.Save()


            ' remove the comment from below to display the message before sending
            'objMail.Display()

            'Automatically Send the ticket
            objMail.Send()

            objItem = Nothing
            objMail = Nothing
        Else

        End If
    End Sub

    Function GetCurrentItem() As Object
        Dim objApp As Outlook.Application
        objApp = Application
        On Error Resume Next
        Select Case TypeName(objApp.ActiveWindow)
            Case "Explorer"
                GetCurrentItem =
objApp.ActiveExplorer.Selection.Item(1)
            Case "Inspector"
                GetCurrentItem =
objApp.ActiveInspector.CurrentItem
            Case Else
        End Select
    End Function



    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class

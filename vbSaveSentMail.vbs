' This sample displays the folder picker for each outgoing email. Press Cancel on the dialog if you want to store the message in the default folder.
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
  If TypeOf Item Is Outlook.MailItem Then
    Cancel = Not SaveSentMail(Item)
  End If
End Sub

Private Function SaveSentMail(Item As Outlook.MailItem) As Boolean
  Dim F As Outlook.MAPIFolder

  If Item.DeleteAfterSubmit = False Then
    Set F = Application.Session.PickFolder
    If Not F Is Nothing Then
      Set Item.SaveSentMessageFolder = F
      SaveSentMail = True
    End If
  End If
End Function

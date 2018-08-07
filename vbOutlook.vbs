
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

Function DeleteCompletedTasks()
    Dim ofolder As Outlook.Folder
    Dim i As Long
    i = 0
    'Get the Task items
    Set ofolder = Outlook.Session.GetDefaultFolder(olFolderTasks)
    For Each Item In ofolder.Items
        If Item.Status = olTaskComplete Then
            Item.Delete
            i = i + 1
        End If
    Next
    MsgBox Str(i) + " completed tasks deleted."
End Function

Function DeleteExpiredMails()
    Dim ofolder As Outlook.Folder
    Dim i As Long
    i = 0
    
    Set ofolder = Outlook.Session.GetDefaultFolder(olFolderInbox).Folders("ToDo")
    For Each Item In ofolder.Items
        If Item.Class = olMail Then
            If Item.ExpiryTime < Now Then
                Item.Delete
                i = i + 1
            End If
        End If
    Next
    MsgBox Str(i) + " expired mails deleted."
End Function

Sub HouseKeeping()
    Call DeleteCompletedTasks
    Call DeleteExpiredMails
End Sub
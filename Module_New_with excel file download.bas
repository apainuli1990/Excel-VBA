Attribute VB_Name = "Module_New"
Sub Attachment()

    
    Call New01
    
    If OutOpened Then App.Quit
    Set App = Nothing
            
End Sub

Private Sub New01()

    Dim App As Outlook.Application
    Dim ns As Outlook.NameSpace
    Dim MsgInbox As Outlook.MAPIFolder
    Dim Attach As Outlook.Attachment
    Dim Item As Object
    Dim MailItem As Outlook.MailItem
    Dim subject As String
    Dim saveFolder As String
    Dim DateFormat As String
    Dim Filter As String
   Const Filetype As String = "xlsx"
    
    saveFolder = "D:\Test02"
    If Right(saveFolder, 1) <> "\" Then saveFolder = saveFolder & "\"
    
         
    
    
    Set MsgInbox = Outlook.Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    
    
    subject = "Attachment"
      
        If MsgInbox.Items.Count > 0 Then
            For Each Item In MsgInbox.Items
                If Item.Class = Outlook.OlObjectClass.olMail Then
                    Set MailItem = Item
                    If MailItem.subject = subject Then
                        Debug.Print MailItem.subject
                        For Each Attach In MailItem.Attachments
                        If Right(LCase(Attach.fileName), Len(Filetype)) = Filetype Then
                            DateFormat = Format(MailItem.ReceivedTime(), "yyyy-mm-dd hh-mm")
                            Attach.SaveAsFile saveFolder & "(" & DateFormat & ")" & " " & Attach
                        End If
                        Next
                    End If
                End If
            Next
        End If
MsgBox "All Atttachments are downloaded"

End Sub

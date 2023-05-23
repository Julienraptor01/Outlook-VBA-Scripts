Sub BulkTransferEmailsAsAttachements()
	Dim olMsgToSend As Outlook.MailItem
	Dim olItem As Outlook.MailItem

	On Error Resume Next
	If Application.ActiveExplorer.Selection.Count = 0 Then
		MsgBox ("No item selected")
		Exit Sub ' TODO : Verify if exiting sub automatically clean up variables
	End If

	For Each olItem In Application.ActiveExplorer.Selection
		Set olMsgToSend = Application.CreateItem(olMailItem)

		With olMsgToSend
			.Attachments.Add olItem, olEmbeddeditem
			.Subject = "Macro-TR:" + olItem.Subject
			.To = "Sender" ' TODO : Change the mail address in the macro before you run it.
			.Display
		End With
	Next

	Set olItem = Nothing
	Set olMsgToSend = Nothing
End Sub
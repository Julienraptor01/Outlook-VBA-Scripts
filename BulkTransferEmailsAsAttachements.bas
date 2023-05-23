Sub BulkTransferEmailsAsAttachements()
	Dim olMsgToSend As Outlook.MailItem
	Dim olItem As Outlook.MailItem

	On Error Resume Next
	If Application.ActiveExplorer.Selection.Count = 0 Then
	'show the user that he didn't select anything if he didn't and exit the sub
		MsgBox ("No item selected")
		Exit Sub ' TODO : Verify if exiting sub automatically clean up variables
	End If

	For Each olItem In Application.ActiveExplorer.Selection
		Set olMsgToSend = Application.CreateItem(olMailItem) ' create a new mail item

		With olMsgToSend
			.Attachments.Add olItem, olEmbeddeditem 'add the selected item we are iterating onto as an attachment
			.Subject = "Macro - TR: " & olItem.Subject 'keep the subject of the original mail but add a prefix to it
			.To = "Receiver" ' TODO : Change the mail address in the macro before you run it since this define the receiver email address
			.Display 'display the email whitout sending for debugging purposes
			'.Send 'send the mail
		End With

		olItem.categories = "OKiADocs" 'add a category to the original mail to keep track of the mails that have been processed
	Next

	'both next lines are to clean up the variables
	Set olItem = Nothing
	Set olMsgToSend = Nothing
End Sub
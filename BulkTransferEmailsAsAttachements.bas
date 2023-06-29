Sub BulkTransferEmailsAsAttachements()
	Const Receiver As String = "Receiver" ' TODO : Change the email address in the macro when setting it up
	Const CategoryName As String = "CategoryName" ' TODO : Change the category name in the macro when setting it up

	On Error Resume Next
	If Application.ActiveExplorer.Selection.Count = 0 Then
	'show the user that he didn't select anything if he didn't and exit the sub
		MsgBox ("No item selected")
		Exit Sub
	End If

	Dim olMsgToSend As Outlook.MailItem
	Dim olItem As Outlook.MailItem

	For Each olItem In Application.ActiveExplorer.Selection
		Set olMsgToSend = Application.CreateItem(olMailItem) 'create a new email item

		With olMsgToSend
			.Attachments.Add olItem, olEmbeddeditem 'add the selected item we are iterating onto as an attachment
			.Subject = "Macro - TR: " & olItem.Subject 'keep the subject of the original email but add a prefix to it
			.To = Receiver 'set the receiver of the email
			'.Display 'display the email without sending for debugging purposes
			.Send 'send the email
		End With

		'add a category to the original email to keep track of the emails that have been processed
		If olItem.categories = "" Then
			olItem.categories = CategoryName
		Else
			olItem.categories = olItem.categories & "," & CategoryName
		End If
	Next

	'both next lines are to clean up the variables
	Set olMsgToSend = Nothing
	Set olItem = Nothing
End Sub
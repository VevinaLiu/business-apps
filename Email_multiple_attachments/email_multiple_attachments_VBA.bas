Attribute VB_Name = "Module1"

Public Function GetEmailFolder() As String

Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select Folder"
    .AllowMultiSelect = False
    .InitialFileName = ActiveWorkbook.path
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetEmailFolder = sItem
Set fldr = Nothing

End Function


Sub Emailing()

On Error Resume Next


CarryOn = MsgBox("Send reports?" & vbNewLine & vbNewLine & "NOTE:" & vbNewLine & "Emailing can take up to several minutes to complete.", vbYesNo + vbQuestion)
    
If CarryOn = vbYes Then

Dim sdFolder As String

sdFolder = GetEmailFolder

If sdFolder = "" Then
    Exit Sub
End If

Application.ScreenUpdating = False
Application.DisplayAlerts = False


'--------------------------------------------'
'Emailing'


Dim i As Integer
Dim end_row As Integer
Dim Client_code As String
Dim month As String
Dim signature As String
Dim Filename As String
Dim path As String
Dim objOutlook As New Outlook.Application
Dim objMail As MailItem

month = Sheets("Start").Range("A1")
signature = "XXX Ltd"

path = sdFolder & "\"


end_row = Sheets("Summary").Range("D1").CurrentRegion.Rows.Count


Set objOutlook = New Outlook.Application

For i = 2 To end_row Step 1
    Client_code = Sheets("Summary").Cells(i, "E")
    Filename = Dir(path & Client_code & "*.pdf")
    Set objMail = objOutlook.CreateItem(olMailItem)
    
    Do While Len(Filename) > 0
        DoEvents
        A = Application.WorksheetFunction.Clean(path & Filename)
        objMail.Attachments.Add A
        Filename = Dir
        
    Loop
    
    If Sheets("Summary").Cells(i, "V") <> "" Then
        objMail.To = Sheets("Summary").Cells(i, "V")
        objMail.Subject = Client_code & " - Reports " & month
        objMail.HTMLBody = "<html><body><p>Hi,</p><p>Please find attached reports.</p><p>Kind regards,</p>" & signature & "</body></html>"
        objMail.Send
    End If
    Set objMail = Nothing
    
Next i

'----------------------------------------------------'


Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox ("Emailing Complete")

End If

End Sub





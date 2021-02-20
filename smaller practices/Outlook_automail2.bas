Attribute VB_Name = "automail"
Option Explicit
Private Sub automail()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim toRecipient As String
    Dim todaysdate As Date
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    toRecipient = "xz@gmail.com; yt@gmail.com"
    todaysdate = Date
    
    On Error Resume Next
    With OutMail
        .Display
        .To = toRecipient
        .Subject = todaysdate & " - Daily data"
        .HTMLBody = "Dear Colleagues," & "<br/><br/>" & "Please find today's data attached. " & "<br/><br/>" & "Have a nice day, " & .HTMLBody
        .Attachments.Add (ActiveWorkbook.FullName)
    End With
End Sub


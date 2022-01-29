Private Sub Worksheet_Change(ByVal Target As Range)

'如果儲存格A1>=100則Notify
If (Target.Address = "$A$1" And Val([A1]) >= 100) Then
    Call notify
End If

End Sub


Sub notify()

Dim xhr As Object
Set xhr = CreateObject("Microsoft.XMLHTTP")
xhr.Open "POST", "https://notify-api.line.me/api/notify", False
Dim Token As String
Token = Line的Token 'ibhx3..................Bv7A5LW
'注意 "Bearer{有一個空白字元}" + Token
xhr.setRequestHeader "Authorization", "Bearer " + Token
xhr.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;"
Dim postdata As String

postdata = "message=測試訊息" + CStr(Now)
xhr.send (postdata)
Set xhr= Nothing
End Sub

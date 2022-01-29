Private Function CreateObjectx86(Optional sProgID, Optional bClose = False)
'取得Javascrip Encoded 建立物件
    Static oWnd As Object
    Dim bRunning As Boolean

    #If Win64 Then
        bRunning = InStr(TypeName(oWnd), "HTMLWindow") > 0
        If bClose Then
            If bRunning Then oWnd.Close
            Exit Function
        End If
        If Not bRunning Then
            Set oWnd = CreateWindow()
            oWnd.execScript "Function CreateObjectx86(sProgID): Set CreateObjectx86 = CreateObject(sProgID): End Function", "VBScript"
        End If
        Set CreateObjectx86 = oWnd.CreateObjectx86(sProgID)
    #Else
        Set CreateObjectx86 = CreateObject(sProgID)
    #End If

End Function



Private Function CreateWindow()
'取得Javascrip Encoded 建立物件
    ' source http://forum.script-coding.com/viewtopic.php?pid=75356#p75356
    Dim sSignature, oShellWnd, oProc

    On Error Resume Next
    sSignature = Left(CreateObject("Scriptlet.TypeLib").GUID, 38)
    CreateObject("WScript.Shell").Run "%systemroot%\syswow64\mshta.exe about:""about:<head><script>moveTo(-32000,-32000);document.title='x86Host'</script><hta:application showintaskbar=no /><object id='shell' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shell.putproperty('" & sSignature & "',document.parentWindow);</script></head>""", 0, False
    Do
        For Each oShellWnd In CreateObject("Shell.Application").Windows
             Set CreateWindow = oShellWnd.GetProperty(sSignature)
             If err.Number = 0 Then Exit Function
             err.Clear
        Next
    Loop

End Function


Public Function encodeURL(str As String) As String
    Dim ScriptEngine As Object
    Dim encoded As String

    Set ScriptEngine = CreateObjectx86("scriptcontrol")
    ScriptEngine.Language = "JScript"

    encoded = ScriptEngine.Run("encodeURIComponent", str)

    encodeURL = encoded
End Function

Public Sub writefile(ByVal str)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile("1.txt")
    oFile.WriteLine str
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End Sub

Public Sub writefileto(ByVal str, ByVal filename)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(filename & ".txt")
    oFile.WriteLine str
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End Sub

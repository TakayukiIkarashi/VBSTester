Option Explicit

Const VBS_TESTER_LOG_NAME = ".\VbsTester.log"
Const VBS_TESTER_ADR_NAME = ".\Address.ini"

Dim vt
Set vt = New VbsTester
vt.Test()
Set vt = Nothing

Class VbsTester
    Private fso

    Private Sub Class_Initialize()
        Set fso = CreateObject("Scripting.FileSystemObject")
    End Sub

    Private Sub Class_Terminate()
        Set fso = Nothing
    End Sub

    Public Sub Test()
        Call CreateLog()
        Dim dirObj
        Set dirObj = fso.GetFolder(".")
        Dim fileObj
        For Each fileObj In dirObj.Files
            Call TestMethod(fileObj.Name)
        Next
        Set fileObj = Nothing
        Set dirObj = Nothing
        Call SendLog()
        Call DeleteLog()
    End Sub

    Private Sub TestMethod(ByVal fileName)
        If Not (LCase(Right(fileName, 4)) = ".inc") Then
            Exit Sub
        End If
        Execute Include(fileName)
        Dim tt
        Set tt = New Tester
        Call tt.ExecuteTest()
        Call WriteText(tt.Log)
        Set tt = Nothing
    End Sub

    Private Function Include(ByVal fileName)
        Include = fso.OpenTextFile(fileName, 1, False).ReadAll()
    End Function

    Private Function ReadText(ByVal filePath)
        ReadText = ""
        Dim fileStream
        Set fileStream = fso.OpenTextFile(filePath, 1, False)
        Dim buf
        buf = fileStream.ReadAll
        If (CStr(buf) = "") Then
            buf = ""
        End If
        Set fileStream = Nothing
        ReadText = buf
    End Function

    Private Sub WriteText(ByVal value)
        Dim fileStream
        Set fileStream = fso.OpenTextFile(VBS_TESTER_LOG_NAME, 8, False)
        fileStream.WriteLine value
        Set fileStream = Nothing
    End Sub

    Private Sub CreateLog()
        Dim ts
        Set ts = fso.CreateTextFile(VBS_TESTER_LOG_NAME)
        ts.Close
        Set ts = Nothing
    End Sub

    Private Sub SendLog()
        Dim olApp
        Set olApp = CreateObject("Outlook.Application")
        Dim myItem
        Set myItem = olApp.CreateItem(0)
        Set olApp = Nothing
        myItem.To = Replace(ReadText(VBS_TESTER_ADR_NAME), vbCrLf, ";")
        myItem.Subject = "From VbsTester"
        myItem.Body = ReadText(VBS_TESTER_LOG_NAME)
        myItem.Send
    End Sub

    Private Sub DeleteLog()
        Dim fileObj
        Set fileObj = fso.GetFile(VBS_TESTER_LOG_NAME)
        fileObj.Delete
        Set fileObj = Nothing
    End Sub
End Class

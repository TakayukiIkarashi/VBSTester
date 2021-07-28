Option Explicit

Class Tester
    Private mLog
    Private Sub Class_Initialize()
        mLog = ""
    End Sub
    Public Property Get Log()
        log = mLog
    End Property
    Private Property Let Log(ByVal sLog)
        mLog = sLog
    End Property
    Private Function Include(ByVal fileName)
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        Include = fso.OpenTextFile(fileName, 1, False).ReadAll()
        Set fso = Nothing
    End Function
    Private Sub Test(ByVal bIdn, ByVal msg)
        Dim sLog
        If (bIdn) Then
            sLog = msg & " ---> OK"
        Else
            sLog = msg & " ---> NG"
        End If
        If Not (Log = "") Then
            Log = Log & vbCrLf
        End If
        Log = Log & sLog
    End Sub

    '************************************************************
    ' Test Pattern
    '************************************************************
    Public Sub ExecuteTest()
        Execute Include(".\sample_source\Sample2.vbs")

        Dim cl
        Set cl = New Calc2

        '+++ Calc2 +++
        Call Test(cl.Calc2(1, 2) = 3, "Calc2")
        Call Test(cl.Calc2(3, 4) = 7, "Calc2")
        Call Test(cl.Calc2(5, 6) = 7, "Calc2")

        Set cl = Nothing
    End Sub
End Class

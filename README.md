# VBS Tester Ver.1.0.１

## 概要
このスクリプトは、VBScriptの自動単体テストを行うためのVBScriptです。
テストしたいVBScriptモジュールのテストケースを作成し、VbsTester.vbsと同一ディレクトリに格納したら、VbsTester.vbsを実行します。テスト結果は、Microsoft Outlookにて、指定のアドレスにメール送信します。

## 使い方
まずは、Address.iniをテキストエディタで開きます。
このファイルには、テスト結果を送信したいメールアドレスを登録します。
登録は、つぎのようにメールアドレスを1行ずつ入力します。

  mail_address1@gmail.com
  mail_address2@gmail.com
  mail_address3@gmail.com

VbsTester.vbsは、Microsoft Outlookによるメール送信機能を持っています。
もし、この機能が不要であれば、VbsTester.vbsを修正する必要があります。
網がけされている部分を、削除もしくはコメントアウトしてください。


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

この修正を行うと、VbsTester.vbsと同じディレクトリに、VbsTester.logというテキストファイルでテスト結果を出力します。
次に、テストケースを作成します。
テストケースは、拡張子”.inc”ファイルです。

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
            Execute Include(".\sample_source\Sample1.vbs")

            '+++ Calc1 +++
            Call Test(Calc1(1, 2) = 3, "Calc1")
            Call Test(Calc1(3, 4) = 7, "Calc1")
            Call Test(Calc1(5, 6) = 7, "Calc1")
        End Sub
    End Class

ExecuteTestメソッドないの1行目で、テストしたいVBSモジュールを指定します。

    Execute Include(Extension)

    Extension: 「VBSモジュールのパス」

次に、VBSモジュールないのメソッドのテストケースを記述します。
テストケースは、Testメソッドで記述します。
Testメソッドは、次のように指定します。

    Test(TestCase, TestCaseName)

    TestCase:テストケース
    TestCaseName:テストケース名

Testメソッドは、パラメータに指定されたテストケースが真となるように記述します。
このサンプルをみると、sample_sourceフォルダにSample1.vbsというVBSがあります。このVBSには、Calc1というメソッドがあります。Calc1は、パラメータに指定された2つの値を加算します。
ExecuteTestメソッドないのテストケースでは、

        Call Test(Calc1(1, 2) = 3, "Calc1")
        Call Test(Calc1(3, 4) = 7, "Calc1")
        Call Test(Calc1(5, 6) = 7, "Calc1")

となっていますので、1行目は「１＋２＝３」で真、２行目も「３＋４＝７」で真ですが、３行目は「５＋６＝７」となっていますので、偽です。
（テストケースが正しいのなら、Calc1メソッドは、「１＋２＝３」「３＋４＝７」「５＋６＝７」となるメソッドでなくてはならない!!）
また、VbsTester.vbsと同じディレクトリにある、拡張子”.inc”ファイルに記述されているテストケースすべてが、テストの対象となります。
このサンプルでは、「TestSample1.inc」と「TesSample2.inc」がテストファイルです。
このテストファイルに記述されているテストケースが、VbsTester.vbsの実行によってテストされます。
テスト結果は、Microsoft Outlookにてメールで指定のアドレスに送信されます。

    件名: From VbsTester
    本文:
    Calc1 ---> OK 
    Calc1 ---> OK 
    Calc1 ---> NG 
    Calc2 ---> OK 
    Calc2 ---> OK 
    Calc2 ---> NG


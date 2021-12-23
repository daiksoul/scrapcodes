Public fPath As String

Dim Serial As Integer
Dim go As Boolean
Dim fe As Boolean

Sub Printer()
'fPath = "C:\Users\" + Environ("Username") + "\Desktop\test.txt"
'Set fs = CreateObject("Scripting.FileSystemObject")

'If fs.FileExists(fPath) Then
'Set ba = fs.OpenTextFile(fPath, 1)
'Dim init As String
'init = ba.ReadLine
'ba.Close
'Serial = CInt(init)
'fe = True
'Else
Serial = 1
'fe = False
'End If

'MsgBox init
go = True

With Word.ActiveDocument.Content.Find

.Text = "<#SERIAL#>"
.Forward = True
.Execute

While go = True

If .Found = True Then
.Replacement.Text = CStr(Serial)
.Execute Replace:=wdReplaceOne, Forward:=True, Wrap:=wdFindContinue
Serial = Serial + 1
Else
go = False
End If

Wend

End With


If fe = True Then
Set ba = fs.OpenTextFile(fPath, 2)
ba.Write CStr(Serial)
ba.Close
Else
Set ba = fs.CreateTextFile(fPath)
ba.Write CStr(Serial)
ba.Close
End If

End Sub

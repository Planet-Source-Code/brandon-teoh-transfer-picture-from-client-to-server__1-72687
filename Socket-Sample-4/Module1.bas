Attribute VB_Name = "Module1"
Option Explicit

Public Sub writeFileData(MyData As Variant, transType As String, Optional filePathName As String)
'On Error GoTo m_err2:
Dim strm As TextStream
Dim fso As New FileSystemObject

If filePathName = "" Then
   filePathName = App.Path & "\Data" & "\" & "DataFile.txt"
End If


With fso

'check if file exist

If .FileExists(filePathName) Then
     If transType = "Append" Then
        'update file
        Set strm = .OpenTextFile(filePathName, ForAppending, False)
        strm.Write (MyData)
     ElseIf transType = "Overwrite" Then
       '=== Kill the file=====
        Kill CStr(filePathName)
       '=== create a new file
        Set strm = .CreateTextFile(filePathName, True)
       
       '==== write new data ====================
       strm.Write (MyData)
     End If
Else
     'create new file
     Set strm = .CreateTextFile(filePathName, True)
     strm.Write (MyData)
End If

End With

m_quit:
  Set strm = Nothing
  Set fso = Nothing
  Exit Sub
m_err:
   GoTo m_quit:
m_err2:
   GoTo m_quit:
End Sub

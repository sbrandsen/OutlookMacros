Option Explicit

Const strPath = "c:\temp\deleted msg.csv"
Sub DeleteDuplicateEmails()

Dim lngCnt As Long
Dim objMail As Object
Dim objFSO As Object
Dim objTF As Object

Dim objDic As Object
Dim objItem As Object
Dim olApp As Outlook.Application
Dim olNS As NameSpace
Dim olFolder As Folder
Dim olFolder2 As Folder
Dim strCheck As String

Set objDic = CreateObject("scripting.dictionary")
Set objFSO = CreateObject("scripting.filesystemobject")
Set objTF = objFSO.CreateTextFile(strPath)
objTF.WriteLine "Subject"

Set olApp = Outlook.Application
Set olNS = olApp.GetNamespace("MAPI")
Set olFolder = olNS.PickFolder

If olFolder Is Nothing Then Exit Sub

On Error Resume Next
Set olFolder2 = olFolder.Folders("removed items")
On Error GoTo 0

If olFolder2 Is Nothing Then Set olFolder2 = olFolder.Folders.Add("removed items")


For lngCnt = olFolder.Items.Count To 1 Step -1

Set objItem = olFolder.Items(lngCnt)

strCheck = objItem.Subject & "," & objItem.Body & ","
strCheck = Replace(strCheck, ", ", Chr(32))

    If objDic.Exists(strCheck) Then
       objItem.Move olFolder2
       objTF.WriteLine Replace(objItem.Subject, ", ", Chr(32))
    Else
        objDic.Add strCheck, True
    End If
Next

If objTF.Line > 2 Then
    MsgBox "duplicate items were removed to ""removed items""", vbCritical, "See " & strPath & " for details"
Else
    MsgBox "No duplicates found"
End If
End Sub

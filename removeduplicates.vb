Option Explicit
Dim deletedcount As Integer

Private Sub Main()
    Dim objNameSpace As Outlook.NameSpace
    Dim objMainFolder As Outlook.Folder
    Dim objDeletedFolder As Outlook.Folder
    
    Set objNameSpace = Application.GetNamespace("MAPI")
    
    MsgBox "Which folder should I search recusively?"
    Set objMainFolder = objNameSpace.PickFolder
    
    MsgBox "Which folder should the duplicates be moved to?"
    Set objDeletedFolder = objNameSpace.PickFolder
    
    If objMainFolder Is Nothing Then Exit Sub
    If objDeletedFolder Is Nothing Then Exit Sub
    
    Call ProcessCurrentFolder(objMainFolder, objDeletedFolder)
    MsgBox "Deleted " & deletedcount & " duplicates"
    
End Sub
 
Private Sub ProcessCurrentFolder(ByVal objParentFolder As Outlook.MAPIFolder, objDeletedFolder As Folder)
    Dim objCurFolder As Outlook.MAPIFolder
    Dim objMail As Outlook.MailItem
    
    On Error Resume Next
    
    Call DeleteDuplicateEmails(objParentFolder, objDeletedFolder)
    
    '    Process the subfolders in the folder recursively
    If (objParentFolder.Folders.Count > 0) Then
        For Each objCurFolder In objParentFolder.Folders
            Call ProcessCurrentFolder(objCurFolder, objDeletedFolder)
        Next
    End If
End Sub


Sub DeleteDuplicateEmails(olFolder As Folder, olFolder2 As Folder)

Dim lngCnt As Long
Dim objMail As Object
Dim objFSO As Object
Dim objTF As Object

Dim objDic As Object
Dim objItem As Object
Dim olApp As Outlook.Application
Dim olNS As NameSpace
Dim strCheck As String

Set objDic = CreateObject("scripting.dictionary")
Set objFSO = CreateObject("scripting.filesystemobject")

For lngCnt = olFolder.Items.Count To 1 Step -1

Set objItem = olFolder.Items(lngCnt)

strCheck = objItem.Subject & "," & objItem.Body & ","
strCheck = Replace(strCheck, ", ", Chr(32))

    If objDic.Exists(strCheck) Then
       objItem.Move olFolder2
       deletedcount = deletedcount + 1
    Else
        objDic.Add strCheck, True
    End If
Next
End Sub

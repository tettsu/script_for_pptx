msgbox "PowerPoint�͏I�����Ă��瑀����s�Ȃ��ĉ������B"

Dim folderPath
folderPath = InputBox("�f�B���N�g���̃p�X����͂��ĉ������B���͂������ꍇ�́A�X�N���v�g�̂���f�B���N�g����Ώۂɂ��܂��B","title")

Dim sysObj
Set sysObj = CreateObject("Scripting.FileSystemObject")
if sysObj.FolderExists(folderPath) = false then
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    folderPath = fso.getParentFolderName(WScript.ScriptFullName)
end if




Dim powerPoint
Set powerPoint = CreateObject("PowerPoint.Application")
powerPoint.Visible = True
Dim Target

'�w��t�H���_�̒��̃t�@�C��
For Each oFile In sysObj.GetFolder(folderPath).files
  Target =  oFile.Name
  '�g���q�̔���
  If LCase(sysObj.GetExtensionName(Target)) = "ppt" Or LCase(sysObj.GetExtensionName(Target)) = "pptx" Then
    ''Target�ɑ΂��鏈��
    Call repSub(folderPath & "\" & Target, fromStr,toStr, powerPoint)
  End If
Next

powerPoint.Quit
Set powerPoint = Nothing



Sub repSub(filePath, fromStr, toStr, powerPoint)
On Error Resume Next
  With powerPoint.Presentations.Open(filePath)
  For Each myS In powerPoint.ActivePresentation.Slides
     For Each mySP In myS.Shapes
       mySP.TextFrame.TextRange = Replace(mySP.TextFrame.TextRange, fromStr, toStr)
     Next
  Next
 .Save
 .Close
  End With
End Sub
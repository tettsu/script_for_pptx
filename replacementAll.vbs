msgbox "PowerPoint�͏I�����Ă��瑀����s�Ȃ��ĉ������B"

Dim folderPath
folderPath = InputBox("�f�B���N�g���̃p�X����͂��ĉ������B","title")

Dim sysObj
Set sysObj = CreateObject("Scripting.FileSystemObject")
if sysObj.FolderExists(folderPath) = false then
    msgbox "���݂��Ȃ��f�B���N�g���ł��B"
    WScript.Quit
end if

Dim fromStr
Dim toStr

fromStr = InputBox("�u�����錳�̕��������͂��ĉ������B","title")
if fromStr = "" then
    msgbox "�������͂����������̂ŏI�����܂��B"
    WScript.Quit
end if

toStr = InputBox("�u����̕��������͂��ĉ������B","title")

if msgbox(folderPath & "�ȉ��̃p���|�t�@�C����" & fromStr & "����" & toStr  & "�ɒu�����Ă������ł����H",vbYesNo + vbQuestion) = vbYes then

    Dim poworPoint
    Set poworPoint = CreateObject("PowerPoint.Application")
    poworPoint.Visible = True
    Dim Target

    '�w��t�H���_�̒��̃t�@�C��
    For Each oFile In sysObj.GetFolder(folderPath).files
     Target =  oFile.Name
     '�g���q�̔���
      If LCase(sysObj.GetExtensionName(Target)) = "ppt" Or LCase(sysObj.GetExtensionName(Target)) = "pptx" Then
       ''Target�ɑ΂��鏈��
       Call repSub(folderPath & "\" & Target, fromStr,toStr, poworPoint)
      End If
    Next

    poworPoint.Quit
    Set poworPoint = Nothing
end if

Sub repSub(filePath, fromStr, toStr, poworPoint)
On Error Resume Next
  With poworPoint.Presentations.Open(filePath)
  For Each myS In poworPoint.ActivePresentation.Slides
     For Each mySP In myS.Shapes
       mySP.TextFrame.TextRange = Replace(mySP.TextFrame.TextRange, fromStr, toStr)
     Next
  Next
 .Save
 .Close
  End With
End Sub

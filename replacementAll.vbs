msgbox "PowerPoint�͏I�����Ă��瑀����s�Ȃ��ĉ������B"

Dim path
path = InputBox("�f�B���N�g���̃p�X����͂��ĉ������B���͂������ꍇ�́A�X�N���v�g�̂���f�B���N�g����Ώۂɂ��܂��B","title")

Dim so
Set so = CreateObject("Scripting.FileSystemObject")
if so.FolderExists(path) = false then
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    path = fso.getParentFolderName(WScript.ScriptFullName)
end if

Dim fromStr
Dim toStr

fromStr = InputBox("�u���Ώۂ̕��������͂��ĉ������B�������͂��Ȃ��ƏI�����܂��B","title")
if fromStr = "" then
    msgbox "�������͂����������̂ŏI�����܂��B"
    WScript.Quit
end if

toStr = InputBox("�u����̕��������͂��ĉ������B","title")

if msgbox(path & "�z����pptx���́h" & fromStr & "�h���h" & toStr  & "�h�ɒu�����Ă悢�ł����H�Ȃ��A�t�@�C�����͑ΏۊO�ł��B", vbYesNo + vbQuestion) = vbYes then

    Dim powerPoint
    Set powerPoint = CreateObject("PowerPoint.Application")
    powerPoint.Visible = True
    Dim target

    '�w��t�H���_�̒��̃t�@�C��
    For Each oFile In so.GetFolder(path).files
     target =  oFile.Name
     '�g���q�̔���
      If LCase(so.GetExtensionName(target)) = "ppt" Or LCase(so.GetExtensionName(target)) = "pptx" Then
       ''Target�ɑ΂��鏈��
       Call repSub(path & "\" & target, fromStr, toStr, powerPoint)
      End If
    Next

    powerPoint.Quit
    Set powerPoint = Nothing
end if

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

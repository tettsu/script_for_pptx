msgbox "PowerPoint�͏I�����Ă��瑀����s�Ȃ��ĉ������B"

Dim path
path = InputBox("�f�B���N�g���̃p�X����͂��ĉ������B���͂������ꍇ�́A�X�N���v�g�̂���f�B���N�g����Ώۂɂ��܂��B","title")

Dim so
Set so = CreateObject("Scripting.FileSystemObject")

'�f�B���N�g�������ꍇ�̓J�����g�f�B���N�g����Ώۂɂ���
if so.FolderExists(path) = false then
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    path = fso.getParentFolderName(WScript.ScriptFullName)
end if

if msgbox(path & "�z����pptx���̃��^�����폜���Ă悢�ł����H", vbYesNo + vbQuestion) = vbYes then
    Dim powerPoint
    Set powerPoint = CreateObject("PowerPoint.Application")
    Dim target

    '�w��t�H���_�̒��̃t�@�C��
    For Each oFile In so.GetFolder(path).files
     target =  oFile.Name
     '�g���q�̔���
      If LCase(so.GetExtensionName(target)) = "ppt" Or LCase(so.GetExtensionName(target)) = "pptx" Then
        '���^�����폜("99"�͑S�č폜)
        powerPoint.Presentations.Open(oFile)
        With powerPoint.ActivePresentation
            .RemoveDocumentInformation(99)
            .Save
            .Close
        End With
      End If
    Next

    powerPoint.Quit
    Set powerPoint = Nothing
end if

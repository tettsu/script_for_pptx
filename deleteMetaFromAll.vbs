msgbox "PowerPointは終了してから操作を行なって下さい。"

Dim path
path = InputBox("ディレクトリのパスを入力して下さい。入力が無い場合は、スクリプトのあるディレクトリを対象にします。","title")

Dim so
Set so = CreateObject("Scripting.FileSystemObject")

'ディレクトリ無い場合はカレントディレクトリを対象にする
if so.FolderExists(path) = false then
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    path = fso.getParentFolderName(WScript.ScriptFullName)
end if

if msgbox(path & "配下のpptx内のメタ情報を削除してよいですか？", vbYesNo + vbQuestion) = vbYes then
    Dim powerPoint
    Set powerPoint = CreateObject("PowerPoint.Application")
    Dim target

    '指定フォルダの中のファイル
    For Each oFile In so.GetFolder(path).files
     target =  oFile.Name
     '拡張子の判別
      If LCase(so.GetExtensionName(target)) = "ppt" Or LCase(so.GetExtensionName(target)) = "pptx" Then
        'メタ情報を削除("99"は全て削除)
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

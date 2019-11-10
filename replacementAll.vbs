msgbox "PowerPointは終了してから操作を行なって下さい。"

Dim path
path = InputBox("ディレクトリのパスを入力して下さい。入力が無い場合は、スクリプトのあるディレクトリを対象にします。","title")

Dim so
Set so = CreateObject("Scripting.FileSystemObject")
if so.FolderExists(path) = false then
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    path = fso.getParentFolderName(WScript.ScriptFullName)
end if

Dim fromStr
Dim toStr

fromStr = InputBox("置換対象の文字列を入力して下さい。何も入力しないと終了します。","title")
if fromStr = "" then
    msgbox "何も入力が無かったので終了します。"
    WScript.Quit
end if

toStr = InputBox("置換後の文字列を入力して下さい。","title")

if msgbox(path & "配下のpptx内の”" & fromStr & "”を”" & toStr  & "”に置換してよいですか？なお、ファイル名は対象外です。", vbYesNo + vbQuestion) = vbYes then

    Dim powerPoint
    Set powerPoint = CreateObject("PowerPoint.Application")
    powerPoint.Visible = True
    Dim target

    '指定フォルダの中のファイル
    For Each oFile In so.GetFolder(path).files
     target =  oFile.Name
     '拡張子の判別
      If LCase(so.GetExtensionName(target)) = "ppt" Or LCase(so.GetExtensionName(target)) = "pptx" Then
       ''Targetに対する処理
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

msgbox "PowerPointは終了してから操作を行なって下さい。"

Dim folderPath
folderPath = InputBox("ディレクトリのパスを入力して下さい。入力が無い場合は、スクリプトのあるディレクトリを対象にします。","title")

Dim sysObj
Set sysObj = CreateObject("Scripting.FileSystemObject")
if sysObj.FolderExists(folderPath) = false then
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    folderPath = fso.getParentFolderName(WScript.ScriptFullName)
end if

Dim fromStr
Dim toStr

fromStr = InputBox("置換対象の文字列を入力して下さい。何も入力しないと終了します。","title")
if fromStr = "" then
    msgbox "何も入力が無かったので終了します。"
    WScript.Quit
end if

toStr = InputBox("置換後の文字列を入力して下さい。","title")

if msgbox(folderPath & "配下のpptx内の”" & fromStr & "”を”" & toStr  & "”に置換してよいですか？なお、ファイル名は対象外です。",vbYesNo + vbQuestion) = vbYes then

    Dim powerPoint
    Set powerPoint = CreateObject("PowerPoint.Application")
    powerPoint.Visible = True
    Dim Target

    '指定フォルダの中のファイル
    For Each oFile In sysObj.GetFolder(folderPath).files
     Target =  oFile.Name
     '拡張子の判別
      If LCase(sysObj.GetExtensionName(Target)) = "ppt" Or LCase(sysObj.GetExtensionName(Target)) = "pptx" Then
       ''Targetに対する処理
       Call repSub(folderPath & "\" & Target, fromStr,toStr, powerPoint)
      End If
    Next

    powerPoint.Quit
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
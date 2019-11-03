msgbox "PowerPointは終了してから操作を行なって下さい。"

Dim folderPath
folderPath = InputBox("ディレクトリのパスを入力して下さい。","title")

Dim sysObj
Set sysObj = CreateObject("Scripting.FileSystemObject")
if sysObj.FolderExists(folderPath) = false then
    msgbox "存在しないディレクトリです。"
    WScript.Quit
end if

Dim fromStr
Dim toStr

fromStr = InputBox("置換する元の文字列を入力して下さい。","title")
if fromStr = "" then
    msgbox "何も入力が無かったので終了します。"
    WScript.Quit
end if

toStr = InputBox("置換後の文字列を入力して下さい。","title")

if msgbox(folderPath & "以下のパワポファイルを" & fromStr & "から" & toStr  & "に置換してもいいですか？",vbYesNo + vbQuestion) = vbYes then

    Dim poworPoint
    Set poworPoint = CreateObject("PowerPoint.Application")
    poworPoint.Visible = True
    Dim Target

    '指定フォルダの中のファイル
    For Each oFile In sysObj.GetFolder(folderPath).files
     Target =  oFile.Name
     '拡張子の判別
      If LCase(sysObj.GetExtensionName(Target)) = "ppt" Or LCase(sysObj.GetExtensionName(Target)) = "pptx" Then
       ''Targetに対する処理
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

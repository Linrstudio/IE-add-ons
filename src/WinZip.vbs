' 此脚本用于将文件夹压缩成 zip
' @Author: Vickey Chen
' @Copyright: (C)Linr 2013 http://www.xiaogezi.cn/
' 命令行参数：WinZip.vbs d:\folder\ d:\folder\filename.zip
' 常用于打包工作
' Notepad++ 参数 $(FULL_CURRENT_PATH)
'

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Sub Zip(ByVal mySourceDir, ByVal myZipFile)
    If fso.GetExtensionName(myZipFile) <> "zip" Then
        Exit Sub
    ElseIf fso.FolderExists(mySourceDir) Then
        FType = "Folder"
    ElseIf fso.FileExists(mySourceDir) Then
        FType = "File"
        FileName = fso.GetFileName(mySourceDir)
        FolderPath = Left(mySourceDir, Len(mySourceDir) - Len(FileName))
    Else
        Exit Sub
    End If
    Set f = fso.CreateTextFile(myZipFile, True)
        f.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
        f.Close
    Set objShell = CreateObject("Shell.Application")
    Select Case Ftype
        Case "Folder"
            Set objSource = objShell.NameSpace(mySourceDir)
            Set objFolderItem = objSource.Items()
        Case "File"
            Set objSource = objShell.NameSpace(FolderPath)
            Set objFolderItem = objSource.ParseName(FileName)
    End Select
    Set objTarget = objShell.NameSpace(myZipFile)
    intOptions = 256
    objTarget.CopyHere objFolderItem, intOptions
    Do
        WScript.Sleep 1000
    Loop Until objTarget.Items.Count > 0
End Sub

Function getCommand(index)
	on error resume next
	getCommand = WScript.Arguments(index)
End Function

Sub goZip()

	dim srcStr
		srcStr = getCommand(0)
	
	dim tgtStr
		tgtStr = getCommand(1)
	
	dim cmdStr
		cmdStr = getCommand(2)
		
	dim targetDir, fileName
	fileName = fso.GetFileName(srcStr)
	targetDir = replace(srcStr, fileName, "")
	
	If srcStr <> "" And tgtStr <> "" And fso.FolderExists(targetDir) And fso.FolderExists(srcStr) Then
	
		Zip targetDir, tgtStr
		
		If cmdStr <> "" Then callback(cmdStr)
	
	End if

End Sub

Sub callback(str)

	dim wsh
	set wsh = CreateObject("WScript.Shell")
		wsh.Run str, 1, false

End Sub

Call goZip
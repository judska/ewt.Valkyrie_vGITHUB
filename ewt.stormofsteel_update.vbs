' source folder
SourceFolder = "C:\Program Files (x86)\Steam\steamapps\workshop\content\244450\759785045" 
' destination folder
DestinationFolder = "C:\Program Files (x86)\Steam\steamapps\common\Men of War Assault Squad 2\mods\ewt.valkyrie_v0.69.2" 

On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShellApp = CreateObject("Shell.Application")

Set objFSO1 = CreateObject("Scripting.FileSystemObject")
Set objShellApp1 = CreateObject("Shell.Application")

LogPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
Set LogStream = objFSO.OpenTextFile(LogPath & "\CopyLog.log", 8, True)
LogStream.WriteLine "Updating start " & Now()
CopyFiles SourceFolder, DestinationFolder 
LogStream.WriteLine "Updating finish " & Now()
LogStream.Close
var WSHShell = WScript.CreateObject("WScript.Shell")
MsgBox "Mod has been updated", vbOKOnly, "Valour update"

Sub CopyFiles(FolderPath, FolderPath1)
    On Error Resume Next
    Set objFolderItems = objShellApp.NameSpace(FolderPath).Items()
	Set objFolderItems1 = objShellApp1.NameSpace(FolderPath1).Items()
	
    For Each objFolderItem In objFolderItems
        SourceSubPath = Mid(objFolderItem.Path, Len(FolderPath) + 2)
		If objFolderItem.IsFolder Then
			SubPath = Mid(objFolderItem.Path, Len(SourceFolder) + 1)
			TargetPath = DestinationFolder & SubPath
            CopyFiles objFolderItem.Path, TargetPath
        Else
            Set objFile = objFSO.GetFile(objFolderItem.Path)
			isFound = False
			For Each objFolderItem1 In objFolderItems1
				DestinationSubPath = Mid(objFolderItem1.Path, Len(FolderPath1) + 2)
				IF SourceSubPath =DestinationSubPath Then
					isFound = True
					Set objFile1 = objFSO1.GetFile(objFolderItem1.Path)
					IF objFile1.DateLastModified < objFile.DateLastModified Then
						CopyFile objFolderItem.Path
					End If
				End If	
			Next
			IF NOT isFound Then
				CopyFile objFolderItem.Path
			End If
        End If
    Next
End Sub

Sub CopyFile(FilePath)
    On Error Resume Next
    SubPath = Mid(FilePath, Len(SourceFolder) + 1)
    TargetPath = DestinationFolder & SubPath
    FolderPath = objFSO.GetParentFolderName(TargetPath)
    If Not objFSO.FolderExists(FolderPath) Then
        CreateFolder FolderPath
    End If

    If objFSO.FileExists(TargetPath) Then
        Set objFile = objFSO.GetFile(TargetPath)
        If objFile.Attributes And 1 Then
            objFile.Attributes = objFile.Attributes - 1
        End If
    End If
    objFSO.CopyFile FilePath, TargetPath, True
    If Err.Number <> 0 Then
        LogStream.WriteLine
        LogStream.WriteLine FilePath
        LogStream.WriteLine Err.Description
        LogStream.WriteLine
        Err.Clear
    Else
        LogStream.WriteLine TargetPath
    End If
End Sub

Sub CreateFolder (FolderPath)
    On Error Resume Next
    ParentFolder = objFSO.GetParentFolderName(FolderPath)
    If Not objFSO.FolderExists(ParentFolder) Then
        CreateFolder ParentFolder
    End If
    objFSO.CreateFolder FolderPath
End Sub
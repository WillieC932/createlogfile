'*************************************************
'Created By:  William Collins
'Date:  11/21/2008
'Script Name:  CreateLogFile.vbs
'Version:  2.0
'Key Concepts are listed below:
'This script will recursively read files in folders & subfolders and write them to a text file

'*************************************************


'+++++HEADER & REFERENCE INFORMATION SECTION+++++

Option Explicit
'On Error Resume Next 'UNCOMMENT AS NEEDED

Dim FolderPath		'path to the folder to be searched for files
Dim objFSO			'the FileSystemObject
Dim objSubFolder	'individual subfolder
Dim objFolder		'the folder object
Dim colFiles		'collection of files from files method
Dim objFile			'individual file object
Dim strOut			'single output variable from Browse dialog
Dim objLogFile		'File in which results are written
Dim Subfolder		'single subfolder
Const ForAppending = 8

objLogFile = "\twg_script_log.txt"

'+++++WORKER INFORMATION SECTION+++++

subCheckWscript	'ensures script is running under wscript
subGetFolder	' calls the BrowseForFolder method

set objFSO = CreateObject("Scripting.FileSystemObject")
set objFolder = objFSO.GetFolder(FolderPath)
Set colFiles = objFolder.Files

For Each objFile in colFiles
	strOut = strOut & objFile.Path & vbCrLf
Next

subSubFolders objFSO.GetFolder(FolderPath)

subGetFolder	' calls the BrowseForFolder method

If objFSO.FileExists(FolderPath & objLogFile) Then
	Set objFile = objFSO.OpenTextFile(FolderPath & objLogFile, ForAppending)
	objFile.Write "File Appended " & Now & vbCrLf
Else
	Set objFile = objFSO.CreateTextFile(FolderPath & objLogFile)
	objfile.write "File created " & now & vbCrLf
End If



'+++++OUTPUT INFORMATION SECTION+++++


objfile.write strOut



'+++++Functions and Subroutines+++++


Sub subCheckWscript
If UCase(Right(WScript.FullName, 11)) = "CSCRIPT.EXE" Then
	wScript.Echo "This script must be run under WScript."
	WScript.Quit
End If
End Sub




Sub subGetFolder
Dim objShell, objFolder, objFolderItem, ObjPath
Const windowHandle = 0
Const folderOnly = 0
Const folderAndFiles = &H4200&
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.BrowseForFolder(WindowHandle, "Select a folder:", folderAndFiles)
Set objFolderItem = objFolder.Self
FolderPath = objFolderItem.Path
End sub




Sub subSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
    set colFiles = subfolder.Files
    	For Each objFile in colFiles
		strOut = strOut & objFile.Path & vbCrLf
		Next
    subSubFolders Subfolder
    Next
End Sub
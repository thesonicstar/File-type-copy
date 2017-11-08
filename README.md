# File-type-copy
Option Explicit

' The list of files to copy. Should be a text file with one file on each row. No paths - just file name.
Const strFileList = "c:\filelist.txt"

' Should files be overwriten if they already exist? TRUE or FALSE.
Const blnOverwrite = TRUE

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim objShell
Set objShell = CreateObject("Shell.Application")

Dim objFolder, objFolderItem

' Get the source path for the copy operation.
Dim strSourceFolder
Set objFolder = objShell.BrowseForFolder(0, "Select source folder", 0 )
If objFolder Is Nothing Then Wscript.Quit
Set objFolderItem = objFolder.Self
strSourceFolder = objFolderItem.Path

' Get the target path for the copy operation.
Dim strTargetFolder
Set objFolder = objShell.BrowseForFolder(0, "Select target folder", 0 )
If objFolder Is Nothing Then Wscript.Quit
Set objFolderItem = objFolder.Self
strTargetFolder = objFolderItem.Path


Const ForReading = 1
Dim objFileList
Set objFileList = objFSO.OpenTextFile(strFileList, ForReading, False)

Dim strFileToCopy, strSourceFilePath, strTargetFilePath
Dim strResults, iSuccess, iFailure
iSuccess = 0
iFailure = 0

On Error Resume Next
Do Until objFileList.AtEndOfStream
    ' Read next line from file list and build filepaths
    strFileToCopy = objFileList.Readline
    strSourceFilePath = objFSO.BuildPath(strSourceFolder, strFileToCopy)
    strTargetFilePath = objFSO.BuildPath(strTargetFolder, strFileToCopy)
    ' Copy file to specified target folder.
    Err.Clear
    objFSO.CopyFile strSourceFilePath, strTargetFilePath, blnOverwrite
    If Err.Number = 0 Then
        ' File copied successfully
        iSuccess = iSuccess + 1
        If Instr(1, Wscript.Fullname, "cscript.exe", 1) > 0 Then
            ' Running cscript, output text to screen
            Wscript.Echo strFileToCopy & " copied successfully"
        End If
    Else
        ' Error copying file
        iFailure = iFailure + 1
        TextOut "Error " & Err.Number & " (" & Err.Description & ") trying to copy " & strFileToCopy
    End If
Loop

strResults = strResults & vbCrLf
strResults = strResults & iSuccess & " files copied successfully." & vbCrLf
strResults = strResults & iFailure & " files generated errors" & vbCrLf
Wscript.Echo strResults

Sub TextOut(strText)
    If Instr(1, Wscript.Fullname, "cscript.exe", 1) > 0 Then
        ' Running cscript, use direct output
        Wscript.Echo strText
    Else
        strResults = strResults & strText & vbCrLf
    End If
End Sub

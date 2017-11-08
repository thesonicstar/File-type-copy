# File-type-copy-recurse
' Read a list of filelist from text file
' and copy those filelist from SourceFolder\SubFolders to TargetFolder

' Should files be overwriten if they already exist? TRUE or FALSE.
Const blnOverwrite = TRUE

Dim objFSO, objShell, WSHshell, objFolder, objFolderItem, strExt, strSubFolder
Dim objFileList, strFileToCopy, strSourceFilePath, strTargetFilePath 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")
Set WSHshell = CreateObject("WScript.Shell")
Const ForReading = 1

' Make the script useable on anyone's desktop without typing in the path
DeskTop = WSHShell.SpecialFolders("Desktop")
strFileList = DeskTop & "\" & "cdmsfiles.txt"

' File Extension type
strExt = InputBox("Please enter the File type" _
& vbcrlf & "For Example: jpg or tif")
If strExt="" Then 
   WScript.Echo "Invalid Input, Script Canceled"
Wscript.Quit
End if

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

Set objFileList = objFSO.OpenTextFile(strFileList, ForReading, False)

On Error Resume Next
Do Until objFileList.AtEndOfStream
    ' Read next line from file list and build filepaths
    strFileToCopy = objFileList.Readline & "." & strExt

    ' Check for files in SubFolders
    For Each strSubFolder in EnumFolder(strSourceFolder)
      For Each strFileToCopy in oFSO.GetFolder(strSubFolder).Files

    strSourceFilePath = objFSO.BuildPath(strSubFolder, strFileToCopy)
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
        TextOut "Error " & Err.Number & _
        " (" & Err.Description & ")trying to copy " & strFileToCopy
    End If
   Next
Next
Loop

strResults = strResults + 0 '& vbCrLf
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

Function EnumFolder(ByRef vFolder)
Dim oFSO, oFolder, sFldr, oFldr
Set oFSO = CreateObject("Scripting.FileSystemObject")
If Not IsArray(vFolder) Then
If Not oFSO.FolderExists(vFolder) Then Exit Function
sFldr = vFolder
ReDim vFolder(0)
vFolder(0) = oFSO.GetFolder(sFldr).Path
Else sFldr = vFolder(UBound(vFolder))
End If
Set oFolder = oFSO.GetFolder(sFldr)
For Each oFldr in oFolder.Subfolders
ReDim Preserve vFolder(UBound(vFolder) + 1)
vFolder(UBound(vFolder)) = oFldr.Path
EnumFolder vFolder
Next
EnumFolder = vFolder
End Function

option explicit

dim strFileName,strNewDrive,strDrive,objFso,objFile,strText,strReplace

Const ForReading = 1
Const ForWriting = 2

strFileName = Wscript.Arguments(0)
strNewDrive = Wscript.Arguments(1)
strDrive = "PATH"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(strFileName, ForReading)

strText = objFile.ReadAll
objFile.Close
strReplace = Replace(strText, strDrive, strNewDrive)

Set objFile = objFSO.OpenTextFile(strFileName, ForWriting)
objFile.WriteLine strReplace
objFile.Close

wscript.quit()
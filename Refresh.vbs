' VB Script Document
option explicit

const ForReading = 1, ForWriting = 2

dim strCurDIR, objFSO, objShell, whShell, strInstanceQ, strInstance, strInstancePath, strDatabasePath, strInterSystemsPath, strRefreshPath, strFileName, strDrive, strNewDrive
dim strText, strMessage, strReferenceNumber, strResult, strMessageB

set objShell = CreateObject("Wscript.Shell")
set objFSO =   CreateObject("Scripting.FileSystemObject")

strCurDIR = ObjShell.CurrentDirectory

wscript.echo strCurDIR

strReferenceNumber = InputBox("Enter Reference Number (F00xxxx)")' Enter reference number,no checking, can amend it later

strInstanceQ = InputBox("ACCEPT or TRAIN?","Enter Instance")

strInstanceQ = UCase(strInstanceQ)

  If strInstanceQ = "ACCEPT" Then
    strInstance = "ACCEPT"
  ElseIf strInstanceQ = "TRAIN" Then
    strInstance = "TRAIN"
  Else
    strMessageB = MsgBox("Invalid Input",vbOKOnly+vbCritical,"Error")
    WScript.Quit
  End If  

' check paths
 
Sub IntersystemsPath
  strInterSystemsPath = InputBox("Default intersystems path not found, please specify (i.e C:\Intersystems" & vbCrLf & "i.e C:\intersystems")        
  strMessage="You entered -" & strInterSystemsPath & " - is this correct?"
  strResult = MsgBox(strMessage,vbYesNo,"Please Confirm")
    If (strResult = 7) Then' input equals no then
      Call IntersystemsPath        
    Else
    End If
End Sub

Sub InstancePath
  strInstancePath = InputBox("Default Instance installation path not found please specify" & vbCrLf & "i.e C:\intersystems\InstanceName")
  strMessage="You entered - " & strInstancePath & " - is this correct?"
  strResult = MsgBox(strMessage,vbYesNo,"Please Confirm")
    If (strResult = 7) Then
      Call InstancePath
    Else 
    End If
End Sub

Sub DatabasesPath
  strDatabasePath = InputBox("Default Database path not found, please specifiy" & vbCrLf & "i.e D:\JACDatabasesInstance")
  strMessage="You entered - "  &strDatabasePath & " - is this correct?"
  strResult = MsgBox(strMessage,vbYesNo,"Please Confirm")
    If (strResult = 7) Then
      Call DatabasesPath
    Else
    End If
End Sub

  If Not (objFSO.FolderExists("C:\InterSystems")) Then
    call IntersystemsPath 
  Else
    strInterSystemsPath="C:\InterSystems"
  End If

         
  If Not (objFSO.FolderExists("C:\InterSystems\" & strInstance )) Then
    call InstancePath
  Else
    strInstancePath="C:\InterSystems\" & strInstance
  End If

' Path checking done, now find databases folder

      If (objFSO.FolderExists("D:\JACDatabases" & strInstance)) Then
          strDatabasePath="D:\JACDatabases" & strInstance
      ElseIf (objFSO.FolderExists("E:\JACDatabases" & strInstance)) Then
          strDatabasePath="E:\JACDatabases" & strInstance
      Else
          call DatabasesPath
      End If
            
' Create Refresh directory

strRefreshPath = strDatabasePath & "\Refresh" & " - " & Day(Now) & "." & Month(Now) & "." & Year(Now) 

If (objFSO.FolderExists(strRefreshPath)) Then
  strRefreshPath = strRefreshPath & "_" & Second(Now)
End If

objFSO.CreateFolder(strRefreshPath)

' Correct Paths on script files
  
strFileName = strCurDIR & "\Backup.txt"
strDrive =  Left(strCurDIR,2)        ' Get first two characters of current path (c:)
strNewDrive = Left(strInstancePath,2)  ' As above
  call RenameFile

sub RenameFile
  Set objFile = objFSO.OpenTextFile(strFileName, ForReading)  
    strText = objFile.ReadAll
      objFile.Close
      strReplace = Replace(strText, strDrive, strNewDrive)

  Set objFile = objFSO.OpenTextFile(strFileName, ForWriting)
    objFile.WriteLine strReplace
      objFile.Close
End sub
      
objShell.Run(strInstancePath & "\bin\css run " & strInstance & " " & strFileName & " USER") ' run direct into terminal and run script

objShell.Run(strInstancePath & "\bin\css run " & strInstance & "^BACKUP %SYS") ' Run Backup program

objShell.Run(strInstancePath & "\bin\css stop" & strInstance)

objShell.Run(strInstancePath & "\bin\css start" & strInstance)

' Need to test this rename stuff?

objFSO.MoveFile strRefreshPath & "\refresh.cbk", "refresh" & strReferenceNumber & ".cbk"
objFSO.MoveFile strRefreshPath & "\Settings.GOGEN", "Settings" & strReferenceNumber & ".GOGEN"
objFSO.MoveFile strRefreshPath & "\APBackup.ro", "APBackup" & strReferenceNumber & ".ro"
objFSO.MoveFile strRefreshPath & "\GLBackup.ro", "GLBackup" & strReferenceNumber & ".ro"
objFSO.MoveFile strRefreshPath & "\PASADTBackup.ro", "PASADTBackup" & strReferenceNumber & ".ro"
objFSO.MoveFile strRefreshPath & "\Backup.log", "Backup" & strReferenceNumber & ".log"

strDrive = "C:"
strNewDrive = "C:"

call RenameFile 'reset paths

strMessage = strInstance & " has been refreshed"

strMessageB = MsgBox(strMessage,vbOK,"Refresh Complete")

wscript.Quit






        


    

          
            

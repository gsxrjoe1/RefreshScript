' VB Script Document
option explicit

dim strCurDIR, whShell, strInstanceQ, strInstance, strInstancePath, strDatabasePath, strInterSystemsPath, strRefreshPath, strFileName, strDrive, strNewDrive
dim strText, strMessage, strReferenceNumber, strResult

set ObjShell = CreateObject("Wscript.shell")
set ObjFSO =   CreateObject("Wscript.FileSystemObject")

set strCurDIR=ObjShell.CurrentDirectory

strReferenceNumber = InputBox("Enter Reference Number (F00xxxx)")' Enter reference number,no checking, can amend it later

strInstanceQ = InputBox("ACCEPT or TRAIN?","Enter Instance")

UCase(strInstance)

  If strInstanceQ = "ACCEPT" Then
    Set strInstance = "ACCEPT"
  ElseIf strInstanceQ = "TRAIN" Then
    Set strInstance = "TRAIN"
  Else
    MessageBox("Invalid Input")
    WScript.Quit
  End If  

' check paths
 
Sub IntersystemsPath
  strInterSystemsPath = InputBox("Default intersystems path not found, please specify (i.e C:\Intersystems")        
  strMessage="You entered" & strInterSystemsPath & " is this correct?"
  strResult = MsgBox(strMessage,vbYes+vbNo,"Please Confirm")
    If (strResult = 7) Then' input equals no then
      Call IntersystemsPath        
    Else 
      UCase(strInterSystemsPath)
    End If
End Sub

Sub InstancePath
  strInstancePath = InputBox("Default Instance installation path not found please specify")
  strMessage="You entered " & strInstancePath & " is this correct?"
  strResult = MsgBox(strMessage,vbYes+vbNo,"Please Confirm")
    If (strResult = 7) Then
      Call InstancePath
    Else 
      UCase(strInstancePath)
    End If
End Sub

Sub DatabasesPath
  strDatabasePath = InputBox("Default Database path not found, please specifiy")
  strMessage="You entered "  &strDatabasePath & " is this correct?"
  strResult = MsgBox(strMessage,vbYes+vbNo,"Please Confirm")
    If (strResult = 7) Then
      Call DatabasesPath
    Else 
      UCase(strDatabasePath)
    End If
End Sub

  If Not (objFSO.FolderExists("C:\Intersystems")) Then
    call IntersystemsPath 
  Else
    strInterSystemsPath="C:\Intersystems"
    UCase(strInterSystemsPath)
  End If

         
  If Not (objFSO.FolderExists("C:\Intersystems\" & strInstance )) Then
    call InstancePath
  Else
    strInstancePath="C:\Intersystems\" & strInstance
    UCase(strInstancePath)
  End If

' Path checking done, now find databases folder

      If (objFso.FolderExists("D:\JACDatabases" & strInstance)) Then
          strDatabasePath="D:\JACDatabases" & strInstance
          UCase(strDatabasePath)
      ElseIf (objFso.FolderExists("E:\JACDatabases" & strInstance)) Then
          strDatabasePath="E:\JACDatabases" & strInstance
          Ucase(strDatabasePath)
      Else
          call DatabasePath
      End If
            
' Create Refresh directory

strRefreshPath = strDatabasePath & "\Refresh" & "." & Day() & "." & Month() & "." & Year() 

objFso.CreateFolder(strRefreshPath)

' Correct Paths on script files
  
strFileName = strCurDIR & "\Backup.txt"
strDrive =  mid(strCurDIR,0,2)        ' Get first two characters of current path (c:)
strNewDrive = mid(strInstancePath,0,2)  ' As above
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

objFso.MoveFile strRefreshPath & "\refresh.cbk", "refresh" & strReferenceNumber & ".cbk"
objFso.MoveFile strRefreshPath & "\Settings.GOGEN", "Settings" & strReferenceNumber & ".GOGEN"
objFso.MoveFile strRefreshPath & "\APBackup.ro", "APBackup" & strReferenceNumber & ".ro"
objFso.MoveFile strRefreshPath & "\GLBackup.ro", "GLBackup" & strReferenceNumber & ".ro"
objFso.MoveFile strRefreshPath & "\PASADTBackup.ro", "PASADTBackup" & strReferenceNumber & ".ro"
objFso.MoveFile strRefreshPath & "\Backup.log", "Backup" & strReferenceNumber & ".log"

strDrive = "C:"
strNewDrive = "C:"

call RenameFile 'reset paths

strMessage= strInstance & " has been refreshed"

MsgBox(strMessage,options....)

wscript.Quit






        


    

          
            

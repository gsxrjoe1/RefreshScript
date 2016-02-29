' VB Script Document
option explicit

dim strCurDIR, whShell, strInstanceQ, strInstance, strInstancePath, strDatabasePath, strInterSystemsPath, strRefreshPath, strFileName, strDrive, strNewDrive
dim strText, strMessage, strReferenceNumber

set ObjShell = CreateObject("Wscript.shell")
set ObjFSO =   CreateObject("Wscript.FileSystemObject")

set strCurDIR=ObjShell.CurrentDirectory

strReferenceNumber = InputBox("Enter Reference Number (F00xxxx)")' Enter reference number,no checking, can amend it later

strInstanceQ = InputBox("ACCEPT or TRAIN?")

  If strInstanceQ = "ACCEPT" then
  	Set strInstance = "ACCEPT"
  Else if strInstanceQ = "TRAIN" then
    Set strInstance = "TRAIN"
  Else
    MessageBox("Invalid Input")
    WScript.Quit
  End If  

' check paths
 
Sub IntersystemsPath
  strInterSystemsPath = InputBox("Default intersystems path not found, please specify (i.e C:\Intersystems")        
  MsgBox("You entered" & strInterSystemsPath & " is this correct?")
  	If () ' input equals no then
  		Call IntersystemsPath        
  	Else 
  		UCase(strInterSystemsPath)
  	End If
End Sub

Sub InstancePath
  strInstancePath = InputBox("Default Instance installation path not found please specify")
  MsgBox("You entered " & strInstancePath & " is this correct?")
  	If () then
    	Call InstancePath
    Else 
    	UCase(strInstancePath)
    End If
End Sub

Sub DatabasesPath
  strDatabasePath = InputBox("Default Database path not found, please specifiy")
  MsgBox("You entered "  &strDatabasePath & " is this correct?")
  	If () then
  		Call DatabasesPath
  	Else 
    	UCase(strDatabasePath)
  	End If
End Sub

  If Not objFSO.FolderExists("C:\Intersystems") then
    call Intersystems Path 
  Else
  	strInterSystemsPath="C:\Intersystems"
  End If

          
  If Not objFSO.FolderExists("C:\Intersystems\" & strInstance ) then
              call InstancePath
  End If

' Path checking done, now find databases folder

      if objFso.FolderExists("D:\JACDatabases" & strInstance) then
        strDatabasePath="D:\JACDatabases" & strInstance
        ElseIf objFso.FolderExists("E:\JACDatabases" & strInstance) then
          strDatabasePath="E:\JACDatabases" & strInstance
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
      
objShell.Run(strInstancePath & "\bin\cterm.exe /CONSOLE=cn_ap:"& strInstance &"[USER] " & strFileName) ' run direct into terminal and run script

objShell.Run(strInstancePath & "\bin\ceterm.exe CControl.exe CTERMINAL" & strInstance & "^BACKUP %SYS" ) ' Run Backup program

objShell.Run(strInstancePath & "\bin\css stop" & strInstance)

objShell.Run(strInstancePath & "\bin\css start" & strInstance)

' Need to test this rename stuff?

objFso.MoveFile strRefreshPath & "\refresh.cbk" "\refresh" & strReferenceNumber & ".cbk"
objFso.MoveFile strRefreshPath & "\Settings.GOGEN" "\Settings" & strReferenceNumber & ".GOGEN"
objFso.MoveFile strRefreshPath & "\APBackup.ro" "\APBackup" & strReferenceNumber & ".ro"
objFso.MoveFile strRefreshPath & "\GLBackup.ro" "\GLBackup" & strReferenceNumber & ".ro"
objFso.MoveFile strRefreshPath & "\PASADTBackup.ro" "\PASADTBackup" & strReferenceNumber & ".ro"
objFso.MoveFile strRefreshPath & "\Backup.log" "\Backup" & strReferenceNumber & ".log"

strDrive = "C:"
strNewDrive = "C:"

call RenameFile 'reset paths

strMessage= strInstance & " has been refreshed"

MsgBox(strMessage,options....)

wscript.Quit






        


    

          
            

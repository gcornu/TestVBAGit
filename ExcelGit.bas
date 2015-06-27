Attribute VB_Name = "ExcelGit"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       28th March 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
Dim mWsh As IWshRuntimeLibrary.WshNetwork
Dim mWshell As IWshRuntimeLibrary.WshShell

Private Function StatusToString(Received As IWshRuntimeLibrary.WshExecStatus) As String
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       28th March 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
Const strWshRunning = "Running"
Const strWshFinished = "Finished"
Const strWshFailed = "Failed"
    
    Select Case Received
        Case IWshRuntimeLibrary.WshRunning
            StatusToString = strWshRunning
        Case IWshRuntimeLibrary.WshFinished
            StatusToString = strWshFinished
        Case IWshRuntimeLibrary.WshFailed
            StatusToString = strWshFailed
    End Select

End Function

Public Function WriteToGit()
    '============================================================================================================================
    '
    '
    '   Author      :       John Greenan
    '   Email       :
    '   Company     :       Alignment Systems Limited
    '   Date        :       28th March 2014
    '
    '   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
    '
    '   References  :       See VB Module FL for list extracted from VBE
    '   References  :
    '============================================================================================================================
    'Const
    Const strSourceDirectory As String = "D:\Users\Gauthier\Downloads\TestVBAGit"
    Const strCMD As String = "cmd /K"
    Const strChangeDirectoryTo As String = "cd"
    Const strGitAdd As String = "git add ."
    Const strGitCommit As String = "git commit -am"
    Const strGitPush As String = "git push"
    Const strGitStatus As String = "git status"
    Const strProcessID As String = "PID="
    Const strTitle As String = "title Alignment-Systems.com Git Integration"
    'Variables
    Dim strExecStatus As String
    Dim dtNow As Date
    Dim strTextFromStdStream As String
    Dim strBuiltCommand As String
    Dim strUserName As String
    Dim commitMessage As String
    
    Set mWshell = New IWshRuntimeLibrary.WshShell
    Set mWsh = New IWshRuntimeLibrary.WshNetwork
    
    dtNow = Now()
    
    commitMessage = InputBox("Commit message:")
    
    Call ExportVBAFiles
    
    '   Change to the correct folder with cmd:>cd folder
    strBuiltCommand = strCMD
    
    With mWshell.Exec(strBuiltCommand)
    
    '----------------------------------
        
        strBuiltCommand = strChangeDirectoryTo
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand
        .StdIn.WriteLine strBuiltCommand
        strExecStatus = StatusToString(.Status)
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand & "=" & strExecStatus
           
        
    '----------------------------------
        
        strBuiltCommand = strTitle
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand
        .StdIn.WriteLine strBuiltCommand
        strExecStatus = StatusToString(.Status)
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand & "=" & strExecStatus
        
    '----------------------------------
        
        strBuiltCommand = strChangeDirectoryTo & Chr(VBA.KeyCodeConstants.vbKeySpace) & strSourceDirectory
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand
        .StdIn.WriteLine strBuiltCommand
        strExecStatus = StatusToString(.Status)
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand & "=" & strExecStatus
        
    '----------------------------------
                
        strBuiltCommand = strGitAdd
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand
        .StdIn.WriteLine strBuiltCommand
        strExecStatus = StatusToString(.Status)
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand & "=" & strExecStatus
        
    '----------------------------------
                
        strBuiltCommand = strGitCommit & Chr(VBA.KeyCodeConstants.vbKeySpace) & """" & dtNow & ":" & Chr(VBA.KeyCodeConstants.vbKeySpace) & commitMessage & """"
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand
        .StdIn.WriteLine strBuiltCommand
        strExecStatus = StatusToString(.Status)
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand & "=" & strExecStatus
        
    '----------------------------------
    
        strBuiltCommand = strGitPush
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand
        .StdIn.WriteLine strBuiltCommand
        strExecStatus = StatusToString(.Status)
        Debug.Print "[" & strProcessID & .ProcessID & "]>" & strBuiltCommand & "=" & strExecStatus
        
    '----------------------------------
    'Cleanup
    '----------------------------------
        .StdIn.Close
    
        If Not .StdOut.AtEndOfStream Then
            Debug.Print "Dumping out Process StdOut"
        End If
        
        Do While Not .StdOut.AtEndOfStream
            strTextFromStdStream = "[" & strProcessID & .ProcessID & "]" & .StdOut.ReadLine()
            Debug.Print strTextFromStdStream
        Loop
        
        If Not .StdErr.AtEndOfStream Then
            Debug.Print "Dumping out Process StdErr"
        End If
        
        
        Do While Not .StdErr.AtEndOfStream
            strTextFromStdStream = "[" & strProcessID & .ProcessID & "]" & .StdErr.ReadLine()
            Debug.Print strTextFromStdStream
        Loop
            
        .StdErr.Close
        .StdOut.Close
        .Terminate
    End With

End Function

Private Sub ExportVBAFiles()

    'Dim pVBAProject As VBIDE.VBProject
    'Dim project As VBIDE.VBProject
    Dim vbComp As Variant  'VBA module, form, etc...
    Dim strSavePath As String
    'Dim vbRef As VBIDE.Reference
   
      
     ' Get the VBA project
     ' If you want to export code for Normal instead, paste this macro into
     ' ThisDocument in the Normal VBA project and change the following line to:
     ' Set pVBAProject = ThisDocument.VBProject

     'Set pVBAProject = Application.VBE.ActiveVBProject

     strSavePath = "D:\Users\Gauthier\Downloads\TestVBAGit"
          
     'For Each vbRef In pVBAProject.References
       'Debug.Print vbRef.description, vbRef.FullPath
     'Next
    
     ' Loop through all the components (modules, forms, etc) in the VBA project
    

    
     For Each vbComp In ThisWorkbook.VBProject.VBComponents
  
       Select Case vbComp.Type
  
           Case vbext_ct_StdModule
               vbComp.Export strSavePath & "\" & vbComp.Name & ".bas"
  
           Case vbext_ct_Document, vbext_ct_ClassModule
               ' ThisDocument and class modules
               Call vbComp.Export(strSavePath & "\" & vbComp.Name & ".cls")
  
           Case vbext_ct_MSForm
               vbComp.Export strSavePath & "\" & vbComp.Name & ".frm"
  
           Case Else
               vbComp.Export strSavePath & "\" & vbComp.Name
  
       End Select

     Next
  
     'MsgBox "VBA files have been exported to: " & strSavePath

  

End Sub

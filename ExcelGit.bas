Attribute VB_Name = "ExcelGit"
Option Explicit
Option Compare Text
Option Base 0

Const strSourceDirectory As String = "D:\Users\Gauthier\Downloads\TestVBAGit"

Dim mWsh As IWshRuntimeLibrary.WshNetwork
Dim mWshell As IWshRuntimeLibrary.WshShell

' Inspired from https://github.com/JohnGreenan/5_ExcelVBE/blob/master/ExcelGit.bas
Public Sub WriteToGit()
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
    
        ' Change directory
        strBuiltCommand = strChangeDirectoryTo & Chr(VBA.KeyCodeConstants.vbKeySpace) & strSourceDirectory
        .StdIn.WriteLine strBuiltCommand
        
        ' Track files (git add .)
        strBuiltCommand = strGitAdd
        .StdIn.WriteLine strBuiltCommand
        
        ' Commit files (git commit -am)
        strBuiltCommand = strGitCommit & Chr(VBA.KeyCodeConstants.vbKeySpace) & """" & dtNow & ":" & Chr(VBA.KeyCodeConstants.vbKeySpace) & commitMessage & """"
        .StdIn.WriteLine strBuiltCommand
        
        ' Push commit (git push)
        strBuiltCommand = strGitPush
        .StdIn.WriteLine strBuiltCommand

        'Cleanup
        .StdIn.Close
        
        Do While Not .StdOut.AtEndOfStream
            strTextFromStdStream = "[" & strProcessID & .ProcessID & "]" & .StdOut.ReadLine()
            Debug.Print strTextFromStdStream
        Loop

        Do While Not .StdErr.AtEndOfStream
            strTextFromStdStream = "[" & strProcessID & .ProcessID & "]" & .StdErr.ReadLine()
            Debug.Print strTextFromStdStream
        Loop
            
        .StdErr.Close
        .StdOut.Close
        .Terminate
    End With

End Sub

' Inspired from http://visguy.com/vgforum/index.php?topic=3815.0
Private Sub ExportVBAFiles()
    Dim vbComp As Variant
    Dim strSavePath As String

    strSavePath = "D:\Users\Gauthier\Downloads\TestVBAGit"

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
End Sub

' Inspired from https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/
Private Sub RemoveAllModules()
    Dim project As VBProject
    Dim moduleName As String
    
    Set project = Application.VBE.ActiveVBProject
    moduleName = Application.VBE.ActiveCodePane.CodeModule.Name
     
    Dim comp As VBComponent
    For Each comp In project.VBComponents
        If Not comp.Name = moduleName And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule Or comp.Type = vbext_ct_MSForm) Then
            project.VBComponents.Remove comp
        End If
    Next
End Sub

' Inspired from https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/
Private Sub ImportVBAFiles(sourcePath As String)
    Dim fso As FileSystemObject
    Dim folder As folder
    Dim file As file
    Dim fileName As String, extension As String
    Dim moduleName As String
    Dim sheet As Worksheet
    Dim isSheet As Boolean
    Dim sheets As Object
    
    moduleName = Application.VBE.ActiveCodePane.CodeModule.Name
    Set sheets = ThisWorkbook.Worksheets
    
    'Create an instance of the FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set folder = fso.GetFolder(sourcePath)
    'loops through each file in the directory and prints their names and path
    For Each file In folder.Files
        extension = Right(file.Name, Len(file.Name) - InStrRev(file.Name, "."))
        fileName = Left(file.Name, Len(file.Name) - Len(extension) - 1)
        isSheet = False
        
        For Each sheet In sheets
            If sheet.Name = fileName Then
                isSheet = True
                Exit For
            End If
        Next sheet
        
        If file.Name <> moduleName & ".bas" And file.Name <> "ThisWorkbook.cls" And Not isSheet And (extension = "bas" Or extension = "cls" Or extension = "frm") Then
            'Common case
            Application.VBE.ActiveVBProject.VBComponents.Import file.Path
        ElseIf file.Name = "ThisWorkbook.cls" Or isSheet Then
            'Special case for ThisWorbook and worksheets
            With Application.VBE.ActiveVBProject.VBComponents(fileName).CodeModule
                .DeleteLines 1, .CountOfLines
                .AddFromFile file.Path
                .DeleteLines 1, 5
            End With
        End If
    Next file
End Sub

Public Sub ReadFromGit()
    Call RemoveAllModules
    Call ImportVBAFiles(strSourceDirectory)
End Sub

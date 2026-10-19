Attribute VB_Name = "zz_gitFunctions"
Option Explicit
'@IgnoreModule

''*********************************************************************************
'' ***   This module deals with the GitHub functionality.                       ***
'' ***   The GitHub functionality gives easy import/export of all code          ***
'' ***   to the local GitHub folder stored on the user's machine                ***
'' ***   GitHub Link: https://github.com/ChristopherSBond/Pseudo_Pandas         ***
''*********************************************************************************
Private Const MODULE_NAME As String = "zz_gitFunctions"

' Local github folder
Private Const GITHUB_REPOSITORY_FOLDER As String = "C:\Users\Public\GitHub\Pseudo_Pandas"

' Workbook name
Private Const GITHUB_LOCAL_FILE_NAME As String = "Pseudo Pandas"

' Folder that stores the workbook
Private Const GITHUB_WB_FOLDER_NAME As String = "Pandas Program Files"

''***** Import Function *************************************************************************************************
' importVBComponents
'   Purpose:
'       Imports vb files stored in the local GitHub folder
'       Deletes all existing vb components
'---------------------------------------------------------------------------------
Public Sub gitImportVBComponents()
'Commented out so I don't accidentally run and delete all unsaved work in wb
'    ' Clear the current workbook of all modules (excluding this git module)
'    deleteAllVbComponents
'
'    ' Import all components to the current workbook
'    Dim gitFolderPath As String
'    gitFolderPath = GITHUB_REPOSITORY_FOLDER
'
'    importComponentsFromGitFolder gitFolderPath
'
'    ' Clean up worksheets that were imported into cls modules
'    delWsInClassModules
    
End Sub

''***** Export Function *************************************************************************************************
' exportVBComponents
'   Purpose:
'       Export all VB components to the local GitHub folder
'---------------------------------------------------------------------------------
Public Sub gitExportVBComponents()
    ' Delete all the code stored in the GitHub folder.
    ' Ensures that deleted modules will also get deleted when uploading to Git.
    deleteAllCodeInLocalFolder
    
    ' Export all the vb components (including WB) to the GitHub folder
    exportComponentsToFolder
    
End Sub

''***** Helper Functions **************************************************************************************************
'**************************************************************************************************************************
' getFolderPaths
'   Purpose:
'       Get the folder paths for each vb component type (in-place)
'---------------------------------------------------------------------------------
Private Sub getFolderPaths(ByRef modFolderPath As String, ByRef frmFolderPath As String, _
                            ByRef clsModFolderPath As String, ByRef wsFolderPath As String, ByRef wbFolderPath As String)
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    
    Dim folderPath As String: folderPath = GITHUB_REPOSITORY_FOLDER
    
    modFolderPath = fso.BuildPath(folderPath, "Modules")
    frmFolderPath = fso.BuildPath(folderPath, "Forms")
    clsModFolderPath = fso.BuildPath(folderPath, "Class_Modules")
    wsFolderPath = fso.BuildPath(folderPath, "Worksheets")
    wbFolderPath = fso.BuildPath(GITHUB_REPOSITORY_FOLDER, GITHUB_WB_FOLDER_NAME)
    
    If Not fso.FolderExists(modFolderPath) Then fso.CreateFolder modFolderPath
    If Not fso.FolderExists(frmFolderPath) Then fso.CreateFolder frmFolderPath
    If Not fso.FolderExists(clsModFolderPath) Then fso.CreateFolder clsModFolderPath
    If Not fso.FolderExists(wsFolderPath) Then fso.CreateFolder wsFolderPath
    If Not fso.FolderExists(wbFolderPath) Then fso.CreateFolder wbFolderPath
        
End Sub

'**************************************************************************************************************************
' exportComponentsToFolder
'   Purpose:
'       Export all vb components to their respective folders in the local GitHub Folder
'       Components include all worksheets, forms, modules, cls modules, and the workbook itself
'---------------------------------------------------------------------------------
Private Sub exportComponentsToFolder()
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    
    Dim modFolderPath As String, frmFolderPath As String, clsModFolderPath As String, wsFolderPath As String, wbFolderPath As String
    getFolderPaths modFolderPath, frmFolderPath, clsModFolderPath, wsFolderPath, wbFolderPath

    Dim vbc As VBComponent
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        Dim fileName As String
        
        deleteNewLinesTopSub vbc    ' fixes a bug when exporting files
        
        Select Case vbc.Type
            Case vbext_ct_StdModule ' Standard Module
                fileName = fso.BuildPath(modFolderPath, vbc.Name) & ".bas"
                
            Case vbext_ct_ClassModule ' Class Module
                fileName = fso.BuildPath(clsModFolderPath, vbc.Name) & ".cls"
                
            Case vbext_ct_MSForm ' UserForm
                fileName = fso.BuildPath(frmFolderPath, vbc.Name) & ".frm"
                    
            Case vbext_ct_Document  ' Worksheet
                fileName = fso.BuildPath(wsFolderPath, vbc.Name) & ".cls"
                
            Case Else   ' Should be covered by above, but skip jic
                GoTo nextIter
        End Select
        
        vbc.Export fileName
    
nextIter:
    Next vbc
    
    Dim gitFolderPath As String
    
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs fso.BuildPath(wbFolderPath, GITHUB_LOCAL_FILE_NAME & ".xlsm")
    Application.DisplayAlerts = True
    
End Sub

'**************************************************************************************************************************
' importComponentsFromGitFolder
'   Purpose:
'       Imports each file (excluding current module) stored in the folder at a given folderPath, including those in subfolders
'   Parameter:
'       folderPath  : path to the folder storing the files to be imported
'---------------------------------------------------------------------------------
Private Sub importComponentsFromGitFolder(ByVal folderPath As String)
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    
    Dim subfolder As FOLDER
    For Each subfolder In fso.GetFolder(folderPath).SubFolders
        If subfolder.Name <> "Class_Modules" And subfolder.Name <> "Forms" And subfolder.Name <> "Modules" And _
           subfolder.Name <> "Worksheets" Then GoTo nextIter
        
        importComponentsFromGitFolder subfolder.Path
nextIter:
    Next subfolder
    
    ' When all subfolders have been looped thru, and folderPath = the original folder path (subfolder.count>0), exit sub
    If fso.GetFolder(folderPath).SubFolders.count > 0 Then Exit Sub
    
    Dim file As file
    For Each file In fso.GetFolder(folderPath).Files
        If (Not file.Name Like "*gitFunctions*") And (fso.GetExtensionName(file.Name) <> "frx") Then
            ThisWorkbook.VBProject.VBComponents.Import file.Path
        End If
    Next file

End Sub

'**************************************************************************************************************************
' deleteAllCodeInLocalFolder
'   Purpose:
'       Deletes all the code stored in the local GitHub folder.
'       Need to do this s.t. if a module is deleted, the module will be deleted when committing to git
'---------------------------------------------------------------------------------
Private Sub deleteAllCodeInLocalFolder()
    Dim gitPath As String: gitPath = GITHUB_REPOSITORY_FOLDER
    
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    
    Dim subfolder As FOLDER
    For Each subfolder In fso.GetFolder(gitPath).SubFolders
        If subfolder.Name <> ".git" And subfolder.Name <> GITHUB_WB_FOLDER_NAME Then
            Call fso.DeleteFolder(folderspec:=subfolder, force:=True) '---Force= True ignores folders (or files within folders) that have Read-Only attribute turned on.
        End If
    Next subfolder
    
End Sub

'**************************************************************************************************************************
' deleteAllVbComponents
'   Purpose:
'       Deletes all vb components in the current excel workbook (excluding ws Objects and git module)
'       These include modules, userforms, and class modules
'---------------------------------------------------------------------------------
Private Sub deleteAllVbComponents()

    Dim vbComponentTypes As Scripting.Dictionary: Set vbComponentTypes = New Scripting.Dictionary
    vbComponentTypes(vbext_ct_ClassModule) = vbNull
    vbComponentTypes(vbext_ct_StdModule) = vbNull
    vbComponentTypes(vbext_ct_MSForm) = vbNull
        
    Dim vbc As VBComponent
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        If (vbc.Name <> MODULE_NAME) And vbComponentTypes.Exists(vbc.Type) Then
            ThisWorkbook.VBProject.VBComponents.Remove vbc
        End If
    
    Next vbc
End Sub

'**************************************************************************************************************************
' deleteNewLinesTopSub
'   Purpose:
'       Delete the blank lines at the top of subs
'       Created to negate a bug that will append a newline to a userform when exporting a file
'       I hate excel.
'   Parameters:
'       vbc : VBComponent -- code sheet on a form, module, or class module
'---------------------------------------------------------------------------------
Private Sub deleteNewLinesTopSub(ByRef vbc As VBComponent)
    ' Only run if there is code in the code sheet
    If vbc.CodeModule.CountOfLines = 0 Then Exit Sub
        
    ' Define all blank lines to cover each case
    ' Use null dict to access the .exists method for keys
    Dim blankLineCharacters As Scripting.Dictionary: Set blankLineCharacters = New Scripting.Dictionary
    blankLineCharacters(vbNewLine) = vbNull
    blankLineCharacters(vbCrLf) = vbNull
    blankLineCharacters("") = vbNull
    
    Do While blankLineCharacters.Exists(vbc.CodeModule.lines(1, 1))
        vbc.CodeModule.DeleteLines 1
    Loop

End Sub

'**************************************************************************************************************************
' delWsInClassModules
'   Purpose:
'       Delete all worksheets that get imported into class modules
'       Fixes an issue where the worksheets would import in their correct location AND in the class modules
'       This is because worksheet objects share the same .cls extension as class modules
'       I hate excel.
'---------------------------------------------------------------------------------
Private Sub delWsInClassModules()
    Dim vbc As VBComponent
    For Each vbc In ThisWorkbook.VBProject.VBComponents
    
        If vbc.Type = vbext_ct_ClassModule And (left$(vbc.Name, 2) = "sh" Or left$(vbc.Name, 12) = "ThisWorkbook") Then
            ThisWorkbook.VBProject.VBComponents.Remove vbc
        End If
    Next vbc

End Sub
'**************************************************************************************************************************
'**************************************************************************************************************************



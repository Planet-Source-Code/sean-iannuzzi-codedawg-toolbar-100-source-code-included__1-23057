VERSION 5.00
Begin VB.Form FrmSaveChanges 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Changes"
   ClientHeight    =   2160
   ClientLeft      =   2508
   ClientTop       =   2892
   ClientWidth     =   5532
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5532
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdProcessSave 
      Caption         =   "&Options"
      Height          =   348
      Index           =   2
      Left            =   4512
      TabIndex        =   4
      Top             =   1728
      Width           =   924
   End
   Begin VB.CommandButton CmdProcessSave 
      Caption         =   "&Close"
      Height          =   348
      Index           =   1
      Left            =   4512
      TabIndex        =   2
      Top             =   744
      Width           =   924
   End
   Begin VB.CommandButton CmdProcessSave 
      Caption         =   "&Save"
      Height          =   348
      Index           =   0
      Left            =   4512
      TabIndex        =   1
      Top             =   336
      Width           =   924
   End
   Begin VB.ListBox lstFileToSave 
      Height          =   1584
      Left            =   96
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   312
      Width           =   4260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please select the files you wish to save."
      Height          =   192
      Left            =   96
      TabIndex        =   3
      Top             =   72
      Width           =   2808
   End
End
Attribute VB_Name = "FrmSaveChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VBInstance As VBIDE.VBE
Public Connect As Connect
Private mBln_FormLoaded As Boolean
Private mColProjectInfoFromListBox As Collection

Private Sub CmdProcessSave_Click(Index As Integer)
    
    Select Case (Index)
        Case 0
            Call gfIntQuickSaveVBProject(True)
            Unload Me
        Case 1
            Unload Me
        Case 2
            FrmOptions.Show vbModal
    End Select
    
End Sub

Private Sub Form_Activate()
        
    'needed to make sure we have the VBInstance var set
    If mBln_FormLoaded = True Then
        Call mSubLoadListBoxWithProjectInfo
        mBln_FormLoaded = False
    End If
    
End Sub

Private Sub Form_Load()
    mBln_FormLoaded = True
End Sub
Private Sub mSubLoadListBoxWithProjectInfo()
    
    Dim lIntCurrentProject      As Integer
    Dim lIntCurrentComponent    As Integer
    Dim lVbProject              As VBProject
    Dim lVbComponent            As VBComponent
    Dim lBln_AddProject         As Boolean '990826 SBI0
    Dim lBln_ItemNotSaved       As Boolean
        
    'SBI0 990816 Cannot use For each permission will be denied
    lstFileToSave.Visible = False
    lstFileToSave.Clear
    Set mColProjectInfoFromListBox = New Collection
    For lIntCurrentProject = 1 To VBInstance.VBProjects.Count
        Set lVbProject = VBInstance.VBProjects(lIntCurrentProject)
        If lVbProject.IsDirty = True Then
            If lVbProject.FileName = "" Then
                lBln_ItemNotSaved = True
            Else
                lBln_AddProject = False 'used to add project if has dirty component
                For lIntCurrentComponent = 1 To lVbProject.VBComponents.Count
                    Set lVbComponent = lVbProject.VBComponents(lIntCurrentComponent)
                    If lVbComponent.IsDirty = True Then
                        
                        If lVbComponent.FileNames(1) = "" Then
                            lBln_ItemNotSaved = True
                        Else
                            'only add the project is one or more compoenets are dirty
                            'we do not save the project only it's members
                            If lBln_AddProject = False Then
                                mColProjectInfoFromListBox.Add lIntCurrentProject, Mid$(lVbProject.FileName, InStrRev(lVbProject.FileName, "\") + 1)
                                lstFileToSave.AddItem Mid$(lVbProject.FileName, InStrRev(lVbProject.FileName, "\") + 1)
                                lBln_AddProject = True
                            End If
                            lstFileToSave.AddItem "    " & Mid$(lVbComponent.FileNames(1), InStrRev(lVbComponent.FileNames(1), "\") + 1)
                            mColProjectInfoFromListBox.Add lIntCurrentProject, Mid$(lVbComponent.FileNames(1), InStrRev(lVbComponent.FileNames(1), "\") + 1)
                        End If
                    End If
                    DoEvents
                Next
            End If
        End If
        DoEvents
    Next
    lstFileToSave.Visible = True
    
    If lstFileToSave.ListCount = 0 Then
        If lBln_ItemNotSaved = False Then
            MsgBox "There are no files to save!", vbInformation, "No Files Changed!"
        Else
            MsgBox "One or more files have not been saved.  Please save all files before attempting to perform a quick save!", vbInformation, "Files Not Saved!"
        End If
        Unload Me
    End If
    
End Sub

Public Function gfIntQuickSaveVBProject(Optional pvBln_SaveSelected As Boolean = False) As Integer

On Error GoTo gfIntQuickSaveVBProject_ErrorHandler:

    Dim lIntCurrentProject              As Integer
    Dim lIntRetVal                      As Integer
    Dim lBln_SaveActiveCodePaineOnly    As Boolean
    Dim lBln_SaveAllFileNoPrompt        As Boolean
    Dim lBln_SaveReadOnlyFiles          As Boolean
    Dim lBln_SaveReadOnlyPrompt         As Boolean
    Dim lBlnFileIsReadOnly              As Boolean
    Dim lLngSaveCount                   As Long
    Dim lIntProjectToSave               As Integer
    
    lLngSaveCount = 0
    
    'retrive options
    lBln_SaveActiveCodePaineOnly = GetSetting(App.Title, "Settings", "SaveOnlyActiveCodePaine", False)
    lBln_SaveAllFileNoPrompt = GetSetting(App.Title, "Settings", "SaveAllFilesWithoutPrompt", False)
    lBln_SaveReadOnlyFiles = GetSetting(App.Title, "Settings", "SaveReadOnlyFiles", False)
    lBln_SaveReadOnlyPrompt = GetSetting(App.Title, "Settings", "SaveReadOnlyFilesPrompt", False)
        
    lIntRetVal = -1
                    
    If pvBln_SaveSelected = False Then
        'Get the active Code Paine and saves it
        For lIntCurrentProject = 1 To VBInstance.VBProjects.Count
            Call mSubQuickSaveProject(lBln_SaveActiveCodePaineOnly, lBln_SaveAllFileNoPrompt, lBln_SaveReadOnlyFiles, lBln_SaveReadOnlyPrompt, lIntCurrentProject, lLngSaveCount)
            'needed to exit after saving active code window only
            If lBln_SaveActiveCodePaineOnly = True Then
                Exit For
            End If
            DoEvents
        Next
    Else
        For lIntCurrentProject = 1 To Me.lstFileToSave.ListCount - 1
            If lstFileToSave.Selected(lIntCurrentProject) = True Then
                'retrive project info for list item
                lIntProjectToSave = mColProjectInfoFromListBox.Item(Trim$(lstFileToSave.List(lIntCurrentProject)))
                Call mSubQuickSaveProject(lBln_SaveActiveCodePaineOnly, lBln_SaveAllFileNoPrompt, lBln_SaveReadOnlyFiles, lBln_SaveReadOnlyPrompt, lIntProjectToSave, lLngSaveCount)
            End If
            'needed to exit after saving active code window only
            If lBln_SaveActiveCodePaineOnly = True Then
                Exit For
            End If
            DoEvents
        Next
    End If

    
    If lBln_SaveActiveCodePaineOnly = False Then
        If lLngSaveCount = 1 Then
            MsgBox CStr(lLngSaveCount) & " project file has been saved!", vbInformation, "File Saved!"
        ElseIf lLngSaveCount > 1 Then
            MsgBox CStr(lLngSaveCount) & " project files were saved!", vbInformation, "Files Saved!"
        Else
            MsgBox "No project files were saved!", vbInformation, "No Project(s) Saved"
        End If
    Else
        If lLngSaveCount = 1 Then
            MsgBox "The Active Code Window has been saved!", vbInformation, "Code Window Saved"
        Else
            MsgBox "The Active Code Window was not saved!", vbInformation, "Code Window Not Saved"
        End If
    End If
    
    
    Exit Function
    
gfIntQuickSaveVBProject_ErrorHandler:
    If Err Then
        MsgBox Error
        Resume Next
    End If
    
End Function

Private Sub mSubQuickSaveProject(ByVal lBln_SaveActiveCodePaineOnly As Boolean, _
    ByVal lBln_SaveAllFileNoPrompt As Boolean, _
    ByVal lBln_SaveReadOnlyFiles As Boolean, _
    lBln_SaveReadOnlyPrompt, ByVal lIntCurrentProject As Integer, _
    ByRef lLngSaveCount As Long)
    
On Error GoTo mSubQuickSaveProject_ErrorHandler:

    Dim lVbProject                      As VBProject
    Dim lVbComponent                    As VBComponent
    Dim lIntFreeFile                    As Integer
    Dim lStrQuickSaveFileName           As String
    Dim lIntCurrentComponent            As Integer
    Dim lStrOldname                     As String
    Dim lStrPropNameOfComponent         As String
    
    'get vb project Reference
    Set lVbProject = VBInstance.VBProjects(lIntCurrentProject)
    
    'extra check to make sure there is a project object
    If lVbProject.IsDirty = True Then
        
        'Get active Project
        If Not VBInstance.ActiveVBProject Is Nothing Or lBln_SaveActiveCodePaineOnly = False Then
        
            'check to see if we have the correct VB project
            'ONLY USED FOR ACTIVE PROJECT SAVE
            If lVbProject = VBInstance.ActiveVBProject Or lBln_SaveActiveCodePaineOnly = False Then
                
                'check to see if file is read-only if so prompt user if flag is set
                'Not needed we cannot save VBP files
                'If lBln_SaveReadOnlyPrompt = True And mfBln_IsFileReadOnly(lVbProject.FileName) = True Then
                    'prompt user where to save
                    'need dialog to save read-only file
                    'Can't save At Project Level
                    'MsgBox "Cannot Save At Project Level !", vbInformation, "Cannot Save!"
                'End If
                    
                'check to see if project file is read only if so check flag to see if we should overwrite
                If lBln_SaveReadOnlyFiles = True Or mfBln_IsFileReadOnly(lVbProject.FileName) = False Then
                    
                    'loop through project components
                    For lIntCurrentComponent = 1 To lVbProject.VBComponents.Count
                    
                        'get the compoennt reference
                        Set lVbComponent = lVbProject.VBComponents(lIntCurrentComponent)
                        
                        'check if component is dirty
                        If lVbComponent.IsDirty = True Then
                        
                            'extra check to make sure active code paine is avaiable
                            If Not VBInstance.ActiveCodePane Is Nothing Or lBln_SaveActiveCodePaineOnly = False Then
                                
                                'check to see if we have the correct project component
                                'ONLY USED FOR ACTIVE CODE WINDOW
                                If lVbComponent.CodeModule = VBInstance.ActiveCodePane.CodeModule Or lBln_SaveActiveCodePaineOnly = False Then
                                    
                                    'check to see if component file is read-only, if so check flag to see if should show
                                    'prompt
                                    If lBln_SaveReadOnlyPrompt = True And mfBln_IsFileReadOnly(lVbComponent.FileNames(1)) = True Then
                                    
                                        lStrOldname = lVbComponent.FileNames(1)
                                        
                                        'need dialog to save read-only file
                                        lStrQuickSaveFileName = mfStr_PromptToSaveReadOnly(lVbProject.Name, lVbComponent.FileNames(1), lVbComponent.Name)
                                        
                                        'delete any old file
                                        Kill lStrQuickSaveFileName

                                        'get file handle
                                        lIntFreeFile = FreeFile
                                        
                                        If Trim$(lStrQuickSaveFileName) <> "" Then
                                            'put all source code to file
                                            Open lStrQuickSaveFileName For Binary Access Write Lock Read Write As #lIntFreeFile
                                            Put #lIntFreeFile, , lVbComponent.CodeModule.Lines(1, lVbComponent.CodeModule.CountOfLines)
                                            Close #lIntFreeFile
                                            
                                            'rename the VB filename back to the old one
                                            'to make sure that the project does not show that
                                            'we saved the file
                                            lVbComponent.SaveAs lVbComponent.FileNames(1)
                                            
                                            'add to the save file counter
                                            lLngSaveCount = lLngSaveCount + 1
                                        End If
                                    'check to see if we should save read-only components, or show message!
                                    ElseIf lBln_SaveReadOnlyFiles = True Or mfBln_IsFileReadOnly(lVbComponent.FileNames(1)) = False Then
                                        
                                        'create save file name
                                        lStrPropNameOfComponent = Mid$(lVbComponent.FileNames(1), InStrRev(lVbComponent.FileNames(1), "\") + 1, Len(lVbComponent.FileNames(1)) - InStrRev(lVbComponent.FileNames(1), "\") - 4)
                                        
                                        'save old file name NEEDED TO keep the same file name as th one in the
                                        'project
                                        lStrOldname = lVbComponent.FileNames(1)
                                        lStrQuickSaveFileName = Mid$(lVbProject.FileName, 1, InStrRev(lVbProject.FileName, "\")) & "QSF\"
                                        
                                        'create QSF directory
                                        MkDir lStrQuickSaveFileName
                                        
                                        'finished file creation name
                                        lStrQuickSaveFileName = lStrQuickSaveFileName & lVbProject.Name & " - " & lVbComponent.Name & " - " & lStrPropNameOfComponent & ".qsf"
                                        
                                        'delete any old file
                                        Kill lStrQuickSaveFileName
                                        
                                        
                                        'get file handle
                                        lIntFreeFile = FreeFile
                                        
                                        'put all source code to file
                                        Open lStrQuickSaveFileName For Binary Access Write Lock Read Write As #lIntFreeFile
                                        Put #lIntFreeFile, , lVbComponent.CodeModule.Lines(1, lVbComponent.CodeModule.CountOfLines)
                                        Close #lIntFreeFile
                                        
                                        'rename the VB filename back to the old one
                                        'to make sure that the project does not show that
                                        'we saved the file
                                        lVbComponent.SaveAs lStrOldname
                                        
                                        'add to the save file counter
                                        lLngSaveCount = lLngSaveCount + 1
                                        
                                        'needed to exit after saving active code window only
                                        If lBln_SaveActiveCodePaineOnly = True Then
                                            Exit For
                                        End If
                                    ElseIf mfBln_IsFileReadOnly(lVbComponent.FileNames(1)) = True And lBln_SaveReadOnlyFiles = False And lBln_SaveReadOnlyPrompt = False Then
                                        MsgBox "Cannot Save Read-Only File: " & vbCrLf & "'" & lVbComponent.FileNames(1) & "'." & vbCrLf & vbCr & "If you need to save read-only files, select the 'Save Read-Only Files' Option.", vbInformation, "Cannot read-Only File"
                                    End If
                                End If
                            End If
                        End If
                        DoEvents
                    Next
                ElseIf mfBln_IsFileReadOnly(lVbProject.FileName) = True And lBln_SaveReadOnlyFiles = False And lBln_SaveReadOnlyPrompt = False Then
                    MsgBox "Cannot Save Read-Only File: " & vbCrLf & "'" & lVbProject.FileName & "'." & vbCrLf & vbCr & "If you need to save read-only files, select the 'Save Read-Only Files' Option.", vbInformation, "Cannot Read-Only File"
                End If
            End If
        End If
    End If
    
    'should be nothing, but kill anyway to be save
    Set lVbProject = Nothing
    Set lVbComponent = Nothing
    
    Exit Sub
    
mSubQuickSaveProject_ErrorHandler:
    
    'errors to ignore
    'File not found, path error, or object ver not set
    If Err = 53 Or Err = 75 Or Err = 91 Then
        Resume Next
    End If
    
    If Err Then
        MsgBox "Err: " & Error
        Resume Next
    End If

End Sub

Private Function mfBln_IsFileReadOnly(ByVal pvStrFileName As String) As Boolean
    Select Case GetAttr(pvStrFileName)
        Case vbReadOnly, vbReadOnly + vbArchive
            'prompt user for new file location and/or name
            mfBln_IsFileReadOnly = True
    End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
            
    Set mColProjectInfoFromListBox = Nothing
    Set FrmSaveChanges = Nothing

End Sub
Private Function mfStr_PromptToSaveReadOnly(ByVal pvStrProjectName As String, Optional ByVal pvStrComponentName As String = "", Optional ByVal pvStrPropNameOfComponent As String) As String

    On Error Resume Next
    FrmCodeDawgToolbar.dlgSave.FileName = Mid$(pvStrComponentName, 1, InStrRev(pvStrComponentName, "\")) & "QSF\" & pvStrProjectName & " - " & pvStrPropNameOfComponent & " - " & Mid$(pvStrComponentName, InStrRev(pvStrComponentName, "\") + 1, Len(pvStrComponentName) - InStrRev(pvStrComponentName, "\") - 4) & ".qsf"
    FrmCodeDawgToolbar.dlgSave.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
    FrmCodeDawgToolbar.dlgSave.ShowSave
    If Err = 0 Then
        'return String of file to save
        mfStr_PromptToSaveReadOnly = FrmCodeDawgToolbar.dlgSave.FileName
    End If
    On Error GoTo 0

End Function


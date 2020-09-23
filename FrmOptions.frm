VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assistant Options"
   ClientHeight    =   2940
   ClientLeft      =   2004
   ClientTop       =   852
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   2244
      Left            =   48
      TabIndex        =   2
      Top             =   72
      Width           =   4836
      _ExtentX        =   8530
      _ExtentY        =   3958
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   420
      WordWrap        =   0   'False
      TabCaption(0)   =   "Startup"
      TabPicture(0)   =   "FrmOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraStartup"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "FrmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkOptions(0)"
      Tab(1).Control(1)=   "FraQSFOPtion"
      Tab(1).Control(2)=   "ChkOptions(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Save"
      TabPicture(2)   =   "FrmOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "FraOptions"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Auto Save"
      TabPicture(3)   =   "FrmOptions.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FraAutoSave"
      Tab(3).ControlCount=   1
      Begin VB.CheckBox ChkOptions 
         Caption         =   "Always Keep Toolbar on top."
         Height          =   276
         HelpContextID   =   1007
         Index           =   0
         Left            =   -74880
         TabIndex        =   23
         Top             =   1704
         Width           =   4380
      End
      Begin VB.Frame Frame1 
         Caption         =   "Read-Only Options"
         Height          =   888
         HelpContextID   =   1008
         Left            =   -74928
         TabIndex        =   16
         Top             =   1272
         Width           =   4632
         Begin VB.OptionButton OptReadOnlyOptions 
            Caption         =   "Prompt for save of Read-Only Files"
            Height          =   276
            HelpContextID   =   1008
            Index           =   1
            Left            =   168
            TabIndex        =   18
            Top             =   516
            Width           =   3690
         End
         Begin VB.OptionButton OptReadOnlyOptions 
            Caption         =   "Save Read-Only Files"
            Height          =   276
            HelpContextID   =   1008
            Index           =   0
            Left            =   168
            TabIndex        =   17
            Top             =   240
            Width           =   3810
         End
      End
      Begin VB.Frame FraAutoSave 
         Caption         =   "Extra Save Features"
         Height          =   1812
         HelpContextID   =   1010
         Left            =   -74928
         TabIndex        =   14
         Top             =   336
         Width           =   4692
         Begin VB.CheckBox ChkAutoSaveOptions 
            Caption         =   "Auto Save Breakpoints when exiting Visual Basic."
            Height          =   315
            HelpContextID   =   1010
            Index           =   3
            Left            =   144
            TabIndex        =   22
            Top             =   1392
            Width           =   4092
         End
         Begin VB.CheckBox ChkAutoSaveOptions 
            Caption         =   "Auto Save Bookmarks when exiting Visual Basic."
            Height          =   315
            HelpContextID   =   1010
            Index           =   2
            Left            =   144
            TabIndex        =   21
            Top             =   1008
            Width           =   4092
         End
         Begin VB.CheckBox ChkAutoSaveOptions 
            Caption         =   "Auto Save Find Criteria when exiting Visual Basic."
            Height          =   315
            HelpContextID   =   1010
            Index           =   0
            Left            =   144
            TabIndex        =   20
            Top             =   600
            Width           =   4092
         End
         Begin VB.TextBox TxtSaveTime 
            Height          =   288
            HelpContextID   =   1010
            Left            =   2664
            TabIndex        =   15
            Text            =   "25"
            Top             =   240
            Width           =   300
         End
         Begin VB.CheckBox ChkAutoSaveOptions 
            Caption         =   "Auto Save changed files every              minutes."
            Height          =   228
            HelpContextID   =   1010
            Index           =   1
            Left            =   144
            TabIndex        =   19
            Top             =   264
            Width           =   4092
         End
      End
      Begin VB.Frame FraQSFOPtion 
         Caption         =   "QSF File Options"
         Height          =   996
         HelpContextID   =   1007
         Left            =   -74904
         TabIndex        =   9
         Top             =   312
         Width           =   4644
         Begin VB.OptionButton OptQSFFiles 
            Caption         =   "Remove QSF Files from hard drive when project is removed."
            Height          =   324
            HelpContextID   =   1007
            Index           =   0
            Left            =   48
            TabIndex        =   11
            Top             =   240
            Width           =   4548
         End
         Begin VB.OptionButton OptQSFFiles 
            Caption         =   "Always remove QSF files when exiting Visual Basic."
            Height          =   324
            HelpContextID   =   1007
            Index           =   1
            Left            =   48
            TabIndex        =   10
            Top             =   552
            Width           =   4212
         End
      End
      Begin VB.CheckBox ChkOptions 
         Caption         =   "Minimize Toolbar when main window is minimized."
         Height          =   276
         HelpContextID   =   1007
         Index           =   1
         Left            =   -74880
         TabIndex        =   8
         Top             =   1392
         Width           =   4380
      End
      Begin VB.Frame FraStartup 
         Caption         =   "Auto Retrieval "
         Height          =   1824
         HelpContextID   =   1006
         Left            =   72
         TabIndex        =   6
         Top             =   312
         Width           =   4668
         Begin VB.CheckBox ChkStartupOptions 
            Caption         =   "Auto Retrieve Breakpoints at Startup"
            Height          =   315
            HelpContextID   =   1006
            Index           =   2
            Left            =   168
            TabIndex        =   13
            Top             =   936
            Width           =   4245
         End
         Begin VB.CheckBox ChkStartupOptions 
            Caption         =   "Auto Retrieve Bookmarks at Startup"
            Height          =   315
            HelpContextID   =   1006
            Index           =   1
            Left            =   168
            TabIndex        =   12
            Top             =   600
            Width           =   4245
         End
         Begin VB.CheckBox ChkStartupOptions 
            Caption         =   "Auto Retrieve Find Criteria at Startup"
            Height          =   315
            HelpContextID   =   1006
            Index           =   0
            Left            =   168
            TabIndex        =   7
            Top             =   264
            Width           =   4245
         End
      End
      Begin VB.Frame FraOptions 
         Caption         =   "QSF Save Options"
         Height          =   924
         HelpContextID   =   1008
         Left            =   -74928
         TabIndex        =   3
         Top             =   336
         Width           =   4656
         Begin VB.OptionButton OptSaveOptions 
            Caption         =   "Save all changed files without prompt."
            Height          =   276
            HelpContextID   =   1008
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Top             =   240
            Width           =   3276
         End
         Begin VB.OptionButton OptSaveOptions 
            Caption         =   "Save ONLY the Active Code Window."
            Height          =   276
            HelpContextID   =   1008
            Index           =   1
            Left            =   180
            TabIndex        =   4
            Top             =   528
            Width           =   3060
         End
      End
   End
   Begin VB.CommandButton CmdProcessSave 
      Caption         =   "&Apply"
      Height          =   348
      Index           =   0
      Left            =   3876
      TabIndex        =   1
      Top             =   2472
      Width           =   924
   End
   Begin VB.CommandButton CmdProcessSave 
      Caption         =   "&Close"
      Height          =   348
      Index           =   1
      Left            =   2856
      TabIndex        =   0
      Top             =   2472
      Width           =   924
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkAutoSaveOptions_Click(Index As Integer)
    If Index = 1 Then
        If ChkAutoSaveOptions(1).Value = vbUnchecked Then
            TxtSaveTime.Enabled = False
        Else
            TxtSaveTime.Enabled = True
        End If
    End If
    
End Sub

Private Sub ChkOptions_Click(Index As Integer)
    CmdProcessSave(0).Enabled = True
End Sub

Private Sub CmdProcessSave_Click(Index As Integer)
    
    Select Case (Index)
        Case 0
            Call mSubsaveAllSettings
            CmdProcessSave(0).Enabled = False
        Case 1
            Unload Me
    End Select
End Sub
Private Sub mSubsaveAllSettings()
    
    SaveSetting App.Title, "Settings", "SaveAllFilesWithoutPrompt", OptSaveOptions(gcIntSaveAllFileNoPrompt).Value
    SaveSetting App.Title, "Settings", "SaveOnlyActiveCodePaine", OptSaveOptions(gcIntSaveactiveCodePaineOnly).Value
    SaveSetting App.Title, "Settings", "SaveReadOnlyFiles", OptReadOnlyOptions(gcIntSaveReadOnlyFiles).Value
    SaveSetting App.Title, "Settings", "SaveReadOnlyFilesPrompt", OptReadOnlyOptions(gcIntSaveReadOnlyPrompt).Value
    
    SaveSetting App.Title, "Settings", "RmQSFProjectRm", OptQSFFiles(gcIntRemoveQSFWhenProjectRemoved).Value
    SaveSetting App.Title, "Settings", "AlwaysRmQSF", OptQSFFiles(gcIntAlwaysRemoveQSF).Value
    
    SaveSetting App.Title, "Settings", "MinToolMainMin", ChkOptions(gcIntMinToolMainMin).Value
    SaveSetting App.Title, "Settings", "ToolOnTop", ChkOptions(gcIntAlwaysOnTop).Value
    
    SaveSetting App.Title, "Settings", "FindRetrieveStartup", ChkStartupOptions(0).Value
    
    SaveSetting App.Title, "Settings", "FindAutoSave", ChkAutoSaveOptions(gcInt_ChkAutoSaveOptions).Value
    
    SaveSetting App.Title, "Settings", "AutoSave", ChkAutoSaveOptions(gcInt_ChkAutoSaveOptions_AUTOSAVE).Value
    SaveSetting App.Title, "Settings", "AutoSaveTime", CInt(TxtSaveTime)
    
    SaveSetting App.Title, "Settings", "AutoSaveBKM", ChkAutoSaveOptions(gcInt_ChkAutoSaveOptions_BKM).Value
    SaveSetting App.Title, "Settings", "AutoSaveBKP", ChkAutoSaveOptions(gcInt_ChkAutoSaveOptions_BKP).Value
    SaveSetting App.Title, "Settings", "StartupBKM", ChkStartupOptions(ChkStartupOptions_BKM).Value
    SaveSetting App.Title, "Settings", "StartupBKP", ChkStartupOptions(ChkStartupOptions_BKP).Value
    
    
End Sub
Private Sub Form_Load()
    
    OptSaveOptions(gcIntSaveactiveCodePaineOnly).Value = GetSetting(App.Title, "Settings", "SaveOnlyActiveCodePaine", False)
    OptSaveOptions(gcIntSaveAllFileNoPrompt).Value = GetSetting(App.Title, "Settings", "SaveAllFilesWithoutPrompt", False)
    OptReadOnlyOptions(gcIntSaveReadOnlyFiles).Value = GetSetting(App.Title, "Settings", "SaveReadOnlyFiles", False)
    OptReadOnlyOptions(gcIntSaveReadOnlyPrompt).Value = GetSetting(App.Title, "Settings", "SaveReadOnlyFilesPrompt", True)
    
    OptQSFFiles(gcIntRemoveQSFWhenProjectRemoved).Value = GetSetting(App.Title, "Settings", "RmQSFProjectRm", True)
    OptQSFFiles(gcIntAlwaysRemoveQSF).Value = GetSetting(App.Title, "Settings", "AlwaysRmQSF", False)
    ChkOptions(gcIntMinToolMainMin).Value = GetSetting(App.Title, "Settings", "MinToolMainMin", vbChecked)
    
    ChkOptions(gcIntAlwaysOnTop).Value = GetSetting(App.Title, "Settings", "ToolOnTop", vbChecked)
    
    ChkStartupOptions(0).Value = GetSetting(App.Title, "Settings", "FindRetrieveStartup", vbChecked)
    ChkAutoSaveOptions(gcInt_ChkAutoSaveOptions).Value = GetSetting(App.Title, "Settings", "FindAutoSave", vbChecked)
    
    ChkAutoSaveOptions(gcInt_ChkAutoSaveOptions_AUTOSAVE).Value = GetSetting(App.Title, "Settings", "AutoSave", vbChecked)
    If ChkAutoSaveOptions(gcInt_ChkAutoSaveOptions_AUTOSAVE).Value = vbUnchecked Then
        TxtSaveTime.Enabled = False
    End If
    TxtSaveTime = CInt(GetSetting(App.Title, "Settings", "AutoSaveTime", 25))
    ChkAutoSaveOptions(gcInt_ChkAutoSaveOptions_BKM).Value = GetSetting(App.Title, "Settings", "AutoSaveBKM", vbChecked)
    ChkAutoSaveOptions(gcInt_ChkAutoSaveOptions_BKP).Value = GetSetting(App.Title, "Settings", "AutoSaveBKP", vbChecked)
    ChkStartupOptions(ChkStartupOptions_BKM).Value = GetSetting(App.Title, "Settings", "StartupBKM", vbChecked)
    ChkStartupOptions(ChkStartupOptions_BKP).Value = GetSetting(App.Title, "Settings", "StartupBKP", vbChecked)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call mSubsaveAllSettings
    
End Sub

Private Sub OptQSFFiles_Click(Index As Integer)
    CmdProcessSave(0).Enabled = True
End Sub

Private Sub OptQSFFiles_DblClick(Index As Integer)
    If OptQSFFiles.Item(Index).Value = True Then
        OptQSFFiles.Item(Index).Value = False
    End If
End Sub

Private Sub OptReadOnlyOptions_Click(Index As Integer)
    CmdProcessSave(0).Enabled = True
End Sub

Private Sub OptReadOnlyOptions_DblClick(Index As Integer)
    If OptReadOnlyOptions.Item(Index).Value = True Then
        OptReadOnlyOptions.Item(Index).Value = False
    End If
End Sub

Private Sub OptSaveOptions_Click(Index As Integer)
    CmdProcessSave(0).Enabled = True
End Sub

Private Sub OptSaveOptions_DblClick(Index As Integer)
    If OptSaveOptions.Item(Index).Value = True Then
        OptSaveOptions.Item(Index).Value = False
    End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TxtSaveTime_Change()
    
    If Not IsNumeric(TxtSaveTime) Then
        MsgBox "Only numbers can be entered!", vbInformation, "Value is not a number!"
        TxtSaveTime = "1"
    Else
        If CInt(TxtSaveTime) > 25 Then
            MsgBox "You can not exceed 25 minutes for a maximum saving time!", vbInformation, "Max Time"
            TxtSaveTime = "25"
        ElseIf CInt(TxtSaveTime) = 0 Then
            MsgBox "Setting this to 0 will result in the autosave feature not being enabled!", vbInformation
        End If
    End If
    
End Sub


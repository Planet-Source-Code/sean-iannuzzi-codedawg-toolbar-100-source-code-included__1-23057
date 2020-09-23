VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{389B19AA-9A87-11D1-B77F-00001C1AD1F8}#6.0#0"; "DWSHK36.OCX"
Begin VB.Form FrmCodeDawgToolbar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CodeDawg Toolbar"
   ClientHeight    =   324
   ClientLeft      =   4356
   ClientTop       =   2868
   ClientWidth     =   2064
   Icon            =   "FrmCodeDawgToolbar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   324
   ScaleWidth      =   2064
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2592
      Top             =   1656
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   2952
      Top             =   1128
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save Read-Only File as"
      Filter          =   "*.*"
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2952
      Top             =   672
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Options"
      Height          =   324
      Left            =   1320
      TabIndex        =   8
      Top             =   0
      Width           =   744
   End
   Begin VB.Frame FraMenuItems 
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   315
      HelpContextID   =   1000
      Index           =   0
      Left            =   72
      TabIndex        =   6
      Top             =   408
      Width           =   1932
      Begin VB.Image ImgMenuItem 
         Height          =   192
         Index           =   0
         Left            =   36
         Picture         =   "FrmCodeDawgToolbar.frx":030A
         Top             =   60
         Width           =   192
      End
      Begin VB.Label LblMenuDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Quick Save"
         Height          =   192
         Index           =   0
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "Quick Saves to QSF directory within the project path"
         Top             =   60
         Width           =   840
      End
   End
   Begin VB.Frame FraMenuItems 
      BorderStyle     =   0  'None
      Height          =   315
      HelpContextID   =   1003
      Index           =   1
      Left            =   72
      TabIndex        =   4
      Top             =   720
      Width           =   1932
      Begin VB.Image ImgMenuItem 
         Height          =   192
         Index           =   1
         Left            =   36
         Picture         =   "FrmCodeDawgToolbar.frx":0454
         Top             =   60
         Width           =   192
      End
      Begin VB.Label LblMenuDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieve Bookmar&ks"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Right Click to Toggle Save/Retrieve"
         Top             =   60
         Width           =   1440
      End
   End
   Begin VB.Frame FraMenuItems 
      BorderStyle     =   0  'None
      Height          =   315
      HelpContextID   =   1004
      Index           =   2
      Left            =   72
      TabIndex        =   2
      Top             =   1008
      Width           =   1932
      Begin VB.Label LblMenuDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieve Breakpo&ints"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   3
         ToolTipText     =   "Right Click to Toggle Save/Retrieve"
         Top             =   60
         Width           =   1485
      End
      Begin VB.Image ImgMenuItem 
         Height          =   192
         Index           =   2
         Left            =   36
         Picture         =   "FrmCodeDawgToolbar.frx":059E
         Top             =   60
         Width           =   192
      End
   End
   Begin VB.Frame FraMenuItems 
      BorderStyle     =   0  'None
      Height          =   315
      HelpContextID   =   1005
      Index           =   3
      Left            =   72
      TabIndex        =   0
      Top             =   1296
      Width           =   1932
      Begin VB.Label LblMenuDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieve Find Crit&eria"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   1
         ToolTipText     =   "Right Click to Toggle Save/Retrieve"
         Top             =   60
         Width           =   1470
      End
      Begin VB.Image ImgMenuItem 
         Height          =   192
         Index           =   3
         Left            =   36
         Picture         =   "FrmCodeDawgToolbar.frx":06E8
         Top             =   60
         Width           =   192
      End
   End
   Begin DWSHK36Lib.WinHook WinHook1 
      Left            =   3000
      Top             =   1632
      MessagesAndKeys =   "FrmCodeDawgToolbar.frx":0832
      Notify          =   1
      Monitor         =   0
      HookType        =   0
      HookEnabled     =   0   'False
      KeyboardHook    =   3
      KeyboardNotify  =   1
      KeyboardEvent   =   1
      KeyIgnoreCapsLock=   -1  'True
      KeyViewPeeked   =   0   'False
      RegMessage1     =   ""
      RegMessage2     =   ""
      RegMessage3     =   ""
      RegMessage4     =   ""
      RegMessage5     =   ""
      PostOnFreeze    =   0   'False
      PostOnFreezeMax =   20
      CrossTaskTimeout=   5000
      UseDirectInterface=   0   'False
   End
   Begin VB.Label LblDrop 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   360
      TabIndex        =   9
      Top             =   24
      Width           =   156
   End
   Begin VB.Image ImgTemp 
      Height          =   480
      Left            =   2952
      Picture         =   "FrmCodeDawgToolbar.frx":0C96
      Stretch         =   -1  'True
      Top             =   48
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   7
      Visible         =   0   'False
      X1              =   504
      X2              =   384
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   6
      Visible         =   0   'False
      X1              =   480
      X2              =   408
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Image Image1 
      Height          =   228
      Index           =   0
      Left            =   72
      Picture         =   "FrmCodeDawgToolbar.frx":0FA0
      Stretch         =   -1  'True
      Top             =   48
      Width           =   228
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   504
      X2              =   504
      Y1              =   0
      Y2              =   336
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   2040
      X2              =   24
      Y1              =   1632
      Y2              =   1632
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   2040
      X2              =   2040
      Y1              =   360
      Y2              =   1632
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   360
      Y2              =   1608
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   24
      X2              =   2064
      Y1              =   348
      Y2              =   348
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   336
      X2              =   24
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   360
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   336
   End
   Begin VB.Image Image1 
      Height          =   156
      Index           =   1
      Left            =   408
      Picture         =   "FrmCodeDawgToolbar.frx":12AA
      Stretch         =   -1  'True
      Top             =   96
      Width           =   48
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   336
      X2              =   336
      Y1              =   24
      Y2              =   312
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   2
      Visible         =   0   'False
      X1              =   360
      X2              =   360
      Y1              =   312
      Y2              =   0
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "mnuOptions"
      Visible         =   0   'False
      Begin VB.Menu mniButtonEnabled 
         Caption         =   "Button Enabled"
         Checked         =   -1  'True
      End
      Begin VB.Menu mniOPtions 
         Caption         =   "&Options"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmCodeDawgToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

Public mcbMenuBreakPointToggle      As Office.CommandBarControl
Public mcbMenuBookmarkToggle        As Office.CommandBarControl
Public VBWndHandle                  As Long
Public VBInstance                   As VBIDE.VBE
Public Connect                      As Connect
Private mBln_DoNotChangeColor       As Boolean
Private mBln_MouseDown              As Boolean
Private mBln_ProcessClicked         As Boolean
Private mInt_ProcessClicked         As EnmAssistActions
Dim mColFindCriteria                As Collection
Private mBln_FindWindowShown        As Boolean

Private Enum EnmAssistActions
    EnmDefault = -1
    EnmSaveProject = 0
    EnmSaveBoomarks = 1
    EnmSaveBreakpoints = 2
    EnmSaveFind = 3
End Enum

Private Sub mSubChanegFrameColor(ByVal pvControlIndex As Integer, ByVal pvLng_BKColor As Long, Optional ByVal pvBln_MouseUp = False, Optional ByVal pvIntOpt_MouseButton As Integer = 0)
    
    Dim lIntXX As Integer
        
    If pvIntOpt_MouseButton = 2 Then
        If mniButtonEnabled.Checked = False Then
            mniOPtions.Visible = True
        End If
        PopupMenu mnuOptions
        mniOPtions.Visible = False
        Exit Sub
    End If
        
    If mBln_DoNotChangeColor = True Then
        Exit Sub
    End If
    
    If FraMenuItems(pvControlIndex).BackColor = pvLng_BKColor Then
        Exit Sub
    End If
    
    For lIntXX = FraMenuItems.LBound To FraMenuItems.UBound
        If FraMenuItems(lIntXX).BackColor <> &H8000000D Or pvBln_MouseUp = False Then
            FraMenuItems(lIntXX).BackColor = &H8000000F ' normal
        End If
        DoEvents
    Next
    
    If FraMenuItems(pvControlIndex).BackColor <> &H8000000D Or pvBln_MouseUp = False Then
        FraMenuItems(pvControlIndex).BackColor = pvLng_BKColor
    End If
    
End Sub

Private Sub Command1_Click()
    FrmOptions.Show vbModal
    
End Sub

Private Sub Form_Activate()
    
    'get Thread ID to catch Keyboard event for this instance of VB only
    Dim lRt             As Long
    Dim lLngPRocessID   As Long
    
    VBInstance.MainWindow.SetFocus
    lRt = GetWindowThreadProcessId(VBWndHandle, lLngPRocessID)
    DoEvents
    VBInstance.MainWindow.SetFocus
    'set the spyworks control Task Param
    WinHook1.TaskParam = lLngPRocessID
    DoEvents
    VBInstance.MainWindow.SetFocus
    
    'if auto retrieve option is on then
    If Connect.FormActivated = False Then
        If GetSetting(App.Title, "Settings", "StartupBKM", vbChecked) = vbChecked Then
            Call Connect.mfInt_RetrieveBookmarks
        End If
        
        If GetSetting(App.Title, "Settings", "StartupBKP", vbChecked) = vbChecked Then
            Call Connect.mfInt_RetrieveBreakpoints
        End If
        Connect.FormActivated = True
    End If

End Sub

Private Sub Form_Load()

    Dim lLngPRocessID As Long
        
    'set the help file location
    App.HelpFile = App.Path & "\cd_tb32.hlp"
            
    Me.Top = GetSetting(App.Title, "Settings", "FormTop", 120)
    Me.Left = GetSetting(App.Title, "Settings", "FormLeft", 120)
    
    If Trim$(GetSetting(App.Title, "Settings\FindCriteria", "ItemData0", "")) = "" Or GetSetting(App.Title, "Settings", "FindRetrieveStartup", vbChecked) = vbChecked Then
        Me.LblMenuDescription(3).Caption = "Save Find Crit&eria"
    End If
    
    If GetSetting(App.Title, "Settings", "ButtonEnabled", True) = False Then
        mniButtonEnabled_Click
    End If
    
    If GetSetting(App.Title, "Settings", "StartupBKM", vbChecked) = vbChecked Then
        LblMenuDescription(1) = "Save Bookmar&ks"
    End If
    If GetSetting(App.Title, "Settings", "StartupBKP", vbChecked) = vbChecked Then
        LblMenuDescription(2) = "Save Breakpo&ints"
    End If
    
    mInt_ProcessClicked = EnmDefault
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call mSubTurnOnLines(False)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    SaveSetting App.Title, "Settings", "FormTop", Me.Top
    SaveSetting App.Title, "Settings", "FormLeft", Me.Left
    SaveSetting App.Title, "Settings", "ButtonEnabled", mniButtonEnabled.Checked
End Sub

Private Sub FraMenuItems_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSubChanegFrameColor Index, &H8000000D, pvIntOpt_MouseButton:=Button
    If Button <> 2 Then
        mBln_DoNotChangeColor = True
    End If

End Sub

Private Sub FraMenuItems_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSubChanegFrameColor Index, &H80000014
End Sub

Private Sub FraMenuItems_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSubChanegFrameColor Index, &H80000014, True
    mBln_DoNotChangeColor = False
    
    Call mSubProcessClick(Index)
    
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        If mniButtonEnabled.Checked = False Then
            mniOPtions.Visible = True
        End If
        PopupMenu mnuOptions
        mniOPtions.Visible = False
        Exit Sub
    End If

    'if no item has been selected, then
    'change prompt to show menu items
    If mInt_ProcessClicked = EnmDefault Then
        Index = 1
    End If
    
    Select Case (Index)
        Case 1
            Line1(2).BorderColor = &H808080
            Line1(2).BorderWidth = 2
            Line1(5).BorderColor = &HE0E0E0
            Line1(5).BorderWidth = 1
            Line1(6).BorderColor = &HE0E0E0
            Line1(6).BorderWidth = 1
            Line1(7).BorderColor = &H808080
            Line1(7).BorderWidth = 2
        
            If mBln_MouseDown = True Then
                mBln_MouseDown = False
                'mouse up toggle button
                Line1(2).BorderColor = &HE0E0E0
                Line1(2).BorderWidth = 1
                Line1(5).BorderColor = &H808080
                Line1(5).BorderWidth = 2
                Line1(6).BorderColor = &H808080
                Line1(6).BorderWidth = 2
                Line1(7).BorderColor = &HE0E0E0
                Line1(7).BorderWidth = 1
                ' remove menu
                Me.Height = 630
                If mInt_ProcessClicked = EnmDefault Then
                    Set Image1(0).Picture = ImgTemp.Picture
                End If
            Else
                mBln_MouseDown = True
                'expand form, and show menu
                Me.Height = 2250
                
                'reset menu icon
                mInt_ProcessClicked = EnmDefault
            End If
            
    Case 0
        Line1(4).BorderColor = &HE0E0E0
        Line1(3).BorderColor = &HE0E0E0
        Line1(0).BorderColor = &H808080
        Line1(1).BorderColor = &H808080
        Call mSubProcessClick(mInt_ProcessClicked)
    End Select
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call mSubTurnOnLines(True)
    
End Sub
Private Sub mSubTurnOnLines(ByVal pvBln_On As Boolean)
    Dim lIntXX As Integer
    
    If mBln_MouseDown = True Then
        Exit Sub
    End If
    
    For lIntXX = Line1.LBound To Line1.UBound
        Line1(lIntXX).Visible = pvBln_On
    Next

End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mBln_MouseDown = False Then
        Line1(0).BorderColor = &HE0E0E0
        Line1(1).BorderColor = &HE0E0E0
        Line1(3).BorderColor = &H808080
        Line1(4).BorderColor = &H808080
    End If
End Sub

Private Sub ImgMenuItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSubChanegFrameColor Index, &H8000000D, pvIntOpt_MouseButton:=Button
    If Button <> 2 Then
        mBln_DoNotChangeColor = True
    End If

End Sub

Private Sub ImgMenuItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        mSubChanegFrameColor Index, &H80000014
End Sub

Private Sub ImgMenuItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSubChanegFrameColor Index, &H80000014, True
    mBln_DoNotChangeColor = False
    Call mSubProcessClick(Index)
End Sub

Private Sub LblDrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Image1_MouseDown(1, Button, Shift, X, Y)
End Sub

Private Sub LblDrop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call mSubTurnOnLines(True)
    
End Sub

Private Sub LblDrop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If mBln_MouseDown = False Then
        Line1(0).BorderColor = &HE0E0E0
        Line1(1).BorderColor = &HE0E0E0
        Line1(3).BorderColor = &H808080
        Line1(4).BorderColor = &H808080
    End If
End Sub

Private Sub LblMenuDescription_DblClick(Index As Integer)
    Call mSubProcessClick(Index)
End Sub

Private Sub LblMenuDescription_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 2 Then
        mSubChanegFrameColor Index, &H8000000D, pvIntOpt_MouseButton:=Button
        mBln_DoNotChangeColor = True
    Else
        Select Case (Index)
            Case 1
                If Left$(LblMenuDescription(1).Caption, 4) = "Save" Then
                    LblMenuDescription(1).Caption = "Retrieve Bookmar&ks"
                Else
                    LblMenuDescription(1).Caption = "Save Bookmar&ks"
                End If
            Case 2
                If Left$(LblMenuDescription(2).Caption, 4) = "Save" Then
                    LblMenuDescription(2).Caption = "Retrieve Breakpo&ints"
                Else
                    LblMenuDescription(2).Caption = "Save Breakpo&ints"
                End If
            Case 3
                If Left$(LblMenuDescription(3).Caption, 4) = "Save" Then
                    LblMenuDescription(3).Caption = "Retrieve Find Crit&eria"
                Else
                    LblMenuDescription(3).Caption = "Save Find Crit&eria"
                End If
        End Select
    End If
End Sub

Private Sub LblMenuDescription_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSubChanegFrameColor Index, &H80000014
End Sub

Private Sub LblMenuDescription_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSubChanegFrameColor Index, &H80000014, True
    mBln_DoNotChangeColor = False
    If Button <> 2 Then
        Call mSubProcessClick(Index)
    End If
End Sub
Private Sub mSubProcessClick(ByVal pvIntIndex As Integer)
    
    Dim lBln_SaveActiveCodePaineOnly    As Boolean
    Dim lBln_SaveAllFileNoPrompt        As Boolean
    Dim lLngSavedCountCount             As Long
    
    Timer1.Enabled = False
    
    mInt_ProcessClicked = pvIntIndex
    Set Image1(0).Picture = ImgMenuItem(pvIntIndex).Picture
    Select Case (pvIntIndex)
        Case 0
            'retrive options
            lBln_SaveActiveCodePaineOnly = GetSetting(App.Title, "Settings", "SaveOnlyActiveCodePaine", False)
            lBln_SaveAllFileNoPrompt = GetSetting(App.Title, "Settings", "SaveAllFilesWithoutPrompt", False)

            Load FrmSaveChanges
            Set FrmSaveChanges.Connect = Me.Connect
            Set FrmSaveChanges.VBInstance = Me.VBInstance
            If lBln_SaveAllFileNoPrompt = True Then
                Call FrmSaveChanges.gfIntQuickSaveVBProject
            ElseIf lBln_SaveActiveCodePaineOnly = True Then
                'save this only window only
                Call FrmSaveChanges.gfIntQuickSaveVBProject
            Else
                'show form
                FrmSaveChanges.Show vbModal
            End If
            Unload FrmSaveChanges
        Case 1
            If Left$(LblMenuDescription(1).Caption, 4) = "Save" Then
                Call Connect.mfInt_SaveBookMark_BreakPointFile(gcStr_BookMarkFileExtension)
                MsgBox "Bookmarks have been saved!"
            Else
                Call Connect.mfInt_RetrieveBookmarks
                LblMenuDescription(1).Caption = "Save Bookmar&ks"
                MsgBox "Bookmarks have been retrieved!"
            End If
        Case 2
            If Left$(LblMenuDescription(2).Caption, 4) = "Save" Then
                Call Connect.mfInt_SaveBookMark_BreakPointFile(gcStr_BreakpointFileExtension)
                MsgBox "Breakpoints have been saved!"
            Else
                Call Connect.mfInt_RetrieveBreakpoints
                LblMenuDescription(2).Caption = "Save Breakpo&ints"
                MsgBox "Breakpoints have been retrieved!"
            End If
        Case 3
            If Left$(LblMenuDescription(3).Caption, 4) = "Save" Then
                Call pfBln_SaveFindCriteria
            Else
                If mfBln_RetrieveFindCriteria = True Then
                    LblMenuDescription(3).Caption = "Save Find Crit&eria"
                End If
                
            End If
            FindWindowShown = True
    End Select
    
    If mniButtonEnabled.Checked = True Then
        If mInt_ProcessClicked <> EnmDefault Then
             Me.Height = 630
        End If
    End If
    Timer1.Enabled = True
End Sub
Private Function mfBln_RetrieveFindCriteria(Optional ByVal pvBln_DisplaySuccess As Boolean = True) As Boolean

    Dim lLngComboHandle     As Long
    Dim lLngCount           As Long
    Dim lStrData            As String
    Dim lBln_DoneReading    As Boolean
    
    Const CB_ADDSTRING = &H143
        
    lLngComboHandle = mfLngGetFindWindowComboForVB()
    
    '990903 SBI0 New way only retrieves up to 15
     'retrieve info
    For lLngCount = 1 To gcIntMaxNumFindSaved
        lStrData = GetSetting(App.Title, "Settings\FindCriteria", "ItemData" & CStr(lLngCount), "")
        If Trim$(lStrData) <> "" Then
             'check to see if the value is in the combo already
             'if so then skip it will be true
             Call mfBln_ShouldSkipAdd(lLngComboHandle, Trim$(lStrData))
        Else
            Exit For
        End If
    Next
        
    If Not mColFindCriteria Is Nothing Then
        If mColFindCriteria.Count > 0 Then
            If lLngComboHandle > 0 Then
                'put all source code to file
                For lLngCount = mColFindCriteria.Count To 1 Step -1
                    If mfBln_ShouldSkipAdd(lLngComboHandle, Trim$(mColFindCriteria(lLngCount))) = False Then
                        Call SendMessageStr(lLngComboHandle, CB_ADDSTRING, 0&, ByVal Trim$(mColFindCriteria(lLngCount)))
                    End If
                    DoEvents
                Next
                If pvBln_DisplaySuccess = True Then
                    MsgBox "Find Criteria Restored!"
                End If
            End If
            mfBln_RetrieveFindCriteria = True
        End If
    Else
        If pvBln_DisplaySuccess = True Then
            MsgBox "There are no find criteria items to retrieve!"
        End If
    End If
        
End Function
Private Function mfBln_ShouldSkipAdd(ByVal lLngComboHandle As Long, ByVal lStrDataIn As String) As Boolean

    'before we add it to the list check to see if it is on the combo
    Dim lLngComboCount   As Long
    Dim lLngCountIndex   As Long
    Dim lStrComboItems   As String
    Dim lRtn             As Long
    Dim pos              As Long      '***NEW
    Dim lStrDataFrom     As String  '***NEW
    Dim lBln_SkipIt      As Boolean
    
    Const CB_GETCOUNT = &H146
    Const CB_GETLBTEXT = &H148

    lLngComboCount = SendMessageLong(lLngComboHandle, CB_GETCOUNT, 0&, 0&)

    lBln_SkipIt = False
    
    If lLngComboCount = 0 Then
        'If in collection then
        If Trim$(lStrDataIn) <> "" Then
            If gfBln_DoesItemExistInCol(mColFindCriteria, Trim$(lStrDataIn)) = False Then
                If mColFindCriteria Is Nothing Then
                    Set mColFindCriteria = New Collection
                End If
                mColFindCriteria.Add Trim$(lStrDataIn), Trim$(lStrDataIn)
            End If
        End If
    End If
    
    'If the combo has item them kill the old file
    'put all source code to file
    For lLngCountIndex = 0 To lLngComboCount - 1

        'pad the string
        lStrDataFrom = Space$(255)
        
        'retrieve the text
        lRtn = SendMessageStr(lLngComboHandle, CB_GETLBTEXT, lLngCountIndex, ByVal lStrDataFrom)
        
        'trim off trailing nulls        ***NEW
        pos = InStr(lStrDataFrom, Chr$(0))
        
        If pos Then     ' ***NEW
            lStrDataFrom = Left$(lStrDataFrom, InStr(lStrDataFrom, Chr$(0)) - 1)
            
            'If in collection then
            If gfBln_DoesItemExistInCol(mColFindCriteria, Trim$(lStrDataFrom)) = False Then
                If mColFindCriteria Is Nothing Then
                    Set mColFindCriteria = New Collection
                End If
                If Trim$(lStrDataFrom) <> "" Then
                    mColFindCriteria.Add Trim$(lStrDataFrom), Trim$(lStrDataFrom)
                End If
            End If

            If Trim$(lStrDataFrom) = Trim$(lStrDataIn) Then
                lBln_SkipIt = True
                Exit For
            End If
            
        End If
        DoEvents
    Next
    
    mfBln_ShouldSkipAdd = lBln_SkipIt
    
End Function
Private Function mfLngGetFindWindowComboForVB() As Long
    
    Dim lFindWndHandles() As Long
    Dim r As Long
    Dim lIntXX As Integer
    Dim lLngComboHandle As Long
    
    'now check for find window using VB as the owner to start with
    r = FindWindowLike(lFindWndHandles(), 0, "Find", "*")
    If r > 0 Then
        'then get combo ref on Find Window
        For lIntXX = 0 To UBound(lFindWndHandles)
            If r = 1 And lIntXX = 1 Then
                'find window, only one
                lLngComboHandle = GetComboRefFromFindWindow(lFindWndHandles(lIntXX))
                If lLngComboHandle <> 0 Then
                    Exit For
                End If
            Else
                If GetParent(lFindWndHandles(lIntXX)) = VBWndHandle Then
                    lLngComboHandle = GetComboRefFromFindWindow(lFindWndHandles(lIntXX))
                    If lLngComboHandle <> 0 Then
                        'Find window found
                        Exit For
                    End If
                End If
            End If
        Next
    End If
    
    'return handle
    mfLngGetFindWindowComboForVB = lLngComboHandle
    
End Function
Public Function pfBln_SaveFindCriteria() As Boolean

    Dim lLngComboHandle As Long
    
    If FindWindowShown = True Then
        If pfLng_SaveFindComboInfoUsingCol > 0 Then
            MsgBox "Find Criteria Saved!", vbInformation, "Find Criteria Saved!"
            pfBln_SaveFindCriteria = True
        Else
            MsgBox "There is no find data to save!", vbInformation
        End If
    Else
        MsgBox "There is no find data to save!", vbInformation
    End If
    
End Function
Private Sub mniButtonEnabled_Click()
    Dim lIntXX As Integer
    If mniButtonEnabled.Checked = False Then
        mniButtonEnabled.Checked = True
        For lIntXX = Line1.LBound To Line1.UBound
            Line1(lIntXX).Visible = True
        Next
        Image1(0).Visible = True
        Image1(1).Visible = True
        
        Line2(0).X1 = 0
        Line2(0).X2 = 2040
        Line2(0).Y1 = 420
        Line2(0).Y2 = 420
        Line2(1).X1 = 0
        Line2(1).X2 = 0
        Line2(1).Y1 = 420
        Line2(1).Y2 = 1860
        Line2(2).X1 = 2040
        Line2(2).X2 = 2040
        Line2(2).Y1 = 450
        Line2(2).Y2 = 1950
        Line2(3).X1 = 2040
        Line2(3).X2 = 0
        Line2(3).Y1 = 1935
        Line2(3).Y2 = 1935
        
        FraMenuItems(0).Top = 495
        FraMenuItems(1).Top = 855
        FraMenuItems(2).Top = 1215
        FraMenuItems(3).Top = 1575
        Me.Height = 630
        Command1.Visible = True
    Else
        mniButtonEnabled.Checked = False
        For lIntXX = Line1.LBound To Line1.UBound
            Line1(lIntXX).Visible = False
        Next
        Image1(0).Visible = False
        Image1(1).Visible = False
        
        Line2(0).X1 = 0
        Line2(0).X2 = 2040
        Line2(0).Y1 = 0
        Line2(0).Y2 = 0
        Line2(1).X1 = 0
        Line2(1).X2 = 0
        Line2(1).Y1 = 30
        Line2(1).Y2 = 1470
        Line2(2).X1 = 2040
        Line2(2).X2 = 2040
        Line2(2).Y1 = 30
        Line2(2).Y2 = 1530
        Line2(3).X1 = 2040
        Line2(3).X2 = 0
        Line2(3).Y1 = 1515
        Line2(3).Y2 = 1515
        FraMenuItems(0).Top = 45
        FraMenuItems(1).Top = 405
        FraMenuItems(2).Top = 765
        FraMenuItems(3).Top = 1125
        Me.Height = 1850
        Command1.Visible = False
    End If
    
End Sub


Private Function mfLng_SaveFindComboInfo(ByVal pvLng_CmbHWnd As Long)

On Error GoTo mfLng_SaveFindComboInfo_ErrorHandler:

    Dim hwndEdit         As Long
    Dim lLngComboCount   As Long
    Dim lLngCount        As Long
    Dim lStrComboItems   As String
    Dim lRtn             As Long
    Dim pos              As Long      '***NEW
    Dim lStrData         As String  '***NEW
    
    Const CB_GETCOUNT = &H146
    Const CB_GETLBTEXT = &H148

    lLngComboCount = SendMessageLong(pvLng_CmbHWnd, CB_GETCOUNT, 0&, 0&)

    If lLngComboCount > 0 Then 'if item found copy them
        
        'delete the Last saved find entries
        DeleteSetting App.Title, "Settings\FindCriteria"

        'If the combo has item them kill the old file
        'put all source code to file
        For lLngCount = lLngComboCount To 0 Step -1 'save in reverse order when brough back will be right
            
            '990903 SBI added to limit the number of find items saved
            If lLngCount <= gcIntMaxNumFindSaved Then
                'save info
            
                'pad the string
                lStrData = Space$(255)
                
                'retrieve the text
                lRtn = SendMessageStr(pvLng_CmbHWnd, CB_GETLBTEXT, lLngCount, ByVal lStrData)
                
                'trim off trailing nulls        ***NEW
                pos = InStr(lStrData, Chr$(0))
                
                If pos Then     ' ***NEW
                    lStrData = Left$(lStrData, InStr(lStrData, Chr$(0)) - 1)
                    If Trim$(lStrData) <> "" Then
                        SaveSetting App.Title, "Settings\FindCriteria", "ItemData" & CStr(lLngCount), lStrData
                    End If
                End If
            End If
            DoEvents
        Next
   End If
   'return saved count
   mfLng_SaveFindComboInfo = lLngComboCount
   Exit Function
    
mfLng_SaveFindComboInfo_ErrorHandler:

    If Err Then Resume Next
    
End Function

Public Function pfLng_SaveFindComboInfoUsingCol()

On Error GoTo pfLng_SaveFindComboInfoUsingCol_ErrorHandler:

    Dim lLngCount        As Long

    If Not mColFindCriteria Is Nothing Then
        If mColFindCriteria.Count > 0 Then 'if item found copy them
            
            'delete the Last saved find entries
            DeleteSetting App.Title, "Settings\FindCriteria"
    
            'If the combo has item them kill the old file
            'put all source code to file
            For lLngCount = mColFindCriteria.Count To 1 Step -1
                SaveSetting App.Title, "Settings\FindCriteria", "ItemData" & CStr(lLngCount), mColFindCriteria(lLngCount)
                DoEvents
                If Err Then
                    Exit For
                End If
            Next
       End If
   End If
   
   'return saved count
   pfLng_SaveFindComboInfoUsingCol = mColFindCriteria.Count
   Exit Function
    
pfLng_SaveFindComboInfoUsingCol_ErrorHandler:

    If Err Then Resume Next
    
End Function

Private Sub mniOPtions_Click()
FrmOptions.Show vbModal
End Sub

Private Sub Timer1_Timer()

 'IMPORTANT ORDER MATTERS FOR PROCESSING TIME **********
 
On Error GoTo Timer1_Timer_ErrorHandler:
    
    Dim lBlnWinMinHit               As Boolean
        
    If GetSetting(App.Title, "Settings", "ToolOnTop", vbChecked) = vbChecked Then
        FrmCodeDawgToolbar.ZOrder 0
    End If
    
    'check the list to see if we need to add the data from the reg to te window
    If VBInstance.Windows.Item("Find").Visible = True Then
        If FindWindowShown = False Then
            If GetSetting(App.Title, "Settings", "FindRetrieveStartup", vbChecked) = vbChecked Then
                Call mfBln_RetrieveFindCriteria(False)
            Else
                Call mSubLoadFindCol
            End If
        Else
            Call mSubLoadFindCol
            Call mfBln_RetrieveFindCriteria(False)
        End If
        FindWindowShown = True
    End If
    
WindowMin:
    lBlnWinMinHit = True
    If GetSetting(App.Title, "Settings", "MinToolMainMin", vbChecked) = vbChecked Then
        If VBInstance.MainWindow.WindowState = vbext_ws_Minimize Then
            'if reg setting is on then
                Connect.ShowWindow vbMinimized, False
            'end if
        ElseIf (VBInstance.MainWindow.WindowState = vbext_ws_Normal Or VBInstance.MainWindow.WindowState = vbext_ws_Maximize) Then
            Connect.ShowWindow vbNormal
        End If
    End If
    Me.Caption = VBInstance.MainWindow.Caption
            
      
    'check to see if we need to save
    If GetSetting(App.Title, "Settings", "AutoSave", vbChecked) = vbChecked Then
        
        'check time to save
        Dim lIntTimeToSave              As Integer
        Static lStcIntTimeCount         As Integer
        Dim lIntNumberOfMinutesPassed   As Integer
        
        lIntTimeToSave = CInt(GetSetting(App.Title, "Settings", "AutoSaveTime", 25))
        lStcIntTimeCount = lStcIntTimeCount + 1
        If lStcIntTimeCount = 120 Then
            lIntNumberOfMinutesPassed = lIntNumberOfMinutesPassed + 1
            lStcIntTimeCount = 0
        End If
        If lIntNumberOfMinutesPassed >= lIntTimeToSave Then
            lStcIntTimeCount = 0
            lIntNumberOfMinutesPassed = 0
            Call mSubProcessClick(0)
        End If
    Else
        'Reset time to 0
        lStcIntTimeCount = 0
    End If
    
    'check if in run mode if so, save breakpoints, abd bookmarks
    Static lBln_SaveOnce As Boolean
    If InRunMode Then
        If lBln_SaveOnce = True Then
            'save bookmarks, and breakpoints
            Call Connect.mfInt_SaveBookMark_BreakPointFile(gcStr_BookMarkFileExtension)
            Call Connect.mfInt_SaveBookMark_BreakPointFile(gcStr_BreakpointFileExtension)
            If FindWindowShown = True Then
                Call pfLng_SaveFindComboInfoUsingCol
            End If
            lBln_SaveOnce = False
        End If
    Else
        lBln_SaveOnce = True
    End If
    
    Exit Sub
Timer1_Timer_ErrorHandler:
    If Err = 91 Then
        If lBlnWinMinHit = False Then
            Resume WindowMin:
        End If
    End If
    
    Exit Sub
End Sub
Private Sub mSubLoadFindCol()


    Dim lLngComboHandle     As Long
    Dim lLngCount           As Long
    Dim lStrData            As String
    Dim lBln_DoneReading    As Boolean
    
    Const CB_ADDSTRING = &H143
    
    lLngComboHandle = mfLngGetFindWindowComboForVB()
    Call mfBln_ShouldSkipAdd(lLngComboHandle, "")

End Sub
Public Property Get FindWindowShown() As Boolean
    FindWindowShown = mBln_FindWindowShown
End Property

Public Property Let FindWindowShown(ByVal vNewValue As Boolean)
    mBln_FindWindowShown = vNewValue
End Property

Private Sub Timer2_Timer()
On Error GoTo Timer2_Timer_ErrorHandler:
    If Not VBInstance.ActiveCodePane Is Nothing Then
        Call mSubHandleKeyPress
    End If

Timer2_Timer_ErrorHandler:
    If Err Then Exit Sub
End Sub

Private Sub WinHook1_KbdHook(keycode As Long, keystate As Long, ByVal shiftstate As Integer, discard As Integer)

On Error GoTo WinHook1_KbdHook_ErrorHandler:
    
    Static lStcBln_DidIt        As Boolean
    
    'Here becuase this method is envokle twice, toggle it
    If lStcBln_DidIt = True Then
        lStcBln_DidIt = False
        Exit Sub
    Else
        lStcBln_DidIt = True
    End If
    
    '990911 SBI0 Place in for speed keys
    Select Case (keycode)
        Case 120 'Break Point Clear
            If shiftstate = 3 Then
                'Break Point Clear
                Connect.pSub_BreakPointRemoveAllKeyDown
                lStcBln_DidIt = False
            ElseIf shiftstate = 0 Then
                Call Connect.pSub_BreakPointKeyDownToggle
            End If
        Case 81 'Control Q = Quick Save
            Call mSubProcessClick(0)
            lStcBln_DidIt = False
        Case 75 'Control K = Bookmarks
            Call mSubProcessClick(1)
            lStcBln_DidIt = False
        Case 73 'Control I = Breakpoints
            Call mSubProcessClick(2)
            lStcBln_DidIt = False
        Case 69 'Control E = Criteria
            Call mSubProcessClick(3)
            lStcBln_DidIt = False
            
    End Select
        
WinHook1_KbdHook_ErrorHandler:
    If Err Then Exit Sub

End Sub
Private Sub mSubHandleKeyPress()

On Error GoTo mSubHandleKeyPress_ErrorHandler

    If Not VBInstance.ActiveCodePane Is Nothing Then

        Dim lLngCurLineNumber   As Long
        Dim lLngDummy1          As Long
        Dim lLngDummy2          As Long
        Dim lLngDummy3          As Long
        Dim lBln_AddORDelete    As Boolean
        Dim lLngCurCountOfLines As Long
        Dim lLngNumberOfLineDiff As Long
        Static lLngPreviousCountofLines As Long
        Static lStcStr_PreviousCodePaineCaption As String
        
        If lStcStr_PreviousCodePaineCaption <> VBInstance.ActiveCodePane.Window.Caption Then
            lStcStr_PreviousCodePaineCaption = VBInstance.ActiveCodePane.Window.Caption
            lLngPreviousCountofLines = 0
        End If
                
        'get count of lines
        lLngCurCountOfLines = VBInstance.ActiveCodePane.CodeModule.CountOfLines
        
        'get current line pos
        VBInstance.ActiveCodePane.GetSelection lLngCurLineNumber, lLngDummy1, lLngDummy2, lLngDummy3
        
        'lLngPreviousCountofLines = Connect.TotalNumLines
        If lLngPreviousCountofLines = 0 Then
            lLngPreviousCountofLines = lLngCurCountOfLines
        End If
        
        'check if lines have been added
        If lLngPreviousCountofLines < lLngCurCountOfLines And lLngPreviousCountofLines <> 0 Then
            lBln_AddORDelete = True
            lLngNumberOfLineDiff = lLngCurCountOfLines - lLngPreviousCountofLines
            Call Connect.pSub_LineChange(lBln_AddORDelete, lLngCurLineNumber, lLngNumberOfLineDiff)
        End If
        
        'check if lines have been removed
        If lLngPreviousCountofLines > lLngCurCountOfLines And lLngPreviousCountofLines <> 0 Then
            lBln_AddORDelete = False
            lLngNumberOfLineDiff = lLngPreviousCountofLines - lLngCurCountOfLines
            Call Connect.pSub_LineChange(lBln_AddORDelete, lLngCurLineNumber, lLngNumberOfLineDiff)
        End If
        lLngPreviousCountofLines = lLngCurCountOfLines
        'if no change exit
        If lLngNumberOfLineDiff = 0 Then
            Exit Sub
        End If
    End If
    Exit Sub

mSubHandleKeyPress_ErrorHandler:
    If Err Then
        Exit Sub
    End If
End Sub

Private Function mfLng_GetLineCountFromClipboard() As Long
     mfLng_GetLineCountFromClipboard = Len(Clipboard.GetText) - Len(Replace(Clipboard.GetText, vbCr, Mid$(vbCr, 2)))
     Debug.Print "Clip count " & mfLng_GetLineCountFromClipboard
End Function

Private Function InRunMode() As Boolean
  InRunMode = (VBInstance.CommandBars("File").Controls(1).Enabled = False)
End Function

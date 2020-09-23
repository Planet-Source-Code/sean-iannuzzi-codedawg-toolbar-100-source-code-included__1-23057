VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9000
   ClientLeft      =   2460
   ClientTop       =   924
   ClientWidth     =   7068
   _ExtentX        =   12467
   _ExtentY        =   15875
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "CodeDawg Toolbar Assistant"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mFrmCodeDawgToolbar       As New FrmCodeDawgToolbar
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Dim mcbMenuBreakPointToggle     As Office.CommandBarControl
Dim mcbMenuBookmarkToggle       As Office.CommandBarControl

Dim mLngTotalNumLines As Long

'New Code
Dim mClsBookMarks           As ClsBookmarks
Dim mClsBreakPoints         As ClsBreakPoints
 
Dim mvarFormActivated As Boolean
 
'used to keep the vb instance
Private mLng_VBWndHandle As Long

'''
Private Enum EnmBarType
    EnmBreakPointToggle = 51
    EnmBreakPointRemoveALL = 579
    EnmBookmarkToggle = 2525
    EnmBookmarkremoveAll = 2528
End Enum

'needed to subclass the all events from the toolbars
'MENU TITLE 'MENU BAR, SUB MENU DEBUG
Public WithEvents Subclass_MenuMainDebug_BreakPointToggle    As CommandBarEvents
Attribute Subclass_MenuMainDebug_BreakPointToggle.VB_VarHelpID = -1
Public WithEvents Subclass_MenuMainDebug_BreakPointRemoveAll As CommandBarEvents
Attribute Subclass_MenuMainDebug_BreakPointRemoveAll.VB_VarHelpID = -1

'"MENU TITLE 'Debug'
Public WithEvents Subclass_MenuDebug_BreakPointToggle    As CommandBarEvents
Attribute Subclass_MenuDebug_BreakPointToggle.VB_VarHelpID = -1
Public WithEvents Subclass_MenuDebug_BreakPointRemoveAll As CommandBarEvents
Attribute Subclass_MenuDebug_BreakPointRemoveAll.VB_VarHelpID = -1

'Menu Title 'Edit'
Public WithEvents Subclass_Edit_BreakPointToggle       As CommandBarEvents
Attribute Subclass_Edit_BreakPointToggle.VB_VarHelpID = -1
Public WithEvents Subclass_Edit_BookmarkToggle         As CommandBarEvents
Attribute Subclass_Edit_BookmarkToggle.VB_VarHelpID = -1
Public WithEvents Subclass_Edit_BookmarkRemoveAll      As CommandBarEvents
Attribute Subclass_Edit_BookmarkRemoveAll.VB_VarHelpID = -1

'Menu Title 'Toggle'
Public WithEvents Subclass_Toggle_BreakPointToggle       As CommandBarEvents
Attribute Subclass_Toggle_BreakPointToggle.VB_VarHelpID = -1
Public WithEvents Subclass_Toggle_BookmarkToggle         As CommandBarEvents
Attribute Subclass_Toggle_BookmarkToggle.VB_VarHelpID = -1

'MenuBookmarks
Public WithEvents Subclass_MenuBookmarks_BookmarkToggle         As CommandBarEvents
Attribute Subclass_MenuBookmarks_BookmarkToggle.VB_VarHelpID = -1
Public WithEvents Subclass_MenuBookmarks_BookmarkRemoveAll      As CommandBarEvents
Attribute Subclass_MenuBookmarks_BookmarkRemoveAll.VB_VarHelpID = -1

'For Project Events
Public WithEvents Subclass_VBProjectEvents As VBProjectsEvents
Attribute Subclass_VBProjectEvents.VB_VarHelpID = -1
'''

Sub ShowWindow(ByVal pvEnmWindowState As FormWindowStateConstants, Optional ByVal pvBln_HideWindow As Boolean = True)
    If mFrmCodeDawgToolbar.WindowState <> pvEnmWindowState Then
        mFrmCodeDawgToolbar.WindowState = pvEnmWindowState
        If mFrmCodeDawgToolbar.Visible <> pvBln_HideWindow Then
            mFrmCodeDawgToolbar.Visible = pvBln_HideWindow
        End If
    End If
End Sub
Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mFrmCodeDawgToolbar.Hide
   
End Sub

Sub Show()
  
    On Error Resume Next
    
    If mFrmCodeDawgToolbar Is Nothing Then
        Set mFrmCodeDawgToolbar = New FrmCodeDawgToolbar
    End If
        
    Set mFrmCodeDawgToolbar.VBInstance = VBInstance
    Set mFrmCodeDawgToolbar.Connect = Me
    Set mFrmCodeDawgToolbar.mcbMenuBreakPointToggle = mcbMenuBreakPointToggle
    Set mFrmCodeDawgToolbar.mcbMenuBookmarkToggle = mcbMenuBookmarkToggle
    mFrmCodeDawgToolbar.VBWndHandle = VBWndHandle
        
    If mClsBookMarks Is Nothing Then
        Set mClsBookMarks = New ClsBookmarks
    End If
    If mClsBreakPoints Is Nothing Then
        Set mClsBreakPoints = New ClsBreakPoints
    End If
    
    FormDisplayed = True
    
    mFrmCodeDawgToolbar.Show
   
End Sub

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
    
    mFrmCodeDawgToolbar.Timer1.Enabled = False
    
    'remove all QSF Files for all projects if options checked
    If GetSetting(App.Title, "Settings", "AlwaysRmQSF", False) = True Then
        Dim lVbProject As VBProject
        For Each lVbProject In VBInstance.VBProjects
            Call mSubRemoveQSFFiles(lVbProject)
        Next
    End If
    
    'check to see if we should auto save the find criteria
    
    If GetSetting(App.Title, "Settings", "FindAutoSave", vbChecked) = vbChecked Then
        'check to see if the user even opened the find window
        If mFrmCodeDawgToolbar.FindWindowShown = True Then
            Call mFrmCodeDawgToolbar.pfLng_SaveFindComboInfoUsingCol
        End If
    End If
    
    If GetSetting(App.Title, "Settings", "AutoSaveBKM", vbChecked) = vbChecked Then
        Call mfInt_SaveBookMark_BreakPointFile(gcStr_BookMarkFileExtension)
    End If
    If GetSetting(App.Title, "Settings", "AutoSaveBKP", vbChecked) = vbChecked Then
        Call mfInt_SaveBookMark_BreakPointFile(gcStr_BreakpointFileExtension)
    End If
        
    Set mClsBookMarks = Nothing
    Set mClsBreakPoints = Nothing
        
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    
    'FrmUnlicensedCopyNotice.Show vbModal
    
    'SBI0 Added to show the form on connection
    'Moved to show on startup complete
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("CodeDawg Toolbar Assistant")
        
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
        
        'Sub Class all neccessary buttons on menus
        Call mSub_InitMenuSubClassing("Menu Bar", 30165, "MenubarDebug")
        Call mSub_InitMenuSubClassing("Bookmarks")
        Call mSub_InitMenuSubClassing("Debug")
        Call mSub_InitMenuSubClassing("Edit")
        Call mSub_InitMenuSubClassing("Toggle")
            
        'get VB Handle and set property of handle
        Call mSubGetVBHandle
                    
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
            
    Unload mFrmCodeDawgToolbar
    Set mFrmCodeDawgToolbar = Nothing
        
    'kill everything
    Call mSubSetAllSubClassItemsToNothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    
    
    'after the projects have been loaded sub class the VB events
    Set Subclass_VBProjectEvents = VBInstance.Events.VBProjectsEvents
    If VBInstance.VBProjects.Count >= 1 Then
        Show

    End If
    
    
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

Private Sub mSub_InitMenuSubClassing(ByVal pvStrMenuToSubClass As String, Optional ByVal pvLngSubMenuTopID As Long, Optional ByVal pvStrSubIDSubName As String)
'''''''''''''''''''''''
    'Routine handle subclassing all of the menu bar passed in
    ' Break points and bookmarks
    
'''''''''''''''''''''''
On Error GoTo mSub_InitMenuSubClassing_ErrorHandler:

    Dim lObj_MenuBar        As Object
    Dim lobj_MenuControl    As Office.CommandBarControl
    Dim lObjSubMenuControl  As Office.CommandBarControl
    Dim lBln_FoundIt        As Boolean
    Dim lObj_SubMenuBar     As Object
    
    
    Set lObj_MenuBar = VBInstance.CommandBars(pvStrMenuToSubClass)
    
    For Each lobj_MenuControl In lObj_MenuBar.Controls
        If pvStrSubIDSubName = "" Then
            Call mSub_SetEventToSink(lobj_MenuControl, pvStrMenuToSubClass)
            lobj_MenuControl.Enabled = True
        Else
            If lobj_MenuControl.Id = pvLngSubMenuTopID Then
                For Each lObj_SubMenuBar In lobj_MenuControl.Controls
                    'Debug.Print lObj_SubMenuBar.Caption & "|| " & lObj_SubMenuBar.Id
                    Call mSub_SetEventToSink(lObj_SubMenuBar, pvStrSubIDSubName)
                    lObj_SubMenuBar.Enabled = True
                Next
            End If
'            Debug.Print lobj_MenuControl.Caption & "|| " & lobj_MenuControl.Id
        End If
    Next
    
    'kill alll objects
    Set lObj_MenuBar = Nothing
    Set lobj_MenuControl = Nothing
    Set lObjSubMenuControl = Nothing
    
mSub_InitMenuSubClassing_ErrorHandler:
    
    If Err Then
        MsgBox "There was an error trying to initialize one or more of the Menu Bar buttons!", vbCritical
        Resume Next
    End If
    
End Sub
Private Sub mSub_SetEventToSink(ByVal lObjCommandBar As Office.CommandBarControl, ByVal pvStrMenuToSubClass As String)

On Error GoTo mSub_SetEventToSink_ErrorHandler:
    Select Case (lObjCommandBar.Id)
        Case EnmBreakPointToggle
            Select Case (UCase$(pvStrMenuToSubClass))
                Case "DEBUG"
                    'used to set breakpoints, only need one reference to this
                    If mcbMenuBreakPointToggle Is Nothing Then
                        Set mcbMenuBreakPointToggle = lObjCommandBar
                        mcbMenuBreakPointToggle.Enabled = True
                    End If
                    Set Subclass_MenuDebug_BreakPointToggle = VBInstance.Events.CommandBarEvents(lObjCommandBar)
                Case "EDIT"
                    'used to set breakpoints, only need one reference to this
                    If mcbMenuBreakPointToggle Is Nothing Then
                        Set mcbMenuBreakPointToggle = lObjCommandBar
                        mcbMenuBreakPointToggle.Enabled = True
                    End If
                    Set Subclass_Edit_BreakPointToggle = VBInstance.Events.CommandBarEvents(lObjCommandBar)
                Case "TOGGLE"
                    'used to set breakpoints, only need one reference to this
                    If mcbMenuBreakPointToggle Is Nothing Then
                        Set mcbMenuBreakPointToggle = lObjCommandBar
                        mcbMenuBreakPointToggle.Enabled = True
                    End If
                    Set Subclass_Toggle_BreakPointToggle = VBInstance.Events.CommandBarEvents(lObjCommandBar)
                Case "MENUBARDEBUG"
                    'MENU TITLE 'MENU BAR, SUB MENU DEBUG
                    Set Subclass_MenuMainDebug_BreakPointToggle = VBInstance.Events.CommandBarEvents(lObjCommandBar)
                    'used to set breakpoints, only need one reference to this
                    If mcbMenuBreakPointToggle Is Nothing Then
                        Set mcbMenuBreakPointToggle = lObjCommandBar
                        mcbMenuBreakPointToggle.Enabled = True
                    End If
            End Select
        Case EnmBreakPointRemoveALL
            'Can't get any
            Select Case (UCase$(pvStrMenuToSubClass))
            Case "MENUBARDEBUG"
                Set Subclass_MenuMainDebug_BreakPointRemoveAll = VBInstance.Events.CommandBarEvents(lObjCommandBar)
            End Select
        Case EnmBookmarkToggle
            Select Case (UCase$(pvStrMenuToSubClass))
                Case "EDIT"
                    'used to set bookmarks, only need one reference to this
                    If mcbMenuBookmarkToggle Is Nothing Then
                        Set mcbMenuBookmarkToggle = lObjCommandBar
                        mcbMenuBookmarkToggle.Enabled = True
                    End If
                    Set Subclass_Edit_BookmarkToggle = VBInstance.Events.CommandBarEvents(lObjCommandBar)
                    
                Case "TOGGLE"
                    'used to set bookmarks, only need one reference to this
                    If mcbMenuBookmarkToggle Is Nothing Then
                        Set mcbMenuBookmarkToggle = lObjCommandBar
                        mcbMenuBookmarkToggle.Enabled = True
                    End If
                    Set Subclass_Toggle_BookmarkToggle = VBInstance.Events.CommandBarEvents(lObjCommandBar)
                Case "BOOKMARKS"
                    'used to set bookmarks, only need one reference to this
                    If mcbMenuBookmarkToggle Is Nothing Then
                        Set mcbMenuBookmarkToggle = lObjCommandBar
                        mcbMenuBookmarkToggle.Enabled = True
                    End If
                    Set Subclass_MenuBookmarks_BookmarkToggle = VBInstance.Events.CommandBarEvents(lObjCommandBar)
            End Select
        Case EnmBookmarkremoveAll
            Select Case (UCase$(pvStrMenuToSubClass))
                Case "EDIT"
                    Set Subclass_Edit_BookmarkRemoveAll = VBInstance.Events.CommandBarEvents(lObjCommandBar)
                Case "BOOKMARKS"
                    Set Subclass_MenuBookmarks_BookmarkRemoveAll = VBInstance.Events.CommandBarEvents(lObjCommandBar)
            End Select
    End Select
    Exit Sub

mSub_SetEventToSink_ErrorHandler:

    If Err Then
        Resume Next
    End If
    
End Sub

Private Sub mSubSetAllSubClassItemsToNothing()
    
    Set Subclass_MenuDebug_BreakPointToggle = Nothing
    Set Subclass_MenuDebug_BreakPointRemoveAll = Nothing
    
    'Menu Title 'Edit'
    Set Subclass_Edit_BreakPointToggle = Nothing
    Set Subclass_Edit_BookmarkToggle = Nothing
    Set Subclass_Edit_BookmarkRemoveAll = Nothing
    
    'Menu Title 'Toggle'
    Set Subclass_Toggle_BreakPointToggle = Nothing
    Set Subclass_Toggle_BookmarkToggle = Nothing
    
    'MenuBookmarks
    Set Subclass_MenuBookmarks_BookmarkToggle = Nothing
    Set Subclass_MenuBookmarks_BookmarkRemoveAll = Nothing
    
    Set Subclass_VBProjectEvents = Nothing
    
    Set VBInstance = Nothing
    Set mcbMenuCommandBar = Nothing
    Set MenuHandler = Nothing

End Sub

Private Sub Subclass_Edit_BookmarkRemoveAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call mfLng_BookmarkRemoveAll(CommandBarControl, handled, CancelDefault)
End Sub

Private Sub Subclass_Edit_BookmarkToggle_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call mfLng_BookmarkToggle(CommandBarControl, handled, CancelDefault)
End Sub

Private Sub Subclass_Edit_BreakPointToggle_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call mfLng_BreakPointToggle(CommandBarControl, handled, CancelDefault)
End Sub

Private Sub Subclass_MenuBookmarks_BookmarkRemoveAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call mfLng_BookmarkRemoveAll(CommandBarControl, handled, CancelDefault)
End Sub

Private Sub Subclass_MenuBookmarks_BookmarkToggle_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call mfLng_BookmarkToggle(CommandBarControl, handled, CancelDefault)
End Sub

Private Sub Subclass_MenuDebug_BreakPointRemoveAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call mfLng_BreakPointRemoveAll(CommandBarControl, handled, CancelDefault)
End Sub

Private Sub Subclass_MenuDebug_BreakPointToggle_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call mfLng_BreakPointToggle(CommandBarControl, handled, CancelDefault)
End Sub

Private Sub Subclass_MenuMainDebug_BreakPointRemoveAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call mfLng_BreakPointRemoveAll(CommandBarControl, handled, CancelDefault)
End Sub

Private Sub Subclass_MenuMainDebug_BreakPointToggle_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call mfLng_BreakPointToggle(CommandBarControl, handled, CancelDefault)
End Sub

Private Sub Subclass_Toggle_BookmarkToggle_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call mfLng_BookmarkToggle(CommandBarControl, handled, CancelDefault)
End Sub

Private Sub Subclass_Toggle_BreakPointToggle_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call mfLng_BreakPointToggle(CommandBarControl, handled, CancelDefault)
End Sub

Private Sub Subclass_VBProjectEvents_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    
    mFrmCodeDawgToolbar.Timer1.Enabled = False
    'This means there were no projects and they just added 1
    If VBInstance.VBProjects.Count >= 1 Then
        If FormDisplayed = False Then
            Show
        End If
    End If
    mFrmCodeDawgToolbar.Timer1.Enabled = True
    
End Sub
Private Sub Subclass_VBProjectEvents_ItemRemoved(ByVal VBProject As VBIDE.VBProject)

    If GetSetting(App.Title, "Settings", "RmQSFProjectRm", True) = True Then
        Call mSubRemoveQSFFiles(VBProject)
    End If
    
    'menas they are removing the last project
    If VBInstance.VBProjects.Count = 1 Then
        Hide
        Set mClsBookMarks = Nothing
        Set mClsBreakPoints = Nothing
    End If
    
End Sub
Private Sub mSubRemoveQSFFiles(ByVal VBProject As VBIDE.VBProject)
    
On Error GoTo mSubRemoveQSFFiles_ErrorHandler:

    Dim lStrFile            As String
    Dim lColFilesToDelete   As New Collection
    Dim lStrFilePath        As String
    Dim lVarFile            As Variant
        
    lStrFilePath = Mid$(VBProject.FileName, 1, InStrRev(VBProject.FileName, "\")) & "QSF\"
    lStrFile = Dir$(lStrFilePath)
    Do Until lStrFile = ""
        'do not delete bookmarks, or breakpoints
        If Right$(lStrFile, 3) <> "BMK" And Right$(lStrFile, 3) <> "BKP" Then
            lColFilesToDelete.Add lStrFilePath & lStrFile
        End If
        lStrFile = Dir$
        DoEvents
    Loop
    For Each lVarFile In lColFilesToDelete
        Kill lVarFile
        DoEvents
    Next
    RmDir Mid$(VBProject.FileName, 1, InStrRev(VBProject.FileName, "\")) & "QSF"
    
    Exit Sub
    
mSubRemoveQSFFiles_ErrorHandler:
    
    If Err = 76 Or Err = 75 Then
        Resume Next
    ElseIf Err Then
        MsgBox "There was an error trying to remove 1 or more of the QSF files.  Please verify that the files located at: " & vbCr & "'" & _
               Mid$(VBProject.FileName, 1, InStrRev(VBProject.FileName, "\")) & "QSF\'" & " are not being used by any other applications." & vbCrLf & vbCr & _
               "The QSF Directory, and files still will remain after VB has closed!", vbInformation, "QSF Directory not deleted"
    End If
    
    Exit Sub

End Sub
Public Sub pSub_BreakPointKeyDownToggle()

On Error GoTo pSub_BreakPointKeyDownToggle_ErrorHandler:

    Dim lLng_StartPos   As Long
    Dim lLngDummy1      As Long
    Dim lLngDummy2      As Long
    Dim lLngDummy3      As Long
    Dim lStrColKey      As String
    
'    Static lStcBln_DidIt        As Boolean
    
    'Here becuase this method is envokle twice, toggle it
'    If lStcBln_DidIt = True Then
'        lStcBln_DidIt = False
'        Exit Sub
'    Else
'        lStcBln_DidIt = True
'    End If

    With VBInstance
        .ActiveCodePane.Window.SetFocus
        'we just need to current cursor position
        .ActiveCodePane.GetSelection lLng_StartPos, lLngDummy1, lLngDummy2, lLngDummy3
        
        lStrColKey = .ActiveVBProject.FileName & "@" & .ActiveCodePane.CodeModule & "@" & CStr(lLng_StartPos)
    
        If gfBln_DoesItemExistInCol(mClsBreakPoints, lStrColKey) Then
            mClsBreakPoints.Remove lStrColKey
            Debug.Print "Off"
        Else
            If mClsBreakPoints Is Nothing Then
                Set mClsBreakPoints = New ClsBreakPoints
            End If
            If mcbMenuBreakPointToggle.Enabled = True Then
                mClsBreakPoints.Add .ActiveVBProject.FileName, .ActiveVBProject.Name, .ActiveCodePane.CodeModule, lLng_StartPos, lStrColKey
            End If
        End If
    End With
    
    If Not mClsBreakPoints Is Nothing Then
        If mClsBreakPoints.Count = 0 Then
            Set mClsBreakPoints = Nothing
        End If
    End If
    VBInstance.ActiveCodePane.Window.SetFocus
    Exit Sub

pSub_BreakPointKeyDownToggle_ErrorHandler:

    If Err Then Resume Next

End Sub

Private Function mfLng_BreakPointToggle(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

On Error GoTo mfLng_BreakPointToggle_ErrorHandler:

    Dim lLng_StartPos   As Long
    Dim lLngDummy1      As Long
    Dim lLngDummy2      As Long
    Dim lLngDummy3      As Long
    Dim lStrColKey      As String

    With VBInstance
        .ActiveCodePane.Window.SetFocus
        'we just need to current cursor position
        .ActiveCodePane.GetSelection lLng_StartPos, lLngDummy1, lLngDummy2, lLngDummy3
        
        lStrColKey = .ActiveVBProject.FileName & "@" & .ActiveCodePane.CodeModule & "@" & CStr(lLng_StartPos)
    
        If gfBln_DoesItemExistInCol(mClsBreakPoints, lStrColKey) Then
            mClsBreakPoints.Remove lStrColKey
        Else
            If mClsBreakPoints Is Nothing Then
                Set mClsBreakPoints = New ClsBreakPoints
            End If
            If CommandBarControl.Enabled = True Then
                mClsBreakPoints.Add .ActiveVBProject.FileName, .ActiveVBProject.Name, .ActiveCodePane.CodeModule, lLng_StartPos, lStrColKey
            End If
        End If
    End With
    
    If Not mClsBreakPoints Is Nothing Then
        If mClsBreakPoints.Count = 0 Then
            Set mClsBreakPoints = Nothing
        End If
    End If
    VBInstance.ActiveCodePane.Window.SetFocus
    Exit Function
    
mfLng_BreakPointToggle_ErrorHandler:
    If Err Then Resume Next
    
End Function

Public Function mfInt_SaveBookMark_BreakPointFile(ByVal pvStrFileTypeExtenstion As String) As Integer

On Error GoTo mfInt_SaveBookMark_BreakPointFile_ErrorHandler:
    
    Dim lClsBookMark_BreakPoint As ClsBookmark_BreakPoint
    Dim lClsObj_SaveMe          As Object
    Dim lIntFileHandle          As Integer
    Dim lStrFileNames         As String
    
    If pvStrFileTypeExtenstion = gcStr_BookMarkFileExtension Then
        Set lClsObj_SaveMe = mClsBookMarks
    Else
        Set lClsObj_SaveMe = mClsBreakPoints
    End If
    
    If Not lClsObj_SaveMe Is Nothing Then
        For Each lClsBookMark_BreakPoint In lClsObj_SaveMe
            'kill file before starting
            If Not lClsBookMark_BreakPoint Is Nothing Then
                If InStr(lStrFileNames, lClsBookMark_BreakPoint.SaveFileName) = 0 Then
                    Kill lClsBookMark_BreakPoint.SaveFileName
                    lStrFileNames = lStrFileNames & "~~" & lClsBookMark_BreakPoint.SaveFileName
                    MkDir Mid$(lClsBookMark_BreakPoint.SaveFileName, 1, InStrRev(lClsBookMark_BreakPoint.SaveFileName, "\"))
                End If
                
                lIntFileHandle = FreeFile
                Open lClsBookMark_BreakPoint.SaveFileName For Append Lock Read Write As #lIntFileHandle
                Print #lIntFileHandle, lClsBookMark_BreakPoint.Key
                Close #lIntFileHandle
            End If
        Next
    End If
        
    On Error Resume Next
    If lClsObj_SaveMe.Count = 0 Then
        'remove old files is break points or bookmarks have been cleared
        Dim lVbProject As VBProject
        Dim lPrjIdx    As Integer
        For lPrjIdx = 1 To VBInstance.VBProjects.Count
            Set lVbProject = VBInstance.VBProjects(lPrjIdx)
            Kill Mid$(lVbProject.FileName, 1, InStrRev(lVbProject.FileName, "\")) & "QSF\" & lVbProject.Name & "." & pvStrFileTypeExtenstion
        Next
    End If
    On Error GoTo mfInt_SaveBookMark_BreakPointFile_ErrorHandler
    
    Set lClsBookMark_BreakPoint = Nothing
    Set lClsObj_SaveMe = Nothing
    
    Exit Function
    
mfInt_SaveBookMark_BreakPointFile_ErrorHandler:
    
    If Err = 53 Or Err = 76 Or Err = 75 Then  'ignore these
        Resume Next
    Else
        MsgBox "There was an error saving the Bookmark/Breakpoint file!" & vbCrLf & vbCr & "Error: " & Error, vbCritical, "Not Saved!"
        Resume Next
    End If
    
End Function

Private Function mfLng_BookmarkToggle(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
On Error GoTo mfLng_BookmarkToggle_ErrorHandler:
    
    Dim lLng_StartPos   As Long
    Dim lLngDummy1      As Long
    Dim lLngDummy2      As Long
    Dim lLngDummy3      As Long
    Dim lStrColKey      As String
    
    With VBInstance
        .ActiveCodePane.Window.SetFocus
        'we just need to current cursor position
        .ActiveCodePane.GetSelection lLng_StartPos, lLngDummy1, lLngDummy2, lLngDummy3
        
        lStrColKey = .ActiveVBProject & "@" & .ActiveCodePane.CodeModule & "@" & CStr(lLng_StartPos)
    
        If gfBln_DoesItemExistInCol(mClsBookMarks, lStrColKey) Then
            mClsBookMarks.Remove lStrColKey
        Else
            'init if nothing
            If mClsBookMarks Is Nothing Then
                Set mClsBookMarks = New ClsBookmarks
            End If
            If CommandBarControl.Enabled = True Then
                mClsBookMarks.Add .ActiveVBProject.FileName, .ActiveVBProject.Name, .ActiveCodePane.CodeModule, lLng_StartPos, lStrColKey
            End If
        End If
    End With
    
    If Not mClsBookMarks Is Nothing Then
        If mClsBookMarks.Count = 0 Then
            Set mClsBookMarks = Nothing
        End If
    End If
    VBInstance.ActiveCodePane.Window.SetFocus
    
    FormActivated = True
    
    Exit Function
    
mfLng_BookmarkToggle_ErrorHandler:
    If Err Then Resume Next
    
End Function

Private Function mfLng_BreakPointRemoveAll(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
On Error GoTo mfLng_BreakPointRemoveAll_ErrorHandler:
    
    Set mClsBreakPoints = Nothing
    
    Exit Function
    
mfLng_BreakPointRemoveAll_ErrorHandler:
    If Err Then Resume Next
    
End Function

Private Function mfLng_BookmarkRemoveAll(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
On Error GoTo mfLng_BookmarkRemoveAll_ErrorHandler:
    
    Set mClsBookMarks = Nothing
    
    Exit Function
mfLng_BookmarkRemoveAll_ErrorHandler:
    If Err Then Resume Next
    
End Function

Public Property Get VBWndHandle() As Long
    VBWndHandle = mLng_VBWndHandle
End Property

Public Property Let VBWndHandle(ByVal vNewValue As Long)
    mLng_VBWndHandle = vNewValue
End Property

Private Sub mSubGetVBHandle()
    
On Error GoTo mSubGetVBHandle_ErrorHandler:

    Dim lVbWndHandles() As Long
    Dim r As Long
    Dim lIntVBHwndCount As Integer
    
    'if the collection has not been set
    'create it, means it is the first one
    If gColVBHandles Is Nothing Then
       Set gColVBHandles = New Collection
    End If
    
    'get VB window
    r = FindWindowLike(lVbWndHandles(), 0, "*Microsoft Visual Basic*", "IDEOwner")
    
    If r > 0 Then
        'add the handles to our collection if we get more then one that indicates
        'that the user started up a copy of VB without the toolbar,
        'just take the first one that we don't have in our collection
        For lIntVBHwndCount = 0 To UBound(lVbWndHandles)
            'add the handle to our global collection if it is not there
            If lVbWndHandles(lIntVBHwndCount) > 0 Then
                If gfBln_DoesItemExistInCol(gColVBHandles, lVbWndHandles(lIntVBHwndCount)) = False Then
                    gColVBHandles.Add lVbWndHandles(lIntVBHwndCount), CStr(lVbWndHandles(lIntVBHwndCount))
                    'set property in connect object for the handle in the
                    'collection, use the handle as the key
                    Me.VBWndHandle = lVbWndHandles(lIntVBHwndCount)
                    Exit For
                End If
            End If
        Next
    End If
    
    Exit Sub
mSubGetVBHandle_ErrorHandler:

    If Err Then Resume Next
    
End Sub
Public Function mfInt_RetrieveBookmarks(Optional ByVal pvOptRetrieveAll As Boolean = True) As Integer
        
On Error GoTo mfInt_RetrieveBookmarks_ErrorHandler:

    Dim lIntCurProjectIndex As Integer
    
    For lIntCurProjectIndex = 1 To VBInstance.VBProjects.Count
        If pvOptRetrieveAll = True Then
            'add them all
            Call mSubRetrieveBookMark(lIntCurProjectIndex)
        Else
            'loop until we get the one that was just added
            If VBInstance.VBProjects.Item(lIntCurProjectIndex).Name = VBInstance.ActiveVBProject.Name Then
                Call mSubRetrieveBookMark(lIntCurProjectIndex)
            End If
        End If
        DoEvents
    Next
    
    Exit Function
mfInt_RetrieveBookmarks_ErrorHandler:

    If Err Then Resume Next
    Exit Function
    
End Function
Private Sub mSubRetrieveBookMark(ByVal pvIntProjectIndex As Integer)
    
On Error GoTo mSubRetrieveBookMark_ErrorHandler:

    Dim lVbProject          As VBProject
    Dim lStrFileName        As String
    Dim lIntFileHandle      As Integer
    Dim lStrDataIn          As String
    Dim lVarParsedData      As Variant
    Dim lColOpenWindows     As New Collection
    
    'get vb project Reference
    Set lVbProject = VBInstance.VBProjects(pvIntProjectIndex)
    lIntFileHandle = FreeFile
    lStrFileName = Mid$(lVbProject.FileName, 1, InStrRev(lVbProject.FileName, "\")) & "QSF\" & lVbProject.Name & "." & gcStr_BookMarkFileExtension
    Open lStrFileName For Input Lock Read Write As #lIntFileHandle
    Do Until EOF(lIntFileHandle)
        Line Input #lIntFileHandle, lStrDataIn
        If Trim$(lStrDataIn) <> "" Then
            'process data
            If gfBln_DoesItemExistInCol(mClsBookMarks, lStrDataIn) = False Then
                lVarParsedData = Split(Trim$(lStrDataIn), "@")
                lVbProject.VBComponents(CStr(lVarParsedData(1))).CodeModule.CodePane.Window.Visible = True
                DoEvents
                lVbProject.VBComponents(CStr(lVarParsedData(1))).CodeModule.CodePane.Window.SetFocus
                DoEvents
                lVbProject.VBComponents(CStr(lVarParsedData(1))).CodeModule.CodePane.SetSelection CLng(lVarParsedData(2)), 1, CLng(lVarParsedData(2)), 1
                DoEvents
                lVbProject.VBComponents(CStr(lVarParsedData(1))).CodeModule.CodePane.Window.SetFocus
                DoEvents
                If mcbMenuBookmarkToggle.Enabled = True Then
                    mcbMenuBookmarkToggle.Execute
                End If
                DoEvents
                lVbProject.VBComponents(CStr(lVarParsedData(1))).CodeModule.CodePane.Window.SetFocus
                DoEvents
            End If
        Else
            Exit Do
        End If
        If Err Then Exit Do
        DoEvents
    Loop
    Close #lIntFileHandle
    
    '990922 SBI0 Added to kill old data
    Kill lStrFileName
    
    Exit Sub
mSubRetrieveBookMark_ErrorHandler:
    If Err Then Resume Next
    Exit Sub
End Sub

Public Function mfInt_RetrieveBreakpoints(Optional ByVal pvOptRetrieveAll As Boolean = True) As Integer
        
On Error GoTo mfInt_RetrieveBreakpoints_ErrorHandler:

    Dim lIntCurProjectIndex As Integer
    
    For lIntCurProjectIndex = 1 To VBInstance.VBProjects.Count
        If pvOptRetrieveAll = True Then
            'add them all
            Call mSubRetrieveBreakpoint(lIntCurProjectIndex)
        Else
            'loop until we get the one that was just added
            If VBInstance.VBProjects.Item(lIntCurProjectIndex).Name = VBInstance.ActiveVBProject.Name Then
                Call mSubRetrieveBreakpoint(lIntCurProjectIndex)
            End If
        End If
        DoEvents
    Next
    
    Exit Function
mfInt_RetrieveBreakpoints_ErrorHandler:

    If Err Then Resume Next
    Exit Function
    
End Function

Private Sub mSubRetrieveBreakpoint(ByVal pvIntProjectIndex As Integer)
    
On Error GoTo mSubRetrieveBreakpoint_ErrorHandler:

    Dim lVbProject          As VBProject
    Dim lStrFileName        As String
    Dim lIntFileHandle      As Integer
    Dim lStrDataIn          As String
    Dim lVarParsedData      As Variant

    'get vb project Reference
    Set lVbProject = VBInstance.VBProjects(pvIntProjectIndex)
    lIntFileHandle = FreeFile
    lStrFileName = Mid$(lVbProject.FileName, 1, InStrRev(lVbProject.FileName, "\")) & "QSF\" & lVbProject.Name & "." & gcStr_BreakpointFileExtension
    Open lStrFileName For Input Lock Read Write As #lIntFileHandle
    Do Until EOF(lIntFileHandle)
        Line Input #lIntFileHandle, lStrDataIn
        If Trim$(lStrDataIn) <> "" Then
            'process data
            If gfBln_DoesItemExistInCol(mClsBreakPoints, lStrDataIn) = False Then
                lVarParsedData = Split(Trim$(lStrDataIn), "@")
                lVbProject.VBComponents(CStr(lVarParsedData(1))).CodeModule.CodePane.Window.Visible = True
                DoEvents
                lVbProject.VBComponents(CStr(lVarParsedData(1))).CodeModule.CodePane.Window.SetFocus
                DoEvents
                lVbProject.VBComponents(CStr(lVarParsedData(1))).CodeModule.CodePane.SetSelection CLng(lVarParsedData(2)), 1, CLng(lVarParsedData(2)), 1
                DoEvents
                lVbProject.VBComponents(CStr(lVarParsedData(1))).CodeModule.CodePane.Window.SetFocus
                DoEvents
                If mcbMenuBreakPointToggle.Enabled = True Then
                    mcbMenuBreakPointToggle.Execute
                End If
                DoEvents
                lVbProject.VBComponents(CStr(lVarParsedData(1))).CodeModule.CodePane.Window.SetFocus
                DoEvents
            End If
        Else
            Exit Do
        End If
        If Err Then Exit Do
        DoEvents
    Loop
    Close #lIntFileHandle
    
    '990922 SBI0 Added to kill old data
    Kill lStrFileName

    Exit Sub
    
mSubRetrieveBreakpoint_ErrorHandler:
    If Err Then Resume Next
    Exit Sub
End Sub


Public Property Get FormActivated() As Boolean
    FormActivated = mvarFormActivated
End Property

Public Property Let FormActivated(ByVal vNewValue As Boolean)
    mvarFormActivated = vNewValue
End Property

Public Sub pSub_LineChange(ByVal pvBln_AddLines As Boolean, ByVal pvLngCurLineNumber As Long, ByVal pvLngNumLinesToChange As Long)

On Error GoTo pSub_LineChange_ErrorHandler:

    Dim lobjItems               As ClsBookmark_BreakPoint
    Dim lclsBookmarks           As New ClsBookmarks
    Dim lclsBreakpoints         As New ClsBreakPoints
    Dim lColKeysToRemove        As Collection
    Dim lLngNewLineNumber       As Long
    Dim lVarItemsToRemove       As Variant
    
    If mClsBookMarks Is Nothing And mClsBreakPoints Is Nothing Then
        Exit Sub
    End If
    
    If mClsBookMarks.Count = 0 And mClsBreakPoints.Count = 0 Then
        Exit Sub
    End If
    
''''''  BOOKMARKS STARTS HERE

    For Each lobjItems In mClsBookMarks
        If pvBln_AddLines = True Then
            If lobjItems.LineNumber >= pvLngCurLineNumber - pvLngNumLinesToChange Then
                lLngNewLineNumber = lobjItems.LineNumber + pvLngNumLinesToChange
            End If
        Else
            If lobjItems.LineNumber >= pvLngCurLineNumber + pvLngNumLinesToChange Then
                lLngNewLineNumber = lobjItems.LineNumber - pvLngNumLinesToChange
            End If
        End If
        
        If lLngNewLineNumber <> 0 Then
        
            lclsBookmarks.Add lobjItems.ProjectFileName, _
                  lobjItems.ProjectName, _
                  lobjItems.ModuleName, _
                  lLngNewLineNumber, _
                  Left$(lobjItems.Key, InStrRev(lobjItems.Key, "@")) & CStr(lLngNewLineNumber)
            
            If lColKeysToRemove Is Nothing Then
                Set lColKeysToRemove = New Collection
            End If
            lColKeysToRemove.Add lobjItems.Key
        End If
        lLngNewLineNumber = 0
    Next
    Set lobjItems = Nothing
    
    If Not lColKeysToRemove Is Nothing Then
        'remove all old entries
        If lColKeysToRemove.Count > 0 Then
            For Each lVarItemsToRemove In lColKeysToRemove
                mClsBookMarks.Remove CStr(lVarItemsToRemove)
            Next
        End If
        
        If lclsBookmarks.Count > 0 Then
            For Each lobjItems In lclsBookmarks
                mClsBookMarks.Add lobjItems.ProjectFileName, _
                                lobjItems.ProjectName, _
                                lobjItems.ModuleName, _
                                lobjItems.LineNumber, _
                                lobjItems.Key
            Next
            Set lobjItems = Nothing
        End If
    End If
    Set lclsBookmarks = Nothing
'''''''''  BOOKMARKS ENDS HERE
    
    'Kill Old Data
    Set lColKeysToRemove = Nothing
    Set lColKeysToRemove = New Collection
    lLngNewLineNumber = 0
    
''''''''' BREAKPOINTS STARTS HERE
    For Each lobjItems In mClsBreakPoints
        If pvBln_AddLines = True Then
            If lobjItems.LineNumber >= pvLngCurLineNumber - pvLngNumLinesToChange Then
                lLngNewLineNumber = lobjItems.LineNumber + pvLngNumLinesToChange
            End If
        Else
            If lobjItems.LineNumber >= pvLngCurLineNumber + pvLngNumLinesToChange Then
                lLngNewLineNumber = lobjItems.LineNumber - pvLngNumLinesToChange
            End If
        End If
       
        If lLngNewLineNumber <> 0 Then
        
            lclsBreakpoints.Add lobjItems.ProjectFileName, _
                  lobjItems.ProjectName, _
                  lobjItems.ModuleName, _
                  lLngNewLineNumber, _
                  Left$(lobjItems.Key, InStrRev(lobjItems.Key, "@")) & CStr(lLngNewLineNumber)
            
            If lColKeysToRemove Is Nothing Then
                Set lColKeysToRemove = New Collection
            End If
            lColKeysToRemove.Add lobjItems.Key
        End If
        lLngNewLineNumber = 0
    Next
    Set lobjItems = Nothing
    
    Debug.Print lLngNewLineNumber
    
    If Not lColKeysToRemove Is Nothing Then
        'remove all old entries
        If lColKeysToRemove.Count > 0 Then
            For Each lVarItemsToRemove In lColKeysToRemove
                mClsBreakPoints.Remove CStr(lVarItemsToRemove)
            Next
        End If
        
        If lclsBreakpoints.Count > 0 Then
            For Each lobjItems In lclsBreakpoints
                mClsBreakPoints.Add lobjItems.ProjectFileName, _
                                lobjItems.ProjectName, _
                                lobjItems.ModuleName, _
                                lobjItems.LineNumber, _
                                lobjItems.Key
            Next
            Set lobjItems = Nothing
        End If
    End If
    Set lclsBreakpoints = Nothing
    
''''''''' BREAKPOINTS ENDS HERE
    
    Exit Sub
    
pSub_LineChange_ErrorHandler:
    
    'SBI0 990911 Change form exiting on error to resume next
    If Err Then Resume Next
    
End Sub

Public Property Get TotalNumLines() As Long
    TotalNumLines = mLngTotalNumLines
End Property

Public Property Let TotalNumLines(ByVal vNewValue As Long)
    mLngTotalNumLines = vNewValue
End Property
Public Sub pSub_BreakPointRemoveAllKeyDown()

On Error GoTo pSub_BreakPointRemoveAllKeyDown:
    
    Set mClsBreakPoints = Nothing
    
    Exit Sub
    
pSub_BreakPointRemoveAllKeyDown:
    If Err Then Resume Next

End Sub

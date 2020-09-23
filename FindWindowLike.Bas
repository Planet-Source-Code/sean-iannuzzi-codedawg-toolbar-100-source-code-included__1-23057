Attribute VB_Name = "basFindWindow"
Option Explicit
Global gColVBHandles As Collection


Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_RESETCONTENT = &H14B

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
   (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
   
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function GetWindow Lib "user32" _
   (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
   (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
   (ByVal hwnd As Long, ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Long) As Long
    
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As String) As Long

Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

'New for Keyboard Hooks
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Function GetComboRefFromFindWindow(ByVal hWndStart As Long) As Long
   
    Dim hwnd            As Long
    Dim sClassname      As String
    Dim r               As Long
    Dim lBlnFoundIt     As Boolean
    Dim lStrData        As String
    Dim lPos            As Long
    Dim lLngItemCount   As Long
    Dim sWindowText As String
    'need to start the window parent as VB
    
    'Get first child window
    hwnd = GetWindow(hWndStart, GW_CHILD)

    'Search children by recursion
    Do Until hwnd = 0
        
        sClassname = Space(255)
        r = GetClassName(hwnd, sClassname, 255)
        sClassname = Left(sClassname, r)
        
               'Get the window text and class name
        sWindowText = Space(255)
        r = GetWindowText(hwnd, sWindowText, 255)
        sWindowText = Left(sWindowText, r)
            
        If sClassname = "ComboBox" Then
            lBlnFoundIt = True
            'get the handle to the edit portion  'of the combo control
            'check value could be wrong combo
            lLngItemCount = SendMessageLong(hwnd, CB_GETCOUNT, 0&, 0&)
            If lLngItemCount = 3 Then
                lStrData = ""
                lStrData = Space$(255)
                r = SendMessageStr(hwnd, CB_GETLBTEXT, 0&, ByVal lStrData)
                If mfStrStripSpaces(lStrData) = "All" Then
                    lStrData = ""
                    lStrData = Space$(255)
                    r = SendMessageStr(hwnd, CB_GETLBTEXT, 1&, ByVal lStrData)
                    If mfStrStripSpaces(lStrData) = "Down" Then
                        lStrData = ""
                        lStrData = Space$(255)
                        r = SendMessageStr(hwnd, CB_GETLBTEXT, 2&, ByVal lStrData)
                        If mfStrStripSpaces(lStrData) = "Up" Then
                            'wrong one
                            lBlnFoundIt = False
                        End If
                    End If
                End If
                
                If lBlnFoundIt = True Then
                    GetComboRefFromFindWindow = hwnd
                End If
            Else
                'It the one we want
                GetComboRefFromFindWindow = hwnd
                lBlnFoundIt = True
            End If
        End If
        
        'if we found it then exit the loop
        If lBlnFoundIt = True Then
            Exit Do
        End If
        
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        DoEvents
    Loop

End Function
Private Function mfStrStripSpaces(ByVal pvStrIn As String) As String

    Dim lPos As Long
    lPos = InStr(pvStrIn, Chr$(0))
    If lPos Then
        mfStrStripSpaces = Left$(pvStrIn, InStr(pvStrIn, Chr$(0)) - 1)
    End If
    
End Function
Function FindWindowLike(hWndArray() As Long, _
                        ByVal hWndStart As Long, _
                        WindowText As String, _
                        Classname As String, Optional ByVal pvBln_StartOver As Boolean = False) As Long
   
On Error GoTo FindWindowLike_ErrorHandler:

    Dim hwnd As Long
    Dim sWindowText As String
    Dim sClassname As String
    Dim r As Long
      
   'Hold the level of recursion and
   'hold the number of matching windows
    Static level As Long
    Static found As Long
  
   'Initialize if necessary
    If level = 0 Then
      found = 0
      ReDim hWndArray(0 To 0)
      If hWndStart = 0 Then hWndStart = GetDesktopWindow()
    End If
  
   'Increase recursion counter
    level = level + 1
  
   'Get first child window
    hwnd = GetWindow(hWndStart, GW_CHILD)

    Do Until hwnd = 0
      
       'Search children by recursion
        Call FindWindowLike(hWndArray(), hwnd, WindowText, Classname)
      
       'Get the window text and class name
        sWindowText = Space(255)
        r = GetWindowText(hwnd, sWindowText, 255)
        If r > 0 Then
            sWindowText = Left(sWindowText, r)
            
            sClassname = Space(255)
            r = GetClassName(hwnd, sClassname, 255)
            sClassname = Left(sClassname, r)
          
           'Check that window matches the search parameters
            If (sWindowText Like WindowText) And (sClassname Like Classname) Then
                found = found + 1
                ReDim Preserve hWndArray(0 To found)
                hWndArray(found) = hwnd
            End If
        End If
       'Get next child window
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        
    Loop
  
   'Decrement recursion counter
    level = level - 1
  
   'Return the number of windows found
    FindWindowLike = found
    
    Exit Function

FindWindowLike_ErrorHandler:

    If Err Then Resume Next

End Function

Function gfBln_DoesItemExistInCol(ByVal pvColToCheck As Object, ByVal pvVarKey As Variant) As Boolean
    On Error Resume Next
    Dim lVar As Variant
    lVar = TypeName(pvColToCheck.Item(pvVarKey))
    If Err = 0 Then
        gfBln_DoesItemExistInCol = True
    End If
    On Error GoTo 0
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBookmarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClsBookmark_BreakPoint"
Attribute VB_Ext_KEY = "Member0" ,"ClsBookmark_BreakPoint"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
DefLng A-Z

Private mCol As Collection

Public Function Add(ProjectFileName As String, ProjectName As String, ModuleName As String, LineNumber As Long, Optional sKey As String) As ClsBookmark_BreakPoint
    
    'create a new object
    Dim objNewMember As ClsBookmark_BreakPoint
    Set objNewMember = New ClsBookmark_BreakPoint

    'set the properties passed into the method
    objNewMember.ProjectName = ProjectName
    objNewMember.ModuleName = ModuleName
    objNewMember.LineNumber = LineNumber
    objNewMember.Key = sKey
    objNewMember.ProjectFileName = ProjectFileName
    objNewMember.SaveFileName = Mid$(ProjectFileName, 1, InStrRev(ProjectFileName, "\")) & "QSF\" & ProjectName & "." & gcStr_BookMarkFileExtension
        
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClsBookmark_BreakPoint
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
     Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBookmark_BreakPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
DefLng A-Z

Private mvarProjectName As String
Private mvarModuleName  As String
Private mvarLineNumber  As Long
Private mvarProjectFilename As String
Private mvarSaveFileName As String

Public Key              As Variant


Public Property Let LineNumber(ByVal vData As Long)
    mvarLineNumber = vData
End Property

Public Property Get LineNumber() As Long
    LineNumber = mvarLineNumber
End Property

Public Property Let ModuleName(ByVal vData As String)
    mvarModuleName = vData
End Property

Public Property Get ModuleName() As String
    ModuleName = mvarModuleName
End Property

Public Property Let ProjectName(ByVal vData As String)
    mvarProjectName = vData
End Property

Public Property Get ProjectName() As String
    ProjectName = mvarProjectName
End Property

Public Property Get ProjectFileName() As String
    ProjectFileName = mvarProjectFilename
End Property

Public Property Let ProjectFileName(ByVal vNewValue As String)
    mvarProjectFilename = vNewValue
End Property

Public Property Get SaveFileName() As String
    SaveFileName = mvarSaveFileName
End Property

Public Property Let SaveFileName(ByVal vNewValue As String)
    mvarSaveFileName = vNewValue
End Property

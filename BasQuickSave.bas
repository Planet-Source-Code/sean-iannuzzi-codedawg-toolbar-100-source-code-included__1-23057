Attribute VB_Name = "BasQuickSaveRoutine"
Option Explicit
DefInt A-Z

'constants for save options
Global Const gcIntSaveAllFileNoPrompt = 0
Global Const gcIntSaveactiveCodePaineOnly = 1
Global Const gcIntSaveReadOnlyFiles = 0
Global Const gcIntSaveReadOnlyPrompt = 1

Global Const gcIntRemoveQSFWhenProjectRemoved = 0
Global Const gcIntAlwaysRemoveQSF = 1
Global Const gcIntMinToolMainMin = 1

Global Const gcIntAlwaysOnTop = 0

Global Const gcStr_BookMarkFileExtension = "BMK"
Global Const gcStr_BreakpointFileExtension = "BKP"

Global Const gcInt_ChkAutoSaveOptions = 0
Global Const gcInt_ChkAutoSaveOptions_BKM = 2
Global Const gcInt_ChkAutoSaveOptions_BKP = 3
Global Const gcInt_ChkAutoSaveOptions_AUTOSAVE = 1

Global Const ChkStartupOptions_BKM = 1
Global Const ChkStartupOptions_BKP = 2

Global Const gcIntMaxNumFindSaved = 15

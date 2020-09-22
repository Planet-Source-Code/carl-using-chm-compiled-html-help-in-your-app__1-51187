Attribute VB_Name = "mdlSupportHTMLHelp"
'Note: to use these functions the projects help file must be set
'up correctly (project-properties-help file name). This must point to
'a valid HTML help file (*.chm)
'There may be some incompatibilities on systems using IE version less than 5.5
Option Explicit

' UDT for accessing the Search tab
Private Type t_HH_Search
  lSzStruct          As Long
  lUnicodeStrings   As Long
  sSearchQuery      As String
  lProximity        As Long
  lStemmedSearch    As Long
  lTitleOnly        As Long
  lExecute          As Long
  sWindow         As String
End Type

' HTML Help Constants
Private Const HH_DISPLAY_TOPIC = &H0            ' WinHelp equivalent
Private Const HH_DISPLAY_TOC = &H1              ' WinHelp equivalent
Private Const HH_DISPLAY_INDEX = &H2            ' WinHelp equivalent
Private Const HH_DISPLAY_SEARCH = &H3           ' WinHelp equivalent
Private Const HH_KEYWORD_LOOKUP = &HD           ' WinHelp equivalent
Private Const HH_HELP_CONTEXT = &HF             ' WinHelp equivalent
Private Const HH_CLOSE_ALL = &H12               ' WinHelp equivalent

' HTML Help API declarations
Private Declare Function HTMLHelp Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As Long) As Long
    
Private Declare Function HTMLHelpCallSearch Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByRef dwData As t_HH_Search) As Long
    
'/=============================================================================
' Name:     HTMLHelp_Contents
' Purpose:  Displays HTML Help contents
' Level:    0
' History:
'\=============================================================================
Public Sub HTMLHELP_Contents(f As Form)

    Dim hwndHelp As Long
    
    hwndHelp = HTMLHelp(f.hwnd, App.HelpFile, HH_DISPLAY_TOC, 0)
    
End Sub

'/=============================================================================
' Name:     HTMLHelp_ContextID
' Purpose:  Displays HTML Help for item with related ContextID
' Level:    0
' History:
'\=============================================================================
Public Sub HTMLHELP_ContextID(f As Form, ContextID As Long)

    Dim hwndHelp As Long
    
    hwndHelp = HTMLHelp(f.hwnd, App.HelpFile, HH_HELP_CONTEXT, ContextID)
    
End Sub

'/=============================================================================
' Name:     HTMLHelp_Search
' Purpose:  Displays HTML Search Tab
' Level:    0
' History:
'\=============================================================================
Public Sub HTMLHelp_Search(f As Form)

    Dim hwnd As Long

    Dim HH_Search As t_HH_Search

    With HH_Search
        .lSzStruct = Len(HH_Search)
        .lUnicodeStrings = 0&
        .sSearchQuery = ""
        .lProximity = 0&
        .lStemmedSearch = 0&
        .lTitleOnly = 0&
        .lExecute = 1&
        .sWindow = ""
    End With
    
    hwnd = HTMLHelpCallSearch(f.hwnd, App.HelpFile, HH_DISPLAY_SEARCH, HH_Search)
      
End Sub



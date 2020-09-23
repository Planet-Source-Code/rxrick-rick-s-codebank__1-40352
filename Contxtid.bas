Attribute VB_Name = "ContextIDs"
'
' This BAS module was created by VB Helpwriter:
'     http://www.helpwriter.com/
'
' I highly recommend this program, it is an IDE for creating help files.
' VB Helpwrite is shareware and well worth the modest purchase price.
'     -- Rick
'
'
Option Explicit
'=====================================================================
'=====================================================================
'
'This source code contains the following routines:
'  o SetAppHelp() 'Called in the main Form_Load event to register your
'                 'program with WINHELP.EXE
'  o QuitHelp()    'Deregisters your program with WINHELP.EXE. Should
'                  'be called in your main Form_Unload event
'  o ShowHelpTopic(Topicnum) 'Brings up context sensitive help based on
'                  'any of the following CONTEXT IDs
'  o ShowContents  'Displays the startup topic
'  o HelpWindowSize(x,y,dx,dy) ' Position help window in a screen
'                              ' independent manner
'  o SearchHelp()  'Brings up the windows help KEYWORD SEARCH dialog box
'***********************************************************************
'
'=====================================================================
'List of Context IDs for <codebank>
'=====================================================================
Global Const Hlp_VB_Code = 20    'Main Help Window
'=====================================================================
'
'
'  Help engine section.

' Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
Global Const HELP_QUIT = &H2    '  Terminate help
Global Const HELP_FINDER = &HB  '  Display Contents tab
Global Const HELP_INDEX = &H3   '  Display index
Global Const HELP_HELPONHELP = &H4      '  Display help on using help
Global Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
Global Const HELP_KEY = &H101           '  Display topic for keyword in offabData
Global Const HELP_MULTIKEY = &H201
Global Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
Global Const HELP_COMMAND = &H102    ' Execute Help macro
Global Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
Global Const HELP_SETWINPOS = &H203  ' Display and position Help window

#If Win32 Then
    Type HELPWININFO
      wStructSize As Long
      x As Long
      y As Long
      dX As Long
      dY As Long
      wMax As Long
      rgChMember As String * 2
    End Type
    Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
    Declare Function WinHelpByInfo Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As HELPWININFO) As Long
    Declare Function WinHelpByStr Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData$) As Long
    Declare Function WinHelpByNum Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData&) As Long
    Dim m_hWndMainWindow As Long ' hWnd to tell WINHELP the helpfile owner

#Else
    Type HELPWININFO
        wStructSize As Integer
        x As Integer
        y As Integer
        dX As Integer
        dY As Integer
        wMax As Integer
        rgChMember As String * 2
    End Type
    Declare Function WinHelp Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Any) As Integer
    Declare Function WinHelpByInfo Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As HELPWININFO) As Integer
    Declare Function WinHelpByStr Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$) As Integer
    Declare Function WinHelpByNum Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&) As Integer
    Dim m_hWndMainWindow As Integer ' hWnd to tell WINHELP the helpfile owner

#End If
Dim MainWindowInfo As HELPWININFO
Sub SetAppHelp(ByVal hWndMainWindow)
'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================
    m_hWndMainWindow = hWndMainWindow
    If Right$(Trim$(App.Path), 1) = "\" Then
        App.HelpFile = App.Path + "codebank.HLP"
    Else
        App.HelpFile = App.Path + "\codebank.HLP"
    End If
#If Win32 Then
    MainWindowInfo.wStructSize = 26
#Else
    MainWindowInfo.wStructSize = 14
#End If
    MainWindowInfo.x = 256
    MainWindowInfo.y = 256
    MainWindowInfo.dX = 512
    MainWindowInfo.dY = 512
    MainWindowInfo.rgChMember = Chr$(0) + Chr$(0)
End Sub
Sub QuitHelp()
    Dim Result As Variant
    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_QUIT, Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0))
End Sub
Sub ShowHelpTopic(ByVal ContextID As Long)
'=====================================================================
'  FOR CONTEXT SENSITIVE HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic(<any Hlpxxx entry above>)
'=====================================================================
'  TO ADD FORM LEVEL CONTEXT SENSITIVE HELP...
'=====================================================================
'     o  For FORM level context sensetive help, you should set each
'        Me.HelpContext=<any Hlp_xxx entry above>
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXT, CLng(ContextID))

End Sub
Sub ShowHelpTopic2(ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 2 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic2(<any Hlpxxx entry above>)
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile & ">HlpWnd02", HELP_CONTEXT, CLng(ContextID))

End Sub
Sub ShowHelpTopic3(ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 3 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic3(<any Hlpxxx entry above>)
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile & ">HlpWnd03", HELP_CONTEXT, CLng(ContextID))

End Sub
Sub ShowGlossary()
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXT, CLng(64000))

End Sub
Sub ShowPopupHelp(ByVal ContextID As Long)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXTPOPUP, CLng(ContextID))

End Sub
Sub DoHelpMacro(ByVal MacroString As String)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result As Variant

    Result = WinHelpByStr(m_hWndMainWindow, App.HelpFile, HELP_COMMAND, ByVal (MacroString))

End Sub
Sub ShowHelpContents()
'=====================================================================
'  DISPLAY STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTENTS, CLng(0))

End Sub
Sub ShowContentsTab()
'=====================================================================
'  DISPLAY Contents tab (*.CNT)
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_FINDER, CLng(0))

End Sub
Sub ShowHelpOnHelp()
'=====================================================================
'  DISPLAY HELP for WINHELP.EXE  ...
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_HELPONHELP, CLng(0))

End Sub

Sub SearchHelp()
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result As Variant

    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_PARTIALKEY, ByVal "")

End Sub

Sub SearchHelpKeyWord(Argument As String)
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result As Variant

    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_PARTIALKEY, ByVal Trim$(Argument))

End Sub
Sub HelpWindowSize(x As Integer, y As Integer, wx As Integer, wy As Integer)
'=====================================================================
'  TO SET THE SIZE AND POSITION OF THE MAIN HELP WINDOW...
'=====================================================================
'     o   Call HelpWindowSize(x, y, dx, dy), where:
'             x = 1-1024 (position from left edge of screen)
'             y = 1-1024 (position from top of screen)
'             dx= 1-1024 (width)
'             dy= 1-1024 (height)
'
    Dim Result As Variant
    MainWindowInfo.x = x
    MainWindowInfo.y = y
    MainWindowInfo.dX = wx
    MainWindowInfo.dY = wy
    Result = WinHelpByInfo(m_hWndMainWindow, App.HelpFile, HELP_SETWINPOS, MainWindowInfo)
End Sub

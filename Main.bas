Attribute VB_Name = "Main"
Option Explicit

'Declare Global function
'Get color of active window
 Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'Declare global constant
'Color of active window
 Public Const COLOR_ACTIVECAPTION = 2

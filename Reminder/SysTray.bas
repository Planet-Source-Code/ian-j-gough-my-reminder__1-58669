Attribute VB_Name = "SysTray"
'Taken from http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=11893&lngWId=1
'Please vote for the above code if you make use of this function
'
'
'
'The below is for the System Tray Declarations (Don't worry about it!, Just have a look)
'
Public INTRAY As Boolean 'Boolean value to detect App Status[Max or Min] Remember Boolean can be only 2 states true or false!
'
'Declare Tray Icon
Type NOTIFYICONDATA
     cbSize As Long
     hwnd As Long
     uID As Long
     uFlags As Long
     uCallbackMessage As Long
     hIcon As Long
     szTip As String * 64
End Type
'
'Tray Return values
Public Const trayLBUTTONDOWN = 7695 'Left mouse button down
Public Const trayLBUTTONUP = 7710 'Left mouse button up
Public Const trayLBUTTONDBLCLK = 7725 'Left mouse button double click
'
Public Const trayRBUTTONDOWN = 7740 'As above but right instead of left
Public Const trayRBUTTONUP = 7755
Public Const trayRBUTTONDBLCLK = 7770
'
Public Const trayMOUSEMOVE = 7680 'When you move mouse over tray icon ( It helps it know when to show tool tip)
'
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_LBUTTONDBLCLK = &H203
'
Global Const NIM_ADD = &H0& 'Constants & flags for NotifyIcons
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const WM_MOUSEMOVE = &H200
'
Global NI As NOTIFYICONDATA
'
'The API for the System tray
Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'

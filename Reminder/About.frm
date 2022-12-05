VERSION 5.00
Begin VB.Form form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "About.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   3075
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   "K.A.D. Software VB Reminder.  Made by Ian.J.Gough for PlanetSourceCode"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   3
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: This program is free to distribute and under no circumstances is to be sold!"
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   4
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Taken from visual basic 6.0 professional add in form
Option Explicit
'
' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
'
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number
'
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
'
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'
Private Sub cmdSysInfo_Click()
100 On Error GoTo LocalErrors
110  Call StartSysInfo
120 Exit Sub
130 LocalErrors:
140 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '024' and program 'Reminder Service'", vbExclamation, "Error Control"
150 Unload Me
End Sub
'
Private Sub cmdOK_Click()
100 On Error GoTo LocalErrors
110   Unload Me
120 Form1.Enabled = True
130 Form1.Show
140 Exit Sub
150 LocalErrors:
160 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '025' and program 'Reminder Service'", vbExclamation, "Error Control"
170 Unload Me
End Sub
'
Private Sub Form_Load()
100 On Error GoTo LocalErrors
110    Me.Caption = "About " & App.Title
120    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
130    lblTitle.Caption = App.Title
140 Exit Sub
150 LocalErrors:
160 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '026' and program 'Reminder Service'", vbExclamation, "Error Control"
170 Unload Me
End Sub
'
Public Sub StartSysInfo()
100    On Error GoTo SysInfoErr
110
120    Dim rc As Long
130    Dim SysInfoPath As String
'
    ' Try To Get System Info Program Path\Name From Registry...
140    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
150    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
160        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
170            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
'
        ' Error - File Can Not Be Found...
180        Else
190            GoTo SysInfoErr
200        End If
    ' Error - Registry Entry Can Not Be Found...
210    Else
220        GoTo SysInfoErr
230    End If
'
240    Call Shell(SysInfoPath, vbNormalFocus)
'
250    Exit Sub
260 SysInfoErr:
270    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub
'
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
100 On Error GoTo LocalErrors
110    Dim i As Long                                           ' Loop Counter
120    Dim rc As Long                                          ' Return Code
130    Dim hKey As Long                                        ' Handle To An Open Registry Key
140    Dim hDepth As Long                                      '
150    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
160    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
170    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
180    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
'
190    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
'
200    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
210    KeyValSize = 1024                                       ' Mark Variable Size
'
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
220    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
'
230    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
'
240    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
250        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
260    Else                                                    ' WinNT Does NOT Null Terminate String...
270        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
280    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
'
290
    Select Case KeyValType                                  ' Search Data Types...

    Case REG_SZ                                             ' String Registry Key Data Type
300        KeyVal = tmpVal                                     ' Copy String Value
310    Case REG_DWORD                                          ' Double Word Registry Key Data Type
320        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
330            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
340        Next
350        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
360    End Select
'
370    GetKeyValue = True                                      ' Return Success
380    rc = RegCloseKey(hKey)                                  ' Close Registry Key
390    Exit Function                                           ' Exit
    
400 GetKeyError:     ' Cleanup After An Error Has Occured...
410    KeyVal = ""                                             ' Set Return Val To Empty String
420    GetKeyValue = False                                     ' Return Failure
430    rc = RegCloseKey(hKey) ' Close Registry Key
440 Exit Function
450 LocalErrors:
460 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '027' and program 'Reminder Service'", vbExclamation, "Error Control"
470 Unload Me
End Function
'
'
Private Sub lblDescription_Click()
100
End Sub

Private Sub lblVersion_Click()
100
End Sub

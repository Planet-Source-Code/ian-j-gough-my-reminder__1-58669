VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5565
   ControlBox      =   0   'False
   Icon            =   "Calendar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2955
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Date"
      Height          =   495
      Left            =   1155
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _Version        =   524288
      _ExtentX        =   9340
      _ExtentY        =   5318
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2005
      Month           =   1
      Day             =   28
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fairly simple just take the calendar values and put them in form1 textbox5
'
Private Sub Calendar1_Click()
'
End Sub
'
Private Sub Command1_Click()
100 On Error GoTo LocalErrors
110 Form1.Text5.Text = Calendar1.Value 'What ever calendar date you have clicked on show it in form1 textbox5
120 Unload Form2
130 Form1.Show
140 LocalErrors:
150 Exit Sub
160 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '017' and program 'Reminder Service'", vbExclamation, "Error Control"
170 Unload Me
End Sub
'
Private Sub Command2_Click()
100 On Error GoTo LocalErrors
110 Unload Form2 ' As it says
120 Form1.Show
130 LocalErrors:
140 Exit Sub
150 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '018' and program 'Reminder Service'", vbExclamation, "Error Control"
160 Unload Me
End Sub
'
Private Sub Form_Load()
100 On Error GoTo LocalErrors
110 Calendar1.Value = Date 'Make the calendar show todays date
120 LocalErrors:
130 Exit Sub
140 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '019' and program 'Reminder Service'", vbExclamation, "Error Control"
150 Unload Me
End Sub
'

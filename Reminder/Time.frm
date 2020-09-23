VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   Caption         =   "Alarm Time"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   1830
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2513
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Time"
      Height          =   495
      Left            =   713
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   945
      Left            =   2333
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "00"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   945
      Left            =   1133
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "00"
      Top             =   120
      Width           =   975
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   945
      Left            =   3300
      TabIndex        =   2
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1667
      _Version        =   393216
      Max             =   59
      Wrap            =   -1  'True
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   945
      Left            =   2100
      TabIndex        =   3
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1667
      _Version        =   393216
      Max             =   23
      Wrap            =   -1  'True
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Just takes the values from the textbox's and puts them in form1 textbox 4 and 5
'
'As you can see this form uses text2.text and text3.text as you don't have to start with text1.text!
Private Sub Command1_Click()
100 On Error GoTo LocalErrors
110 Form1.Text4.Text = Text3.Text & ":" & Text2.Text & ":00" 'We use form1 before text4.text as that tells the App to goto that form before looking for the textbox
120 Form1.Show 'The above code takes both textbox values and puts them into 1 textbox we use the ":" and ":00" just for the end user only (Looks better!)
130 Unload Form3
140 Exit Sub
150 LocalErrors:
160 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '020' and program 'Reminder Service'", vbExclamation, "Error Control"
170 Unload Me
End Sub
'
Private Sub Command2_Click()
100 On Error GoTo LocalErrors
110 Form1.Show 'Show form1
120 Unload Form3 'Unload form3
130 Exit Sub
140 LocalErrors:
150 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '021' and program 'Reminder Service'", vbExclamation, "Error Control"
160 Unload Me
End Sub
'
Private Sub Form_Load()
100
End Sub

Private Sub Text3_Change()
100
End Sub
'
Private Sub UpDown1_Change()
100 On Error GoTo LocalErrors
110 If UpDown1.Value < 10 Then 'If text3.text is less than 10 add a "0" then the value from the updown button
120 Text3.Text = "0" & CStr(UpDown1.Value)
130 Else 'If text3.text is more than 9 then just add the value from the up down button
140 Text3.Text = UpDown1.Value
150 End If
160 Exit Sub
170 LocalErrors:
180 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '022' and program 'Reminder Service'", vbExclamation, "Error Control"
190 Unload Me
End Sub
'
Private Sub UpDown2_Change()
100 On Error GoTo LocalErrors
110 If UpDown2.Value < 10 Then 'Read the above comments it's the same just we use text2.text and updown2 instead of 1
120 Text2.Text = "0" & CStr(UpDown2.Value)
130 Else
140 Text2.Text = UpDown2.Value
150 End If
160 Exit Sub
170 LocalErrors:
180 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '023' and program 'Reminder Service'", vbExclamation, "Error Control"
190 Unload Me
End Sub
'

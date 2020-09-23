VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reminder Service"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4770
   Icon            =   "Reminder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox TrayIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "Reminder.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   0
      Width           =   540
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1920
      Top             =   5040
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2520
      Top             =   5040
   End
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Exit"
      Height          =   495
      Left            =   225
      TabIndex        =   8
      Top             =   5520
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Reminder"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   5040
      WhatsThisHelpID =   1455
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Reminder"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Date"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Time"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2445
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Date"
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   165
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Time"
      Top             =   2760
      Width           =   2175
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2655
      Left            =   165
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   4683
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
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   105
      TabIndex        =   0
      Top             =   3360
      Width           =   4575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Your Chosen Date and Time"
      Height          =   195
      Left            =   1350
      TabIndex        =   10
      Top             =   4440
      Width           =   2010
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Your Reminder"
      Height          =   195
      Left            =   1905
      TabIndex        =   9
      Top             =   3120
      Width           =   1065
   End
   Begin VB.Menu mnuOP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuCTRL 
         Caption         =   ""
      End
      Begin VB.Menu spe1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###########################################################
'#   Reminder - A General Use Utility
'#      By: Ian.J.Gough
'#
'#      1.0 (Feb 03, 2005):
'#          Initial Release
'#          In version 2.0 coming soon your be able to set multiple reminders
'#
'#      Copyright © 2005 Ian.J.Gough  (iangough7@aol.com)
'#          This source code is provided 'as-is', without any express or implied warranty. In no event will the author(s) be held liable for any damages arising from the use of this source code. Permission is granted to anyone to use this source code for any purpose, including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:
'#          1. The origin of this source code must not be misrepresented; you must not claim that you wrote the original source code. If you use this source code in a product, an acknowledgment in the product documentation would be appreciated but is not required.
'#          2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original source code.
'#          3. This notice may not be removed or altered from any source distribution.
'#              (NOTE: This license is borrowed from zLib.)
'#
'#  Please remember to vote on PSC.com if you like this code!
'#
'###########################################################
'
'###########################################################
'Most of the line numbers are not essential but i have used them only to help with this tutrial
'I have used error handlers everywhere!  As if there are any bugs the program will unload and not crash.  Plus i have used error codes everywhere too so i'll know where to look if this is an error.
'I hope you can learn from this code and i have submitted it as i have downloaded and learned alot from planetsourcecode and wanted to give something back so here it is.
'I have commented where i can and where i feel appropiate.
'###########################################################

Public Function NoSysIcon(maxIcon As Boolean)   'This says that maxIcon is a Boolean variables and so can only be true or false nothing more! (xxx as Boolean = xxx can be true or false only!)xxx is the variable Name (name of the thing which you want to be only true or false)
100 On Error GoTo LocalErrors   'If we have an error goto "localErrors"
110 Select Case maxIcon     'The name of the case
     Case False   'Case program in Min Mode
120        Me.Visible = False 'Hide the form
130        ShowProgramInTray    'Now show TrayIcon  Picture in SysTray as an icon
140        mnuCTRL.Caption = "E&xpand Application"  'This is the text for the menu when you right click on the tray icon
150     Case Else   'Case program in Max Mode
160        Me.Visible = True    'Show the form
170        DeleteIcon TrayIcon  'Delete the tray icon form the system tray
180        mnuCTRL.Caption = "Minimize App to System Tray"  'This is the text for the menu in this form You could also use it for when people exit your form giving them the option of exiting or minimising
190     End Select  'End choice
200 Exit Function   'Exit the above and as there is no more below it will then end the function  (providing its error free of course!
210 LocalErrors:    'If we have an error this will tell us!  Very usefull when you distribute your program!
220         MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '001' and program 'Reminder Service'", vbExclamation, "Error Control"   'Message box for error!
230         Unload Me   'Unload the form (Quit the program!)
End Function    'As it says!
'###########################################################
' This Function Show the TrayIcon Picture in the System Tray
' Don't Bother with what it says
' To change the TrayIcon's ToolTip goto line 190
'###########################################################
Public Function ShowProgramInTray()
100 On Error GoTo LocalErrors
110 INTRAY = True   'Means App is now in Tray
120    NI.cbSize = Len(NI) 'set the length of this structure
130    NI.hwnd = TrayIcon.hwnd 'control to receive messages from
140    NI.uID = 0 'uniqueID
150    NI.uID = NI.uID + 1
160    NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP 'operation flags
170    NI.uCallbackMessage = WM_MOUSEMOVE 'recieve messages from mouse activities
180    NI.hIcon = TrayIcon.Picture  'the location of the icon to display
190    NI.szTip = "Reminder Service" + Chr$(0) 'LoadResString(Language) + Chr$(0)  'the tool tip to display"' Change System Tray Icon's Tool Tip Here but don't delete chr$(0) [its line carriage here]
200    result = Shell_NotifyIconA(NIM_ADD, NI) 'add the icon to the system tray
210 Exit Function
220 LocalErrors:
230     MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '002' and program 'Reminder Service'", vbExclamation, "Error Control"
240     Unload Me
End Function
'###########################################################
' This Function Deletes the TrayIcon.Picture from the System Tray
' Don't Bother what it says
' On remove, we only have to give enough information for Windows
' to locate the icon, then tell the system to delete it.
'###########################################################
Private Sub DeleteIcon(pic As Control)
100 On Error GoTo LocalErrors
110 INTRAY = False  'Means app is unloaded or Max mode
120    NI.uID = 0 'uniqueID
130    NI.uID = NI.uID + 1
140    NI.cbSize = Len(NI)
150    NI.hwnd = pic.hwnd
160    NI.uCallbackMessage = WM_MOUSEMOVE
170    result = Shell_NotifyIconA(NIM_DELETE, NI)
180 Exit Sub
190 LocalErrors:
200     MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '003' and program 'Reminder Service'", vbExclamation, "Error Control"
210     Unload Me
End Sub

'
Private Sub cmdMenu_Click()
100 On Error GoTo LocalErrors
110     Me.Visible = False     ' Hide the form
120     ShowProgramInTray               'Now show TrayIcon PictureBox's Picture in SysTray as icon
130     mnuCTRL.Caption = "E&xpand Application" ' The menu caption when you right click on the icon in the tray
140 Exit Sub
150 LocalErrors:
160     MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '004' and program 'Reminder Service'", vbExclamation, "Error Control"
170     Unload Me
End Sub
'
Private Sub Command1_Click()
100 On Error GoTo LocalErrors
110     Dim filename As String
120     Dim drive As String
130     Dim reminderdate As String
140     Dim remindertext As String
150     Dim remindertime As String
160
170     drive = App.Path & "\"
180     filename = drive + "Reminder.inf"
190
200     reminderdate = Text5.Text ' Tell the program where it should put each value
230     remindertext = Text1.Text
240     remindertime = Text4.Text
250 Open filename For Output As #1 ' Open the Reminder.inf file so the App can write to it
260 reminderdate = Text5.Text
270 remindertext = Text1.Text
280 remindertime = Text4.Text
290 Write #1, reminderdate, remindertext, remindertime ' Write the values from the above text files into the Reminder.inf file
300 Close #1 ' Close the file the Reminder.inf file
310
320     If Text1.Text = "" Then ' If the textbox is emtpy goto message box else move on
330         MsgBox "Please enter some text to remind you of your message", vbInformation, "Reminder Service!"
340     Else
350
360     If Text4.Text = "" Then
370         MsgBox "Please Choose a time", vbInformation, "Reminder Service!"
380     Else
390
400     If Text5.Text = "" Then
410         MsgBox "Please Choose a Date", vbInformation, "Reminder Service!"
420     Else
430
440     If Text5.Text < Calendar1.Value Then ' If the text in the textbox(date) is older than todays date goto message box else move on
450         MsgBox "This Date has passed please choose a new one", vbInformation, "Reminder Service!"
460     Else
470
480     If Text5.Text = Date And Text4.Text < Time Then ' If the date is today and the time has passed goto message box else move on
490         MsgBox "This Time has passed please choose a new one", vbInformation, "Reminder Service!"
500     Else
510
520     MsgBox "Reminder Set", vbInformation, "Reminder Service!" ' Now we know got through all the above with no problems show the message box
530 Timer2.Interval = 100 ' Set the timer2 time to 100 (Like were resetting it as if a message box from timer2 has been displayed timer2 will be stopped!
540 Me.Visible = False
550 ShowProgramInTray               'Now show TrayIcon PictureBox's Picture in SysTray as icon
560 mnuCTRL.Caption = "E&xpand Application"
570     End If ' We have used if now we need to stop the if (Like if = start and End if = stop) We have to do this for all of them
580     End If
590     End If
600     End If
610     End If
620 Exit Sub
630 LocalErrors:
640     MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '005' and program 'Reminder Service'", vbExclamation, "Error Control"
650     Unload Me
End Sub
'
Private Sub Command2_Click()
100 On Error GoTo LocalErrors
110 Dim filename As String
120 Dim drive As String
130 Dim reminderdate As String
140 Dim remindertext As String
150 Dim remindertime As String
160 Dim a As Integer
180
190 Beep ' Make the system speaker beep
2000 a = MsgBox("Are you sure you want to clear this reminder", vbOKCancel, "Reminder Service!")
210 If a = 1 Then 'If ok was pressed then do the below otherwise goto else (Line 400)
220 Text5.Text = ""
230 Text1.Text = ""
240 Text4.Text = ""
250
260 drive = App.Path & "\"
270 filename = drive + "Reminder.inf"
280 reminderdate = Text5.Text
290 remindertext = Text1.Text
300 remindertime = Text4.Text
310
320 Open filename For Output As #1 'Open the Reminder.inf file
330 reminderdate = Text5.Text
340 remindertext = Text1.Text
350 remindertime = Text4.Text
360 Write #1, reminderdate, remindertext, remindertime 'Write the text from the textbox's to the Reminder.inf file (There is no text as we just changed it so there is not any so it will delete the text from the Reminder.inf file
370 Close #1 'Close the file now we have wrote to it
380
390 MsgBox "Reminder Deleted", vbInformation, "Reminder Service!"
400     Else
410
MsgBox "Reminder Not Deleted", vbInformation, "Reminder Service!"
420 Exit Sub
430 End If
440
450 Exit Sub
460 LocalErrors:
470 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '006' and program 'Reminder Service'", vbExclamation, "Error Control"
480 Unload Me
End Sub
'
Private Sub Command3_Click()
100     On Error GoTo LocalErrors
110         Unload Me
120     Exit Sub
130 LocalErrors:
140     MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '007' and program 'Reminder Service'", vbExclamation, "Error Control"
150         Unload Me
End Sub
'
Private Sub Command4_Click()
100     On Error GoTo LocalErrors
110         Form1.Hide
120         Form3.Show
130     Exit Sub
140 LocalErrors:
150         MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '008' and program 'Reminder Service'", vbExclamation, "Error Control"
160         Unload Me
End Sub
'
Private Sub Command5_Click()
100     On Error GoTo LocalErrors
110         Form1.Hide
120         Form2.Show
130     Exit Sub
140 LocalErrors:
150         MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '009' and program 'Reminder Service'", vbExclamation, "Error Control"
160         Unload Me
End Sub
'
Private Sub Form_Load()
100 On Error GoTo LocalErrors
110 Dim filename As String
120 Dim drive As String
130 Dim reminderdate As String
140 Dim remindertext As String
150 Dim remindertime As String
160
170 Calendar1.Value = Date
180 drive = App.Path & "\"
190 filename = drive + "Reminder.inf"
'
200 ShowAtStartup "Reminder.exe" 'Check the RegStuff.bas as "ShowAtStartup" is based there
'
210 Open filename For Input As #1 'Open the reminder.inf file
220 Input #1, reminderdate, remindertext, remindertime 'Insert the text into the appropiate text box's
230 Do Until EOF(1) 'Continue untill the end of file has been reached
240 Loop
250 Close #1
260
270 Text5.Text = reminderdate
280 Text1.Text = remindertext
290 Text4.Text = remindertime
300
310 TrayIcon.Top = Me.Height + 1000 'Set TrayIcon PictureBox top to such limit that its not visible
320
330 If Text1.Text = "" Then
340 NoSysIcon True ' Check the nosysIcon function for information
350 Else: NoSysIcon False
360 End If
370 If Text4.Text = "" Or Text5.Text = " " Or Text1.Text = " " Then
380 Exit Sub 'To make sure we don't get reminded about something we have no set
390 If Text5.Text < Calendar1.Value And Text4.Text < Time Then
400      If MsgBox("A Reminder was set for today! Would you like to view your Reminder Message?", vbYesNo, "Reminder Service!") = vbYes Then
420 Me.Show
430 Else
440 If MsgBox("Would you like to be reminded again in 10 minutes?", vbYesNo, "Reminder Service!") = vbYes Then
450 timedPause 600
460 GoTo 390
470 Else
480 Exit Sub
490 End If
5000 End If
510 End If
520 End If
530
540 Exit Sub
550 LocalErrors:
560 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '010' and program 'Reminder Service'", vbExclamation, "Error Control"
570 Unload Me
End Sub

Private Sub mnuAbout_Click()
100 On Error GoTo LocalErrors
110 Load form4
120 form4.Show
130 Form1.Enabled = False
140 Exit Sub
150 LocalErrors:
160 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error letter 'MAbout' and program 'Reminder Service'", vbExclamation, "Error Control"
170 Unload Me
End Sub

'###########################################################
'mnuCTRL controls the Expansion and Minimizing function
'###########################################################
Private Sub mnuCTRL_Click()
100 On Error GoTo LocalErrors
110 NoSysIcon INTRAY  'Call Menu Function
120 Exit Sub
130 LocalErrors:
140 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '011' and program 'Reminder Service'", vbExclamation, "Error Control"
150 Unload Me
End Sub
'
Private Sub mnuExit_Click()
100 On Error GoTo LocalErrors
110 DeleteIcon TrayIcon 'As we are exiting App
120 End
130 Exit Sub
140 LocalErrors:
150 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '012' and program 'Reminder Service'", vbExclamation, "Error Control"
160 Unload Me
End Sub
'
Private Sub Text1_Change()
100 On Error GoTo LocalErrors
110 Exit Sub
120 LocalErrors:
130 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '013' and program 'Reminder Service'", vbExclamation, "Error Control"
140 Unload Me
End Sub
'
Private Sub Timer1_Timer()
100 On Error GoTo LocalErrors
110 Text2.Text = Time 'Gets the time from the pc
120 Text3.Text = Date 'Gets the date from the pc
130 Exit Sub
140 LocalErrors:
150 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '014' and program 'Reminder Service'", vbExclamation, "Error Control"
160 Unload Me
End Sub
'
Private Sub Timer2_Timer()
100 On Error GoTo LocalErrors
110 If Text4.Text = Time And Text5.Text = Date Then 'Fairly simple if time and date match the ones the user selected tell them!
120
130      If MsgBox("This is your Reminder! Would you like to view your Reminder Message?", vbYesNo, "Reminder Service!") = vbYes Then
140 Me.Show
150 Timer2.Interval = 0 'Stop the timer so we don't get duplicate messages
160 Exit Sub
170 Else
180 If MsgBox("Would you like to be reminded again in 10 minutes?", vbYesNo, "Reminder Service!") = vbYes Then
190 timedPause 600 'Pause for x amount of seconds then goto line 130 and ask again
200 GoTo 130
210 Else
220 Timer2.Interval = 0
230 Exit Sub
240 End If
250 End If
260 End If
270
280
290
300
310 Exit Sub
320 LocalErrors:
330 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '015' and program 'Reminder Service'", vbExclamation, "Error Control"
340 Unload Me
End Sub
'
Private Sub Trayicon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
100 On Error GoTo LocalErrors
110 Dim Msg As Long
120 Msg = (X And &HFF) * &H100
130 Select Case Msg
 Case 0 'mouse moves
150    Case &HF00  'left mouse button down
160 Case &H1E00 'left mouse button up
170 Case &H3C00  'right mouse button down
180 PopupMenu mnuOP, 2, , , mnuCTRL 'show the popoup menu
190 Case &H2D00 'left mouse button double click
200 NoSysIcon True    'Show App on double clicking Mouse's Left Button
210 Case &H4B00 'right mouse button up
220 Case &H5A00 'right mouse button double click
230 End Select
240 Exit Sub
250 LocalErrors:
260 MsgBox "There was an Error starting this program!  Please re-install the program or contact 'iangough7@aol.com' quotng error number '016' and program 'Reminder Service'", vbExclamation, "Error Control"
270 Unload Me
End Sub
'

VERSION 5.00
Begin VB.Form DayOfWeek 
   Caption         =   "Day Of Week"
   ClientHeight    =   2835
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   3090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   10
      Text            =   "01"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txt2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "01"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      MaxLength       =   2
      TabIndex        =   8
      Text            =   "01"
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdExe 
      Caption         =   "Get Day"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdExpl 
      Caption         =   "Algorithms"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame DateStandard 
      Caption         =   "Date Format"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1815
      Begin VB.OptionButton OptMM 
         Caption         =   "MM-DD-YYYY"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton OptDD 
         Caption         =   "DD-MM-YYYY"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "YYYY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      Caption         =   "DD"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "MM"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   375
   End
   Begin VB.Menu Tools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuMkCalendar 
         Caption         =   "&View calendar"
      End
   End
   Begin VB.Menu mnuAlgorithms 
      Caption         =   "&Algorithms"
      Visible         =   0   'False
      Begin VB.Menu mnuZellersAlgorithm 
         Caption         =   "&Zeller's Algorithm"
      End
      Begin VB.Menu mnuLeapYear 
         Caption         =   "&Leap Year"
      End
   End
End
Attribute VB_Name = "DayOfWeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TheDayToday As String
Dim ALeapYear As String

Private Sub cmdExe_Click()
'overall validation of date
Dim checkdate As String
'If Me.OptDD = True Then
    checkdate = Me.txt1.Text & "/" & Me.txt2.Text & "/" & Me.txt3.Text
'Else
'    checkdate = Me.txt2.Text & "/" & Me.txt1.Text & "/" & Me.txt3.Text
'End If
Debug.Print checkdate
'Debug.Print ZellersAlgorithm(txt1.Text, txt2.Text, txt3.Text)
If Not IsDate(checkdate) Then
    MsgBox "The date that you have entered is not valid. It could be that you have entered more days than there are in that month", vbExclamation, "Error"
Else 'if date is correct, run the algorithm
    'see if it is a leap year
    Debug.Print Calendar.LeapYearTest(Me.txt3)
    If Calendar.LeapYearTest(Me.txt3.Text) = True Then
        ALeapYear = Me.txt3.Text & " is also a leap year"
    Else: ALeapYear = Me.txt3.Text & " is NOT a leap year"
    End If
    Debug.Print ALeapYear
    'display a msgbox with what day of the week it is
    If Me.OptDD = True Then
        MsgBox checkdate & " is a " & LoadResString(ZellersAlgorithm(txt1.Text, txt2.Text, txt3.Text)) & vbCrLf & vbCrLf & ALeapYear
    Else
        MsgBox checkdate & " is a " & LoadResString(ZellersAlgorithm(txt2.Text, txt1.Text, txt3.Text)) & vbCrLf & vbCrLf & ALeapYear
    End If
End If
End Sub

Private Sub cmdExpl_Click()
'Algorithm.Show
PopupMenu mnuAlgorithms
End Sub


Private Sub Form_Load()
'make sure the options are activated
Me.OptDD = True
'Me.optGregorian = True

'Put todays date in the text boxes
TheDayToday = Date
txt1.Text = Day(TheDayToday)
txt2.Text = Month(TheDayToday)
txt3.Text = Year(TheDayToday)
End Sub

Private Sub mnuMkCalendar_Click()
Calendar.Show
Me.Hide

End Sub
Private Sub mnuZellersAlgorithm_click()
frmZellersAlgorithm.Show
End Sub

Private Sub mnuLeapYear_click()
LeapYearAlgorithm.Show

End Sub
Private Sub OptDD_Click()
'idleness
Call OptMM_Click
End Sub

Private Sub OptMM_Click()
Dim temp1 As String
Dim temp2 As String

'swap date around
temp1 = Me.txt1.Text
temp2 = Me.txt2.Text

txt1 = temp2
txt2 = temp1

'swap captions above date
temp1 = Me.lbl1.Caption
temp2 = Me.lbl2.Caption

lbl1.Caption = temp2
lbl2.Caption = temp1


End Sub




Private Sub txt1_KeyPress(KeyAscii As Integer)
'only allow numercal input
If IsNumeric(Chr(KeyAscii)) Then
    KeyAscii = KeyAscii
ElseIf KeyAscii = 8 Or KeyAscii = 127 Then 'allow delete keys
    KeyAscii = KeyAscii
Else: KeyAscii = 0
End If

End Sub

Private Sub txt1_LostFocus()
'validate it acordingly to which format is selected
If Me.OptDD = True Then
    Call CheckDay(Me.txt1)
Else
    Call CheckMonth(Me.txt1)
End If
End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
Call txt1_KeyPress(KeyAscii) 'idleness
End Sub

Private Sub txt2_LostFocus()
'validate it acordingly to which format is selected
If Me.OptDD = True Then
    Call CheckMonth(Me.txt2)
Else
    Call CheckDay(Me.txt2)
End If
End Sub

Sub CheckDay(DayControl As TextBox)
'validate the day
    If Not IsNumeric(DayControl.Text) Then
        MsgBox "The day must be a number not letters", vbExclamation, "Error"
        DayControl.SetFocus
        Exit Sub
    ElseIf DayControl.Text > 31 Or DayControl < 1 Then
        MsgBox "The day must lie between 1 and 31", vbExclamation, "Error"
        DayControl.SetFocus
        Exit Sub
    End If
End Sub

Sub CheckMonth(MonthControl As TextBox)
'validate the month
    If Not IsNumeric(MonthControl.Text) Then
        MsgBox "The month must be a number not letters", vbExclamation, "Error"
        MonthControl.SetFocus
        Exit Sub
    ElseIf MonthControl.Text > 12 Or MonthControl < 1 Then
        MsgBox "The month must lie between 1 and 12", vbExclamation, "Error"
        MonthControl.SetFocus
        Exit Sub
    End If
End Sub


Private Sub txt3_KeyPress(KeyAscii As Integer)
Call txt1_KeyPress(KeyAscii) 'idleness
End Sub

Private Sub txt3_LostFocus()
Call CheckYear(Me.txt3)
End Sub

Sub CheckYear(YearControl As TextBox)

    If Not IsNumeric(YearControl.Text) Then
        YearControl.SetFocus 'some strange bug means it has to be here
        MsgBox "The Year must be a number not letters", vbExclamation, "Error"
        Exit Sub
    ElseIf YearControl.Text < 0 Then
        YearControl.SetFocus 'some strange bug means it has to be here
        MsgBox "The year must be greater than 0", vbExclamation, "Error"
        Exit Sub
    End If
End Sub

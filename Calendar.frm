VERSION 5.00
Begin VB.Form Calendar 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   46
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtYear 
      Height          =   390
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "2000"
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox cboMonth 
      Height          =   390
      Left            =   360
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   36
      Left            =   960
      TabIndex        =   45
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   35
      Left            =   360
      TabIndex        =   44
      Top             =   3000
      Width           =   615
   End
   Begin VB.Line Line16 
      X1              =   360
      X2              =   4560
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   34
      Left            =   3960
      TabIndex        =   43
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   33
      Left            =   3360
      TabIndex        =   42
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   32
      Left            =   2760
      TabIndex        =   41
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   31
      Left            =   2160
      TabIndex        =   40
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   30
      Left            =   1560
      TabIndex        =   39
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   29
      Left            =   960
      TabIndex        =   38
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   28
      Left            =   360
      TabIndex        =   37
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   27
      Left            =   3960
      TabIndex        =   36
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   26
      Left            =   3360
      TabIndex        =   35
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   25
      Left            =   2760
      TabIndex        =   34
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   24
      Left            =   2160
      TabIndex        =   33
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   23
      Left            =   1560
      TabIndex        =   32
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   22
      Left            =   960
      TabIndex        =   31
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   21
      Left            =   360
      TabIndex        =   30
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   20
      Left            =   3960
      TabIndex        =   29
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   19
      Left            =   3360
      TabIndex        =   28
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   18
      Left            =   2760
      TabIndex        =   27
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   17
      Left            =   2160
      TabIndex        =   26
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   16
      Left            =   1560
      TabIndex        =   25
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   15
      Left            =   960
      TabIndex        =   24
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   14
      Left            =   360
      TabIndex        =   23
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   13
      Left            =   3960
      TabIndex        =   22
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   12
      Left            =   3360
      TabIndex        =   21
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   11
      Left            =   2760
      TabIndex        =   20
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   10
      Left            =   2160
      TabIndex        =   19
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   9
      Left            =   1560
      TabIndex        =   18
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   8
      Left            =   960
      TabIndex        =   17
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   16
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   6
      Left            =   3960
      TabIndex        =   15
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   5
      Left            =   3360
      TabIndex        =   14
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   4
      Left            =   2760
      TabIndex        =   13
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   12
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   10
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.Line Line15 
      X1              =   360
      X2              =   4560
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line14 
      X1              =   360
      X2              =   4560
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line13 
      X1              =   360
      X2              =   4560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line12 
      X1              =   360
      X2              =   4560
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line11 
      X1              =   3960
      X2              =   3960
      Y1              =   840
      Y2              =   3360
   End
   Begin VB.Line Line10 
      X1              =   3360
      X2              =   3360
      Y1              =   840
      Y2              =   3360
   End
   Begin VB.Line Line9 
      X1              =   2760
      X2              =   2760
      Y1              =   840
      Y2              =   3360
   End
   Begin VB.Line Line8 
      X1              =   2160
      X2              =   2160
      Y1              =   840
      Y2              =   3360
   End
   Begin VB.Line Line7 
      X1              =   1560
      X2              =   1560
      Y1              =   840
      Y2              =   3360
   End
   Begin VB.Line Line6 
      X1              =   960
      X2              =   960
      Y1              =   840
      Y2              =   3360
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   360
      X2              =   4560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   840
      Y2              =   3360
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4560
      X2              =   360
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   360
      X2              =   360
      Y1              =   840
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   360
      X2              =   4560
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mon"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Sun"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Sat"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Fri"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Thu"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Wed"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Tue"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000F&
      FillColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   360
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DayStarts As Integer 'what day does the day start on
Public LengthOfMonth As Integer
Public i As Integer



Private Sub cboMonth_Click() 'immediate affect

'change forms caption acordingly
Me.Caption = Me.cboMonth.List(Me.cboMonth.ListIndex) & " " & Me.txtYear.Text & "'s calendar"


Call GenerateCalendar(Me.cboMonth.List(Me.cboMonth.ListIndex), Me.txtYear.Text)
Debug.Print "Monthlentgh ="; LengthOfMonth

End Sub

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub cmdGenerate_Click()
Me.Caption = Me.cboMonth.List(Me.cboMonth.ListIndex) & " " & Me.txtYear.Text & "'s calendar"
Call DayOfWeek.CheckYear(Me.txtYear)
GenerateCalendar Me.cboMonth.List(Me.cboMonth.ListIndex), Me.txtYear.Text
End Sub

Private Sub Form_Load()
'fill up the combobox with the months
With Me.cboMonth
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
End With

'auto select the month from the combo box, will trigger the generate calender event, thus filling the days in
If DayOfWeek.OptDD = True Then
    Me.cboMonth.ListIndex = DayOfWeek.txt2.Text - 1
ElseIf DayOfWeek.OptMM = True Then
    Me.cboMonth.ListIndex = DayOfWeek.txt1.Text - 1
End If
'autofill in the year
Me.txtYear.Text = DayOfWeek.txt3.Text
Me.Caption = Me.cboMonth.List(Me.cboMonth.ListIndex) & " " & Me.txtYear.Text & "'s calendar"
End Sub


Private Sub Form_Unload(Cancel As Integer)
DayOfWeek.Show
Unload Me
End Sub



Private Sub txtYear_LostFocus()
'check that it is a valid year/input
Call DayOfWeek.CheckYear(Me.txtYear)
'change forms caption accordingly
Me.Caption = Me.cboMonth.List(Me.cboMonth.ListIndex) & " " & Me.txtYear.Text & "'s calendar"

Call GenerateCalendar(Me.cboMonth.List(Me.cboMonth.ListIndex), Me.txtYear.Text)
End Sub


Sub GenerateCalendar(TheMonth As String, TheYear As Integer)
    
'clear calendar
For i = 0 To 36
    Me.lblDay(i) = ""
Next
'determine number of days in the month
Select Case TheMonth
    Case "January", "March", "May", "July", "August", "October", "December"
        LengthOfMonth = 31
    Case "April", "June", "September", "November"
        LengthOfMonth = 30
    Case "February"
        If LeapYearTest(Me.txtYear.Text) = True Then
            LengthOfMonth = 29
        Else
            LengthOfMonth = 28
        End If

End Select
    
    
    'find what day the month starts
    DayStarts = ZellersAlgorithm(1, Me.cboMonth.ListIndex + 1, Me.txtYear)
    'put the number one in the day that the month starts
    Me.lblDay(DayStarts) = 1
    
    'fill in the calendar
    For i = 1 To LengthOfMonth
        If i = LengthOfMonth Then Exit For 'otherwise it will go to 32 and onwards regardless
        'fill calendar with day numbers
        Me.lblDay(DayStarts + i) = i + 1
        
    Next

End Sub

Function LeapYearTest(TheYear As Integer) As Boolean
Dim LeapYear As Integer
LeapYear = 0
        'test for leap years
        If (TheYear) - (4 * Int(TheYear / 4)) = 0 Then
            'divisible by 4
            LeapYear = 1
        End If
        If (TheYear) - (100 * Int(TheYear / 100)) = 0 Then
            'divisible by 100
            LeapYear = LeapYear + 2
        End If
        If (TheYear) - (400 * Int(TheYear / 400)) = 0 Then
            'divisible by 400
            LeapYear = LeapYear + 4
        End If
        Debug.Print "leap year: "; LeapYear
        Select Case LeapYear
            Case 1, 5, 7 'leap year (see bottom)
                LeapYearTest = True
            Case Else
                LeapYearTest = False
        End Select

End Function
'____________________________________________________
'LEAP YEARS
'1. Every year divisible by 4 is a leap year.
'2. But every year divisible by 100 is NOT a leap year
'3. Unless the year is also divisible by 400, then it is still a leap year
'____________________________________________________

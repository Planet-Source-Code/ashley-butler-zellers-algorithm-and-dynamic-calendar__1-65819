VERSION 5.00
Begin VB.Form frmZellersAlgorithm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                    How is the day calculated?"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   ControlBox      =   0   'False
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
   ScaleHeight     =   6075
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCloseExpl 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Zeller's Algorithm (or Zeller's Congruence)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label10 
      Caption         =   "Generally: (Int(2.6M - 5.39) + Int(N/4) + Int(C/4) + D + N - (2C)) mod 7"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4920
      Width           =   6975
   End
   Begin VB.Label Label9 
      Caption         =   "Remainder: 0 = Sunday, 1 = Monday, 2 = Tuesday etc"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   6855
   End
   Begin VB.Label Label8 
      Caption         =   "Find the remainder when the above is divided by 7."
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   7095
   End
   Begin VB.Label Label7 
      Caption         =   "Let N be the last 2 digits of Y"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   6975
   End
   Begin VB.Label Label6 
      Caption         =   "Let C be the first 2 digits of Y"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   7095
   End
   Begin VB.Label Label5 
      Caption         =   "Add the integer parts of (2.6M - 5.39), (N/4) and (C/4) then add on D and N and subtract 2C"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   7095
   End
   Begin VB.Label Label4 
      Caption         =   "If M is less than 3 add 12 to M and subtract 1 from Y"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   7095
   End
   Begin VB.Label Label3 
      Caption         =   "Let year number = Y"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Let month number = M"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Let day number = D"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmZellersAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCloseExpl_Click()
Unload Me
End Sub


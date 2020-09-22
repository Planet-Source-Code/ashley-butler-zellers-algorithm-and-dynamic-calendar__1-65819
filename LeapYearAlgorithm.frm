VERSION 5.00
Begin VB.Form LeapYearAlgorithm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "            How is the Leap Year Calculated?"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
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
   ScaleHeight     =   3510
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
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
      Left            =   840
      TabIndex        =   5
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "1900 isn't a leap year, but 2000 is"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "UNLESS it is also divisble by 400"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "BUT if it is divisible by 100 as well, then it isn't"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Every year divisble by 4 is a leap year"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Leap Year Algorithm"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "LeapYearAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub


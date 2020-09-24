VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CLASS METHOD DELEGATOR - Version 0.99 BETA"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "Stop Timer on Both"
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Timer on Form And Button"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   3735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   2040
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "UnSubClass Both"
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SubClass Form And Button"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00CBDEED&
      Caption         =   "Look at Debug (Immediate) Window!!!"
      ForeColor       =   &H00AE693C&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   6135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "DO NOT USE BREAK POINT IN CLASS DELEGATED METHOD!!!  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   5880
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLASSDELEGATOR1 As New Class1
Dim CLASSDELEGATOR2 As New Class1

Private Sub Command1_Click()
'SUBCLASS FORM1
CLASSDELEGATOR1.PerformSubclassCalling
CLASSDELEGATOR1.SubClass Me.hwnd

'SUBCLASS COMMAND1
CLASSDELEGATOR2.PerformSubclassCalling
CLASSDELEGATOR2.SubClass Command1.hwnd
End Sub

Private Sub Command2_Click()
CLASSDELEGATOR1.UnSubClass
CLASSDELEGATOR2.UnSubClass
End Sub

Private Sub Command3_Click()
'TIMER ON FORM1
CLASSDELEGATOR1.PerformTimerCalling
CLASSDELEGATOR1.StartTimer Form1.hwnd, 1

'TIMER ON COMMAND1
CLASSDELEGATOR2.PerformTimerCalling
CLASSDELEGATOR2.StartTimer Command1.hwnd, 1
End Sub

Private Sub Command4_Click()

CLASSDELEGATOR1.StopTimer
CLASSDELEGATOR2.StopTimer
End Sub

Private Sub Form_Load()
Text1 = LoadResString(1)
End Sub

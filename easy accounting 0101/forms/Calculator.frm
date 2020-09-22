VERSION 5.00
Begin VB.Form Calculator 
   Caption         =   "Ajay Patel's calculator"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5550
   Icon            =   "Calculator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "="
      Height          =   975
      Left            =   3960
      TabIndex        =   16
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "x"
      Height          =   855
      Left            =   3960
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "/"
      Height          =   855
      Left            =   2760
      TabIndex        =   14
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "-"
      Height          =   855
      Left            =   1560
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "+"
      Height          =   855
      Left            =   480
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "AC"
      Height          =   975
      Left            =   3960
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      Height          =   975
      Left            =   3960
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   975
      Left            =   2760
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   975
      Left            =   1560
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   975
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   975
      Left            =   2760
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   975
      Left            =   1560
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   975
      Left            =   2760
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   975
      Left            =   1560
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   4560
      Width           =   1095
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private t As String
Private pre_amt As Double
Private now_amt As Double
Private procedure As Integer

Private Sub Command1_Click()

t = t & "1"
Text1.Text = t

End Sub

Private Sub Command11_Click()
'+
now_amt = Val(t)
If procedure = 1 Then
    Text1.Text = now_amt + pre_amt
End If

End Sub

Private Sub Command12_Click()

Text1.Text = ""
pre_amt = 0
now_amt = 0

End Sub

Private Sub Command13_Click()

pre_amt = Val(t)
Text1.Text = ""
t = ""
procedure = 1

End Sub

Private Sub Command2_Click()

t = t & "2"
Text1.Text = t

End Sub
Private Sub Command3_Click()

t = t & "3"
Text1.Text = t

End Sub


VERSION 5.00
Begin VB.Form path_sel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select your path"
   ClientHeight    =   4575
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8130
   Icon            =   "path_sel.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8130
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   7335
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   7335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   7335
   End
End
Attribute VB_Name = "path_sel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim change_sure

change_sure = MsgBox("You want to change path....?", vbQuestion + vbYesNo, "Are You Sure !!!!")

If change_sure = 6 Then
selected_path = my_path & "\co.mdb"
selected_backup_path = my_path & "\backup\co.mdb"
FileCopy App.Path & "\main.txt", App.Path & "\main_old_path.txt"
Kill App.Path & "\main.txt"
Open App.Path & "\main.txt" For Output As #5
    Write #5, selected_path
    Write #5, selected_backup_path
Close #5
MsgBox "You have to restart your application again...!!!", vbOKOnly, "Your company path have been changed..,"
End
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
    my_path = Dir1.Path
    Text1.Text = my_path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    my_path = Dir1.Path
    Text1.Text = my_path
End Sub

Private Sub Form_Load()
    Dir1.Path = Drive1.Drive
    my_path = Dir1.Path
    Text1.Text = my_path
End Sub


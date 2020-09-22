VERSION 5.00
Begin VB.Form path_selection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select file (Open)"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   Icon            =   "select_path.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Ok"
      Height          =   855
      Left            =   8640
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancel"
      Height          =   975
      Left            =   8640
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   3000
      Pattern         =   "*.JPG"
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   5520
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "path_selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f_1 As String
Dim f_2 As String
Dim f_3 As String
Private Sub Command4_Click()
    f_1 = File1.Path
    f_2 = File1.FileName
    f_3 = f_1 & "\" & f_2
End Sub
Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_Click()
    f_1 = File1.Path
    f_2 = File1.FileName
    f_3 = f_1 & "\" & f_2
End Sub

Private Sub Form_Load()

'ap.App
'Me.Icon = LoadPicture(my_app_path & "esw.ico")
Me.Height = 2850
Me.Width = Screen.Width
Me.Left = 0
Me.Top = Screen.Height - (2850 + 400)
Me.BorderStyle = 1

'SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    Dir1.Path = App.Path
    File1.Pattern = "*.mdb"
End Sub

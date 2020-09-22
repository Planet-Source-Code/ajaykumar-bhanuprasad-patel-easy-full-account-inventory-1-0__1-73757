VERSION 5.00
Begin VB.Form frm_usr_creat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Creation Form"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   Icon            =   "user_creat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   13
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   4320
      TabIndex        =   10
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   9
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   8
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Text            =   "Here"
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Password Hint...,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   12
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "User Control selection...,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   11
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Authorization Password...,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm Password...,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Enter User Password...,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Enter User Name ...,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "frm_usr_creat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================
Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Combo1.AddItem "Data Entry Only..,"
    Combo1.AddItem "Show Report Only..,"
    Combo1.AddItem "Data Entry & change..,"
    Combo1.ListIndex = 0
End Sub
Private Sub Command1_Click()
    Call all_text_is_ok
    Call authorization_is_ok
    Call save_new_user
    Call back_to_user_menu
End Sub
Private Sub Command2_Click()
    frm_usr.Enabled = True
    Unload Me
End Sub
Public Sub all_text_is_ok()
    If Text1.Text <> "" Or Text2.Text = Text3.Text Then all_text_is_ok_code = 1
End Sub
Public Sub authorization_is_ok()
authorization_code = 0
    If Text4.Text = "ajaypatel" And Text2.Text = Text3.Text Then
        authorization_code = 1
    Else
        MsgBox "Sorry...!!! There is a wrong Authorization code or both password is not match...!!!"
        Exit Sub
    End If
End Sub
Public Sub save_new_user()
Call open_database
Call open_rs_co_user_dtl
        rs_co_user_dtl.AddNew
        rs_co_user_dtl!co_user_dtl_name = Text1.Text
        rs_co_user_dtl!co_user_dtl_pwrd = Text2.Text
        rs_co_user_dtl!co_user_dtl_hint = Text5.Text
        rs_co_user_dtl!co_user_dtl_ctrl = Combo1.ListIndex
        rs_co_user_dtl.UpdateBatch
End Sub
Public Sub back_to_user_menu()
    Close All
    frm_usr.Enabled = True
    frm_usr.Show
    Unload Me
End Sub



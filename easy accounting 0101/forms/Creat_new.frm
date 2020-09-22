VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Creat Group"
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   840
      TabIndex        =   33
      Text            =   "Combo2"
      Top             =   9120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   32
      Text            =   "Combo1"
      Top             =   8760
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   400
      Left            =   4080
      TabIndex        =   31
      Text            =   "Text13"
      Top             =   8280
      Width           =   5500
   End
   Begin VB.TextBox Text12 
      Height          =   400
      Left            =   4080
      TabIndex        =   30
      Text            =   "Text12"
      Top             =   7800
      Width           =   5500
   End
   Begin VB.TextBox Text11 
      Height          =   400
      Left            =   4080
      TabIndex        =   29
      Text            =   "Text11"
      Top             =   7320
      Width           =   5500
   End
   Begin VB.TextBox Text10 
      Height          =   400
      Left            =   4080
      TabIndex        =   28
      Text            =   "Text10"
      Top             =   6840
      Width           =   5500
   End
   Begin VB.TextBox Text9 
      Height          =   400
      Left            =   4080
      TabIndex        =   27
      Text            =   "Text9"
      Top             =   6360
      Width           =   5500
   End
   Begin VB.TextBox Text8 
      Height          =   400
      Left            =   4080
      TabIndex        =   26
      Text            =   "Text8"
      Top             =   5880
      Width           =   5500
   End
   Begin VB.TextBox Text7 
      Height          =   400
      Left            =   4080
      TabIndex        =   25
      Text            =   "Text7"
      Top             =   5400
      Width           =   5500
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6480
      TabIndex        =   17
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   9120
      Width           =   1575
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save"
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   9120
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   400
      Left            =   4080
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   4920
      Width           =   5500
   End
   Begin VB.TextBox Text5 
      Height          =   400
      Left            =   4080
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   4440
      Width           =   5500
   End
   Begin VB.TextBox Text4 
      Height          =   400
      Left            =   4080
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   3960
      Width           =   5500
   End
   Begin VB.TextBox Text3 
      Height          =   400
      Left            =   4080
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   3480
      Width           =   5500
   End
   Begin VB.TextBox Text2 
      Height          =   400
      Left            =   4080
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   3000
      Width           =   5500
   End
   Begin VB.TextBox Text1 
      Height          =   400
      Left            =   4080
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2520
      Width           =   5500
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   9900
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   24
      Top             =   8160
      Width           =   3000
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   23
      Top             =   7680
      Width           =   3000
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   22
      Top             =   7200
      Width           =   3000
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   21
      Top             =   6720
      Width           =   3000
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   20
      Top             =   6240
      Width           =   3000
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   5880
      Width           =   3000
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   18
      Top             =   5400
      Width           =   3000
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      TabIndex        =   8
      Top             =   3000
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Width           =   3000
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      TabIndex        =   6
      Top             =   4920
      Width           =   3000
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      TabIndex        =   5
      Top             =   4440
      Width           =   3000
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      TabIndex        =   4
      Top             =   3960
      Width           =   3000
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      TabIndex        =   3
      Top             =   3480
      Width           =   3000
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   9360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      Caption         =   "Address"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Top             =   0
      Width           =   8295
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      Caption         =   "Name of company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
lbl_name.Caption = co_name
lbl_add.Caption = co_add1 & ", " & co_add2 & ", " & co_pincode & ", " & co_city & ", " & co_contry

Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")

If selected_path = "" Or selected_path = Null Then
    selected_path = App.Path & "\data\1000\co.mdb;"
End If

Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""
Label13.Caption = ""

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""

Combo1.Text = ""
Combo2.Text = ""
End Sub

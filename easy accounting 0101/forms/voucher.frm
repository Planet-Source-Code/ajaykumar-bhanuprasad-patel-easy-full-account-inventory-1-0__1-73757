VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form voucher 
   Caption         =   "Form1"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13095
   ScaleWidth      =   21480
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   16560
      TabIndex        =   27
      Top             =   10440
      Width           =   1215
   End
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
      Left            =   1920
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   2400
      Width           =   5775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   5175
      Left            =   360
      TabIndex        =   4
      Top             =   5160
      Width           =   17415
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4815
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   16935
         _ExtentX        =   29871
         _ExtentY        =   8493
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3495
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   17415
      Begin VB.ComboBox Combo0 
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
         Left            =   15240
         TabIndex        =   31
         Text            =   "Combo0"
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Pirnt"
         Height          =   495
         Left            =   11880
         TabIndex        =   29
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmd_save_n_exit 
         Caption         =   "Save and exit"
         Height          =   495
         Left            =   11880
         TabIndex        =   26
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   11880
         TabIndex        =   25
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   11880
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_sv_n_new 
         Caption         =   "&Save and New"
         Height          =   495
         Left            =   11880
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text5 
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
         Left            =   15240
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "voucher.frx":0000
         Top             =   2160
         Width           =   9375
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9240
         TabIndex        =   16
         Text            =   "Text3"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   1560
         TabIndex        =   14
         Text            =   "Combo2"
         Top             =   1440
         Width           =   5775
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   9240
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text1 
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
         Left            =   15240
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   720
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   15240
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   15925249
         CurrentDate     =   40166
      End
      Begin VB.Label Label0 
         Caption         =   "Label0"
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
         Left            =   14280
         TabIndex        =   30
         Top             =   240
         Width           =   735
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
         Left            =   7800
         TabIndex        =   28
         Top             =   1440
         Width           =   1215
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
         Left            =   7800
         TabIndex        =   22
         Top             =   840
         Width           =   1215
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
         Height          =   495
         Left            =   14280
         TabIndex        =   20
         Top             =   2760
         Width           =   975
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
         Height          =   855
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   1335
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
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
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
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
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
         Height          =   495
         Left            =   15240
         TabIndex        =   9
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Height          =   495
         Left            =   15240
         TabIndex        =   8
         Top             =   1800
         Width           =   1815
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
         Height          =   375
         Left            =   14280
         TabIndex        =   7
         Top             =   1320
         Width           =   975
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
         Height          =   375
         Left            =   14280
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   7335
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   7095
   End
   Begin VB.Label lbl_head 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   8415
   End
End
Attribute VB_Name = "voucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call set_form_headings
Call set_vourcher_detail
Text5.Text = selected_user
End Sub
Public Sub set_form_headings()
lbl_name.Width = Me.Width
lbl_name.Left = 0
lbl_name.Caption = co_name
lbl_add.Width = Me.Width
lbl_add.Left = 0
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
lbl_head.Width = Me.Width
lbl_head.Left = 0
lbl_head.Caption = UCase(selected_procedure)
Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)
End Sub
Public Sub set_vourcher_detail()
DTPicker1.Value = Date
Label0.Caption = "Type"
Label1.Caption = "No"
Label2.Caption = "Date"
Label3.Caption = "Day"
Label4.Caption = "Time"
Label5.Caption = "Paid by"
Label6.Caption = "To"
Label7.Caption = "Narration"
Label8.Caption = "User"
Label9.Caption = "Amount"
Label10.Caption = "Amount"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

Text5.Enabled = False

Combo0.Text = ""
Combo1.Text = ""
Combo2.Text = ""

Frame1.Caption = "Current Transaction Detail"
selected_date = DTPicker1.Value
Frame2.Caption = selected_date & "s Transactions Detail"

Call add_account_combo0
Call add_account_combo1
Call add_account_combo2
End Sub

Private Sub DTPicker1_Change()
selected_date = DTPicker1.Value
Frame2.Caption = selected_date & "s Transactions Detail"
Call read_all_dated_transaction
End Sub
Public Sub read_all_dated_transaction()

End Sub
Private Sub MSFlexGrid1_Click()
Call read_current_transaction
End Sub
Public Sub read_current_transaction()

End Sub
Public Sub add_account_combo0()
Combo0.AddItem "1"
Combo0.AddItem "2"
Combo0.Text = "2"

End Sub

Public Sub add_account_combo1()
Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
    If rs_lgr_main_dtl!lgr_main_dtl_grup = "Cash-on-hand" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Balances" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Loans" Then
        Combo1.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
        If rs_lgr_main_dtl!lgr_main_dtl_alis <> "" Then Combo1.AddItem rs_lgr_main_dtl!lgr_main_dtl_alis
    End If
rs_lgr_main_dtl.MoveNext
Loop
End Sub
Public Sub add_account_combo2()
Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
    If rs_lgr_main_dtl!lgr_main_dtl_grup = "Cash-on-hand" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Balances" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Loans" Then
    Else
        Combo2.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
        If rs_lgr_main_dtl!lgr_main_dtl_alis <> "" Then Combo2.AddItem rs_lgr_main_dtl!lgr_main_dtl_alis
    End If
rs_lgr_main_dtl.MoveNext
Loop
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text2.Text = Format(Text2.Text, "0.00")
Text3.Text = Format(Text2.Text, "0.00")
End If
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text3.Text = Format(Text3.Text, "0.00")
End If
End Sub

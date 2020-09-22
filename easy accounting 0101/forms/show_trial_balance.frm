VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form show_trial_balance 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Trial_balance_sheet"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "show_trial_balance.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to Exit"
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
      Left            =   360
      TabIndex        =   4
      Top             =   6720
      Width           =   11055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   111673345
      CurrentDate     =   40194
   End
   Begin MSFlexGridLib.MSFlexGrid grid_report 
      Height          =   5175
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorSel    =   -2147483648
      ForeColorSel    =   8388608
      SelectionMode   =   1
   End
   Begin VB.Label lbl_head 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1995
      TabIndex        =   7
      Top             =   600
      Width           =   7995
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1995
      TabIndex        =   6
      Top             =   360
      Width           =   7995
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name of company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2000
      TabIndex        =   5
      Top             =   0
      Width           =   8000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   315
      TabIndex        =   1
      Top             =   960
      Width           =   11100
   End
End
Attribute VB_Name = "show_trial_balance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
selected_date = DTPicker1.Value
Call make_trail_balance_summary
Call set_grid_report
End Sub

Private Sub DTPicker1_Change()
selected_date = DTPicker1.Value
Call make_trail_balance_summary
Call set_grid_report
End Sub

Private Sub Form_Load()

'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'this is a code for sizing===================================

Call set_form_headings
selected_date = Date
Label1.Visible = False
'Label1.Caption = selected_procedure
DTPicker1.Value = Date
Call make_trail_balance_summary
Call set_grid_report
End Sub

Public Sub set_form_headings()
'lbl_name.Width = Me.Width
'lbl_name.Left = 0
lbl_name.Caption = co_name
'lbl_add.Width = Me.Width
'lbl_add.Top = -1000
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
'lbl_head.Width = Me.Width
'lbl_head.Left = 0
lbl_head.Caption = UCase(selected_procedure)
Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)
End Sub
Public Sub set_grid_report()
Dim total_cr_balance
Dim total_dr_balance
total_cr_balance = 0
total_dr_balance = 0

    rep_ending_date = DTPicker1.Value
    grid_report.Clear
    grid_report.Rows = 1
    grid_report.Cols = 4
    grid_report.Font.Size = 12
    
    b = 0
    
    grid_report.Font.Size = 12
    grid_report.TextMatrix(b, 0) = ""
    grid_report.TextMatrix(b, 1) = "Ledger"
    grid_report.TextMatrix(b, 2) = "Dr.Amount"
    grid_report.TextMatrix(b, 3) = "Cr.Amount"
    
    grid_report.ColWidth(0) = 500
    grid_report.ColWidth(1) = 6000
    grid_report.ColWidth(2) = 2000
    grid_report.ColWidth(3) = 2000
    
    'Dim x_grid_col
    'Dim total_grid_width
    'total_grid_width = 500
    'For x_grid_col = 0 To grid_report.Cols - 1
    '    total_grid_width = total_grid_width + grid_report.ColWidth(x_grid_col)
    'Next

'grid_report.Width = total_grid_width
b = 1
Call open_rs_lgr_clsg_smr
rs_lgr_clsg_smr.Sort = "lgr_clsg_dtl_name"
Do Until rs_lgr_clsg_smr.EOF
                
                'rs_lgr_clsg_smr!lgr_clsg_dtl_grup = rs_lgr_main_dtl!lgr_main_dtl_grup
                'rs_lgr_clsg_smr!lgr_clsg_dtl_pgrp = selected_primary_group
                'rs_lgr_clsg_smr!lgr_clsg_dtl_crpd = rs_lgr_main_dtl!lgr_main_dtl_crpd
                'rs_lgr_clsg_smr!lgr_clsg_dtl_cram = rs_lgr_main_dtl!lgr_main_dtl_cram
                'rs_lgr_clsg_smr!lgr_clsg_dtl_bal1 = rs_lgr_main_dtl!lgr_main_dtl_obl1
                'rs_lgr_clsg_smr!lgr_clsg_dtl_bal2 = rs_lgr_main_dtl!lgr_main_dtl_obl2
                'rs_lgr_clsg_smr!lgr_clsg_dtl_sid1 = rs_lgr_main_dtl!lgr_main_dtl_osd1
                'rs_lgr_clsg_smr!lgr_clsg_dtl_sid2 = rs_lgr_main_dtl!lgr_main_dtl_osd2
                'rs_lgr_clsg_smr!lgr_clsg_dtl_slun = rs_lgr_main_dtl!lgr_main_dtl_slun
                
                If rs_lgr_clsg_smr!lgr_clsg_dtl_tsid = "dr" And rs_lgr_clsg_smr!lgr_clsg_dtl_tbal <> 0 Then
                
                grid_report.Rows = grid_report.Rows + 1
                grid_report.TextMatrix(b, 0) = b
                grid_report.TextMatrix(b, 1) = rs_lgr_clsg_smr!lgr_clsg_dtl_name
                
                grid_report.TextMatrix(b, 2) = Format(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal, "0.00")
                total_dr_balance = total_dr_balance + Val(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal)
                b = b + 1

                ElseIf rs_lgr_clsg_smr!lgr_clsg_dtl_tsid = "cr" And rs_lgr_clsg_smr!lgr_clsg_dtl_tbal <> 0 Then
                
                grid_report.Rows = grid_report.Rows + 1
                grid_report.TextMatrix(b, 0) = b
                grid_report.TextMatrix(b, 1) = rs_lgr_clsg_smr!lgr_clsg_dtl_name
                
                
                grid_report.TextMatrix(b, 3) = Format(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal, "0.00")
                total_cr_balance = total_cr_balance + Val(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal)
                b = b + 1

                End If
                
rs_lgr_clsg_smr.MoveNext
Loop
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 2) = "==================="
grid_report.TextMatrix(b, 3) = "==================="
b = b + 1
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 2) = Format(total_dr_balance, "0.00")
grid_report.TextMatrix(b, 3) = Format(total_cr_balance, "0.00")
b = b + 1
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 2) = "==================="
grid_report.TextMatrix(b, 3) = "==================="
End Sub

Private Sub grid_report_DblClick()
    selected_ledger = grid_report.TextMatrix(grid_report.Row, 1)
    ledger_clicked_from_other = 1
    selected_procedure = "show ledger account"
    shw_sel_lgr_dtl.Show
End Sub

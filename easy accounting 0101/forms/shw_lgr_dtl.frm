VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form shw_lgr_dtl 
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "shw_lgr_dtl.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   200
      Top             =   12000
   End
   Begin VB.ComboBox combo_ledger 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   4
      Text            =   "Select a name"
      Top             =   960
      Width           =   3735
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2040
      TabIndex        =   3
      Text            =   "Select Option"
      Top             =   1560
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid grid_lgr_rep 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8281
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   9960
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   111673345
      CurrentDate     =   40126
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   111673345
      CurrentDate     =   40126
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   200
      TabIndex        =   12
      Top             =   11760
      Width           =   2190
   End
   Begin VB.Label lbl_head 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   12000
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
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   12000
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
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   12000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ledger"
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
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   8880
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   8880
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Type"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
End
Attribute VB_Name = "shw_lgr_dtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub combo_ledger_Click()
    Call read_ledger_report_data
End Sub
Private Sub Combo2_Click()
Dim today_day As Integer
Dim today_weekday As Integer
today_weekday = Weekday(Now)
today_day = Day(Now) - 1
If Combo2.Text = "This Week" Then
    DTPicker1.Value = Date - (today_weekday + 1)
    DTPicker2.Value = Date
ElseIf Combo2.Text = "This Year" Then
    DTPicker1.Value = this_year_starting_date
    DTPicker2.Value = this_year_ending_date
ElseIf Combo2.Text = "This Month" Then

    DTPicker1.Value = Date - today_day
    DTPicker2.Value = Date
ElseIf Combo2.Text = "Last Month" Then
    DTPicker1.Value = Day(Now) - today_day & "/" & Month(Now) - 1 & "/" & Year(Now)
    DTPicker2.Value = Date - (today_day + 1)
ElseIf Combo2.Text = "Last Week" Then
    DTPicker1.Value = Date - (today_weekday + 5)
    DTPicker2.Value = Date - (today_weekday - 1)
End If
Call read_ledger_report_data
End Sub

Private Sub Form_Activate()
show_ledger_detail = 0
End Sub

Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================

Label5.Left = 15500
Combo2.AddItem "This Year"
    Combo2.AddItem "This Month"
    
Combo2.AddItem "This Week"
Combo2.AddItem "Last Month"
Combo2.AddItem "Last Week"

Call add_ledgers
Call collect_all_transaction_in_one_file
Label5.Caption = "Select a account for detail..."

End Sub
Public Sub add_ledgers()
Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
combo_ledger.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
rs_lgr_main_dtl.MoveNext
Loop
End Sub
Public Sub read_ledger_report_data()
    rep_starting_date = DTPicker1.Value
    rep_ending_date = DTPicker2.Value
    selected_ledger = combo_ledger.Text
    grid_lgr_rep.Clear
    grid_lgr_rep.Rows = 1
    grid_lgr_rep.Cols = 8
    grid_lgr_rep.Font.Size = 12
    
    b = 0
    grid_lgr_rep.TextMatrix(b, 0) = ""
    grid_lgr_rep.TextMatrix(b, 1) = "Date"
    grid_lgr_rep.TextMatrix(b, 2) = "Voucher"
    grid_lgr_rep.TextMatrix(b, 3) = "V.No"
    grid_lgr_rep.TextMatrix(b, 4) = "Ledger"
    grid_lgr_rep.TextMatrix(b, 5) = "Dr.Amount"
    grid_lgr_rep.TextMatrix(b, 6) = "Cr.Amount"
    
    grid_lgr_rep.ColWidth(0) = 800
    grid_lgr_rep.ColWidth(1) = 1500
    grid_lgr_rep.ColWidth(2) = 1500
    grid_lgr_rep.ColWidth(3) = 1500
    grid_lgr_rep.ColWidth(4) = 5000
    grid_lgr_rep.ColWidth(5) = 2000
    grid_lgr_rep.ColWidth(6) = 2000
    grid_lgr_rep.ColWidth(7) = 2000
    
    Dim x_grid_col
    Dim total_grid_width
    total_grid_width = 500
    For x_grid_col = 0 To grid_lgr_rep.Cols - 1
        total_grid_width = total_grid_width + grid_lgr_rep.ColWidth(x_grid_col)
    Next
    grid_lgr_rep.Width = total_grid_width
    grid_lgr_rep.Left = (Me.Width - grid_lgr_rep.Width) / 2
    
    ledger_cr_total = 0
    ledger_dr_total = 0
    Dim x_int
    b = 1
    Call open_database
    If rs_acn_tran_all.State = 1 Then rs_acn_tran_all.Close
    rs_acn_tran_all.CursorLocation = adUseClient
    rs_acn_tran_all.Open " Select * From [acn_tran_all] order by fin_acnt_trn_date", db_co, adOpenDynamic, adLockPessimistic
    rs_acn_tran_all.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno" 'time"
    For x_int = 1 To rs_acn_tran_all.RecordCount
    If rs_acn_tran_all!fin_acnt_trn_ldgr = selected_ledger Then
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 0) = b
        grid_lgr_rep.TextMatrix(b, 1) = rs_acn_tran_all!fin_acnt_trn_date
        grid_lgr_rep.TextMatrix(b, 2) = UCase(rs_acn_tran_all!fin_acnt_trn_vchr)
        grid_lgr_rep.TextMatrix(b, 3) = rs_acn_tran_all!fin_acnt_trn_vcno
        grid_lgr_rep.TextMatrix(b, 4) = UCase(rs_acn_tran_all!fin_acnt_trn_ldgr)
        If rs_acn_tran_all!fin_acnt_trn_side = LCase("dr") Then
            grid_lgr_rep.TextMatrix(b, 5) = Format(rs_acn_tran_all!fin_acnt_trn_amnt, "0.00")
            ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_all!fin_acnt_trn_amnt)
        End If
        If rs_acn_tran_all!fin_acnt_trn_side = LCase("cr") Then
            grid_lgr_rep.TextMatrix(b, 6) = Format(rs_acn_tran_all!fin_acnt_trn_amnt, "0.00")
            ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_all!fin_acnt_trn_amnt)
        End If
            If ledger_dr_total > ledger_cr_total Then
                grid_lgr_rep.TextMatrix(b, 7) = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
            End If
            If ledger_cr_total > ledger_dr_total Then
                grid_lgr_rep.TextMatrix(b, 7) = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
            End If
        b = b + 1
    End If
    rs_acn_tran_all.MoveNext
    Next
    
    grid_lgr_rep.AddItem ""
    grid_lgr_rep.TextMatrix(b, 5) = "----------------------------"
    grid_lgr_rep.TextMatrix(b, 6) = "----------------------------"
    b = b + 1
    grid_lgr_rep.AddItem ""
    grid_lgr_rep.TextMatrix(b, 4) = "TOTAL AMT."
    grid_lgr_rep.TextMatrix(b, 5) = Format(ledger_dr_total, "0.00")
    grid_lgr_rep.TextMatrix(b, 6) = Format(ledger_cr_total, "0.00")
    b = b + 1
    
    Label5.Top = grid_lgr_rep.Top + grid_lgr_rep.Height
    'grid_lgr_rep.Sort = 1
    If ledger_dr_total > ledger_cr_total Then
        Label5.Caption = combo_ledger.Text & " Dr. Balance is..... " & Format(ledger_dr_total - ledger_cr_total, "0.00")
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 4) = " Dr. Balance is..... "
        grid_lgr_rep.TextMatrix(b, 6) = Format(ledger_dr_total - ledger_cr_total, "0.00")
        b = b + 1
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 5) = "===================="
        grid_lgr_rep.TextMatrix(b, 6) = "===================="
        
        b = b + 1
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 5) = Format(ledger_dr_total, "0.00")
        grid_lgr_rep.TextMatrix(b, 6) = Format(ledger_dr_total, "0.00")
    End If
    'grid_lgr_rep.TextMatrix(b, 6) = Format(ledger_cr_total, "0.00")
        
    
    If ledger_cr_total > ledger_dr_total Then
        Label5.Caption = combo_ledger.Text & " Cr. Balance is..... " & Format(ledger_cr_total - ledger_dr_total, "0.00")
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 4) = " Cr. Balance is..... "
        grid_lgr_rep.TextMatrix(b, 6) = Format(ledger_cr_total - ledger_dr_total, "0.00")
        b = b + 1
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 5) = "======================"
        grid_lgr_rep.TextMatrix(b, 6) = "======================"
        b = b + 1
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 5) = Format(ledger_cr_total, "0.00")
        grid_lgr_rep.TextMatrix(b, 6) = Format(ledger_cr_total, "0.00")
    End If
    b = b + 1
    grid_lgr_rep.AddItem ""
    grid_lgr_rep.TextMatrix(b, 5) = "====================="
    grid_lgr_rep.TextMatrix(b, 6) = "====================="

'grid_lgr_rep.Sort = 2
End Sub


Private Sub grid_lgr_rep_DblClick()
If grid_lgr_rep.TextMatrix(grid_lgr_rep.Row, 0) = "" Then
    MsgBox "This is a invalid entry....,"
    Exit Sub
Else
    selected_voucher_date = grid_lgr_rep.TextMatrix(grid_lgr_rep.Row, 1)
    selected_voucher_name = LCase(grid_lgr_rep.TextMatrix(grid_lgr_rep.Row, 2))
    selected_voucher_no = grid_lgr_rep.TextMatrix(grid_lgr_rep.Row, 3)
    
    show_ledger_detail = 1
    If selected_voucher_name = "payment" Then
    selected_procedure = "payment voucher"
    vchr_payment.Show
    ElseIf selected_voucher_name = "receipt" Then
    selected_procedure = "Receipt voucher"
    vchr_receipt.Show
    ElseIf selected_voucher_name = "sale" Then
            selected_procedure = "Sales voucher"
        vchr_sales.Show
    ElseIf selected_voucher_name = "purchase" Then
    
    ElseIf selected_voucher_name = "contra" Then
    selected_procedure = "Banking voucher"
    vchr_contra.Show
    ElseIf selected_voucher_name = "journal" Then
    selected_procedure = "Adjustment/Journal voucher"
    vchr_Journal.Show
    'ElseIf selected_voucher_name = "payment" Then
    'ElseIf selected_voucher_name = "payment" Then
    End If
End If

End Sub

Private Sub Timer1_Timer()
If Label5.Left + Label5.Width <= 0 Then Label5.Left = 13500 + Label5.Width
Label5.Left = Label5.Left - 500
End Sub

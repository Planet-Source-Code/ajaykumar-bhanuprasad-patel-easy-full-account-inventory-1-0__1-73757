VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form disp_sales_man_repo 
   BackColor       =   &H00C0FFFF&
   Caption         =   "show_selected_ledger_dtl"
   ClientHeight    =   8430
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   8760
   Icon            =   "disp_sales_man_repo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid grid_report_sales 
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4800
      Top             =   11880
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
      Left            =   4080
      TabIndex        =   5
      Text            =   "Select Option"
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   -500
      Width           =   2655
   End
   Begin VB.ListBox combo_list 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox selected_from_list 
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
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid grid_report 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4048
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   10200
      TabIndex        =   6
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
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
      CurrentDate     =   40126
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   8040
      TabIndex        =   7
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
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
      CurrentDate     =   40126
   End
   Begin VB.Label sales_man_lbl2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Sales Man Transactions....,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   14
      Top             =   3960
      Width           =   5145
   End
   Begin VB.Label sales_man_lbl1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Sales Man customers summery....,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   13
      Top             =   1200
      Width           =   6165
   End
   Begin VB.Label m_label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Sales man Report...,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   11
      Top             =   6960
      Width           =   5220
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   3360
      TabIndex        =   10
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   9720
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Period."
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
      Left            =   7080
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Sales man Report...,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "disp_sales_man_repo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private total_cr_balance
Private total_dr_balance
Private total_credit_limit

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
selected_from_list.Text = ""
Combo2.AddItem "This Year"
    Combo2.AddItem "This Month"
    
Combo2.AddItem "This Week"
Combo2.AddItem "Last Month"
Combo2.AddItem "Last Week"
Call add_combo_list
selected_date = Date
Call make_trail_balance_summary
combo_list.Visible = False
Dim today_day As Integer
Dim today_weekday As Integer
today_weekday = Weekday(Now)
today_day = Day(Now) - 1
DTPicker1.Value = Date - today_day
DTPicker2.Value = Date
End Sub
Private Sub combo_list_LostFocus()
combo_list.Visible = False
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
    If Month(Now) = 1 Then
        DTPicker1.Value = Day(Now) - today_day & "/" & 12 & "/" & Year(Now) - 1
    Else
        DTPicker1.Value = Day(Now) - today_day & "/" & Month(Now) - 1 & "/" & Year(Now)
    End If
    DTPicker2.Value = Date - (today_day + 1)
ElseIf Combo2.Text = "Last Week" Then
    DTPicker1.Value = Date - (today_weekday + 5)
    DTPicker2.Value = Date - (today_weekday - 1)
End If

    Call set_sales_man_report_grid
    Call set_grid_report
    Call set_sales_man_detail

End Sub
Private Sub DTPicker1_Change()
Dim today_day As Integer
Dim today_weekday As Integer
today_weekday = Weekday(Now)
today_day = Day(Now) - 1

    Call set_sales_man_report_grid
    Call set_grid_report
    Call set_sales_man_detail
    
End Sub
Private Sub DTPicker2_Change()
Dim today_day As Integer
Dim today_weekday As Integer
today_weekday = Weekday(Now)
today_day = Day(Now) - 1

    Call set_sales_man_report_grid
    Call set_grid_report
    Call set_sales_man_detail
    
End Sub
Private Sub combo_list_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    selected_from_list.Text = combo_list.Text
    combo_list.Visible = False
    Combo2.Text = "This Month"
    Call set_sales_man_report_grid
    Call set_grid_report
    Call set_sales_man_detail
    Text2.SetFocus

End If
sales_man_lbl1.Caption = UCase(selected_sales_man) & "'s Balance Summary as on " & Date
sales_man_lbl2.Caption = UCase(selected_sales_man) & "'s Sales and Cash collection for the period of " & DTPicker1.Value & " to " & DTPicker2.Value

Label1.Caption = "Report of " & selected_sales_man
End Sub

Private Sub selected_from_list_GotFocus()
    combo_list.Visible = True
    combo_list.Height = 2400
    combo_list.SetFocus
End Sub
Public Sub add_combo_list()
Call open_database
Call open_rs_emp_main_dtl
Do Until rs_emp_main_dtl.EOF
    combo_list.AddItem UCase(rs_emp_main_dtl!emp_main_dtl_name)
    rs_emp_main_dtl.MoveNext
Loop
End Sub
Public Sub set_grid_report()

total_credit_limit = 0
total_cr_balance = 0
total_dr_balance = 0
    
    grid_report.Clear
    grid_report.Rows = 1
    grid_report.Cols = 5
    grid_report.Font.Size = 12
    b = 0
    grid_report.Font.Size = 12
    grid_report.TextMatrix(b, 0) = ""
    grid_report.TextMatrix(b, 1) = "Ledger"
    grid_report.TextMatrix(b, 2) = "Dr.Amount"
    grid_report.TextMatrix(b, 3) = "Cr.Amount"
    grid_report.TextMatrix(b, 4) = "Credit Limit"
    
    grid_report.ColWidth(0) = 800
    grid_report.ColWidth(1) = 8000
    grid_report.ColWidth(2) = 2500
    grid_report.ColWidth(3) = 2500
    grid_report.ColWidth(4) = 2500
    
    Dim x_grid_col
    Dim total_grid_width
    total_grid_width = 500
    For x_grid_col = 0 To grid_report.Cols - 1
        total_grid_width = total_grid_width + grid_report.ColWidth(x_grid_col)
    Next
    grid_report.Width = total_grid_width
    b = 1
    
    rep_starting_date = DTPicker1.Value
    rep_ending_date = DTPicker2.Value
    
    selected_sales_man = selected_from_list.Text
'Call open_rs_lgr_main_dtl
'Do Until rs_lgr_main_dtl.EOF
Call open_rs_lgr_clsg_smr
rs_lgr_clsg_smr.Sort = "lgr_clsg_dtl_name"
Do Until rs_lgr_clsg_smr.EOF
        If LCase(rs_lgr_clsg_smr!lgr_clsg_dtl_slun) = LCase(selected_sales_man) Then
                    grid_report.Rows = grid_report.Rows + 1
                    grid_report.TextMatrix(b, 0) = b
                    grid_report.TextMatrix(b, 1) = rs_lgr_clsg_smr!lgr_clsg_dtl_name
                    'rs_lgr_clsg_smr!lgr_clsg_dtl_grup = rs_lgr_main_dtl!lgr_main_dtl_grup
                    'rs_lgr_clsg_smr!lgr_clsg_dtl_pgrp = selected_primary_group
                    'rs_lgr_clsg_smr!lgr_clsg_dtl_crpd = rs_lgr_main_dtl!lgr_main_dtl_crpd
                    grid_report.TextMatrix(b, 4) = Format(rs_lgr_clsg_smr!lgr_clsg_dtl_cram, "0.00")
                    total_credit_limit = total_credit_limit + Val(rs_lgr_clsg_smr!lgr_clsg_dtl_cram)
                    'rs_lgr_clsg_smr!lgr_clsg_dtl_bal1 = rs_lgr_main_dtl!lgr_main_dtl_obl1
                    'rs_lgr_clsg_smr!lgr_clsg_dtl_bal2 = rs_lgr_main_dtl!lgr_main_dtl_obl2
                    'rs_lgr_clsg_smr!lgr_clsg_dtl_sid1 = rs_lgr_main_dtl!lgr_main_dtl_osd1
                    'rs_lgr_clsg_smr!lgr_clsg_dtl_sid2 = rs_lgr_main_dtl!lgr_main_dtl_osd2
                    'rs_lgr_clsg_smr!lgr_clsg_dtl_slun = rs_lgr_main_dtl!lgr_main_dtl_slun
                    If rs_lgr_clsg_smr!lgr_clsg_dtl_tsid = "dr" Then
                    grid_report.TextMatrix(b, 2) = Format(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal, "0.00")
                    total_dr_balance = total_dr_balance + Val(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal)
                    ElseIf rs_lgr_clsg_smr!lgr_clsg_dtl_tsid = "cr" Then
                    grid_report.TextMatrix(b, 3) = Format(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal, "0.00")
                    total_cr_balance = total_cr_balance + Val(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal)
                    End If
        b = b + 1
        End If
        rs_lgr_clsg_smr.MoveNext
    Loop
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 2) = "======================="
grid_report.TextMatrix(b, 3) = "======================="
grid_report.TextMatrix(b, 4) = "======================="
b = b + 1
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 2) = Format(total_dr_balance, "0.00")
grid_report.TextMatrix(b, 3) = Format(total_cr_balance, "0.00")
grid_report.TextMatrix(b, 4) = Format(total_credit_limit, "0.00")

b = b + 1
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 2) = "======================="
grid_report.TextMatrix(b, 3) = "======================="
grid_report.TextMatrix(b, 4) = "======================="

End Sub
Private Sub grid_report_Click()
    selected_ledger = grid_report.TextMatrix(grid_report.Row, 1)
    ledger_clicked_from_other = 1
    selected_procedure = "show ledger account"
    shw_sel_lgr_dtl.Show
End Sub
Private Sub Timer1_Timer()
If m_label.Left + m_label.Width <= 0 Then
    m_label.Left = Me.Width ' + m_label.Width
    If r_clr >= 250 Then
        g_clr = 50
        b_clr = b_clr + 50
    End If
    If g_clr >= 250 Then
        b_clr = 50
        r_clr = r_clr + 50
    End If
    If b_clr >= 250 Then
        r_clr = 50
        g_clr = g_clr + 50
    End If
End If
m_label.Left = m_label.Left - 250
m_label.ForeColor = RGB(r_clr, g_clr, b_clr)
End Sub
Public Sub set_sales_man_report_grid()
    grid_report.Clear
    grid_report_sales.Clear
    grid_report_sales.Rows = 1
    grid_report_sales.Cols = 10
    grid_report_sales.Font.Size = 12
    b = 0
    grid_report_sales.TextMatrix(b, 0) = ""
    grid_report_sales.TextMatrix(b, 1) = "Date"
    grid_report_sales.TextMatrix(b, 2) = "Voucher"
    grid_report_sales.TextMatrix(b, 3) = "V.No"
    grid_report_sales.TextMatrix(b, 4) = "Ledger"
    grid_report_sales.TextMatrix(b, 5) = "Sales"
    grid_report_sales.TextMatrix(b, 6) = "Cash/card Rtn"
    grid_report_sales.TextMatrix(b, 7) = "Balance"
    grid_report_sales.TextMatrix(b, 8) = "Time"
    grid_report_sales.TextMatrix(b, 9) = "User"
    
    grid_report_sales.ColWidth(0) = 800
    grid_report_sales.ColWidth(1) = 1500
    grid_report_sales.ColWidth(2) = 2000
    grid_report_sales.ColWidth(3) = 700
    grid_report_sales.ColWidth(4) = 4500
    grid_report_sales.ColWidth(5) = 1700
    grid_report_sales.ColWidth(6) = 1700
    grid_report_sales.ColWidth(7) = 1700
    grid_report_sales.ColWidth(8) = 1000
    grid_report_sales.ColWidth(9) = 1000
    Dim x_grid_col
    Dim total_grid_width
    total_grid_width = 500
    For x_grid_col = 0 To grid_report_sales.Cols - 1
        total_grid_width = total_grid_width + grid_report_sales.ColWidth(x_grid_col)
    Next
    grid_report_sales.Width = total_grid_width
    
End Sub
Public Sub set_sales_man_detail()
b = b + 1
Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
    If LCase(rs_lgr_main_dtl!lgr_main_dtl_slun) = LCase(selected_sales_man) Then
    
        selected_ledger = rs_lgr_main_dtl!lgr_main_dtl_name
        selected_voucher_ledger = selected_ledger
        'Call set_selected_ledger_detail
        rep_starting_date = DTPicker1.Value
        rep_ending_date = DTPicker2.Value
        ledger_dr_total = 0
        ledger_cr_total = 0
        
        Call copy_specific_ledger_transaction_to_acn_tran_spc_lgr
        
        Call open_rs_acn_tran_spc_lgr
        rs_acn_tran_spc_lgr.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno" 'time"
        Dim temp_starting_dt As Date
        Dim temp_ending_dt As Date
        
        Do Until rs_acn_tran_spc_lgr.EOF
            If rs_acn_tran_spc_lgr!fin_acnt_trn_date >= rep_starting_date And _
            rs_acn_tran_spc_lgr!fin_acnt_trn_date <= rep_ending_date Then
            grid_report_sales.AddItem ""
            grid_report_sales.TextMatrix(b, 1) = rs_acn_tran_spc_lgr!fin_acnt_trn_date
            grid_report_sales.TextMatrix(b, 2) = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_vchr)
            grid_report_sales.TextMatrix(b, 3) = rs_acn_tran_spc_lgr!fin_acnt_trn_vcno
            grid_report_sales.TextMatrix(b, 4) = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_ldgr)
            If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") Then
                grid_report_sales.TextMatrix(b, 5) = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
                ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
            End If
            If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") Then
                grid_report_sales.TextMatrix(b, 6) = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
                ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
            End If
            If ledger_dr_total > ledger_cr_total Then
                grid_report_sales.TextMatrix(b, 7) = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
            End If
            If ledger_cr_total > ledger_dr_total Then
                grid_report_sales.TextMatrix(b, 7) = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
            End If
            If rs_acn_tran_spc_lgr!fin_acnt_trn_time <> "" Then grid_report_sales.TextMatrix(b, 8) = rs_acn_tran_spc_lgr!fin_acnt_trn_time
            If rs_acn_tran_spc_lgr!fin_acnt_trn_user <> "" Then grid_report_sales.TextMatrix(b, 9) = rs_acn_tran_spc_lgr!fin_acnt_trn_user
            If LCase(grid_report_sales.TextMatrix(b, 2)) = "opening balance 1" Or LCase(grid_report_sales.TextMatrix(b, 2)) = "opening balance 2" Then
                grid_report_sales.TextMatrix(b, 4) = grid_report_sales.TextMatrix(b, 2)
                grid_report_sales.TextMatrix(b, 3) = ""
                grid_report_sales.TextMatrix(b, 2) = ""
                'grid_report_sales.TextMatrix(b, 4) = "Opeining Balance is ..., "
            End If
            b = b + 1
        'ElseIf rs_acn_tran_spc_lgr!fin_acnt_trn_date < rep_starting_date Then
        '    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") Then
        '        ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
        '    End If
        '    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") Then
        '        ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
        '    End If
        '    b = 1
        End If
        rs_acn_tran_spc_lgr.MoveNext
        Loop
End If
rs_lgr_main_dtl.MoveNext
Loop
        
If b >= 2 Then
        ledger_dr_total = 0
        ledger_cr_total = 0
        Dim grid_row_counter
        For grid_row_counter = 1 To (b - 1)
            If grid_report_sales.TextMatrix(grid_row_counter, 5) <> "" Then ledger_dr_total = ledger_dr_total + Val(grid_report_sales.TextMatrix(grid_row_counter, 5))
            If grid_report_sales.TextMatrix(grid_row_counter, 6) <> "" Then ledger_cr_total = ledger_cr_total + Val(grid_report_sales.TextMatrix(grid_row_counter, 6))
        Next
            If ledger_dr_total < ledger_cr_total Then
                
                grid_report_sales.AddItem ""
                grid_report_sales.TextMatrix(b, 4) = "Cash Recovery over sale" 'Closing Balance is.....Cr."
                grid_report_sales.TextMatrix(b, 5) = Format(ledger_cr_total - ledger_dr_total, "0.00")
                b = b + 1
                grid_report_sales.AddItem ""
                grid_report_sales.TextMatrix(b, 5) = "================="
                grid_report_sales.TextMatrix(b, 6) = "================="
                b = b + 1
                grid_report_sales.AddItem ""
                grid_report_sales.TextMatrix(b, 5) = Format(ledger_cr_total, "0.00")
                grid_report_sales.TextMatrix(b, 6) = Format(ledger_cr_total, "0.00")
                b = b + 1
                grid_report_sales.AddItem ""
                grid_report_sales.TextMatrix(b, 5) = "================="
                grid_report_sales.TextMatrix(b, 6) = "================="
            End If
            'grid_report_sales.TextMatrix(b, 6) = Format(ledger_cr_total, "0.00")
            If ledger_cr_total < ledger_dr_total Then
                m_label.Caption = selected_sales_man & "'s Total Balance as on " & Date & "are " & ledger_cr_total - ledger_dr_total & selected_sales_man & "'s Sales for the period of " & DTPicker1.Value & " to " & DTPicker2.Value & " are..." & total_dr_balance & " and Cash collection are ...." & total_cr_balance
                grid_report_sales.AddItem ""
                grid_report_sales.TextMatrix(b, 4) = "Short collection of Cash then Sale....." 'Closing Balance is.....Dr. "
                grid_report_sales.TextMatrix(b, 6) = Format(ledger_dr_total - ledger_cr_total, "0.00")
                b = b + 1
                grid_report_sales.AddItem ""
                grid_report_sales.TextMatrix(b, 5) = "================="
                grid_report_sales.TextMatrix(b, 6) = "================="
                b = b + 1
                grid_report_sales.AddItem ""
                grid_report_sales.TextMatrix(b, 4) = "TOTAL AMOUNT..."
                grid_report_sales.TextMatrix(b, 5) = Format(ledger_dr_total, "0.00")
                grid_report_sales.TextMatrix(b, 6) = Format(ledger_dr_total, "0.00")
                b = b + 1
                grid_report_sales.AddItem ""
                grid_report_sales.TextMatrix(b, 5) = "================="
                grid_report_sales.TextMatrix(b, 6) = "================="
            End If
        End If
        
        m_label.Caption = UCase(selected_sales_man) & "'s Total Due Balance as on " & Date & "are " & Format(total_dr_balance - total_cr_balance, "0.00") & " " & UCase(selected_sales_man) & "'s Sales for the period of " & DTPicker1.Value & " to " & DTPicker2.Value & " are..." & Format(ledger_dr_total, "0.00") & "£ and Cash collection are ...." & Format(ledger_cr_total, "0.00") & "£"
End Sub

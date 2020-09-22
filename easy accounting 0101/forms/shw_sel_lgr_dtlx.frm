VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form shw_sel_lgr_dtlx 
   Caption         =   "show_selected_ledger_dtl"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_print 
      Caption         =   "Print"
      Height          =   615
      Left            =   16200
      TabIndex        =   13
      Top             =   10080
      Width           =   1935
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
      Left            =   8400
      TabIndex        =   5
      Text            =   "Select Option"
      Top             =   1080
      Width           =   3375
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
      Height          =   1560
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Width           =   3855
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
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid grid_report 
      Height          =   8055
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   14208
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   15960
      TabIndex        =   6
      Top             =   1080
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
      Format          =   97189889
      CurrentDate     =   40126
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   13200
      TabIndex        =   7
      Top             =   1080
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
      Format          =   97189889
      CurrentDate     =   40126
   End
   Begin VB.Label Label5 
      Caption         =   "Ledger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label m_label 
      AutoSize        =   -1  'True
      Caption         =   "m_label"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   840
      TabIndex        =   11
      Top             =   10320
      Width           =   1800
   End
   Begin VB.Label Label4 
      Caption         =   "Report option"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label3 
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
      Left            =   15480
      TabIndex        =   9
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Period...,"
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
      Left            =   11880
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Ledger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8715
      TabIndex        =   3
      Top             =   360
      Width           =   1605
   End
End
Attribute VB_Name = "shw_sel_lgr_dtlx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub set_ledger_account_detailx()

rep_starting_date = DTPicker1.Value
rep_ending_date = DTPicker2.Value
selected_ledger = selected_from_list.Text
selected_voucher_ledger = selected_ledger
ledger_dr_total = 0
ledger_cr_total = 0

Call set_report_grid
Call open_database
Call copy_specific_ledger_transaction_to_acn_tran_spc_lgr
Call open_rs_acn_tran_spc_lgr
Call open_rs_acn_tran_spc_lgr_print

rs_acn_tran_spc_lgr.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno" 'time"
Dim temp_starting_dt As Date
Dim temp_ending_dt As Date
b = 1

fin_acnt_trn_date '1
fin_acnt_trn_vchr '2
fin_acnt_trn_vcno '3
fin_acnt_trn_ldgr '4
fin_acnt_trn_dram '5
fin_acnt_trn_cram '6
fin_acnt_trn_blnc '7

Do Until rs_acn_tran_spc_lgr.EOF

If rs_acn_tran_spc_lgr!fin_acnt_trn_date >= rep_starting_date And rs_acn_tran_spc_lgr!fin_acnt_trn_date <= rep_ending_date Then

If b = 1 Then
    rs_acn_tran_spc_lgr_print.AddNew
    With rs_acn_tran_spc_lgr_print
    If ledger_dr_total > ledger_cr_total Then
        temp_opening_balance = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        !fin_acnt_trn_vcno = ""
        !fin_acnt_trn_ldgr = "Opeining Balance is ..., " 'Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        !fin_acnt_trn_dram = Format(ledger_dr_total - ledger_cr_total, "0.00") '& " Dr."
        !fin_acnt_trn_cram = ""
        !fin_acnt_trn_blnc = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
    ElseIf ledger_cr_total > ledger_dr_total Then
        temp_opening_balance = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
        !fin_acnt_trn_vcno = ""
        !fin_acnt_trn_ldgr = "Opeining Balance is ..., "
        !fin_acnt_trn_dram = ""
        !fin_acnt_trn_cram = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
        !fin_acnt_trn_blnc = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
    End If
b = b + 1
End If
    
    grid_report.AddItem ""
    'grid_report.TextMatrix(b, 0) = b
    !fin_acnt_trn_date = rs_acn_tran_spc_lgr!fin_acnt_trn_date
    !fin_acnt_trn_vchr = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_vchr)
    !fin_acnt_trn_vcno = rs_acn_tran_spc_lgr!!fin_acnt_trn_vcno
    !fin_acnt_trn_ldgr = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_ldgr)
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") Then
        !fin_acnt_trn_dram = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
        ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") Then
        !fin_acnt_trn_cram = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
        ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    If ledger_dr_total > ledger_cr_total Then
        !fin_acnt_trn_blnc = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
    End If
    If ledger_cr_total > ledger_dr_total Then
        !fin_acnt_trn_blnc = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
    End If
    If rs_acn_tran_spc_lgr!fin_acnt_trn_time <> "" Then grid_report.TextMatrix(b, 8) = rs_acn_tran_spc_lgr!fin_acnt_trn_time
    If rs_acn_tran_spc_lgr!fin_acnt_trn_user <> "" Then grid_report.TextMatrix(b, 9) = rs_acn_tran_spc_lgr!fin_acnt_trn_user
    
    If LCase(grid_report.TextMatrix(b, 2)) = "opening balance 1" Or LCase(grid_report.TextMatrix(b, 2)) = "opening balance 2" Then
        !fin_acnt_trn_ldgr = grid_report.TextMatrix(b, 2)
        !fin_acnt_trn_vcno = ""
        !fin_acnt_trn_vchr = ""
        '!fin_acnt_trn_ldgr= "Opeining Balance is ..., "
    End If
    b = b + 1
ElseIf rs_acn_tran_spc_lgr!fin_acnt_trn_date < rep_starting_date Then
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") Then
        ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") Then
        ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    b = 1
End If
rs_acn_tran_spc_lgr.MoveNext
Loop

If b = 1 Then
If ledger_dr_total > 0 Or ledger_cr_total > 0 Then
    grid_report.AddItem ""
    If ledger_dr_total > ledger_cr_total Then
        temp_opening_balance = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        !fin_acnt_trn_vcno = ""
        !fin_acnt_trn_ldgr = "Opeining Balance is ..., " 'Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        !fin_acnt_trn_dram = Format(ledger_dr_total - ledger_cr_total, "0.00") '& " Dr."
        !fin_acnt_trn_cram = ""
        !fin_acnt_trn_blnc = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
    ElseIf ledger_cr_total > ledger_dr_total Then
        temp_opening_balance = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
        !fin_acnt_trn_vcno = ""
        !fin_acnt_trn_ldgr = "Opeining Balance is ..., "
        !fin_acnt_trn_dram = ""
        !fin_acnt_trn_cram = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
        !fin_acnt_trn_blnc = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
    End If
    b = b + 1
End If
End If

If b >= 2 Then
ledger_dr_total = 0
ledger_cr_total = 0
Dim grid_row_counter
For grid_row_counter = 1 To (b - 1)
    If grid_report.TextMatrix(grid_row_counter, 5) <> "" Then ledger_dr_total = ledger_dr_total + Val(grid_report.TextMatrix(grid_row_counter, 5))
    If grid_report.TextMatrix(grid_row_counter, 6) <> "" Then ledger_cr_total = ledger_cr_total + Val(grid_report.TextMatrix(grid_row_counter, 6))
Next
If ledger_dr_total < ledger_cr_total Then
        m_label.Caption = selected_voucher_ledger & " Closing Balance is..... Cr. " & Format(ledger_cr_total - ledger_dr_total, "0.00") & " on " & rep_ending_date
        grid_report.AddItem ""
        !fin_acnt_trn_ldgr = "Closing Balance is.....Cr."
        !fin_acnt_trn_dram = Format(ledger_cr_total - ledger_dr_total, "0.00")
        b = b + 1
        grid_report.AddItem ""
        !fin_acnt_trn_dram = "================="
        !fin_acnt_trn_cram = "================="
        b = b + 1
        grid_report.AddItem ""
        !fin_acnt_trn_dram = Format(ledger_cr_total, "0.00")
        !fin_acnt_trn_cram = Format(ledger_cr_total, "0.00")
        b = b + 1
        grid_report.AddItem ""
        !fin_acnt_trn_dram = "================="
        !fin_acnt_trn_cram = "================="

    End If
    '!fin_acnt_trn_cram= Format(ledger_cr_total, "0.00")
    If ledger_cr_total < ledger_dr_total Then
        m_label.Caption = selected_voucher_ledger & " Closing Balance is..... Dr. " & Format(ledger_dr_total - ledger_cr_total, "0.00") & " on " & rep_ending_date
        grid_report.AddItem ""
        !fin_acnt_trn_ldgr = " Closing Balance is.....Dr. "
        !fin_acnt_trn_cram = Format(ledger_dr_total - ledger_cr_total, "0.00")
        b = b + 1
        grid_report.AddItem ""
        !fin_acnt_trn_dram = "================="
        !fin_acnt_trn_cram = "================="
        b = b + 1
        grid_report.AddItem ""
        !fin_acnt_trn_ldgr = "TOTAL AMOUNT..."
        !fin_acnt_trn_dram = Format(ledger_dr_total, "0.00")
        !fin_acnt_trn_cram = Format(ledger_dr_total, "0.00")
        b = b + 1
        grid_report.AddItem ""
        !fin_acnt_trn_dram = "================="
        !fin_acnt_trn_cram = "================="
    End If
End If

Dim temp_entry_no
Dim temp_v_no
Dim temp_v_tp
For temp_entry_no = 1 To grid_report.Rows - 3
If grid_report.TextMatrix(temp_entry_no, 2) <> "" Then
    temp_v_tp = grid_report.TextMatrix(temp_entry_no, 2)
    temp_v_no = Val(grid_report.TextMatrix(temp_entry_no, 3))
    Call open_rs_acn_tran_all
    rs_acn_tran_all.MoveFirst
    Do Until rs_acn_tran_all.EOF
            If LCase(selected_ledger) <> LCase(rs_acn_tran_all!fin_acnt_trn_ldgr) And LCase(temp_v_tp) = LCase(rs_acn_tran_all!fin_acnt_trn_vchr) And temp_v_no = Val(rs_acn_tran_all!!fin_acnt_trn_vcno) Then
                grid_report.TextMatrix(temp_entry_no, 4) = LCase(rs_acn_tran_all!fin_acnt_trn_ldgr)
                Exit Do
            End If
        
        rs_acn_tran_all.MoveNext
    Loop
End If
Next
End Sub

Private Sub cmd_print_Click()

rep_starting_date = DTPicker1.Value
rep_ending_date = DTPicker2.Value
selected_ledger = selected_from_list.Text
selected_voucher_ledger = selected_ledger
ledger_dr_total = 0
ledger_cr_total = 0
Call set_report_grid
Call open_database
Call copy_specific_ledger_transaction_to_acn_tran_spc_lgr
rs_acn_tran_spc_lgr_print

With report_ledger_ac.Sections("section2").Controls
    .item("label3").Caption = sale_vchr_customer
    .item("label4").Caption = sale_vchr_address1 & "," & sale_vchr_address2
    .item("label5").Caption = sale_vchr_city
    '.item("label6").Caption = sale_vchr_telephone & " , " & sale_vchr_mobile
    .item("label7").Caption = sale_vchr_Date
    .item("label8").Caption = sale_vchr_user
    .item("label9").Caption = sale_vchr_day '& "( " & sale_vchr_time & ") "
    .item("label10").Caption = sale_vchr_invoiceno
    .item("label11").Caption = sale_vchr_transport
End With
'Set da_find_card_report.DataSource = rs_dap_main_dtl_temp
'da_find_card_report.Show
'report_ledger_ac.Sections("section1").Height

Call open_database
rs_report_ledger_ac.CursorLocation = adUseClient
rs_report_ledger_ac.Open "Select * From inv_tran_sal WHERE stk_invt_trn_vcno = " & Val(Text5.Text) & "", db_co, adOpenDynamic, adLockPessimistic
While rs_report_ledger_ac.EOF = False
Set report_ledger_ac.DataSource = rs_report_ledger_ac

With report_ledger_ac.Sections("section5").Controls
'.item("label14").Caption = "Balance as on..." & sale_vchr_Date & " (Before this Invoice): " & customer_balance_is
.item("label25").Caption = Format(sales_vchr_totalamt, "0.00")                                          'TOTAL VOUCHER AMOUNT
.item("label24").Caption = lbl_total_qty.Caption                                                        'TOTAL QUANTITY
If show_ledger_detail = 0 Then
.item("label23").Caption = Format(((ledger_dr_total - ledger_cr_total)), "0.00")                        'PREVIOUS BALANCE
.item("label27").Caption = Format((ledger_dr_total - ledger_cr_total) + sales_vchr_totalamt, "0.00")    'TOTAL VOUCHER AMOUNT + PERIVOUS BALANCE
Else
.item("label22").Visible = False
.item("label23").Visible = False
.item("label26").Caption = "Net Due Balance"
.item("label27").Caption = Format((ledger_dr_total - ledger_cr_total), "0.00")    'TOTAL VOUCHER AMOUNT + PERIVOUS BALANCE
.item("line5").Visible = False
'.item("line6").Visible = False
.item("shape10").Visible = False
'.item("shape11").Visible = False
End If
End With
report_ledger_ac.Show
rs_report_ledger_ac.MoveNext
Wend

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
Call set_ledger_account_detail
End Sub
Private Sub DTPicker1_Change()
Dim today_day As Integer
Dim today_weekday As Integer
today_weekday = Weekday(Now)
today_day = Day(Now) - 1
Call set_ledger_account_detail
Label1.Caption = selected_ledger & "   (" & DTPicker1.Value & "  To  " & DTPicker2.Value & ")"
End Sub
Private Sub DTPicker2_Change()
Dim today_day As Integer
Dim today_weekday As Integer
today_weekday = Weekday(Now)
today_day = Day(Now) - 1
Call set_ledger_account_detail
Label1.Caption = selected_ledger & "   (" & DTPicker1.Value & "  To  " & DTPicker2.Value & ")"
End Sub
Private Sub combo_list_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    selected_from_list.Text = combo_list.Text
    combo_list.Visible = False
    Combo2.Text = "This Month"
    Dim today_day As Integer
    Dim today_weekday As Integer
    today_weekday = Weekday(Now)
    today_day = Day(Now) - 1
    DTPicker1.Value = Date - today_day
    DTPicker2.Value = Date
    Call set_ledger_account_detail
    Text2.SetFocus
End If
Label1.Caption = selected_ledger & "   (" & DTPicker1.Value & "  To  " & DTPicker2.Value & ")"
End Sub

Private Sub selected_from_list_GotFocus()
    combo_list.Visible = True
    combo_list.Height = 2400
    combo_list.SetFocus
End Sub
Public Sub add_combo_list()
Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
combo_list.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
rs_lgr_main_dtl.MoveNext
Loop
End Sub
Private Sub Form_Activate()
show_ledger_detail = 0
Call set_ledger_account_detail
End Sub
Private Sub Form_Load()
selected_from_list.Text = ""
Combo2.AddItem "This Month"
Combo2.AddItem "This Week"
Combo2.AddItem "Last Month"
Combo2.AddItem "Last Week"
Call add_combo_list
Call collect_all_transaction_in_one_file
combo_list.Visible = False
If ledger_clicked_from_other = 1 Then
    combo_list.Text = selected_ledger
    selected_from_list.Text = combo_list.Text
    combo_list.Visible = False
    Combo2.Text = "This Month"
    Dim today_day As Integer
    Dim today_weekday As Integer
    today_weekday = Weekday(Now)
    today_day = Day(Now) - 1
    DTPicker1.Value = selected_companies_starting_f_date
    DTPicker2.Value = selected_date
    Call set_ledger_account_detail
End If
End Sub
Private Sub grid_report_DblClick()
If grid_report.TextMatrix(grid_report.Row, 3) = "" Then
    MsgBox "This is a invalid entry....,"
    Exit Sub
Else
    selected_voucher_date = grid_report.TextMatrix(grid_report.Row, 1)
    selected_voucher_name = LCase(grid_report.TextMatrix(grid_report.Row, 2))
    selected_voucher_no = grid_report.TextMatrix(grid_report.Row, 3)
    
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
    selected_procedure = "purchase voucher"
    vchr_purchase.Show
    ElseIf selected_voucher_name = "contra" Then
    selected_procedure = "Banking voucher"
    vchr_contra.Show
    ElseIf selected_voucher_name = "journal" Then
    selected_procedure = "Adjustment/Journal voucher"
    vchr_Journal.Show
    ElseIf selected_voucher_name = "sale return" Then
    selected_procedure = "sale return"
    vchr_sale_return.Show
    ElseIf selected_voucher_name = "purchase return" Then
    selected_procedure = "purchase return"
    vchr_purchase_return.Show
    'ElseIf selected_voucher_name = "payment" Then
    'ElseIf selected_voucher_name = "payment" Then
    End If
End If

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
Public Sub set_report_grid()
    grid_report.Clear
    grid_report.Rows = 1
    grid_report.Cols = 10
    grid_report.Font.Size = 12
    b = 0
    grid_report.TextMatrix(b, 0) = ""
    grid_report.TextMatrix(b, 1) = "Date"
    grid_report.TextMatrix(b, 2) = "Voucher"
    grid_report.TextMatrix(b, 3) = "V.No"
    grid_report.TextMatrix(b, 4) = "Ledger"
    grid_report.TextMatrix(b, 5) = "Dr.Amount"
    grid_report.TextMatrix(b, 6) = "Cr.Amount"
    grid_report.TextMatrix(b, 7) = "Balance"
    grid_report.TextMatrix(b, 8) = "Time"
    grid_report.TextMatrix(b, 9) = "User"
    
    grid_report.ColWidth(0) = 1
    grid_report.ColWidth(1) = 1500
    grid_report.ColWidth(2) = 2000
    grid_report.ColWidth(3) = 700
    grid_report.ColWidth(4) = 4500
    grid_report.ColWidth(5) = 1700
    grid_report.ColWidth(6) = 1700
    grid_report.ColWidth(7) = 1700
    grid_report.ColWidth(8) = 1000
    grid_report.ColWidth(9) = 1000
    Dim x_grid_col
    Dim total_grid_width
    total_grid_width = 500
    For x_grid_col = 0 To grid_report.Cols - 1
        total_grid_width = total_grid_width + grid_report.ColWidth(x_grid_col)
    Next
    grid_report.Width = total_grid_width
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form shw_sel_lgr_dtl 
   BackColor       =   &H00FFC0C0&
   Caption         =   "show_selected_ledger_dtl"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "shw_sel_lgr_dtl.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   6360
      Width           =   5295
   End
   Begin VB.CommandButton cmd_print 
      Caption         =   "Print"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   6360
      Width           =   5775
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4800
      Top             =   11880
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
      Left            =   9600
      TabIndex        =   3
      Text            =   "Select Option"
      Top             =   480
      Width           =   1935
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
      Height          =   360
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3495
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
      Height          =   420
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   675
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid grid_report 
      Height          =   4215
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   7435
      _Version        =   393216
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   9600
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
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
      CurrentDate     =   40126
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
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
      CurrentDate     =   40126
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
      Height          =   375
      Left            =   225
      TabIndex        =   16
      Top             =   0
      Width           =   11310
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
      Left            =   225
      TabIndex        =   15
      Top             =   360
      Width           =   11310
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
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   11445
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label m_label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please select ledger...,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   240
      TabIndex        =   12
      Top             =   6840
      Width           =   4050
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Period"
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
      Left            =   8880
      TabIndex        =   11
      Top             =   480
      Width           =   975
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
      Left            =   9000
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   9000
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ledger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   225
      TabIndex        =   7
      Top             =   1680
      Width           =   11430
   End
End
Attribute VB_Name = "shw_sel_lgr_dtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private report_temp_closing_balance
Public Sub set_ledger_account_detail_for_print()
Call open_rs_acn_tran_spc_lgr_print
Do Until rs_acn_tran_spc_lgr_print.EOF
    rs_acn_tran_spc_lgr_print.Delete
    rs_acn_tran_spc_lgr_print.MoveNext
Loop
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
Dim b
b = 9999
Dim temp_cr_total As String
Dim temp_dr_total As String
Do Until rs_acn_tran_spc_lgr.EOF

If rs_acn_tran_spc_lgr!fin_acnt_trn_date >= rep_starting_date And rs_acn_tran_spc_lgr!fin_acnt_trn_date <= rep_ending_date Then

If b = 1 Then
        If ledger_dr_total > 0 And ledger_cr_total > 0 Then
            rs_acn_tran_spc_lgr_print.AddNew
            rs_acn_tran_spc_lgr_print!fin_acnt_trn_id = 0
            If ledger_dr_total > ledger_cr_total Then
                rs_acn_tran_spc_lgr_print!fin_acnt_trn_date = rep_starting_date
                temp_opening_balance = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
                rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = "Opeining Balance is ..., " 'Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
                rs_acn_tran_spc_lgr_print!fin_acnt_trn_dram = Format(ledger_dr_total - ledger_cr_total, "0.00") '& " Dr."
                rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
            ElseIf ledger_cr_total > ledger_dr_total Then
                temp_opening_balance = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
                rs_acn_tran_spc_lgr_print!fin_acnt_trn_date = rep_starting_date
                rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = "Opeining Balance is ..., "
                rs_acn_tran_spc_lgr_print!fin_acnt_trn_cram = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
                rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
            End If
            rs_acn_tran_spc_lgr_print.Update
            b = 1
        End If
        rs_acn_tran_spc_lgr_print.AddNew
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_id = b
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_date = rs_acn_tran_spc_lgr!fin_acnt_trn_date
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_vchr = UCase(rs_acn_tran_spc_lgr!fin_acnt_trn_vchr)
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_vcno = rs_acn_tran_spc_lgr!fin_acnt_trn_vcno
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_ldgr)
        If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") Then
            rs_acn_tran_spc_lgr_print!fin_acnt_trn_dram = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
            ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
        End If
        If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") Then
            rs_acn_tran_spc_lgr_print!fin_acnt_trn_cram = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
            ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
        End If
        If ledger_dr_total > ledger_cr_total Then
            
            temp_dr_total = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
            rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = temp_dr_total
        End If
        If ledger_cr_total > ledger_dr_total Then
            
            temp_cr_total = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
            rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = temp_cr_total
        End If
        If ledger_cr_total = ledger_dr_total Then
            
            temp_cr_total = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Dr."
            rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = temp_cr_total
        End If
        rs_acn_tran_spc_lgr_print.Update
        b = b + 1
Else 'If b > 1 Then
If b = 2 And ledger_cr_total = 0 And ledger_dr_total = 0 Then b = 1
    rs_acn_tran_spc_lgr_print.AddNew
    rs_acn_tran_spc_lgr_print!fin_acnt_trn_id = b
    rs_acn_tran_spc_lgr_print!fin_acnt_trn_date = rs_acn_tran_spc_lgr!fin_acnt_trn_date
    rs_acn_tran_spc_lgr_print!fin_acnt_trn_vchr = UCase(rs_acn_tran_spc_lgr!fin_acnt_trn_vchr)
    rs_acn_tran_spc_lgr_print!fin_acnt_trn_vcno = rs_acn_tran_spc_lgr!fin_acnt_trn_vcno
    rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_ldgr)
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") Then
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_dram = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
        ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") Then
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_cram = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
        ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    If ledger_dr_total > ledger_cr_total Then
        
        temp_dr_total = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = temp_dr_total
    End If
    If ledger_cr_total > ledger_dr_total Then
        
        temp_cr_total = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = temp_cr_total
    End If
    If ledger_cr_total = ledger_dr_total Then
        
        temp_cr_total = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Dr."
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = temp_cr_total
    End If

'    If rs_acn_tran_spc_lgr!fin_acnt_trn_time <> "" Then grid_report.TextMatrix(b, 8) = rs_acn_tran_spc_lgr!fin_acnt_trn_time
'    If rs_acn_tran_spc_lgr!fin_acnt_trn_user <> "" Then grid_report.TextMatrix(b, 9) = rs_acn_tran_spc_lgr!fin_acnt_trn_user
    'If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_date) = "opening balance 1" Or LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_date) = "opening balance 2" Then
    '    rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = rs_acn_tran_spc_lgr!fin_acnt_trn_date
    '    rs_acn_tran_spc_lgr_print!fin_acnt_trn_vcno = ""
    '    rs_acn_tran_spc_lgr_print!fin_acnt_trn_vchr = ""
        '!fin_acnt_trn_ldgr= "Opeining Balance is ..., "
    'End If
    rs_acn_tran_spc_lgr_print.Update
    b = b + 1
End If
End If
If rs_acn_tran_spc_lgr!fin_acnt_trn_date < rep_starting_date Then
        If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") Then
            ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
        End If
        If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") Then
            ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
        End If
        b = 1
        'b = b + 1
End If

rs_acn_tran_spc_lgr.MoveNext
Loop


If b = 1 Then
If ledger_dr_total > 0 Or ledger_cr_total > 0 Then
    rs_acn_tran_spc_lgr_print.AddNew
    If ledger_dr_total > ledger_cr_total Then
        temp_opening_balance = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_id = b
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_date = rep_starting_date
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_vcno = ""
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = "Opeining Balance is ..., " 'Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_dram = Format(ledger_dr_total - ledger_cr_total, "0.00") ' & " Dr."
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_cram = ""
        temp_cr_total = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = temp_cr_total
    ElseIf ledger_cr_total > ledger_dr_total Then
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_id = b
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_date = rep_starting_date
        temp_opening_balance = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_vcno = ""
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = "Opeining Balance is ..., "
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_dram = ""
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_cram = Format(ledger_cr_total - ledger_dr_total, "0.00") ' & " Cr."
        temp_cr_total = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = temp_cr_total
    End If
    rs_acn_tran_spc_lgr_print.Update
    b = b + 1
End If
End If

If b >= 2 Then
'ledger_dr_total = 0
'ledger_cr_total = 0
Dim grid_row_counter
'For grid_row_counter = 1 To (b - 1)
'    If grid_report.TextMatrix(grid_row_counter, 5) <> "" Then ledger_dr_total = ledger_dr_total + Val(grid_report.TextMatrix(grid_row_counter, 5))
'    If grid_report.TextMatrix(grid_row_counter, 6) <> "" Then ledger_cr_total = ledger_cr_total + Val(grid_report.TextMatrix(grid_row_counter, 6))
'Next

If ledger_dr_total < ledger_cr_total Then
        rs_acn_tran_spc_lgr_print.AddNew
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_id = b
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_date = rep_ending_date
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = "Closing Balance is.....Cr."
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_dram = Format(ledger_cr_total - ledger_dr_total, "0.00")
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
        report_temp_closing_balance = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
        b = b + 1
        rs_acn_tran_spc_lgr_print.Update
        
        rs_acn_tran_spc_lgr_print.AddNew
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_id = b
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_dram = Format(ledger_cr_total, "0.00")
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_cram = Format(ledger_cr_total, "0.00")
        
        b = b + 1
        
        rs_acn_tran_spc_lgr_print.Update

End If
If ledger_cr_total < ledger_dr_total Then
        rs_acn_tran_spc_lgr_print.AddNew
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_id = b
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_date = rep_ending_date
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = " Closing Balance is.....Dr. "
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_cram = Format(ledger_dr_total - ledger_cr_total, "0.00")
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_blnc = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        report_temp_closing_balance = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        b = b + 1
        rs_acn_tran_spc_lgr_print.Update
        
        rs_acn_tran_spc_lgr_print.AddNew
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_id = b
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = "TOTAL AMOUNT..."
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_dram = Format(ledger_dr_total, "0.00")
        rs_acn_tran_spc_lgr_print!fin_acnt_trn_cram = Format(ledger_dr_total, "0.00")
        b = b + 1
        rs_acn_tran_spc_lgr_print.Update
End If

'End With
End If
Dim temp_entry_no
Dim temp_v_no
Dim temp_v_tp
Call open_rs_acn_tran_spc_lgr_print

For temp_entry_no = 1 To rs_acn_tran_spc_lgr_print.RecordCount - 3
If rs_acn_tran_spc_lgr_print!fin_acnt_trn_vcno <> 0 Then
'If grid_report.TextMatrix(temp_entry_no, 2) <> "" Then
    temp_v_tp = rs_acn_tran_spc_lgr_print!fin_acnt_trn_vchr
    temp_v_no = rs_acn_tran_spc_lgr_print!fin_acnt_trn_vcno
    Call open_rs_acn_tran_all
    If rs_acn_tran_all.RecordCount > 0 Then rs_acn_tran_all.MoveFirst
    
    Do Until rs_acn_tran_all.EOF
            If LCase(selected_ledger) <> LCase(rs_acn_tran_all!fin_acnt_trn_ldgr) And LCase(temp_v_tp) = LCase(rs_acn_tran_all!fin_acnt_trn_vchr) And temp_v_no = Val(rs_acn_tran_all!fin_acnt_trn_vcno) Then
                rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = LCase(rs_acn_tran_all!fin_acnt_trn_ldgr)
                rs_acn_tran_spc_lgr_print.UpdateBatch
                Exit Do
            End If
        rs_acn_tran_all.MoveNext
    Loop

End If
rs_acn_tran_spc_lgr_print.MoveNext
Next

End Sub

Public Sub set_ledger_account_detail_for_printxxx()

Call open_rs_acn_tran_spc_lgr_print
Do Until rs_acn_tran_spc_lgr_print.EOF
    rs_acn_tran_spc_lgr_print.Delete
    rs_acn_tran_spc_lgr_print.MoveNext
Loop

rep_starting_date = DTPicker1.Value
rep_ending_date = DTPicker2.Value
selected_ledger = selected_from_list.Text
selected_voucher_ledger = selected_ledger
ledger_dr_total = 0
ledger_cr_total = 0

'Call set_report_grid
'Call open_database
'Call copy_specific_ledger_transaction_to_acn_tran_spc_lgr
'Call open_rs_acn_tran_spc_lgr
'Call open_rs_acn_tran_spc_lgr_print

rs_acn_tran_spc_lgr.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno" 'time"

Dim temp_starting_dt As Date
Dim temp_ending_dt As Date
Dim b
b = 0
Do Until rs_acn_tran_spc_lgr.EOF
'MsgBox rs_acn_tran_spc_lgr!fin_acnt_trn_date
rs_acn_tran_spc_lgr_print.AddNew
rs_acn_tran_spc_lgr_print!fin_acnt_trn_id = b
    rs_acn_tran_spc_lgr_print!fin_acnt_trn_date = rs_acn_tran_spc_lgr!fin_acnt_trn_date
    rs_acn_tran_spc_lgr_print!fin_acnt_trn_vchr = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_vchr)
    rs_acn_tran_spc_lgr_print!fin_acnt_trn_vcno = rs_acn_tran_spc_lgr!fin_acnt_trn_vcno
    rs_acn_tran_spc_lgr_print!fin_acnt_trn_ldgr = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_ldgr)
'    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") Then
'        rs_acn_tran_spc_lgr_print!fin_acnt_trn_dram = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
'        ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
'    End If
'    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") Then
'        rs_acn_tran_spc_lgr_print!fin_acnt_trn_cram = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
'        ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
'    End If
rs_acn_tran_spc_lgr_print.Update
b = b + 1
rs_acn_tran_spc_lgr.MoveNext
Loop
End Sub
Public Sub set_ledger_account_detail()
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
'Call open_rs_acn_tran_spc_lgr_print

rs_acn_tran_spc_lgr.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno" 'time"
Dim temp_starting_dt As Date
Dim temp_ending_dt As Date
b = 1

Do Until rs_acn_tran_spc_lgr.EOF

If rs_acn_tran_spc_lgr!fin_acnt_trn_date >= rep_starting_date And rs_acn_tran_spc_lgr!fin_acnt_trn_date <= rep_ending_date Then

If b = 1 Then
    'rs_acn_tran_spc_lgr_print.AddNew
    grid_report.AddItem ""
    'With rs_acn_tran_spc_lgr_print
    If ledger_dr_total > ledger_cr_total Then
        temp_opening_balance = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        grid_report.TextMatrix(b, 1) = rep_starting_date
        grid_report.TextMatrix(b, 3) = ""
        grid_report.TextMatrix(b, 4) = "Opeining Balance is ..., " 'Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        grid_report.TextMatrix(b, 5) = Format(ledger_dr_total - ledger_cr_total, "0.00") '& " Dr."
        grid_report.TextMatrix(b, 6) = ""
        grid_report.TextMatrix(b, 7) = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
    ElseIf ledger_cr_total > ledger_dr_total Then
        temp_opening_balance = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
        grid_report.TextMatrix(b, 1) = rep_starting_date
        grid_report.TextMatrix(b, 3) = ""
        grid_report.TextMatrix(b, 4) = "Opeining Balance is ..., "
        grid_report.TextMatrix(b, 5) = ""
        grid_report.TextMatrix(b, 6) = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
        grid_report.TextMatrix(b, 7) = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
    End If
    b = b + 1
    'End With
End If
    
    If b = 2 And ledger_cr_total = 0 And ledger_dr_total = 0 Then b = 1
    grid_report.AddItem ""
    'grid_report.TextMatrix(b, 0) = b
    grid_report.TextMatrix(b, 1) = rs_acn_tran_spc_lgr!fin_acnt_trn_date
    grid_report.TextMatrix(b, 2) = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_vchr)
    grid_report.TextMatrix(b, 3) = rs_acn_tran_spc_lgr!fin_acnt_trn_vcno
    grid_report.TextMatrix(b, 4) = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_ldgr)
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") Then
        grid_report.TextMatrix(b, 5) = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
        ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") Then
        grid_report.TextMatrix(b, 6) = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
        ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    If ledger_dr_total > ledger_cr_total Then
        grid_report.TextMatrix(b, 7) = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
    End If
    If ledger_cr_total > ledger_dr_total Then
        grid_report.TextMatrix(b, 7) = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
    End If
    If rs_acn_tran_spc_lgr!fin_acnt_trn_time <> "" Then grid_report.TextMatrix(b, 8) = rs_acn_tran_spc_lgr!fin_acnt_trn_time
    If rs_acn_tran_spc_lgr!fin_acnt_trn_user <> "" Then grid_report.TextMatrix(b, 9) = rs_acn_tran_spc_lgr!fin_acnt_trn_user
    
    If LCase(grid_report.TextMatrix(b, 2)) = "opening balance 1" Or LCase(grid_report.TextMatrix(b, 2)) = "opening balance 2" Then
        grid_report.TextMatrix(b, 4) = grid_report.TextMatrix(b, 2)
        grid_report.TextMatrix(b, 3) = ""
        grid_report.TextMatrix(b, 2) = ""
        'grid_report.TextMatrix(b, 4) = "Opeining Balance is ..., "
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
        grid_report.TextMatrix(b, 1) = rep_starting_date
        grid_report.TextMatrix(b, 3) = ""
        grid_report.TextMatrix(b, 4) = "Opeining Balance is ..., " 'Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        grid_report.TextMatrix(b, 5) = Format(ledger_dr_total - ledger_cr_total, "0.00") '& " Dr."
        grid_report.TextMatrix(b, 6) = ""
        grid_report.TextMatrix(b, 7) = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
    ElseIf ledger_cr_total > ledger_dr_total Then
        temp_opening_balance = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
        grid_report.TextMatrix(b, 1) = rep_starting_date
        grid_report.TextMatrix(b, 3) = ""
        grid_report.TextMatrix(b, 4) = "Opeining Balance is ..., "
        grid_report.TextMatrix(b, 5) = ""
        grid_report.TextMatrix(b, 6) = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
        grid_report.TextMatrix(b, 7) = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
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
        grid_report.TextMatrix(b, 1) = rep_ending_date
        grid_report.TextMatrix(b, 4) = "Closing Balance is.....Cr."
        grid_report.TextMatrix(b, 5) = Format(ledger_cr_total - ledger_dr_total, "0.00")
        b = b + 1
        grid_report.AddItem ""
        grid_report.TextMatrix(b, 5) = "================="
        grid_report.TextMatrix(b, 6) = "================="
        b = b + 1
        grid_report.AddItem ""
        grid_report.TextMatrix(b, 5) = Format(ledger_cr_total, "0.00")
        grid_report.TextMatrix(b, 6) = Format(ledger_cr_total, "0.00")
        b = b + 1
        grid_report.AddItem ""
        grid_report.TextMatrix(b, 5) = "================="
        grid_report.TextMatrix(b, 6) = "================="

    End If
    'grid_report.TextMatrix(b, 6) = Format(ledger_cr_total, "0.00")
    If ledger_cr_total < ledger_dr_total Then
        m_label.Caption = selected_voucher_ledger & " Closing Balance is..... Dr. " & Format(ledger_dr_total - ledger_cr_total, "0.00") & " on " & rep_ending_date
        grid_report.AddItem ""
        grid_report.TextMatrix(b, 1) = rep_ending_date
        grid_report.TextMatrix(b, 4) = " Closing Balance is.....Dr. "
        grid_report.TextMatrix(b, 6) = Format(ledger_dr_total - ledger_cr_total, "0.00")
        b = b + 1
        grid_report.AddItem ""
        grid_report.TextMatrix(b, 5) = "================="
        grid_report.TextMatrix(b, 6) = "================="
        b = b + 1
        grid_report.AddItem ""
        grid_report.TextMatrix(b, 4) = "TOTAL AMOUNT..."
        grid_report.TextMatrix(b, 5) = Format(ledger_dr_total, "0.00")
        grid_report.TextMatrix(b, 6) = Format(ledger_dr_total, "0.00")
        b = b + 1
        grid_report.AddItem ""
        grid_report.TextMatrix(b, 5) = "================="
        grid_report.TextMatrix(b, 6) = "================="
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
    If rs_acn_tran_all.RecordCount > 0 Then rs_acn_tran_all.MoveFirst
    Do Until rs_acn_tran_all.EOF
            If LCase(selected_ledger) <> LCase(rs_acn_tran_all!fin_acnt_trn_ldgr) And LCase(temp_v_tp) = LCase(rs_acn_tran_all!fin_acnt_trn_vchr) And temp_v_no = Val(rs_acn_tran_all!fin_acnt_trn_vcno) Then
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
Call set_ledger_account_detail_for_print

With report_ledger_ac.Sections("section2").Controls
    .item("label11").Caption = selected_ledger
    .item("label2").Caption = "period : " & rep_starting_date & " To " & rep_ending_date
End With
With report_ledger_ac.Sections("section5").Controls
    .item("label5").Caption = selected_ledger & " Closing Balance as on : " & rep_ending_date & "  are  " & report_temp_closing_balance
End With

Call open_database
Call open_rs_acn_tran_spc_lgr_print
Dim xxx_rs_acn_tran_spc_lgr_print

Set report_ledger_ac.DataSource = rs_acn_tran_spc_lgr_print
'rs_acn_tran_spc_lgr_print.MoveNext

report_ledger_ac.Show

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
Call set_ledger_account_detail
End Sub

Private Sub Command1_Click()
Unload Me
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
    Call open_database
    Call open_rs_lgr_main_dtl
    Do Until rs_lgr_main_dtl.EOF
    If combo_list.Text = rs_lgr_main_dtl!lgr_main_dtl_alis Then
        combo_list.Text = rs_lgr_main_dtl!lgr_main_dtl_name
    End If
    rs_lgr_main_dtl.MoveNext
    Loop
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
If rs_lgr_main_dtl!lgr_main_dtl_alis <> "" Then combo_list.AddItem rs_lgr_main_dtl!lgr_main_dtl_alis
rs_lgr_main_dtl.MoveNext
Loop
End Sub
Public Sub set_form_headings()
lbl_name.Width = Me.Width
'lbl_name.Left = 0
lbl_name.Caption = co_name
lbl_add.Width = Me.Width
lbl_add.Top = -1000
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
lbl_head.Width = Me.Width
'lbl_head.Left = 0
lbl_head.Caption = UCase(selected_procedure)
Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)
End Sub
Private Sub Form_Activate()
show_ledger_detail = 0
'Call set_ledger_account_detail
End Sub
Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================

Call set_form_headings
    selected_from_list.Text = ""
    Combo2.AddItem "This Year"
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
        DTPicker1.Value = this_year_starting_date
        DTPicker2.Value = selected_date
        Label1.Caption = selected_ledger & "   (" & DTPicker1.Value & "  To  " & DTPicker2.Value & ")"
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
    grid_report.TextMatrix(b, 2) = "Vchr."
    grid_report.TextMatrix(b, 3) = "V.No"
    grid_report.TextMatrix(b, 4) = "Ledger"
    grid_report.TextMatrix(b, 5) = "Dr.Amt."
    grid_report.TextMatrix(b, 6) = "Cr.Amt."
    grid_report.TextMatrix(b, 7) = "Balance"
    grid_report.TextMatrix(b, 8) = "Time"
    grid_report.TextMatrix(b, 9) = "User"
    
    grid_report.ColWidth(0) = 1
    grid_report.ColWidth(1) = 1200
    grid_report.ColWidth(2) = 1000
    grid_report.ColWidth(3) = 550
    grid_report.ColWidth(4) = 2700
    grid_report.ColWidth(5) = 1150
    grid_report.ColWidth(6) = 1150
    grid_report.ColWidth(7) = 1500
    grid_report.ColWidth(8) = 1000
    grid_report.ColWidth(9) = 700
    'Dim x_grid_col
    'Dim total_grid_width
    'total_grid_width = 500
    'For x_grid_col = 0 To grid_report.Cols - 1
    '    total_grid_width = total_grid_width + grid_report.ColWidth(x_grid_col)
    'Next
    'grid_report.Width = total_grid_width
End Sub

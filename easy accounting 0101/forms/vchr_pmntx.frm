VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form vchr_pmntx 
   Caption         =   "Payment Voucher"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10755
   ScaleWidth      =   14340
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   20040
      TabIndex        =   26
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4695
      Left            =   360
      TabIndex        =   4
      Top             =   4920
      Width           =   17415
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   4095
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   16935
         _ExtentX        =   29871
         _ExtentY        =   7223
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3495
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   17415
      Begin VB.ComboBox combo_lgr 
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
         Index           =   1
         Left            =   1800
         TabIndex        =   34
         Text            =   "Combo1"
         Top             =   720
         Width           =   5775
      End
      Begin VB.ComboBox combo_sub_entry_no 
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
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Text            =   "Paid to"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox combo_sub_entry_no 
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
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Text            =   "Paid by"
         Top             =   720
         Width           =   1215
      End
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
         TabIndex        =   30
         Text            =   "Combo0"
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Pirnt"
         Height          =   495
         Left            =   12720
         TabIndex        =   28
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmd_save_n_exit 
         Caption         =   "Save and exit"
         Height          =   495
         Left            =   12720
         TabIndex        =   25
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   12720
         TabIndex        =   24
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   12720
         TabIndex        =   23
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmd_sv_n_new 
         Caption         =   "&Save and New"
         Height          =   495
         Left            =   12720
         TabIndex        =   22
         Top             =   240
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
         TabIndex        =   20
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
         Height          =   855
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "vchr_pmntx.frx":0000
         Top             =   2280
         Width           =   10815
      End
      Begin VB.TextBox text_amt 
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
         Index           =   2
         Left            =   10800
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox combo_lgr 
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
         Index           =   2
         Left            =   1800
         TabIndex        =   13
         Text            =   "Combo2"
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox text_amt 
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
         Index           =   1
         Left            =   8400
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   720
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
         Format          =   98435073
         CurrentDate     =   40166
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
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1215
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
         TabIndex        =   29
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
         Left            =   11280
         TabIndex        =   27
         Top             =   240
         Width           =   855
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
         Left            =   8760
         TabIndex        =   21
         Top             =   240
         Width           =   855
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
         TabIndex        =   19
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
         Height          =   735
         Left            =   240
         TabIndex        =   18
         Top             =   2040
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
         Height          =   375
         Left            =   11160
         TabIndex        =   14
         Top             =   1920
         Width           =   855
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
         Height          =   375
         Left            =   8760
         TabIndex        =   10
         Top             =   1920
         Width           =   855
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
         Left            =   14640
         TabIndex        =   9
         Top             =   2160
         Width           =   2415
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
         Left            =   14760
         TabIndex        =   8
         Top             =   1800
         Width           =   2295
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
      Left            =   1080
      TabIndex        =   2
      Top             =   0
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
      Top             =   360
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
      Top             =   720
      Width           =   8655
   End
End
Attribute VB_Name = "vchr_pmntx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
voucher_total_cr_amt = 0
voucher_total_dr_amt = 0
sub_entry_no = 1
dr_sub_entry_no = 1
cr_sub_entry_no = 1
Call set_form_headings
Call set_form_labels
Text5.Text = selected_user
Call set_vourcher_detail
End Sub
Public Sub set_vourcher_detail()
transaction_type = "payment"
Text1.Text = "" 'voucher no
text_amt(sub_entry_no).Text = "" 'amount
Text4.Text = "" 'narration
Text5.Text = "" 'user
Text1.Enabled = False
Text5.Enabled = False

Combo0.Text = ""    'transaction type 1/2
'combo_lgr(sub_entry_no).Text = ""    'ledger1
'combo_lgr(sub_entry_no).Text = ""    'ledger2
Frame1.Caption = "Current Transaction Detail"
selected_date = DTPicker1.Value
Frame2.Caption = selected_date & "s Transactions Detail"

Call add_account_combo0

sub_entry_no = 1
dr_sub_entry_no = 1
this_entry = "cr"
Call add_combo_sub_entry_no(sub_entry_no)
Call add_account_combo_lgr(sub_entry_no)
Call add_text_amt(sub_entry_no)

this_entry = "dr"
sub_entry_no = sub_entry_no + 1
cr_sub_entry_no = 1
Call add_combo_sub_entry_no(sub_entry_no)
Call add_account_combo_lgr(sub_entry_no)
Call add_text_amt(sub_entry_no)

Call reset_voucher_detail
End Sub
Public Sub reset_voucher_detail()
'=======set date
'=======set no         ' add one
Call open_database
Call open_rs_acn_tran_pmt
'Do Until rs_acn_tran_pmt.EOF
rs_acn_tran_pmt.MoveLast
selected_voucher_no = rs_acn_tran_pmt!fin_acnt_trn_vcno + 1
Text1.Text = selected_voucher_no
'Loop
If rs_acn_tran_pmt.RecordCount = 0 Then Text1.Text = 1
'======= set_type      ' no change
'======= set_ledger     ' write text
'======= set_amount     ' set null
'text_amt(sub_entry_no).Text = "" 'Format(0, "0.00")
'text_amt(sub_entry_no).Text = "" 'Format(0, "0.00")
'======= set_side       ' set common
'======= set_narration  ' set null
Text4.Text = ""
DTPicker1.Value = Date
Label3.Caption = WeekdayName(Weekday(DTPicker1.Value - 1)) ' Day(Weekday(Now))
Label4.Caption = Time

Call arrange_grid1
Call open_grid1
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
Public Sub add_combo_sub_entry_no(sub_entry_no)
combo_sub_entry_no(sub_entry_no).Left = 600
combo_sub_entry_no(sub_entry_no).Top = (sub_entry_no * 600)

combo_sub_entry_no(sub_entry_no).Visible = True
combo_sub_entry_no(sub_entry_no).AddItem "To"
combo_sub_entry_no(sub_entry_no).AddItem "By"
If LCase(this_entry) = "cr" Then
combo_sub_entry_no(sub_entry_no).Text = "By"
ElseIf LCase(this_entry) = "dr" Then
combo_sub_entry_no(sub_entry_no).Text = "To"
End If

End Sub
Public Sub add_account_combo_lgr(sub_entry_no)
combo_lgr(sub_entry_no).Text = "select a ledger"
combo_lgr(sub_entry_no).Left = Frame1.Left + 2000
combo_lgr(sub_entry_no).Top = (sub_entry_no * 600)
combo_lgr(sub_entry_no).Visible = True
If LCase(this_entry) = "cr" Then
Call add_cr_ledger(sub_entry_no)
ElseIf LCase(this_entry) = "dr" Then
Call add_dr_ledger(sub_entry_no)
End If
End Sub
Public Sub add_cr_ledger(sub_entry_no)
Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
    If rs_lgr_main_dtl!lgr_main_dtl_grup = "Cash-on-hand" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Balances" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Loans" Then
        combo_lgr(sub_entry_no).AddItem rs_lgr_main_dtl!lgr_main_dtl_name
        If rs_lgr_main_dtl!lgr_main_dtl_alis <> "" Then combo_lgr(sub_entry_no).AddItem rs_lgr_main_dtl!lgr_main_dtl_alis
    End If
rs_lgr_main_dtl.MoveNext
Loop
End Sub
Public Sub add_dr_ledger(sub_entry_no)
Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
    If rs_lgr_main_dtl!lgr_main_dtl_grup = "Cash-on-hand" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Balances" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Loans" Then
    Else
        combo_lgr(sub_entry_no).AddItem rs_lgr_main_dtl!lgr_main_dtl_name
        If rs_lgr_main_dtl!lgr_main_dtl_alis <> "" Then combo_lgr(sub_entry_no).AddItem rs_lgr_main_dtl!lgr_main_dtl_alis
    End If
rs_lgr_main_dtl.MoveNext
Loop
End Sub

Public Sub add_text_amt(sub_entry_no)
text_amt(sub_entry_no).Top = (sub_entry_no * 600)
text_amt(sub_entry_no).Visible = True
text_amt(sub_entry_no).Text = "" 'amount
If LCase(this_entry) = "cr" Then
Call cr_amt_text_adjust(sub_entry_no)
ElseIf LCase(this_entry) = "dr" Then
dr_amt_text_adjust (sub_entry_no)
End If
End Sub
Public Sub cr_amt_text_adjust(sub_entry_no)
text_amt(sub_entry_no).Left = 10500
End Sub
Public Sub dr_amt_text_adjust(sub_entry_no)
text_amt(sub_entry_no).Left = 8500
End Sub

Private Sub text_amt_cr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
text_amt(sub_entry_no).Text = Format(text_amt(sub_entry_no).Text, "0.00")
End If
End Sub

Private Sub Label7_Click()
MsgBox Label7.Top
End Sub


Private Sub cmd_sv_n_new_Click()
'Call check_all_data
If Text1.Text = "" Or _
    Val(text_amt(sub_entry_no).Text) < 0 Or _
    Val(text_amt(sub_entry_no).Text) < 0 Or _
    Combo0.Text > 2 Or _
    Combo0.Text < 1 Or _
    combo_lgr(sub_entry_no).Text = "select a ledger" Or _
    combo_lgr(sub_entry_no).Text = "select a ledger" Or _
    DTPicker1.Value > Date Or _
    DTPicker1.Value < selected_companies_starting_f_date Then
        
        MsgBox "You have not entered proper or sufficient detail...!!!"
        Exit Sub
End If

Call open_database
Call open_rs_acn_tran_pmt
MsgBox sub_entry_no
Dim tra_counter
For tra_counter = 1 To sub_entry_no

rs_acn_tran_pmt.AddNew

With rs_acn_tran_pmt
    !fin_acnt_trn_vcno = Text1.Text
    !fin_acnt_trn_vtyp = Combo0.Text
    !fin_acnt_trn_date = DTPicker1.Value
    !fin_acnt_trn_time = Label4.Caption
    !fin_acnt_trn_wday = Label3.Caption
    !fin_acnt_trn_ldgr = combo_lgr(tra_counter).Text
    !fin_acnt_trn_amnt = text_amt(tra_counter).Text
    If LCase(combo_sub_entry_no(tra_counter).Text) = "to" Then
    !fin_acnt_trn_side = "dr"
    ElseIf LCase(combo_sub_entry_no(tra_counter).Text) = "by" Then
    !fin_acnt_trn_side = "cr"
    End If
    !fin_acnt_trn_nrtn = Text4.Text
    !fin_acnt_trn_user = Text5.Text
    !fin_acnt_trn_vchr = "payment"
End With
rs_acn_tran_pmt.UpdateBatch
Next

selected_voucher_no = rs_acn_tran_pmt!fin_acnt_trn_vcno + 1
Text1.Text = selected_voucher_no
Call reset_voucher_detail
Call refresh_grid1
End Sub

Public Sub refresh_grid1()
Call arrange_grid1
Call open_grid1
End Sub
Public Sub arrange_grid1()
    Grid1.RowHeightMin = 250
    Grid1.Clear
    Grid1.Rows = 2
    Grid1.Cols = 12
    Grid1.TextMatrix(0, 1) = "Type"
    Grid1.TextMatrix(0, 2) = "Voucher No"
    Grid1.TextMatrix(0, 3) = "Date"
    Grid1.TextMatrix(0, 4) = "Day"
    Grid1.TextMatrix(0, 5) = "Time"
    Grid1.TextMatrix(0, 6) = "Paid by"
    Grid1.TextMatrix(0, 7) = "Amount"
    Grid1.TextMatrix(0, 8) = "Paid To/For"
    Grid1.TextMatrix(0, 9) = "Amount"
    Grid1.TextMatrix(0, 10) = "Narration"
    Grid1.TextMatrix(0, 11) = "User"
    
    
    Grid1.ColWidth(0) = 500
    
    Grid1.ColWidth(1) = 200
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1300
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    
    Grid1.ColWidth(6) = 2500
    Grid1.ColWidth(7) = 1500
    Grid1.ColWidth(8) = 2500
    Grid1.ColWidth(9) = 1500
    Grid1.ColWidth(10) = 3500
    Grid1.ColWidth(11) = 1000
    
    
    Grid1.Font.Size = 10
 'grid1.Width = grid1.ColWidth(0) + grid1.ColWidth(1) + grid1.ColWidth(2) + grid1.ColWidth(3) + grid1.ColWidth(4)
End Sub
Public Sub open_grid1()

Call open_database
Call open_rs_acn_tran_pmt
Dim data_no As Integer
data_no = 1
Do Until rs_acn_tran_pmt.EOF
With rs_acn_tran_pmt
Grid1.TextMatrix(data_no, 0) = data_no
Grid1.TextMatrix(data_no, 1) = !fin_acnt_trn_vtyp

If selected_voucher_no = !fin_acnt_trn_vcno Then

Grid1.TextMatrix(data_no, 9) = Format(!fin_acnt_trn_amnt, "0.00")
'northing to do
Else
    selected_voucher_no = !fin_acnt_trn_vcno
    Grid1.TextMatrix(data_no, 2) = !fin_acnt_trn_vcno
    Grid1.TextMatrix(data_no, 3) = !fin_acnt_trn_date
    Grid1.TextMatrix(data_no, 4) = !fin_acnt_trn_wday
    Grid1.TextMatrix(data_no, 5) = !fin_acnt_trn_time
    Grid1.TextMatrix(data_no, 7) = Format(!fin_acnt_trn_amnt, "0.00")
    'Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_nrtn
    Grid1.TextMatrix(data_no, 10) = !fin_acnt_trn_nrtn
    Grid1.TextMatrix(data_no, 11) = !fin_acnt_trn_user
End If
If LCase(!fin_acnt_trn_side) = "cr" Then
Grid1.TextMatrix(data_no, 6) = !fin_acnt_trn_ldgr
ElseIf LCase(!fin_acnt_trn_side) = "dr" Then
Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_ldgr
End If
'rs_acn_tran_pmt.MoveNext
'If LCase(fin_acnt_trn_side) = "dr" Then
    'Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_side
'End If
'End If
End With
data_no = data_no + 1
If rs_acn_tran_pmt.RecordCount < Grid1.Rows Then
Exit Sub
End If
Grid1.Rows = Grid1.Rows + 1
rs_acn_tran_pmt.MoveNext
Loop
End Sub

Private Sub Command1_Click()
Unload Me
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

Public Sub set_form_labels()
Label0.Caption = "Type"
Label1.Caption = "No"
Label2.Caption = "Date"
'Label3.Caption = "Day:" & Day(DTPicker1.Value)
'Label4.Caption = "Time:" & Time
Label7.Caption = "Narration"
Label8.Caption = "User"
Label9.Caption = "Amount"
Label10.Caption = "Amount"
Label11.Caption = "Paid"
End Sub
Private Sub text_amt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Index = 1 Then
text_amt(Index).Text = Format(text_amt(Index).Text, "0.00")
text_amt(Index + 1).Text = Format(text_amt(Index).Text, "0.00")
ElseIf KeyCode = 13 And Index > 1 Then
text_amt(Index).Text = Format(text_amt(Index).Text, "0.00")
End If
If KeyCode = 13 Then
    Call refresh_dr_cr_total_amt
    'MsgBox voucher_total_dr_amt
    'MsgBox voucher_total_cr_amt
    If voucher_total_dr_amt < voucher_total_cr_amt Then
            If sub_entry_no > 12 Then
            MsgBox "Only 12 Entries allowed Once....!!!!"
            Exit Sub
            End If
            If combo_lgr(sub_entry_no - 1).Text = "" Or LCase(combo_lgr(sub_entry_no - 1).Text) = " " Or LCase(combo_lgr(sub_entry_no - 1).Text) = "select a ledger" _
            Or Val(text_amt(sub_entry_no - 1)) < 0 Or Val(text_amt(sub_entry_no - 1)) = Null Then
            MsgBox "Please, Enter proper values....!!!! and click again...!!!"
            Exit Sub
            End If
            
            If LCase(combo_sub_entry_no(Index).Text) = "to" And Index > 1 And Index = sub_entry_no Then
                    'write here text to insert row
                    sub_entry_no = sub_entry_no + 1
                    Load combo_sub_entry_no(sub_entry_no)
                    Load combo_lgr(sub_entry_no)
                    Load text_amt(sub_entry_no)
                    Call add_combo_sub_entry_no(sub_entry_no)
                    Call add_text_amt(sub_entry_no)
                    Call add_account_combo_lgr(sub_entry_no)
                    Call move_all_command_to_bottom
                    text_amt(sub_entry_no) = Format(voucher_total_cr_amt - voucher_total_dr_amt, "0.00")
            End If
    ElseIf voucher_total_dr_amt > voucher_total_cr_amt Then
            If sub_entry_no > 12 Then
            MsgBox "Only 12 Entries allowed Once....!!!!"
            Exit Sub
            End If
            If combo_lgr(sub_entry_no - 1).Text = "" Or LCase(combo_lgr(sub_entry_no - 1).Text) = " " Or LCase(combo_lgr(sub_entry_no - 1).Text) = "select a ledger" _
            Or Val(text_amt(sub_entry_no - 1)) < 0 Or Val(text_amt(sub_entry_no - 1)) = Null Then
            MsgBox "Please, Enter proper values....!!!! and click again...!!!"
            Exit Sub
            End If
            this_entry = "cr"
                    'write here text to insert row
                    sub_entry_no = sub_entry_no + 1
                    Load combo_sub_entry_no(sub_entry_no)
                    Load combo_lgr(sub_entry_no)
                    Load text_amt(sub_entry_no)
                    Call add_combo_sub_entry_no(sub_entry_no)
                    Call add_text_amt(sub_entry_no)
                    Call add_account_combo_lgr(sub_entry_no)
                    Call move_all_command_to_bottom
                    text_amt(sub_entry_no) = Format(voucher_total_dr_amt - voucher_total_cr_amt, "0.00")
    
    End If
End If
End Sub
Public Sub refresh_dr_cr_total_amt()
voucher_total_cr_amt = 0
voucher_total_dr_amt = 0

Dim int_i

For int_i = 1 To sub_entry_no
    'MsgBox combo_sub_entry_no(int_i).Text
    If LCase(combo_sub_entry_no(int_i).Text) = "by" Then
    
        voucher_total_cr_amt = voucher_total_cr_amt + Val(text_amt(int_i).Text)
    ElseIf LCase(combo_sub_entry_no(int_i).Text) = "to" Then
        voucher_total_dr_amt = voucher_total_dr_amt + Val(text_amt(int_i).Text)
    End If
Next

Label5.Caption = Format(voucher_total_dr_amt, "0.00")
Label6.Caption = Format(voucher_total_cr_amt, "0.00")

End Sub
Private Sub combo_sub_entry_no_Click(Index As Integer)
If sub_entry_no > 12 Then
MsgBox "Only 12 Entries allowed Once....!!!!"
Exit Sub
End If
If combo_lgr(sub_entry_no - 1).Text = "" Or LCase(combo_lgr(sub_entry_no - 1).Text) = " " Or LCase(combo_lgr(sub_entry_no - 1).Text) = "select a ledger" _
Or Val(text_amt(sub_entry_no - 1)) < 0 Or Val(text_amt(sub_entry_no - 1)) = Null Then
MsgBox "Please, Enter proper values....!!!! and click again...!!!"
Exit Sub
End If

If LCase(combo_sub_entry_no(Index).Text) = "to" And Index > 1 And Index = sub_entry_no Then
        'write here text to insert row
        sub_entry_no = sub_entry_no + 1
        Load combo_sub_entry_no(sub_entry_no)
        Load combo_lgr(sub_entry_no)
        Load text_amt(sub_entry_no)
        Call add_combo_sub_entry_no(sub_entry_no)
        Call add_text_amt(sub_entry_no)
        Call add_account_combo_lgr(sub_entry_no)
        Call move_all_command_to_bottom
End If
End Sub
Public Sub move_all_command_to_bottom()
'MsgBox Frame2.Top '1.Height
Label7.Top = (sub_entry_no * 600) + 1000
Text4.Top = (sub_entry_no * 600) + 800
Frame1.Height = (sub_entry_no * 600) + 2300
Frame2.Top = (sub_entry_no * 600) + 4300
Label5.Top = (sub_entry_no * 600) + 500
Label6.Top = (sub_entry_no * 600) + 500
End Sub

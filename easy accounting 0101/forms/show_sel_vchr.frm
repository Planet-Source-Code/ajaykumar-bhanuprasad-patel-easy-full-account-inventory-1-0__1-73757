VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form show_sel_vchr 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Contra Voucher Summary"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   Icon            =   "show_sel_vchr.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
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
      TabIndex        =   7
      Text            =   "Select Option"
      Top             =   240
      Width           =   1695
   End
   Begin VB.ComboBox cmb_voucher 
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
      Left            =   240
      TabIndex        =   6
      Text            =   "Voucher"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to Exit."
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   6840
      Width           =   11175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Frame2"
      Height          =   5175
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   11175
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   4815
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8493
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483626
         ForeColorSel    =   -2147483635
         SelectionMode   =   1
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   9600
      TabIndex        =   8
      Top             =   1200
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
      CurrentDate     =   40126
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9600
      TabIndex        =   9
      Top             =   720
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
      CurrentDate     =   40126
   End
   Begin VB.Label Label1 
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
      Left            =   8640
      TabIndex        =   13
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Period...,"
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
      Left            =   8640
      TabIndex        =   12
      Top             =   240
      Width           =   1095
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
      Left            =   8640
      TabIndex        =   11
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Voucher Type"
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
      Top             =   600
      Width           =   2295
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
      Left            =   195
      TabIndex        =   2
      Top             =   0
      Width           =   11100
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
      Left            =   195
      TabIndex        =   1
      Top             =   360
      Width           =   11100
   End
   Begin VB.Label lbl_head 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2595
      TabIndex        =   0
      Top             =   1080
      Width           =   6900
   End
End
Attribute VB_Name = "show_sel_vchr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_voucher_Click()
    Combo2.Text = "This Month"
    Dim today_day As Integer
    Dim today_weekday As Integer
    today_weekday = Weekday(Now)
    today_day = Day(Now) - 1
    DTPicker1.Value = Date - today_day
    DTPicker2.Value = Date
    Call open_voucher_grid
End Sub
Public Sub open_voucher_grid()
Call arrange_grid1
    If LCase(cmb_voucher.Text) = "payment" Then
    selected_procedure = "payment voucher"
    Call open_grid_payment
    ElseIf LCase(cmb_voucher.Text) = "receipt" Then
    selected_procedure = "Receipt voucher"
    Call open_grid_receipt
    ElseIf LCase(cmb_voucher.Text) = "sale" Then
    selected_procedure = "Sales voucher"
    Call open_grid_sale
    ElseIf LCase(cmb_voucher.Text) = "purchase" Then
    selected_procedure = "purchase voucher"
    Call open_grid_purchase
    ElseIf LCase(cmb_voucher.Text) = "contra" Then
    Call open_grid_contra
    ElseIf LCase(cmb_voucher.Text) = "journal" Then
    selected_procedure = "Adjustment/Journal voucher"
    Call open_grid_journal
    ElseIf LCase(cmb_voucher.Text) = "sale return" Then
    selected_procedure = "sale return"
    Call open_grid_sale_return
    ElseIf LCase(cmb_voucher.Text) = "purchase return" Then
    selected_procedure = "purchase return"
    Call open_grid_purchase_return
    End If
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
Call open_voucher_grid
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
Dim today_day As Integer
Dim today_weekday As Integer
today_weekday = Weekday(Now)
today_day = Day(Now) - 1
Call open_voucher_grid
End Sub
Private Sub DTPicker2_Change()
Dim today_day As Integer
Dim today_weekday As Integer
today_weekday = Weekday(Now)
today_day = Day(Now) - 1
Call open_voucher_grid
End Sub
Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================
Combo2.AddItem "This Year"
    Combo2.AddItem "This Month"
    
    Combo2.AddItem "This Week"
    Combo2.AddItem "Last Month"
    Combo2.AddItem "Last Week"
    
    cmb_voucher.AddItem "Payment"
    cmb_voucher.AddItem "Receipt"
    cmb_voucher.AddItem "Contra"
    cmb_voucher.AddItem "Journal"
    cmb_voucher.AddItem "Sale"
    cmb_voucher.AddItem "Purchase"
    cmb_voucher.AddItem "Sale Return"
    cmb_voucher.AddItem "Purchase Return"
    'selected_procedure = "Contra voucher"
    xsub_entry_no = 0
    current_sub_entry_no = 0
    sub_entry_no = 1
    dr_sub_entry_no = 1
    cr_sub_entry_no = 1
    cmb_voucher.Text = selected_voucher_name
    Call set_form_headings
    Dim today_day As Integer
    Dim today_weekday As Integer
    Combo2.Text = "This Month"
    today_weekday = Weekday(Now)
    today_day = Day(Now) - 1
    DTPicker1.Value = Date - today_day
    DTPicker2.Value = Date
    selected_date = DTPicker1.Value
    Call open_voucher_grid
End Sub
Public Sub search_vouchers()

change_the_old_voucher = 0
current_sub_entry_no = 0
total_sub_entry_no = 0

If Grid1.TextMatrix(Grid1.Row, 0) = "" Then
    MsgBox "This is a invalid entry....,"
    Exit Sub
End If

Dim entry_no
For entry_no = 1 To 12
If Val(Grid1.TextMatrix(Grid1.Row, 2)) <> 0 Then
    selected_voucher_no = Grid1.TextMatrix(Grid1.Row, 2)
    Exit For
ElseIf Val(Grid1.TextMatrix(Grid1.Row - entry_no, 2)) <> 0 Then
    selected_voucher_no = Grid1.TextMatrix(Grid1.Row - entry_no, 2)
    Exit For
End If
Next

sub_entry_no = 1
show_ledger_detail = 1
'Dim selected_voucher_name
'    selected_voucher_date = grid_report.TextMatrix(grid_report.Row, 1)
'    selected_voucher_name = LCase(grid_report.TextMatrix(grid_report.Row, 2))
'    selected_voucher_no = grid_report.TextMatrix(grid_report.Row, 3)
    selected_voucher_name = LCase(cmb_voucher.Text)
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
End Sub
Public Sub refresh_grid1()
Call arrange_grid1
'Call open_grid1
End Sub
Public Sub arrange_grid1()
    Grid1.RowHeightMin = 250
    Grid1.Clear
    Grid1.Rows = 2
    Grid1.Cols = 12
    Grid1.TextMatrix(0, 1) = "Type"
    Grid1.TextMatrix(0, 2) = "Vch.No"
    Grid1.TextMatrix(0, 3) = "Date"
    Grid1.TextMatrix(0, 4) = "Day"
    Grid1.TextMatrix(0, 5) = "Time"
    Grid1.TextMatrix(0, 6) = "Cr / To"
    Grid1.TextMatrix(0, 7) = "Amount"
    Grid1.TextMatrix(0, 8) = "Dr / By"
    Grid1.TextMatrix(0, 9) = "Amount"
    Grid1.TextMatrix(0, 10) = "Narration"
    Grid1.TextMatrix(0, 11) = "User"
    Grid1.ColWidth(0) = 1
    Grid1.ColWidth(1) = 1
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 1200
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 3500
    Grid1.ColWidth(7) = 1200
    Grid1.ColWidth(8) = 3500
    Grid1.ColWidth(9) = 1200
    Grid1.ColWidth(10) = 3500
    Grid1.ColWidth(11) = 1500
    Grid1.Font.Size = 10
End Sub
Public Sub open_grid_contra()
Dim saw_voucher_no
Call open_database
Call open_rs_acn_tran_cnt
rs_acn_tran_cnt.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno"
Dim data_no As Integer
data_no = 1
Do Until rs_acn_tran_cnt.EOF
With rs_acn_tran_cnt
If DTPicker1.Value <= !fin_acnt_trn_date And DTPicker2.Value >= !fin_acnt_trn_date Then
    Grid1.TextMatrix(data_no, 0) = data_no
    Grid1.TextMatrix(data_no, 1) = !fin_acnt_trn_vtyp
    If saw_voucher_no = !fin_acnt_trn_vcno Then
    Else
        saw_voucher_no = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 2) = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 3) = !fin_acnt_trn_date
        Grid1.TextMatrix(data_no, 4) = !fin_acnt_trn_wday
        Grid1.TextMatrix(data_no, 5) = !fin_acnt_trn_time
        Grid1.TextMatrix(data_no, 10) = !fin_acnt_trn_nrtn
        Grid1.TextMatrix(data_no, 11) = !fin_acnt_trn_user
    End If
    If LCase(!fin_acnt_trn_side) = "cr" Then
        Grid1.TextMatrix(data_no, 7) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 6) = !fin_acnt_trn_ldgr
    ElseIf LCase(!fin_acnt_trn_side) = "dr" Then
        Grid1.TextMatrix(data_no, 9) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_ldgr
    End If
    data_no = data_no + 1
    If rs_acn_tran_cnt.RecordCount < Grid1.Rows Then
    Exit Sub
    End If
    Grid1.Rows = Grid1.Rows + 1
End If
End With
rs_acn_tran_cnt.MoveNext
Loop
End Sub
Public Sub open_grid_payment()
Dim saw_voucher_no
Call open_database
Call open_rs_acn_tran_pmt
rs_acn_tran_pmt.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno"
Dim data_no As Integer
data_no = 1
Do Until rs_acn_tran_pmt.EOF
With rs_acn_tran_pmt
If DTPicker1.Value <= !fin_acnt_trn_date And DTPicker2.Value >= !fin_acnt_trn_date Then
    Grid1.TextMatrix(data_no, 0) = data_no
    Grid1.TextMatrix(data_no, 1) = !fin_acnt_trn_vtyp
    If saw_voucher_no = !fin_acnt_trn_vcno Then
    Else
        saw_voucher_no = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 2) = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 3) = !fin_acnt_trn_date
        Grid1.TextMatrix(data_no, 4) = !fin_acnt_trn_wday
        Grid1.TextMatrix(data_no, 5) = !fin_acnt_trn_time
        Grid1.TextMatrix(data_no, 10) = !fin_acnt_trn_nrtn
        Grid1.TextMatrix(data_no, 11) = !fin_acnt_trn_user
    End If
    If LCase(!fin_acnt_trn_side) = "cr" Then
        Grid1.TextMatrix(data_no, 7) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 6) = !fin_acnt_trn_ldgr
    ElseIf LCase(!fin_acnt_trn_side) = "dr" Then
        Grid1.TextMatrix(data_no, 9) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_ldgr
    End If
    data_no = data_no + 1
    If rs_acn_tran_pmt.RecordCount < Grid1.Rows Then
    Exit Sub
    End If
    Grid1.Rows = Grid1.Rows + 1
End If
End With
rs_acn_tran_pmt.MoveNext
Loop
End Sub
Public Sub open_grid_receipt()
Dim saw_voucher_no
Call open_database
Call open_rs_acn_tran_rct
rs_acn_tran_rct.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno"
Dim data_no As Integer
data_no = 1
Do Until rs_acn_tran_rct.EOF
With rs_acn_tran_rct
If DTPicker1.Value <= !fin_acnt_trn_date And DTPicker2.Value >= !fin_acnt_trn_date Then
    Grid1.TextMatrix(data_no, 0) = data_no
    Grid1.TextMatrix(data_no, 1) = !fin_acnt_trn_vtyp
    If saw_voucher_no = !fin_acnt_trn_vcno Then
    Else
        saw_voucher_no = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 2) = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 3) = !fin_acnt_trn_date
        Grid1.TextMatrix(data_no, 4) = !fin_acnt_trn_wday
        Grid1.TextMatrix(data_no, 5) = !fin_acnt_trn_time
        Grid1.TextMatrix(data_no, 10) = !fin_acnt_trn_nrtn
        Grid1.TextMatrix(data_no, 11) = !fin_acnt_trn_user
    End If
    If LCase(!fin_acnt_trn_side) = "cr" Then
        Grid1.TextMatrix(data_no, 7) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 6) = !fin_acnt_trn_ldgr
    ElseIf LCase(!fin_acnt_trn_side) = "dr" Then
        Grid1.TextMatrix(data_no, 9) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_ldgr
    End If
    data_no = data_no + 1
    If rs_acn_tran_rct.RecordCount < Grid1.Rows Then
    Exit Sub
    End If
    Grid1.Rows = Grid1.Rows + 1
End If
End With
rs_acn_tran_rct.MoveNext
Loop
End Sub
Public Sub open_grid_sale()
Dim saw_voucher_no
Call open_database
Call open_rs_acn_tran_sal
rs_acn_tran_sal.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno"
Dim data_no As Integer
data_no = 1
Do Until rs_acn_tran_sal.EOF
With rs_acn_tran_sal
If DTPicker1.Value <= !fin_acnt_trn_date And DTPicker2.Value >= !fin_acnt_trn_date Then
    Grid1.TextMatrix(data_no, 0) = data_no
    Grid1.TextMatrix(data_no, 1) = !fin_acnt_trn_vtyp
    If saw_voucher_no = !fin_acnt_trn_vcno Then
    Else
        saw_voucher_no = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 2) = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 3) = !fin_acnt_trn_date
        Grid1.TextMatrix(data_no, 4) = !fin_acnt_trn_wday
        Grid1.TextMatrix(data_no, 5) = !fin_acnt_trn_time
'        Grid1.TextMatrix(data_no, 10) = !fin_acnt_trn_nrtn
        Grid1.TextMatrix(data_no, 11) = !fin_acnt_trn_user
    End If
    If LCase(!fin_acnt_trn_side) = "cr" Then
        Grid1.TextMatrix(data_no, 7) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 6) = !fin_acnt_trn_ldgr
    ElseIf LCase(!fin_acnt_trn_side) = "dr" Then
        Grid1.TextMatrix(data_no, 9) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_ldgr
    End If
    data_no = data_no + 1
    If rs_acn_tran_sal.RecordCount < Grid1.Rows Then
    Exit Sub
    End If
    Grid1.Rows = Grid1.Rows + 1
End If
End With
rs_acn_tran_sal.MoveNext
Loop
End Sub
Public Sub open_grid_purchase()
Dim saw_voucher_no
Call open_database
Call open_rs_acn_tran_prs
rs_acn_tran_prs.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno"
Dim data_no As Integer
data_no = 1
Do Until rs_acn_tran_prs.EOF
With rs_acn_tran_prs
If DTPicker1.Value <= !fin_acnt_trn_date And DTPicker2.Value >= !fin_acnt_trn_date Then
    Grid1.TextMatrix(data_no, 0) = data_no
    Grid1.TextMatrix(data_no, 1) = !fin_acnt_trn_vtyp
    If saw_voucher_no = !fin_acnt_trn_vcno Then
    Else
        saw_voucher_no = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 2) = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 3) = !fin_acnt_trn_date
        Grid1.TextMatrix(data_no, 4) = !fin_acnt_trn_wday
        Grid1.TextMatrix(data_no, 5) = !fin_acnt_trn_time
        'Grid1.TextMatrix(data_no, 10) = !fin_acnt_trn_nrtn
        Grid1.TextMatrix(data_no, 11) = !fin_acnt_trn_user
    End If
    If LCase(!fin_acnt_trn_side) = "cr" Then
        Grid1.TextMatrix(data_no, 7) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 6) = !fin_acnt_trn_ldgr
    ElseIf LCase(!fin_acnt_trn_side) = "dr" Then
        Grid1.TextMatrix(data_no, 9) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_ldgr
    End If
    data_no = data_no + 1
    If rs_acn_tran_prs.RecordCount < Grid1.Rows Then
    Exit Sub
    End If
    Grid1.Rows = Grid1.Rows + 1
End If
End With
rs_acn_tran_prs.MoveNext
Loop
End Sub
Public Sub open_grid_purchase_return()
Dim saw_voucher_no
Call open_database
Call open_rs_acn_tran_prt
rs_acn_tran_prt.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno"
Dim data_no As Integer
data_no = 1
Do Until rs_acn_tran_prt.EOF
With rs_acn_tran_prt
If DTPicker1.Value <= !fin_acnt_trn_date And DTPicker2.Value >= !fin_acnt_trn_date Then
    Grid1.TextMatrix(data_no, 0) = data_no
    Grid1.TextMatrix(data_no, 1) = !fin_acnt_trn_vtyp
    If saw_voucher_no = !fin_acnt_trn_vcno Then
    Else
        saw_voucher_no = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 2) = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 3) = !fin_acnt_trn_date
        Grid1.TextMatrix(data_no, 4) = !fin_acnt_trn_wday
        Grid1.TextMatrix(data_no, 5) = !fin_acnt_trn_time
        'Grid1.TextMatrix(data_no, 10) = !fin_acnt_trn_nrtn
        Grid1.TextMatrix(data_no, 11) = !fin_acnt_trn_user
    End If
    If LCase(!fin_acnt_trn_side) = "cr" Then
        Grid1.TextMatrix(data_no, 7) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 6) = !fin_acnt_trn_ldgr
    ElseIf LCase(!fin_acnt_trn_side) = "dr" Then
        Grid1.TextMatrix(data_no, 9) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_ldgr
    End If
    data_no = data_no + 1
    If rs_acn_tran_prt.RecordCount < Grid1.Rows Then
    Exit Sub
    End If
    Grid1.Rows = Grid1.Rows + 1
End If
End With
rs_acn_tran_prt.MoveNext
Loop
End Sub
Public Sub open_grid_sale_return()
Dim saw_voucher_no
Call open_database
Call open_rs_acn_tran_srt
rs_acn_tran_srt.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno"
Dim data_no As Integer
data_no = 1
Do Until rs_acn_tran_srt.EOF
With rs_acn_tran_srt
If DTPicker1.Value <= !fin_acnt_trn_date And DTPicker2.Value >= !fin_acnt_trn_date Then
    Grid1.TextMatrix(data_no, 0) = data_no
    Grid1.TextMatrix(data_no, 1) = !fin_acnt_trn_vtyp
    If saw_voucher_no = !fin_acnt_trn_vcno Then
    Else
        saw_voucher_no = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 2) = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 3) = !fin_acnt_trn_date
        Grid1.TextMatrix(data_no, 4) = !fin_acnt_trn_wday
        Grid1.TextMatrix(data_no, 5) = !fin_acnt_trn_time
        'Grid1.TextMatrix(data_no, 10) = !fin_acnt_trn_nrtn
        Grid1.TextMatrix(data_no, 11) = !fin_acnt_trn_user
    End If
    If LCase(!fin_acnt_trn_side) = "cr" Then
        Grid1.TextMatrix(data_no, 7) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 6) = !fin_acnt_trn_ldgr
    ElseIf LCase(!fin_acnt_trn_side) = "dr" Then
        Grid1.TextMatrix(data_no, 9) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_ldgr
    End If
    data_no = data_no + 1
    If rs_acn_tran_srt.RecordCount < Grid1.Rows Then
    Exit Sub
    End If
    Grid1.Rows = Grid1.Rows + 1
End If
End With
rs_acn_tran_srt.MoveNext
Loop
End Sub
Public Sub open_grid_journal()
Dim saw_voucher_no
Call open_database
Call open_rs_acn_tran_jrn
rs_acn_tran_jrn.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno"
Dim data_no As Integer
data_no = 1
Do Until rs_acn_tran_jrn.EOF
With rs_acn_tran_jrn
If DTPicker1.Value <= !fin_acnt_trn_date And DTPicker2.Value >= !fin_acnt_trn_date Then
    Grid1.TextMatrix(data_no, 0) = data_no
    Grid1.TextMatrix(data_no, 1) = !fin_acnt_trn_vtyp
    If saw_voucher_no = !fin_acnt_trn_vcno Then
    Else
        saw_voucher_no = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 2) = !fin_acnt_trn_vcno
        Grid1.TextMatrix(data_no, 3) = !fin_acnt_trn_date
        Grid1.TextMatrix(data_no, 4) = !fin_acnt_trn_wday
        Grid1.TextMatrix(data_no, 5) = !fin_acnt_trn_time
        Grid1.TextMatrix(data_no, 10) = !fin_acnt_trn_nrtn
        Grid1.TextMatrix(data_no, 11) = !fin_acnt_trn_user
    End If
    If LCase(!fin_acnt_trn_side) = "cr" Then
        Grid1.TextMatrix(data_no, 7) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 6) = !fin_acnt_trn_ldgr
    ElseIf LCase(!fin_acnt_trn_side) = "dr" Then
        Grid1.TextMatrix(data_no, 9) = Format(!fin_acnt_trn_amnt, "0.00")
        Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_ldgr
    End If
    data_no = data_no + 1
    If rs_acn_tran_jrn.RecordCount < Grid1.Rows Then
    Exit Sub
    End If
    Grid1.Rows = Grid1.Rows + 1
End If
End With
rs_acn_tran_jrn.MoveNext
Loop
End Sub
Public Sub set_form_headings()
lbl_name.Width = Me.Width
lbl_name.Left = 0
lbl_name.Caption = co_name
lbl_add.Width = Me.Width
lbl_add.Top = -1000
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
lbl_head.Width = Me.Width
lbl_head.Left = 0
lbl_head.Caption = selected_procedure
Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)
End Sub

Private Sub Grid1_DblClick()
Call search_vouchers

End Sub

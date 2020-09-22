VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form da_Gene_cust_repo_main 
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "da_generate_cust_repo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmb_supplier 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11400
      TabIndex        =   9
      Text            =   "No supplier selected...!!!"
      Top             =   11880
      Width           =   4695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "da_generate_cust_repo.frx":1D2A
      Left            =   480
      List            =   "da_generate_cust_repo.frx":1D2C
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   -500
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid grid_report_cust 
      Height          =   4575
      Left            =   480
      TabIndex        =   8
      Top             =   1800
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   8070
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
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
      Height          =   735
      Left            =   18720
      TabIndex        =   6
      Top             =   11760
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print or Export"
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
      Left            =   16320
      TabIndex        =   5
      Top             =   11760
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   11160
      Width           =   20655
      _ExtentX        =   36433
      _ExtentY        =   450
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to Generate Customer Report"
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
      TabIndex        =   2
      Top             =   1320
      Width           =   10575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
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
      CurrentDate     =   40141
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
      Left            =   1680
      TabIndex        =   0
      Text            =   "Select Customer...!!!"
      Top             =   720
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   9000
      TabIndex        =   12
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
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
      CurrentDate     =   40141
   End
   Begin VB.Label Label4 
      Caption         =   "To"
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
      Left            =   8400
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Date"
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
      Left            =   5280
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Generate Deactivation Report."
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
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label1 
      Caption         =   "Customer"
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
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "da_Gene_cust_repo_main"
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
Me.Caption = "Ajay patel's card Deactivation...!!!  " & user_name

Call set_grid_report_cust
DTPicker1.Value = Date
DTPicker2.Value = Date

'add customer from ledger
Call add_ledgers

'set the combo for click
Combo2.AddItem "1"
Combo2.AddItem "2"
Combo2.AddItem "3"
End Sub
Public Sub add_ledgers()
Combo1.Clear
Call open_database
Call open_rs_lgr_main_dtl

Do Until rs_lgr_main_dtl.EOF
selected_group = rs_lgr_main_dtl!lgr_main_dtl_grup 'Combo1.Text
selected_primary_group = ""
        Call open_rs_lgr_main_grp
        Do Until rs_lgr_main_grp.EOF
            If selected_group = rs_lgr_main_grp!lgr_main_grp_name Or selected_group = rs_lgr_main_grp!lgr_main_grp_alis Then
            selected_primary_group = rs_lgr_main_grp!lgr_main_grp_pgrp
            End If
            rs_lgr_main_grp.MoveNext
        Loop
        
        If selected_primary_group = "" Then
            Call open_rs_lgr_prim_grp
            If rs_lgr_prim_grp.RecordCount > 0 Then rs_lgr_prim_grp.MoveFirst
            Do Until rs_lgr_prim_grp.EOF
            If selected_group = rs_lgr_prim_grp!lgr_prim_grp_name Then
            selected_primary_group = rs_lgr_prim_grp!lgr_prim_grp_name
            End If
            rs_lgr_prim_grp.MoveNext
            Loop
        End If
        
        If LCase(selected_primary_group) = LCase("Sundry Debtors") Then ' if the created ledger is a debtor then
            Combo1.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
        End If
rs_lgr_main_dtl.MoveNext
Loop
Combo1.Text = "Select Customer..,"
End Sub

Private Sub Command1_Click()
Call open_rs_dap_main_dtl_temp
Do Until rs_dap_main_dtl_temp.EOF
    rs_dap_main_dtl_temp.Delete
    rs_dap_main_dtl_temp.MoveNext
Loop
Call set_grid_report_cust
'find which user is operating the computer & doing the work
'selecting customer name through combo & date through date pick button
selected_cust_name = Combo1.Text
selected_starting_date = DTPicker1.Value
selected_ending_date = DTPicker2.Value
Call open_rs_dap_main_dtl_temp
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
If selected_cust_name = rs_lgr_main_dtl!lgr_main_dtl_name Then
    If rs_lgr_main_dtl!lgr_main_dtl_stnm = "" Or rs_lgr_main_dtl!lgr_main_dtl_stnm = Null Then
        MsgBox "you have to change customer shortcode...!!! Thank's"
        Exit Sub
    End If
ledger_sort_code = rs_lgr_main_dtl!lgr_main_dtl_stnm
End If
rs_lgr_main_dtl.MoveNext
Loop
'find which suppliers are in the data report & add them in the combos
'finding a sort code of supplier for ref no on data report
'copy sales of selected customer
Call open_rs_inv_tran_otw
If rs_inv_tran_otw.RecordCount > 0 Then rs_inv_tran_otw.MoveFirst
Dim aa As Integer
aa = 1
Do Until rs_inv_tran_otw.EOF
If rs_inv_tran_otw!stk_invt_trn_ldgr = selected_cust_name And rs_inv_tran_otw!stk_invt_trn_date >= selected_starting_date And rs_inv_tran_otw!stk_invt_trn_date <= selected_ending_date Then
    Call open_rs_lgr_main_dtl
    Do Until rs_lgr_main_dtl.EOF
    If rs_inv_tran_otw!stk_invt_trn_splr = rs_lgr_main_dtl!lgr_main_dtl_name Then
        If rs_lgr_main_dtl!lgr_main_dtl_stnm = "" Or rs_lgr_main_dtl!lgr_main_dtl_stnm = Null Then
            MsgBox "you have to change customer shortcode...!!! Thank's"
            Exit Do
        End If
    supl_sort_code = rs_lgr_main_dtl!lgr_main_dtl_stnm
    End If
    rs_lgr_main_dtl.MoveNext
    Loop
        rs_dap_main_dtl_temp.AddNew
        rs_dap_main_dtl_temp!dap_main_dtl_id = aa
        rs_dap_main_dtl_temp!dap_main_dtl_date = rs_inv_tran_otw!stk_invt_trn_date
        rs_dap_main_dtl_temp!dap_main_dtl_name = rs_inv_tran_otw!stk_invt_trn_ldgr
        rs_dap_main_dtl_temp!dap_main_dtl_card = rs_inv_tran_otw!stk_invt_trn_card
        rs_dap_main_dtl_temp!dap_main_dtl_stsl = rs_inv_tran_otw!stk_invt_trn_stno
        rs_dap_main_dtl_temp!dap_main_dtl_edsl = rs_inv_tran_otw!stk_invt_trn_edno
        rs_dap_main_dtl_temp!dap_main_dtl_qnty = rs_inv_tran_otw!stk_invt_trn_qnty
        rs_dap_main_dtl_temp!dap_main_dtl_rate = rs_inv_tran_otw!stk_invt_trn_rate
        rs_dap_main_dtl_temp!dap_main_dtl_amnt = rs_inv_tran_otw!stk_invt_trn_amnt
        rs_dap_main_dtl_temp!dap_main_dtl_comp = rs_inv_tran_otw!stk_invt_trn_comp
        rs_dap_main_dtl_temp!dap_main_dtl_splr = rs_inv_tran_otw!stk_invt_trn_splr
        rs_dap_main_dtl_temp!dap_main_dtl_csrf = temp_ref_no
        rs_dap_main_dtl_temp!dap_main_dtl_user = user_name
        
        grid_report_cust.AddItem aa
        grid_report_cust.TextMatrix(aa, 1) = rs_inv_tran_otw!stk_invt_trn_date
        grid_report_cust.TextMatrix(aa, 2) = rs_inv_tran_otw!stk_invt_trn_ldgr
        grid_report_cust.TextMatrix(aa, 3) = rs_inv_tran_otw!stk_invt_trn_card
        grid_report_cust.TextMatrix(aa, 4) = rs_inv_tran_otw!stk_invt_trn_stno
        grid_report_cust.TextMatrix(aa, 5) = rs_inv_tran_otw!stk_invt_trn_edno
        grid_report_cust.TextMatrix(aa, 6) = rs_inv_tran_otw!stk_invt_trn_qnty
        grid_report_cust.TextMatrix(aa, 7) = Format(rs_inv_tran_otw!stk_invt_trn_rate, "0.00")
        grid_report_cust.TextMatrix(aa, 8) = Format(rs_inv_tran_otw!stk_invt_trn_amnt, "0.00")
        'grid_report_cust.TextMatrix(aa, 9) = rs_inv_tran_otw!stk_invt_trn_comp
'        grid_report_cust.TextMatrix(aa, 10) = rs_inv_tran_otw!stk_invt_trn_splr
                Call open_rs_lgr_main_dtl
                Do Until rs_lgr_main_dtl.EOF
                    If LCase(rs_lgr_main_dtl!lgr_main_dtl_name) = LCase(rs_inv_tran_otw!stk_invt_trn_splr) Then
                    supl_sort_code = UCase(rs_lgr_main_dtl!lgr_main_dtl_stnm)
                    End If
                rs_lgr_main_dtl.MoveNext
                Loop
        temp_ref_no = ledger_sort_code & "-" & supl_sort_code & Str(Day(Now)) & Str(Month(Now)) & Str(Year(Now))
        grid_report_cust.TextMatrix(aa, 11) = temp_ref_no
        grid_report_cust.TextMatrix(aa, 12) = "" 'suplr ref no
        If rs_inv_tran_otw!stk_invt_trn_user <> "" Then grid_report_cust.TextMatrix(aa, 13) = rs_inv_tran_otw!stk_invt_trn_user
        rs_dap_main_dtl_temp.UpdateBatch
aa = aa + 1
End If
rs_inv_tran_otw.MoveNext
Loop

'add all the suplier in suplier combo list
Dim add_combo
Dim iii
Call open_rs_dap_main_dtl_temp
cmb_supplier.Clear
Do Until rs_dap_main_dtl_temp.EOF
add_combo = 1
    For iii = 1 To cmb_supplier.ListCount - 1
        cmb_supplier.ListIndex = iii
        If cmb_supplier.Text = rs_dap_main_dtl_temp!dap_main_dtl_splr Then add_combo = 0
    Next iii
    If add_combo = 1 Then
        cmb_supplier.AddItem rs_dap_main_dtl_temp!dap_main_dtl_splr
    End If
    rs_dap_main_dtl_temp.MoveNext
Loop
cmb_supplier.Text = "No supplier selected...!!!"
'sort combo of suplier
Call SortList(cmb_supplier, Val(0) \ 1, (Val(cmb_supplier.ListCount) - 1) \ 1, Ascending)
If cmb_supplier.ListCount > 0 Then cmb_supplier.ListIndex = 0
End Sub
Private Sub grid_report_cust_fill()
If rs_dap_main_dtl_temp.RecordCount > 0 Then rs_dap_main_dtl_temp.MoveFirst
Dim aa As Integer
aa = 1
Do Until rs_dap_main_dtl_temp.EOF
        grid_report_cust.AddItem aa
        grid_report_cust.TextMatrix(aa, 1) = rs_dap_main_dtl_temp!dap_main_dtl_date
        grid_report_cust.TextMatrix(aa, 2) = rs_dap_main_dtl_temp!dap_main_dtl_name
        grid_report_cust.TextMatrix(aa, 3) = rs_dap_main_dtl_temp!dap_main_dtl_card
        grid_report_cust.TextMatrix(aa, 4) = rs_dap_main_dtl_temp!dap_main_dtl_stsl
        grid_report_cust.TextMatrix(aa, 5) = rs_dap_main_dtl_temp!dap_main_dtl_edsl
        grid_report_cust.TextMatrix(aa, 6) = rs_dap_main_dtl_temp!dap_main_dtl_qnty
        grid_report_cust.TextMatrix(aa, 7) = rs_dap_main_dtl_temp!dap_main_dtl_rate
        grid_report_cust.TextMatrix(aa, 8) = rs_dap_main_dtl_temp!dap_main_dtl_amnt
        grid_report_cust.TextMatrix(aa, 9) = rs_dap_main_dtl_temp!dap_main_dtl_comp
        grid_report_cust.TextMatrix(aa, 10) = rs_dap_main_dtl_temp!dap_main_dtl_splr
'        grid_report_cust.TextMatrix(aa, 11) = rs_dap_main_dtl_temp!dap_main_dtl_csrf
'        grid_report_cust.TextMatrix(aa, 12) = rs_dap_main_dtl_temp!dap_main_dtl_rprf
        grid_report_cust.TextMatrix(aa, 13) = rs_dap_main_dtl_temp!dap_main_dtl_user
aa = aa + 1
Loop
End Sub

Private Sub Command2_Click()
Dim aa
'select a supleir from combo
Dim selected_supplier
selected_supplier = cmb_supplier.Text
If rs_dap_main_dtl_temp.State = 1 Then rs_dap_main_dtl_temp.Close
'set selected supleir data on report
rs_dap_main_dtl_temp.CursorLocation = adUseClient
rs_dap_main_dtl_temp.Open "Select * From dap_main_dtl_temp where dap_main_dtl_splr = '" & selected_supplier & "'", db_co, adOpenDynamic, adLockPessimistic
Call open_rs_dap_main_dtl_all
Call open_rs_dap_main_dtl
Call open_rs_dap_rspn_dtl
'check the record is already not noted or deactivated
Do Until rs_dap_main_dtl_temp.EOF
    Do Until rs_dap_main_dtl_all.EOF
    If rs_dap_main_dtl_all!dap_main_dtl_card = rs_dap_main_dtl_temp!dap_main_dtl_card And rs_dap_main_dtl_all!dap_main_dtl_stsl = rs_dap_main_dtl_temp!dap_main_dtl_stsl And rs_dap_main_dtl_all!dap_main_dtl_edsl = rs_dap_main_dtl_temp!dap_main_dtl_edsl Then
        MsgBox "Some or More card/cards are already deactivated...!!!"
        Exit Sub
    End If
    rs_dap_main_dtl_all.MoveNext
    Loop
    rs_dap_main_dtl_temp.MoveNext
Loop

If rs_dap_main_dtl_temp.State = 1 Then rs_dap_main_dtl_temp.Close

rs_dap_main_dtl_temp.CursorLocation = adUseClient
rs_dap_main_dtl_temp.Open "Select * From dap_main_dtl_temp where dap_main_dtl_splr = '" & selected_supplier & "'", db_co, adOpenDynamic, adLockPessimistic

Call open_rs_dap_main_dtl_all
Call open_rs_dap_main_dtl
Call open_rs_dap_rspn_dtl
If rs_dap_main_dtl_all.RecordCount > 0 Then rs_dap_main_dtl_all.MoveFirst
If rs_dap_main_dtl_temp.RecordCount > 0 Then rs_dap_main_dtl_temp.MoveFirst
Dim temp_cust_rf
Dim temp_cust_nm
Dim temp_supl_nm
Dim temp_tran_dt
Dim temp_main_us

'add data to main data file of card deactivation
Do Until rs_dap_main_dtl_temp.EOF
    rs_dap_main_dtl_all.AddNew
    rs_dap_main_dtl_all!dap_main_dtl_id = aa
    rs_dap_main_dtl_all!dap_main_dtl_date = rs_dap_main_dtl_temp!dap_main_dtl_date
    rs_dap_main_dtl_all!dap_main_dtl_name = rs_dap_main_dtl_temp!dap_main_dtl_name
    rs_dap_main_dtl_all!dap_main_dtl_card = rs_dap_main_dtl_temp!dap_main_dtl_card
    rs_dap_main_dtl_all!dap_main_dtl_stsl = rs_dap_main_dtl_temp!dap_main_dtl_stsl
    rs_dap_main_dtl_all!dap_main_dtl_edsl = rs_dap_main_dtl_temp!dap_main_dtl_edsl
    rs_dap_main_dtl_all!dap_main_dtl_qnty = rs_dap_main_dtl_temp!dap_main_dtl_qnty
    rs_dap_main_dtl_all!dap_main_dtl_rate = rs_dap_main_dtl_temp!dap_main_dtl_rate
    rs_dap_main_dtl_all!dap_main_dtl_amnt = rs_dap_main_dtl_temp!dap_main_dtl_amnt
    rs_dap_main_dtl_all!dap_main_dtl_comp = rs_dap_main_dtl_temp!dap_main_dtl_comp
    rs_dap_main_dtl_all!dap_main_dtl_splr = rs_dap_main_dtl_temp!dap_main_dtl_splr
    rs_dap_main_dtl_all!dap_main_dtl_csrf = rs_dap_main_dtl_temp!dap_main_dtl_csrf
    rs_dap_main_dtl_all!dap_main_dtl_user = rs_dap_main_dtl_temp!dap_main_dtl_user
    rs_dap_main_dtl_all.UpdateBatch
    aa = aa + 1
    
    temp_cust_rf = rs_dap_main_dtl_temp!dap_main_dtl_csrf
    temp_cust_nm = rs_dap_main_dtl_temp!dap_main_dtl_name
    temp_supl_nm = rs_dap_main_dtl_temp!dap_main_dtl_splr
    temp_tran_dt = rs_dap_main_dtl_temp!dap_main_dtl_date
    temp_main_us = rs_dap_main_dtl_temp!dap_main_dtl_user
    rs_dap_main_dtl_temp.MoveNext
Loop

'make entry in response / confirmation & payment data
rs_dap_rspn_dtl.AddNew
rs_dap_rspn_dtl!dap_main_rsp_trrf = temp_cust_rf
rs_dap_rspn_dtl!dap_main_rsp_csnm = temp_cust_nm
rs_dap_rspn_dtl!dap_main_rsp_spnm = temp_supl_nm
rs_dap_rspn_dtl!dap_main_rsp_trdt = temp_tran_dt
rs_dap_rspn_dtl!dap_main_rsp_trus = temp_main_us
rs_dap_rspn_dtl.UpdateBatch

If rs_dap_main_dtl_temp.State = 1 Then rs_dap_main_dtl_temp.Close
rs_dap_main_dtl_temp.CursorLocation = adUseClient
rs_dap_main_dtl_temp.Open "Select * From dap_main_dtl_temp where dap_main_dtl_splr = '" & selected_supplier & "'", db_co, adOpenDynamic, adLockPessimistic

'set the ref no on the report on section 2
With deact_repo_gen.Sections("section2").Controls
    .item("label13").Caption = temp_ref_no
End With

Set deact_repo_gen.DataSource = rs_dap_main_dtl_temp
deact_repo_gen.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Combo2_Click()
If grid_report_cust.Col = 2 Then
grid_report_cust.Text = Combo2.Text
Combo2.Visible = False
End If
End Sub

Private Sub grid_report_cust_Click()
    'when click on such column how to active a combo
'    If grid_report_cust.Col = 2 Then      ' Position and size the ComboBox, then show it.
'        Combo2.Width = grid_report_cust.CellWidth
'        Combo2.Left = grid_report_cust.CellLeft + grid_report_cust.Left
'        Combo2.Top = grid_report_cust.CellTop + grid_report_cust.Top
'        Combo2.Text = grid_report_cust.Text
'        Combo2.Visible = True
'    End If
End Sub
Public Sub set_grid_report_cust()
        'set data grid
        grid_report_cust.Clear
        grid_report_cust.Rows = 1
        grid_report_cust.Cols = 14
        grid_report_cust.TextMatrix(0, 1) = "Date"
        grid_report_cust.TextMatrix(0, 2) = "Name"
        grid_report_cust.TextMatrix(0, 3) = "Item"
        grid_report_cust.TextMatrix(0, 4) = "Start-No"
        grid_report_cust.TextMatrix(0, 5) = "End-no"
        grid_report_cust.TextMatrix(0, 6) = "Qty."
        grid_report_cust.TextMatrix(0, 7) = "Rate"
        grid_report_cust.TextMatrix(0, 8) = "Amount"
        grid_report_cust.TextMatrix(0, 9) = "Company"
        grid_report_cust.TextMatrix(0, 10) = "Supplier"
        grid_report_cust.TextMatrix(0, 11) = "Cus.Ref No"
        grid_report_cust.TextMatrix(0, 12) = "Res.Ref No"
        grid_report_cust.TextMatrix(0, 13) = "User"
        
grid_report_cust.CellAlignment = Center
grid_report_cust.ColWidth(0) = 300
grid_report_cust.ColWidth(1) = 1000
grid_report_cust.ColWidth(2) = 1500
grid_report_cust.ColWidth(3) = 1200
grid_report_cust.ColWidth(4) = 1500
grid_report_cust.ColWidth(5) = 1500
grid_report_cust.ColWidth(6) = 800
grid_report_cust.ColWidth(7) = 500
grid_report_cust.ColWidth(8) = 800
grid_report_cust.ColWidth(9) = 1500
grid_report_cust.ColWidth(10) = 1500
grid_report_cust.ColWidth(11) = 1500
grid_report_cust.ColWidth(12) = 1500
grid_report_cust.ColWidth(13) = 1000

End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form vchr_purchase 
   Caption         =   "purchase Voucher"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.ListBox cmb_ledger 
      Height          =   450
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ListBox List_card 
      Height          =   450
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Height          =   6015
      Left            =   12240
      TabIndex        =   14
      Top             =   3120
      Width           =   5655
      Begin MSFlexGridLib.MSFlexGrid grid_lgr_rep 
         Height          =   5655
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   9975
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   12240
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
      Height          =   435
      Left            =   2400
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   8880
      TabIndex        =   4
      Top             =   720
      Width           =   2535
      Begin VB.CommandButton cmd_sv_n_new 
         Caption         =   "&Save and New"
         Height          =   435
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmb_edit 
         Caption         =   "Edit"
         Height          =   435
         Left            =   1320
         TabIndex        =   11
         Top             =   1200
         Width           =   1065
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "Delete"
         Height          =   435
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   1065
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         Height          =   435
         Left            =   1320
         TabIndex        =   9
         Top             =   1680
         Width           =   1065
      End
      Begin VB.CommandButton cmd_exit 
         Caption         =   "E&xit"
         Height          =   435
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1185
      End
      Begin VB.CommandButton cmd_save_and_exit 
         Caption         =   "Save and Exit"
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1185
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   435
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New"
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1185
      End
   End
   Begin VB.ComboBox cmb_inv_vat_type 
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
      Left            =   3720
      TabIndex        =   3
      Text            =   "combo_inv_vat_type"
      Top             =   240
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Any Serial No. Not allowed"
      Height          =   255
      Left            =   9000
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid grid_stk_dtl 
      Height          =   3375
      Left            =   240
      TabIndex        =   40
      Top             =   3240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   25
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      ScrollBars      =   2
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Detail"
      Height          =   2415
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   8655
      Begin VB.TextBox text_ledger 
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox cmb_ledger1 
         Height          =   315
         Left            =   960
         TabIndex        =   27
         Text            =   "combo_ledger"
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   960
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   960
         TabIndex        =   25
         Text            =   "Text3"
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   5520
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmb_sales_man 
         Height          =   315
         Left            =   5520
         TabIndex        =   22
         Text            =   "combo_sales_man"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cmb_vat_type 
         Height          =   315
         Left            =   5520
         TabIndex        =   21
         Text            =   "combo_vat_type"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   7320
         TabIndex        =   20
         Text            =   "Text6"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   960
         TabIndex        =   19
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         Height          =   735
         Left            =   5520
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "vchr_sale_return1.frx":0000
         Top             =   720
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   7080
         TabIndex        =   23
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   6488065
         CurrentDate     =   40177
      End
      Begin VB.Label lbl_l_name 
         Caption         =   "Name"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbl_l_add 
         Caption         =   "Address"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lbl_l_trans 
         Caption         =   "Guid"
         Height          =   375
         Left            =   4680
         TabIndex        =   37
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbl_inv_no 
         Caption         =   "Inv.No."
         Height          =   375
         Left            =   4680
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl_date 
         Caption         =   "Date"
         Height          =   375
         Left            =   6720
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbl_user 
         Caption         =   "User"
         Height          =   375
         Left            =   6840
         TabIndex        =   34
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lbl_sales_man 
         Caption         =   "Sales man"
         Height          =   375
         Left            =   4680
         TabIndex        =   33
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lbl_vat_type 
         Caption         =   "Vat type"
         Height          =   375
         Left            =   4680
         TabIndex        =   32
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lbl_tel_no 
         Caption         =   "Telephone"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lbl_mob_no 
         Caption         =   "Mobile"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lbl_time 
         Caption         =   "time"
         Height          =   255
         Left            =   6840
         TabIndex        =   29
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lbl_day 
         Caption         =   "day"
         Height          =   255
         Left            =   7320
         TabIndex        =   28
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.Label lbl_head 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   240
      TabIndex        =   47
      Top             =   360
      Width           =   11295
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name of company"
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
      Left            =   120
      TabIndex        =   46
      Top             =   0
      Width           =   12615
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   255
      Left            =   0
      TabIndex        =   45
      Top             =   240
      Width           =   11295
   End
   Begin VB.Label lbl_customer_marq 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   44
      Top             =   11160
      Width           =   1410
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   240
      Top             =   6840
      Width           =   10935
   End
   Begin VB.Line Line1 
      X1              =   7440
      X2              =   7440
      Y1              =   6840
      Y2              =   7320
   End
   Begin VB.Line Line2 
      X1              =   8760
      X2              =   8760
      Y1              =   6840
      Y2              =   7320
   End
   Begin VB.Line Line3 
      X1              =   10200
      X2              =   10200
      Y1              =   6840
      Y2              =   7320
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL :"
      Height          =   255
      Left            =   4800
      TabIndex        =   41
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Label lbl_total_amt 
      Alignment       =   2  'Center
      Caption         =   "total_amount"
      Height          =   375
      Left            =   8760
      TabIndex        =   42
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lbl_total_qty 
      Alignment       =   1  'Right Justify
      Caption         =   "total_quantity"
      Height          =   375
      Left            =   7680
      TabIndex        =   43
      Top             =   6960
      Width           =   975
   End
End
Attribute VB_Name = "vchr_purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmb_ledger_LostFocus()
cmb_ledger.Visible = False
End Sub

Private Sub Command2_Click()
    Call set_form_detail_zero
    Call set_ledger_grid_zero
    Call set_grid_stk_dtl_data
    Call find_last_voucher_no
End Sub

Private Sub Form_Unload(Cancel As Integer)
temp_selected_procedure = selected_procedure
Dim x_temp_list_item_remove
If MDIForm1.List_opened_procedure.ListCount > 0 Then
For x_temp_list_item_remove = 0 To (MDIForm1.List_opened_procedure.ListCount - 1)
MDIForm1.List_opened_procedure.ListIndex = x_temp_list_item_remove
If MDIForm1.List_opened_procedure.Text = temp_selected_procedure Then
    MDIForm1.List_opened_procedure.RemoveItem (x_temp_list_item_remove)
End If
Next
End If

End Sub
Private Sub cmd_delete_Click()
Dim delete_sure
delete_sure = MsgBox("You want to delete voucher....?", vbQuestion + vbYesNo, "Are You Sure !!!!")
If delete_sure = 6 Then
    Call delete_accouting_transaction
    Call delete_inventory_transaction
    Call set_form_detail_zero
    Call set_ledger_grid_zero
    Call set_grid_stk_dtl_data
    Call find_last_voucher_no
End If
End Sub
Public Sub delete_inventory_transaction()
Dim transaction_counter
Dim this_entry_is_saved
Dim available_tran_no
Dim inventory_sub_entry_no
this_entry_is_saved = 0
Call open_database
Call open_rs_inv_tran_prs
For available_tran_no = 1 To rs_inv_tran_prs.RecordCount
    If rs_inv_tran_prs!stk_invt_trn_vcno = Text5.Text Then ' if the transaction is available
        inventory_transaction_is_available = 0
        rs_inv_tran_prs.Delete
        rs_inv_tran_prs.UpdateBatch
    End If
rs_inv_tran_prs.MoveNext
Next
'==========delete from all transaction
Call open_rs_inv_tran_all
For available_tran_no = 1 To rs_inv_tran_all.RecordCount
    If rs_inv_tran_all!stk_invt_trn_vcno = Text5.Text And rs_inv_tran_all!stk_invt_trn_vchr = "purchase" Then ' if the transaction is available
        inventory_transaction_is_available = 0
        rs_inv_tran_all.Delete
        rs_inv_tran_all.UpdateBatch
    End If
rs_inv_tran_all.MoveNext
Next
'==========delete from in transaction
Call open_rs_inv_tran_inw
For available_tran_no = 1 To rs_inv_tran_inw.RecordCount
    If rs_inv_tran_inw!stk_invt_trn_vcno = Text5.Text And rs_inv_tran_inw!stk_invt_trn_vchr = "purchase" Then ' if the transaction is available
        inventory_transaction_is_available = 0
        rs_inv_tran_inw.Delete
        rs_inv_tran_inw.UpdateBatch
    End If
rs_inv_tran_inw.MoveNext
Next
End Sub
Public Sub delete_accouting_transaction()
Dim transaction_counter
Dim this_entry_is_saved
Dim available_tran_no
this_entry_is_saved = 0
    Call open_database
    Call open_rs_acn_tran_prs
For available_tran_no = 1 To rs_acn_tran_prs.RecordCount
    If rs_acn_tran_prs!fin_acnt_trn_vcno = Text5.Text Then
            rs_acn_tran_prs.Delete
            rs_acn_tran_prs.UpdateBatch
    End If
    rs_acn_tran_prs.MoveNext
Next
End Sub

Private Sub cmb_edit_Click()
    Call enable_all_controls
    cmb_ledger.Height = 2400
    cmb_ledger.Visible = True
    cmb_ledger.SetFocus
    List_card.Left = -2500
    Text1.Left = -2500
    cmb_inv_vat_type.Left = -2500
    text_ledger.Enabled = True
End Sub
Public Sub enable_all_controls()
List1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
grid_stk_dtl.Enabled = True
cmb_ledger.Enabled = True
'cmb_sales_man.Enabled = True
DTPicker1.Enabled = True
cmb_vat_type.Enabled = True
'cmb_inv_vat_type.Visible = True
Text1.Visible = True
List_card.Visible = True
cmb_edit.Enabled = False
End Sub
Private Sub cmb_inv_vat_type_KeyDown(KeyCode As Integer, Shift As Integer)
Exit Sub
keycode_now = KeyCode
If keycode_now = 13 Or keycode_now = 39 Then
                If cmb_inv_vat_type.Text < 2 Or cmb_inv_vat_type < 1 Then
                    Exit Sub
                End If
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = cmb_inv_vat_type.Text
                If grid_stk_dtl.Rows = grid_stk_dtl.Row + 1 Then
                    grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
                End If
                grid_stk_dtl.Row = grid_stk_dtl.Row + 1
                grid_stk_dtl.Col = 1
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 0) = grid_stk_dtl.Row
                cmb_inv_vat_type.Visible = False
                Text1.Visible = False
                List_card.Visible = True
                Call next_cardlist_cell
                Call stock_refresh_total
                Exit Sub
ElseIf keycode_now = 37 Then
        grid_stk_dtl.Col = 6
        cmb_inv_vat_type.Visible = False
        Text1.Visible = True
        List_card.Visible = False
        Call next_text_cell
        Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
        Exit Sub
ElseIf keycode_now = 38 Then
        If grid_stk_dtl.Row = 1 Then
        MsgBox "Not Valid key.....!!!"
        Exit Sub
        Else
            grid_stk_dtl.Row = grid_stk_dtl.Row - 1
            'cmb_inv_vat_type.Visible = True
            Text1.Visible = False
            List_card.Visible = False
            Call next_cmb_inv_vat_type
            Exit Sub
        End If
ElseIf keycode_now = 40 Then
    If grid_stk_dtl.Rows = grid_stk_dtl.Row + 1 Then
        MsgBox "Not Valid key.....!!!"
        Exit Sub
    Else
        grid_stk_dtl.Row = grid_stk_dtl.Row + 1
        Text1.Visible = False
        List_card.Visible = False
        Call next_cmb_inv_vat_type
        Exit Sub
    End If
End If
grid_stk_dtl.Font.Size = 6
keycode_now = 0
End Sub
Private Sub cmb_ledger_DblClick()
    
    Call open_database
    Call open_rs_lgr_main_dtl
    Do Until rs_lgr_main_dtl.EOF
    If cmb_ledger.Text = rs_lgr_main_dtl!lgr_main_dtl_alis Then
        cmb_ledger.Text = rs_lgr_main_dtl!lgr_main_dtl_name
    End If
    rs_lgr_main_dtl.MoveNext
    Loop
    
    selected_voucher_ledger = cmb_ledger.Text
    Call set_selected_ledger_detail
    cmb_ledger.Visible = False
    text_ledger.Text = selected_voucher_ledger
'    text_ledger.Visible = True
    If grid_stk_dtl.Enabled = False Then grid_stk_dtl.Enabled = True
    Call go_to_card_detail
End Sub
Private Sub cmb_ledger_KeyDown(KeyCode As Integer, Shift As Integer)
keycode_now = KeyCode
If keycode_now = 13 Then
    
    
    Call open_database
    Call open_rs_lgr_main_dtl
    Do Until rs_lgr_main_dtl.EOF
    If cmb_ledger.Text = rs_lgr_main_dtl!lgr_main_dtl_alis Then
        cmb_ledger.Text = rs_lgr_main_dtl!lgr_main_dtl_name
    End If
    rs_lgr_main_dtl.MoveNext
    Loop
    
    selected_voucher_ledger = cmb_ledger.Text
    Call set_selected_ledger_detail
    cmb_ledger.Visible = False
    text_ledger.Text = selected_voucher_ledger
    
    If grid_stk_dtl.Enabled = False Then grid_stk_dtl.Enabled = True
    Call go_to_card_detail
End If
keycode_now = 0
End Sub
Private Sub cmd_save_and_exit_Click()
Call save_accouting_transaction
    Call Save_inventory_transaction
Unload Me
End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub


Private Sub cmd_sv_n_new_Click()
    Call save_new_transaction
    If db_co.State = 1 Then db_co.Close
'FileCopy selected_path, selected_backup_path
End Sub
Public Sub save_new_transaction()
If Val(lbl_total_amt.Caption) = 0 Or cmb_ledger.Text = "" Then
    MsgBox "Entered value is zero or incorrect account selected. Try again...,"
    Exit Sub
End If

    Call save_accouting_transaction
    Call Save_inventory_transaction
    
    Call set_form_detail_zero
    Call set_ledger_grid_zero
    Call set_grid_stk_dtl_data
    Call find_last_voucher_no
End Sub
Public Sub Save_inventory_transaction()
Dim transaction_counter
Dim this_entry_is_saved
Dim available_tran_no
Dim inventory_sub_entry_no
this_entry_is_saved = 0
Call open_database
Call open_rs_inv_tran_prs
For available_tran_no = 1 To rs_inv_tran_prs.RecordCount
    If rs_inv_tran_prs!stk_invt_trn_vcno = Text5.Text Then ' if the transaction is available
        inventory_transaction_is_available = 0
        'MsgBox "this transaction is available..."
        this_entry_is_saved = 1
        rs_inv_tran_prs.Delete
        rs_inv_tran_prs.UpdateBatch
    End If
rs_inv_tran_prs.MoveNext
Next

'If this_entry_is_saved <> 1 Then
    Call open_database
    Call open_rs_inv_tran_prs
    For inventory_sub_entry_no = 1 To grid_stk_dtl.Rows - 1
    If Val(grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 6)) <= 0 Then Exit Sub
    With rs_inv_tran_prs
        .AddNew
        !stk_invt_trn_vcno = Text5.Text
        !stk_invt_trn_seno = inventory_sub_entry_no
        !stk_invt_trn_date = DTPicker1.Value
        !stk_invt_trn_time = lbl_time.Caption
        !stk_invt_trn_wday = lbl_day.Caption
        !stk_invt_trn_vtyp = cmb_vat_type
        !stk_invt_trn_ldgr = cmb_ledger.Text
        !stk_invt_trn_card = grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 1)
        !stk_invt_trn_stno = grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 2)
        !stk_invt_trn_edno = grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 3)
        !stk_invt_trn_qnty = Val(grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 4))
        !stk_invt_trn_rate = Val(grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 5))
        !stk_invt_trn_amnt = Val(grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 6))
        !stk_invt_trn_fval = Val(grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 7))
        !stk_invt_trn_splr = cmb_ledger.Text 'grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 8)
        !stk_invt_trn_comp = grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 8)
        !stk_invt_trn_user = Text6.Text
        !stk_invt_trn_vchr = "purchase"
        '!stk_invt_trn_slmn = cmb_sales_man.Text
        !stk_invt_trn_vtyp = Val(grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 10))
        !stk_invt_trn_splr = grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 11)
        .UpdateBatch
    End With
    Next
'End If
If this_entry_is_saved = 1 Then
Unload Me
End If
End Sub
Public Sub save_accouting_transaction()
Dim transaction_counter
Dim this_entry_is_saved
Dim available_tran_no
this_entry_is_saved = 0
    Call open_database
    Call open_rs_acn_tran_prs
If rs_acn_tran_prs.RecordCount > 0 Then
    For available_tran_no = 1 To rs_acn_tran_prs.RecordCount
    If rs_acn_tran_prs!fin_acnt_trn_vcno = Text5.Text Then
            rs_acn_tran_prs!fin_acnt_trn_vcno = Text5.Text
            rs_acn_tran_prs!fin_acnt_trn_seno = 1 'transaction_counter
            rs_acn_tran_prs!fin_acnt_trn_vtyp = cmb_vat_type 'Combo0.Text
            rs_acn_tran_prs!fin_acnt_trn_date = DTPicker1.Value
            rs_acn_tran_prs!fin_acnt_trn_time = lbl_time.Caption
            rs_acn_tran_prs!fin_acnt_trn_wday = lbl_day.Caption
            rs_acn_tran_prs!fin_acnt_trn_ldgr = "purchase account" ' combo_lgr(transaction_counter).Text
            rs_acn_tran_prs!fin_acnt_trn_amnt = lbl_total_amt.Caption ' text_amt(transaction_counter).Text
            rs_acn_tran_prs!fin_acnt_trn_side = "dr"
'            !fin_acnt_trn_nrtn = Text4.Text
            rs_acn_tran_prs!fin_acnt_trn_user = Text6.Text
            rs_acn_tran_prs!fin_acnt_trn_vchr = "purchase"
            'rs_acn_tran_prs!fin_acnt_trn_slmn = cmb_sales_man.Text
            rs_acn_tran_prs.UpdateBatch
            
            rs_acn_tran_prs.MoveNext
            rs_acn_tran_prs!fin_acnt_trn_vcno = Text5.Text
            rs_acn_tran_prs!fin_acnt_trn_seno = 2 'transaction_counter
            rs_acn_tran_prs!fin_acnt_trn_vtyp = cmb_vat_type 'Combo0.Text
            rs_acn_tran_prs!fin_acnt_trn_date = DTPicker1.Value
            rs_acn_tran_prs!fin_acnt_trn_time = lbl_time.Caption
            rs_acn_tran_prs!fin_acnt_trn_wday = lbl_day.Caption
            rs_acn_tran_prs!fin_acnt_trn_ldgr = cmb_ledger.Text ' combo_lgr(transaction_counter).Text"
            rs_acn_tran_prs!fin_acnt_trn_amnt = lbl_total_amt.Caption ' text_amt(transaction_counter).Text
            rs_acn_tran_prs!fin_acnt_trn_side = "cr"
'            !fin_acnt_trn_nrtn = Text4.Text
            rs_acn_tran_prs!fin_acnt_trn_user = Text6.Text
            rs_acn_tran_prs!fin_acnt_trn_vchr = "purchase"
            'rs_acn_tran_prs!fin_acnt_trn_slmn = cmb_sales_man.Text
            rs_acn_tran_prs.UpdateBatch
            this_entry_is_saved = 1
            Exit For
    End If
    rs_acn_tran_prs.MoveNext
    Next
End If
    If this_entry_is_saved <> 1 Then
        Call open_database
        Call open_rs_acn_tran_prs
            rs_acn_tran_prs.AddNew
            rs_acn_tran_prs!fin_acnt_trn_vcno = Text5.Text
            rs_acn_tran_prs!fin_acnt_trn_seno = 1 'transaction_counter
            rs_acn_tran_prs!fin_acnt_trn_vtyp = cmb_vat_type 'Combo0.Text
            rs_acn_tran_prs!fin_acnt_trn_date = DTPicker1.Value
            rs_acn_tran_prs!fin_acnt_trn_time = lbl_time.Caption
            rs_acn_tran_prs!fin_acnt_trn_wday = lbl_day.Caption
            rs_acn_tran_prs!fin_acnt_trn_ldgr = "purchase account" ' combo_lgr(transaction_counter).Text
            rs_acn_tran_prs!fin_acnt_trn_amnt = lbl_total_amt.Caption ' text_amt(transaction_counter).Text
            rs_acn_tran_prs!fin_acnt_trn_side = "dr"
'            !fin_acnt_trn_nrtn = Text4.Text
            rs_acn_tran_prs!fin_acnt_trn_user = Text6.Text
            rs_acn_tran_prs!fin_acnt_trn_vchr = "purchase"
            'rs_acn_tran_prs!fin_acnt_trn_slmn = cmb_sales_man.Text
            rs_acn_tran_prs.UpdateBatch
            rs_acn_tran_prs.AddNew
            rs_acn_tran_prs!fin_acnt_trn_vcno = Text5.Text
            rs_acn_tran_prs!fin_acnt_trn_seno = 2 'transaction_counter
            rs_acn_tran_prs!fin_acnt_trn_vtyp = cmb_vat_type 'Combo0.Text
            rs_acn_tran_prs!fin_acnt_trn_date = DTPicker1.Value
            rs_acn_tran_prs!fin_acnt_trn_time = lbl_time.Caption
            rs_acn_tran_prs!fin_acnt_trn_wday = lbl_day.Caption
            rs_acn_tran_prs!fin_acnt_trn_ldgr = cmb_ledger.Text ' combo_lgr(transaction_counter).Text"
            rs_acn_tran_prs!fin_acnt_trn_amnt = lbl_total_amt.Caption ' text_amt(transaction_counter).Text
            rs_acn_tran_prs!fin_acnt_trn_side = "cr"
'            !fin_acnt_trn_nrtn = Text4.Text
            rs_acn_tran_prs!fin_acnt_trn_user = Text6.Text
            rs_acn_tran_prs!fin_acnt_trn_vchr = "purchase"
            'rs_acn_tran_prs!fin_acnt_trn_slmn = cmb_sales_man.Text
            rs_acn_tran_prs.UpdateBatch
            
    End If
End Sub
Public Sub find_last_voucher_no()
Call open_database
Call open_rs_acn_tran_prs
Dim iflvn
Dim this_voucher_no
Dim biggest_voucher_no
If rs_acn_tran_prs.RecordCount > 0 Then
    For iflvn = 1 To rs_acn_tran_prs.RecordCount
    this_voucher_no = rs_acn_tran_prs!fin_acnt_trn_vcno
    If this_voucher_no > biggest_voucher_no Then
       biggest_voucher_no = this_voucher_no
    End If
    rs_acn_tran_prs.MoveNext
    Next
End If
Text5.Text = biggest_voucher_no + 1
End Sub
Public Sub set_card_detail()
    Call open_database
    Call open_rs_stk_item_lgr
    Do Until rs_stk_item_lgr.EOF
        If List_card.Text = rs_stk_item_lgr!stk_item_lgr_name Then
        grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 7) = Format(rs_stk_item_lgr!stk_item_lgr_fcvl, "0.00")
        grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 8) = rs_stk_item_lgr!stk_item_lgr_disc & " %"
        grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 9) = rs_stk_item_lgr!stk_item_lgr_comp
        grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 10) = selected_voucher_ledger 'selected_voucher_ledger 'cmb_vat_type.Text
        Exit Do
        End If
        rs_stk_item_lgr.MoveNext
    Loop
    
    Call open_database
    Call open_rs_inv_tran_prs
    Dim rs_inv_tran_prs_counter
    If rs_inv_tran_prs.RecordCount > 0 Then rs_inv_tran_prs.MoveLast
    
    For rs_inv_tran_prs_counter = rs_inv_tran_prs.RecordCount To 1 Step -1
        If List_card.Text = rs_inv_tran_prs!stk_invt_trn_card And cmb_ledger.Text = rs_inv_tran_prs!stk_invt_trn_ldgr Then
        grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 5) = Format(rs_inv_tran_prs!stk_invt_trn_rate, "0.00")
        End If
        rs_inv_tran_prs.MovePrevious
    Next
    
    selected_stock_item_name = List_card.Text
    Call separation_of_all_inventory_to_inward_and_outward  'merge all transactions
    Call search_closing_stock                               'calculate closing stock
    Call open_rs_tmp_spec_itm_clg_stk
    Do Until rs_tmp_spec_itm_clg_stk.EOF
        rs_tmp_spec_itm_clg_stk.Delete
        rs_tmp_spec_itm_clg_stk.UpdateBatch
        rs_tmp_spec_itm_clg_stk.MoveNext
    Loop
    Call open_rs_tmp_spec_itm_clg_stk
    Call open_rs_stk_clsg_srl
    
        Dim rs_stk_clsg_srl_counter
        Dim grid_stk_row_no
        Dim total_inward
        Dim total_outward
        Dim temp_stock_balance
        
        grid_stk_row_no = 1
        grid_stk_dtl.Font.Size = 12
        
        Do Until rs_stk_clsg_srl.EOF
        If rs_stk_clsg_srl!stk_invt_clg_card = selected_stock_item_name Then
            With rs_stk_clsg_srl
                    rs_tmp_spec_itm_clg_stk.AddNew
                    rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_id = grid_stk_row_no
                    rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_card = !stk_invt_clg_card
                    If !stk_invt_clg_stno <> "" Then rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_stno = !stk_invt_clg_stno
                    If !stk_invt_clg_edno <> "" Then rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_edno = !stk_invt_clg_edno
                    If !stk_invt_clg_rate <> "" Then rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_rate = Format(!stk_invt_clg_rate, "0.00")
                    If !stk_invt_clg_fval <> "" Then rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_fval = Format(!stk_invt_clg_fval, "0.00")
                    If !stk_invt_clg_comp <> "" Then rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_comp = !stk_invt_clg_comp
                    If !stk_invt_clg_vtyp <> "" Then rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_vtyp = !stk_invt_clg_vtyp
                    If !stk_invt_clg_splr <> "" Then rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_splr = !stk_invt_clg_splr
                    grid_stk_row_no = grid_stk_row_no + 1
                    rs_tmp_spec_itm_clg_stk.UpdateBatch
            End With
        End If
        rs_stk_clsg_srl.MoveNext
        Loop
    
End Sub
Public Sub search_this_stock_serial_no_is_avilable()
Call open_database
'Call open_rs_stk_clsg_srl
Call open_rs_tmp_spec_itm_clg_stk
Do Until rs_tmp_spec_itm_clg_stk.EOF
    If Val(Text1.Text) >= Val(rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_stno) And Val(rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_edno) >= Val(Text1.Text) Then
    MsgBox "Stock is available"
    Exit Sub
    End If
rs_tmp_spec_itm_clg_stk.MoveNext
Loop
MsgBox "Stock is not available"
End Sub

Private Sub grid_stk_dtl_GotFocus()
grid_stk_dtl.Font.Size = 6
End Sub

Private Sub List_card_DblClick()
    Call select_card
    Call set_card_detail
    grid_stk_dtl.Col = grid_stk_dtl.Col + 1
    Call next_text_cell
End Sub
Private Sub List_card_KeyDown(KeyCode As Integer, Shift As Integer)
grid_stk_dtl.Font.Size = 8
keycode_now = KeyCode
If keycode_now = 37 Then
        If grid_stk_dtl.Row = 1 Then
            MsgBox "Not Valid key.....!!!"
        Else
                grid_stk_dtl.Row = grid_stk_dtl.Row - 1
                grid_stk_dtl.Col = 6
                'cmb_inv_vat_type.Visible = False
                Text1.Visible = True
                List_card.Visible = False
                Call next_text_cell
                Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
                Exit Sub


                'grid_stk_dtl.Col = 10
                'grid_stk_dtl.Row = grid_stk_dtl.Row - 1
                'Call next_cmb_inv_vat_type
                'grid_stk_dtl.Col = 5
                'Call next_text_cell
        End If
ElseIf keycode_now = 38 Then
'        If grid_stk_dtl.Row = 1 Then
'            MsgBox "Not Valid key.....!!!"
'        Else
'            grid_stk_dtl.Row = grid_stk_dtl.Row - 1
'            Call next_cardlist_cell
'        End If
ElseIf keycode_now = 39 Or keycode_now = 13 Then
    Call select_card
    Call set_card_detail
    grid_stk_dtl.Col = grid_stk_dtl.Col + 1
    Call next_text_cell
    
    cmb_ledger.Visible = False
    If grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) <> "" Then
        Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
    Else
        Text1.Text = ""
    End If
ElseIf keycode_now = 40 Then
'        If grid_stk_dtl.Rows = (grid_stk_dtl.Row + 1) Then
'            MsgBox "Not Valid key.....!!!"
'        Else
'            grid_stk_dtl.Row = grid_stk_dtl.Row + 1
'            Call next_cardlist_cell
'        End If
End If
keycode_now = 0
End Sub
Public Sub select_card()
    Call open_database
    Call open_rs_stk_item_lgr
    Do Until rs_stk_item_lgr.EOF
        If List_card.Text = rs_stk_item_lgr!stk_item_lgr_name Then
            selected_voucher_card = rs_stk_item_lgr!stk_item_lgr_name
            Exit Do
        ElseIf List_card.Text = rs_stk_item_lgr!stk_item_lgr_alis Then
            selected_voucher_card = rs_stk_item_lgr!stk_item_lgr_name
            grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = List_card.Text
            List_card.Text = selected_voucher_card
            Exit Do
        End If
        rs_stk_item_lgr.MoveNext
    Loop
    Dim card_name_is_available
    card_name_is_available = 0
    Call open_database
    Call open_rs_stk_item_lgr
    Do Until rs_stk_item_lgr.EOF
        If List_card.Text = rs_stk_item_lgr!stk_item_lgr_name Then
            card_name_is_available = 1
            Exit Do
        End If
        rs_stk_item_lgr.MoveNext
    Loop
    If card_name_is_available = 0 Then
    MsgBox "Sorry!!! Not Valid card.....!!!"
    Exit Sub
    End If
    grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = selected_voucher_card
    List_card.Visible = False
End Sub
Private Sub List_card_GotFocus()
grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = List_card.Text
List_card.Height = 1500
End Sub
Private Sub List_card_LostFocus()
Call open_database
Call open_rs_stk_item_lgr
Do Until rs_stk_item_lgr.EOF
    If List_card.Text = rs_stk_item_lgr!stk_item_lgr_alis Then
        selected_voucher_card = rs_stk_item_lgr!stk_item_lgr_name
        grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = List_card.Text
        List_card.Text = selected_voucher_card
        Exit Do
    End If
    rs_stk_item_lgr.MoveNext
Loop
'If grid_stk_dtl.Col = 1 Then grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = selected_voucher_card
'    grid_stk_dtl.Col = grid_stk_dtl.Col + 1
'    Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
'    If Text1.Visible = False Then
'        Text1.Visible = True
'    End If
'    Call next_text_cell
End Sub
Private Sub DTPicker1_Change()
lbl_day.Caption = WeekdayName(Weekday(DTPicker1.Value - 1))
End Sub
Private Sub grid_stk_dtl_Click()
If grid_stk_dtl.Col = 0 Or grid_stk_dtl.Row = 0 And show_ledger_detail <> 1 Then
'        MsgBox "Not Valid key.....!!!"
        Exit Sub
ElseIf grid_stk_dtl.Col = 0 Or grid_stk_dtl.Row = 0 And show_ledger_detail = 1 Then
        Exit Sub
ElseIf grid_stk_dtl.Col = 1 Then
    'If list_card.Visible = False Then
        List_card.Visible = True
        Call next_cardlist_cell
        grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = List_card.Text
    'End If
ElseIf grid_stk_dtl.Col = 2 Or grid_stk_dtl.Col = 3 Or grid_stk_dtl.Col = 5 Or grid_stk_dtl.Col = 6 Then
    If grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 1) = "" Then
    grid_stk_dtl.Col = 1
    If List_card.ListIndex < 0 Then List_card.ListIndex = 0
        Call select_card
        Call set_card_detail
    grid_stk_dtl.Col = 2
    End If
            
        Text1.Visible = True
        Call next_text_cell
        Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
ElseIf grid_stk_dtl.Col = 10 Then
        Call next_cmb_inv_vat_type
        cmb_inv_vat_type.Text = cmb_vat_type.Text
End If
End Sub
Public Sub next_cardlist_cell()
    List_card.Left = grid_stk_dtl.CellLeft + grid_stk_dtl.Left
    List_card.Top = grid_stk_dtl.CellTop + grid_stk_dtl.Top
    List_card.Width = grid_stk_dtl.CellWidth
    cmb_inv_vat_type.Visible = False
    Text1.Visible = False
    List_card.Visible = True
    List_card.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
    List_card.SetFocus
End Sub
Public Sub next_text_cell()
    List_card.Visible = False
    cmb_inv_vat_type.Visible = False
    Text1.Left = grid_stk_dtl.CellLeft + grid_stk_dtl.Left
    Text1.Top = grid_stk_dtl.CellTop + grid_stk_dtl.Top
    Text1.Width = grid_stk_dtl.CellWidth
    Text1.Visible = True
    Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
    Text1.SetFocus
End Sub
Public Sub next_cmb_inv_vat_type()
    'cmb_inv_vat_type.Text = "1"
    'Text1.Visible = False
    'List_card.Visible = False
    'cmb_inv_vat_type.Left = grid_stk_dtl.CellLeft + grid_stk_dtl.Left
    'cmb_inv_vat_type.Top = grid_stk_dtl.CellTop + grid_stk_dtl.Top
    'cmb_inv_vat_type.Width = grid_stk_dtl.CellWidth
    'cmb_inv_vat_type.Visible = True
    'cmb_inv_vat_type.SetFocus
    'cmb_inv_vat_type.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 10)
End Sub

Private Sub text_ledger_GotFocus()
If show_ledger_detail <> 1 Then
    cmb_ledger.Height = 2400
    cmb_ledger.Visible = True
    cmb_ledger.SetFocus
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim card_is_available As Integer
keycode_now = KeyCode
If keycode_now = 37 Then
    If grid_stk_dtl.Col = 1 Then
        MsgBox "Not Valid key.....!!!"
        Exit Sub
    ElseIf grid_stk_dtl.Col = 2 Then
        grid_stk_dtl.Col = grid_stk_dtl.Col - 1
        Text1.Visible = False
        Call next_cardlist_cell
        List_card.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
    ElseIf grid_stk_dtl.Col = 3 Then
        grid_stk_dtl.Col = grid_stk_dtl.Col - 1
        Call next_text_cell
        Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
    ElseIf grid_stk_dtl.Col = 5 Then
        grid_stk_dtl.Col = grid_stk_dtl.Col - 2
        Call next_text_cell
        Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
    ElseIf grid_stk_dtl.Col = 6 Then
        grid_stk_dtl.Col = grid_stk_dtl.Col - 1
        Call next_text_cell
        Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
    End If
ElseIf keycode_now = 38 Then 'And grid_stk_dtl.Row >= 1 Then
    If grid_stk_dtl.Row = 1 Then
        MsgBox "Not Valid key.....!!!"
        Exit Sub
    End If
    grid_stk_dtl.Row = grid_stk_dtl.Row - 1
    Call next_text_cell
    Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
ElseIf keycode_now = 39 Or keycode_now = 13 Then
            If grid_stk_dtl.Col = 6 Then    'enter on the amount column
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = Format(Text1.Text, "0.00")
                If Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)) <= 0 Then
                    MsgBox "Not Valid value.....!!! Please Enter Again"
                    Exit Sub
                End If
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 5) = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 6)) / Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 4)), "0.00")
                grid_stk_dtl.Col = 10
                'Call next_cmb_inv_vat_type
                
                
                If grid_stk_dtl.Rows = grid_stk_dtl.Row + 1 Then
                    grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
                End If
                grid_stk_dtl.Row = grid_stk_dtl.Row + 1
                grid_stk_dtl.Col = 1
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 0) = grid_stk_dtl.Row
                
                List_card.Visible = True
                Call next_cardlist_cell
                Call stock_refresh_total
                Exit Sub
                
                'Text1.Visible = False
                'Call next_cardlist_cell
                'Call stock_refresh_total
                'Exit Sub
            ElseIf grid_stk_dtl.Col = 5 Then    'enter on the Rate column
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = Format(Text1.Text, "0.00")
                If Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)) <= 0 Then
                    MsgBox "Not Valid value.....!!! Please Enter Again"
                    Exit Sub
                End If
                grid_stk_dtl.Col = grid_stk_dtl.Col + 1
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col - 1)) * Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col - 2)), "0.00")
                Text1.Text = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col - 1)) * Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col - 2)), "0.00")
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col - 1)) * Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col - 2)), "0.00")
                Call next_text_cell
                Call stock_refresh_total
            ElseIf grid_stk_dtl.Col = 3 Then    'enter on the ending serial no column
                
                card_is_available = 0
                        'Call open_database
                        'Call open_rs_stk_clsg_srl
                        Call open_rs_tmp_spec_itm_clg_stk
                        Do Until rs_tmp_spec_itm_clg_stk.EOF
                            Dim temp_card_srl_no
                            For temp_card_srl_no = Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 2)) To Val(Text1.Text)
                            If Val(Text1.Text) >= Val(rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_stno) And Val(rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_edno) >= Val(Text1.Text) Then
                                card_is_available = 1
                                Exit Do
                                'Exit Sub
                            End If
                            Next
                            rs_tmp_spec_itm_clg_stk.MoveNext
                        Loop
                        If card_is_available = 1 Then
                            MsgBox "card is available in the stock you can not input again...!!!"
                            Exit Sub
                        End If
                If Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 2)) > Val(Text1.Text) Then
                MsgBox "Ending serial number should be greater then starting serial number. Try Again..,"
                Exit Sub
                End If

                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = Text1.Text
                
                If Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)) <= 0 Then
                    MsgBox "Not Valid value.....!!! Please Enter Again"
                    Exit Sub
                End If
                grid_stk_dtl.Col = grid_stk_dtl.Col + 1
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col - 1)) - Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col - 2)) + 1
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 6) = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 4)) * Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 5)), "0.00")
                grid_stk_dtl.Col = grid_stk_dtl.Col + 1
                Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
                'Call search_this_stock_serial_no_is_avilable
                Call next_text_cell
                Call stock_refresh_total
                
            ElseIf grid_stk_dtl.Col = 2 Then    'enter on the starting serial no column

                If Val(Text1.Text) <= 0 Then
                    MsgBox "Not Valid value.....!!! Please Enter Again"
                    Exit Sub
                End If
                
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = Text1.Text
                
                'search the serial no of supplier
                'grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = ""
                'Call open_database
                'Call open_rs_inv_tran_prs
                'rs_inv_tran_prs!stk_invt_trn_splr = ""
                card_is_available = 0
                        'Call open_database
                        'Call open_rs_stk_clsg_srl
                        Call open_rs_tmp_spec_itm_clg_stk
                        Do Until rs_tmp_spec_itm_clg_stk.EOF
                        
                            If Val(Text1.Text) >= Val(rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_stno) And Val(rs_tmp_spec_itm_clg_stk!stk_xinvt_clg_edno) >= Val(Text1.Text) Then
                                card_is_available = 1
                
                                'Exit Sub
                            End If
                            rs_tmp_spec_itm_clg_stk.MoveNext
                        Loop
                        
                        
                        
                        If card_is_available = 1 Then
                            MsgBox "card is available in the stock you can not input again...!!!"
                            Exit Sub
                        End If
                
                
                
                
                grid_stk_dtl.Col = grid_stk_dtl.Col + 1
                Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
                grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 6) = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 4)) * Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 5)), "0.00")
                If Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 3)) >= 0 Then
                    grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 4) = Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 3)) - Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 2)) + 1
                    If Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 5)) >= 0 Then grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 6) = Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 4)) * Val(grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 5))
                End If
                
                'Call search_this_stock_serial_no_is_avilable
                Call next_text_cell
            End If
            
ElseIf keycode_now = 40 Then
    If grid_stk_dtl.Rows = grid_stk_dtl.Row + 1 Then
        MsgBox "Not Valid key.....!!!"
        Exit Sub
    Else
        grid_stk_dtl.Row = grid_stk_dtl.Row + 1
        Call next_text_cell
        Text1.Text = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col)
    End If
End If
keycode_now = 0
End Sub
Public Sub stock_refresh_total()
Dim xyz_stock_row
total_qnty = 0
total_amnt = 0
For I = 1 To grid_stk_dtl.Rows - 1
total_qnty = total_qnty + Val(grid_stk_dtl.TextMatrix(I, 4))
total_amnt = total_amnt + Val(grid_stk_dtl.TextMatrix(I, 6))
Next
lbl_total_qty.Caption = total_qnty
lbl_total_amt.Caption = Format(total_amnt, "0.00")
grid_stk_dtl.Font.Size = 6
End Sub
'==================form detail set====================
Private Sub Form_Load()

DTPicker1.TabIndex = 1
text_ledger.TabIndex = 2
cmb_ledger.TabIndex = 3
List1.TabIndex = 4
Text1.TabIndex = 5
Text2.TabIndex = 6
Text3.TabIndex = 7
Text4.TabIndex = 8
cmb_sales_man.TabIndex = 9
cmb_vat_type.TabIndex = 10

cmd_sv_n_new.TabIndex = 11
cmd_save_and_exit.TabIndex = 12
cmd_exit.TabIndex = 13
Command1.TabIndex = 14
cmd_delete.TabIndex = 15
cmb_edit.TabIndex = 16
Command3.TabIndex = 17

Text5.Enabled = False
Text6.Enabled = False

lbl_l_name = "NAME"
lbl_l_add = "ADDRESS"
lbl_tel_no = "TELEPHONE"
lbl_mob_no = "MOBILE"
lbl_inv_no = "INVOICE NO."
lbl_l_trans = "TRAVEL DET."
lbl_sales_man = "SALES MAN"
lbl_vat_type = "VAT TYPE"
lbl_user = "USER"

selected_procedure = "purchase Invoice"
lbl_head.Caption = selected_procedure
Call set_form_detail_zero
Call set_grid_stk_dtl_data
Call find_last_voucher_no
    grid_stk_dtl.Col = 4
    lbl_total_qty.Left = grid_stk_dtl.CellLeft + grid_stk_dtl.Left
    Line1.X1 = grid_stk_dtl.CellLeft + grid_stk_dtl.Left
    Line1.X2 = grid_stk_dtl.CellLeft + grid_stk_dtl.Left
    lbl_total_qty.FontSize = 12
    lbl_total_qty.Caption = "Qty"
    lbl_total_qty.Width = grid_stk_dtl.CellWidth
    grid_stk_dtl.Col = 6
    lbl_total_amt.Left = grid_stk_dtl.CellLeft + grid_stk_dtl.Left
    Line2.X1 = grid_stk_dtl.CellLeft + grid_stk_dtl.Left
    Line2.X2 = grid_stk_dtl.CellLeft + grid_stk_dtl.Left
    lbl_total_amt.FontSize = 12
    lbl_total_amt.Caption = "Amt"
    lbl_total_amt.Width = grid_stk_dtl.CellWidth
    Line3.X1 = grid_stk_dtl.CellLeft + grid_stk_dtl.Left + lbl_total_amt.Width
    Line3.X2 = grid_stk_dtl.CellLeft + grid_stk_dtl.Left + lbl_total_amt.Width
    lbl_customer_marq.Caption = "You have not selected any ledger...!!!"
If show_ledger_detail = 1 Then
    cmb_ledger.Text = selected_voucher_ledger
    Call open_selected_voucher
End If
'grid_stk_dtl.Enabled = False
End Sub
Public Sub open_selected_voucher()
Dim rs_acnt_tran_prs_counter
Call open_database
Call open_rs_acn_tran_prs
For rs_acnt_tran_prs_counter = 1 To rs_acn_tran_prs.RecordCount
If rs_acn_tran_prs!fin_acnt_trn_vcno = selected_voucher_no And rs_acn_tran_prs!fin_acnt_trn_seno = 2 Then
    Text5.Text = rs_acn_tran_prs!fin_acnt_trn_vcno
    'rs_acn_tran_prs!fin_acnt_trn_seno = 1 'transaction_counter
    cmb_vat_type.Text = rs_acn_tran_prs!fin_acnt_trn_vtyp 'Combo0.Text
    DTPicker1.Value = rs_acn_tran_prs!fin_acnt_trn_date
    lbl_time.Caption = rs_acn_tran_prs!fin_acnt_trn_time
    lbl_day.Caption = rs_acn_tran_prs!fin_acnt_trn_wday
    lbl_total_amt.Caption = rs_acn_tran_prs!fin_acnt_trn_amnt ' text_amt(transaction_counter).Text
    Text6.Text = rs_acn_tran_prs!fin_acnt_trn_user
    cmb_ledger.Text = rs_acn_tran_prs!fin_acnt_trn_ldgr ' combo_lgr(transaction_counter).Text
    selected_voucher_ledger = cmb_ledger.Text
    text_ledger.Text = cmb_ledger.Text
End If
rs_acn_tran_prs.MoveNext
Next
Call set_selected_ledger_detail
Dim inventory_sub_entry_no
Call open_database

If rs_inv_tran_prs.State = 1 Then rs_inv_tran_prs.Close
rs_inv_tran_prs.CursorLocation = adUseClient
rs_inv_tran_prs.Open "Select * From [inv_tran_prs] order by stk_invt_trn_id", db_co, adOpenDynamic, adLockPessimistic

For rs_inv_tran_prs_counter = 1 To rs_inv_tran_prs.RecordCount 'inventory_sub_entry_no = 1 To grid_stk_dtl.Rows - 1
If rs_inv_tran_prs!stk_invt_trn_vcno = selected_voucher_no Then
    With rs_inv_tran_prs
        grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
        inventory_sub_entry_no = !stk_invt_trn_seno
        grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 0) = inventory_sub_entry_no
        grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 1) = !stk_invt_trn_card
        grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 2) = !stk_invt_trn_stno
        grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 3) = !stk_invt_trn_edno
        grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 4) = !stk_invt_trn_qnty
        grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 5) = Format(!stk_invt_trn_rate, "0.00")
        grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 6) = Format(!stk_invt_trn_amnt, "0.00")
        grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 7) = Format(!stk_invt_trn_fval, "0.00")
        grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 8) = !stk_invt_trn_comp
        grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 10) = !stk_invt_trn_vtyp
        grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 11) = !stk_invt_trn_splr
'       !stk_invt_trn_vcno = Text5.Text
'       !stk_invt_trn_date = DTPicker1.Value
'       !stk_invt_trn_time = lbl_time.Caption
'       !stk_invt_trn_wday = lbl_day.Caption
'       !stk_invt_trn_vtyp = cmb_vat_type
'       !stk_invt_trn_ldgr = cmb_ledger.Text
'       !stk_invt_trn_splr = grid_stk_dtl.TextMatrix(inventory_sub_entry_no, 8)
'       Text6.Text = !stk_invt_trn_user
'       !stk_invt_trn_vchr = "sale"
'       !stk_invt_trn_slmn = cmb_sales_man.Text
'        .UpdateBatch
    End With
End If
rs_inv_tran_prs.MoveNext
Next
'disable all controls of the form
Call disable_all_controls
End Sub
Public Sub disable_all_controls()
text_ledger.Enabled = False
List1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
grid_stk_dtl.Enabled = False
cmb_ledger.Enabled = False
'cmb_sales_man.Enabled = False
DTPicker1.Enabled = False
cmb_vat_type.Enabled = False
cmb_inv_vat_type.Visible = False
Text1.Visible = False
List_card.Visible = False
cmb_edit.Enabled = True
End Sub
Public Sub set_form_detail_zero()
Call add_ledgers
Call add_cards
cmb_inv_vat_type.AddItem "1"
cmb_inv_vat_type.AddItem "2"
Call set_headings
Call set_labels_and_texts
End Sub
Public Sub set_headings()
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
Public Sub set_labels_and_texts()
List1.Font.Size = 12
Text2.Font.Size = 12
Text3.Font.Size = 12
Text4.Font.Size = 12
Text5.Font.Size = 12
Text6.Font.Size = 12
List1.Clear
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = selected_user
cmb_ledger.Font.Size = 12
cmb_sales_man.Font.Size = 12
cmb_vat_type.Font.Size = 12
cmb_vat_type.AddItem "1"
cmb_vat_type.AddItem "2"
cmb_vat_type.Text = "2"
DTPicker1.Font.Size = 12
DTPicker1.Value = Date
lbl_time.Caption = Time
lbl_day.Caption = WeekdayName(Weekday(DTPicker1.Value - 1))
cmb_edit.Enabled = False
'Call add_sales_mans
cmb_sales_man.Enabled = False
End Sub
Public Sub add_sales_mans()
Call open_database
Call open_rs_emp_main_dtl
Do Until rs_emp_main_dtl.EOF
    cmb_sales_man.AddItem rs_emp_main_dtl!emp_main_dtl_name
    rs_emp_main_dtl.MoveNext
Loop
cmb_sales_man.Text = "Sales under"
End Sub
Public Sub add_cards()
Call open_database
Call open_rs_stk_item_lgr
Do Until rs_stk_item_lgr.EOF
    List_card.AddItem rs_stk_item_lgr!stk_item_lgr_name
    If rs_stk_item_lgr!stk_item_lgr_alis <> "" Then List_card.AddItem rs_stk_item_lgr!stk_item_lgr_alis
rs_stk_item_lgr.MoveNext
Loop
'Call SortList(List_card, Val(0) \ 1, (Val(List_card.ListCount) - 1) \ 1, Ascending)
End Sub
Public Sub add_ledgers()
cmb_ledger.Clear
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
        'MsgBox rs_lgr_main_dtl!lgr_main_dtl_name & " group is...." & selected_primary_group
        If LCase(selected_primary_group) = LCase("Sundry Creditors") Or LCase(selected_primary_group) = LCase("Sundry Debtors") Then ' if the created ledger is a debtor then
            cmb_ledger.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
            If rs_lgr_main_dtl!lgr_main_dtl_alis <> "" Then cmb_ledger.AddItem rs_lgr_main_dtl!lgr_main_dtl_alis
        End If
rs_lgr_main_dtl.MoveNext
Loop
cmb_ledger.Text = "Select Customer..,"
End Sub
'============================ form control working reflactions ==============================
Public Sub go_to_card_detail()
    grid_stk_dtl.Row = 1
    grid_stk_dtl.Col = 1
    List_card.Visible = True
    Call next_cardlist_cell
    grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, grid_stk_dtl.Col) = List_card.Text
End Sub
Private Sub cmb_ledger_Click()
'    selected_voucher_ledger = cmb_ledger.Text
'    Call set_selected_ledger_detail
'    If grid_stk_dtl.Enabled = False Then grid_stk_dtl.Enabled = True
'    Call go_to_card_detail
'    text_ledger.Text = selected_voucher_ledger
'    cmb_ledger.Visible = False
End Sub
Public Sub set_selected_ledger_detail()

Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
        If rs_lgr_main_dtl!lgr_main_dtl_name = selected_voucher_ledger Then  'if the created ledger is a debtor then
        List1.Clear
        List1.AddItem rs_lgr_main_dtl!lgr_main_dtl_add1
        List1.AddItem rs_lgr_main_dtl!lgr_main_dtl_add2
        List1.AddItem rs_lgr_main_dtl!lgr_main_dtl_city & " - " & UCase(rs_lgr_main_dtl!lgr_main_dtl_pncd)
        Text2.Text = rs_lgr_main_dtl!lgr_main_dtl_tel1 & " - " & rs_lgr_main_dtl!lgr_main_dtl_tel1
        Text3.Text = rs_lgr_main_dtl!lgr_main_dtl_mobl
        Text4.Text = rs_lgr_main_dtl!lgr_main_dtl_trnp
        'cmb_sales_man.Text = rs_lgr_main_dtl!lgr_main_dtl_slun
        End If
        rs_lgr_main_dtl.MoveNext
Loop
Call set_ledger_account_detail
End Sub
Public Sub set_ledger_grid_zero()
    grid_lgr_rep.Clear
    grid_lgr_rep.Rows = 1
    grid_lgr_rep.Cols = 11
    grid_lgr_rep.Font.Size = 6
    b = 0
    grid_lgr_rep.TextMatrix(b, 0) = ""
    grid_lgr_rep.TextMatrix(b, 1) = "Date"
    grid_lgr_rep.TextMatrix(b, 2) = "Voucher"
    grid_lgr_rep.TextMatrix(b, 3) = "V.No"
    grid_lgr_rep.TextMatrix(b, 4) = "Ledger"
    grid_lgr_rep.TextMatrix(b, 5) = "Dr.Amount"
    grid_lgr_rep.TextMatrix(b, 6) = "Cr.Amount"
    grid_lgr_rep.TextMatrix(b, 7) = "Balance"
    grid_lgr_rep.TextMatrix(b, 8) = "Time"
    grid_lgr_rep.TextMatrix(b, 9) = "User"
    grid_lgr_rep.ColWidth(0) = 100
    grid_lgr_rep.ColWidth(1) = 1000
    grid_lgr_rep.ColWidth(2) = 700
    grid_lgr_rep.ColWidth(3) = 600
    grid_lgr_rep.ColWidth(4) = 1700
    grid_lgr_rep.ColWidth(5) = 950
    grid_lgr_rep.ColWidth(6) = 950
    grid_lgr_rep.ColWidth(7) = 1050
    grid_lgr_rep.ColWidth(8) = 500
    grid_lgr_rep.ColWidth(9) = 500
    ledger_cr_total = 0
    ledger_dr_total = 0
    Dim x_int
    b = 1
End Sub
Public Sub set_ledger_account_detail()
selected_ledger = selected_voucher_ledger
Call set_ledger_grid_zero
Call open_database

Call copy_specific_ledger_transaction_to_acn_tran_spc_lgr
Call open_rs_acn_tran_spc_lgr
rs_acn_tran_spc_lgr.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno" 'time"
Dim temp_starting_no As Integer
Dim temp_ending_no As Integer
Dim b
b = 1
If rs_acn_tran_spc_lgr.RecordCount < 7 Then
    temp_starting_no = 0
    temp_ending_no = rs_acn_tran_spc_lgr.RecordCount
Else
    temp_starting_no = rs_acn_tran_spc_lgr.RecordCount - 6
    temp_ending_no = rs_acn_tran_spc_lgr.RecordCount
End If

For x_int = 1 To temp_ending_no

If x_int >= temp_starting_no + 1 Then
    
    grid_lgr_rep.AddItem ""
    'grid_lgr_rep.TextMatrix(b, 0) = b
    grid_lgr_rep.TextMatrix(b, 1) = rs_acn_tran_spc_lgr!fin_acnt_trn_date
    grid_lgr_rep.TextMatrix(b, 2) = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_vchr)
    grid_lgr_rep.TextMatrix(b, 3) = rs_acn_tran_spc_lgr!fin_acnt_trn_vcno
    grid_lgr_rep.TextMatrix(b, 4) = LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_ldgr)
    
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") Then
        grid_lgr_rep.TextMatrix(b, 5) = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
        ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") Then
        grid_lgr_rep.TextMatrix(b, 6) = Format(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt, "0.00")
        ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    If ledger_dr_total > ledger_cr_total Then
        grid_lgr_rep.TextMatrix(b, 7) = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
    End If
    If ledger_cr_total > ledger_dr_total Then
        grid_lgr_rep.TextMatrix(b, 7) = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
    End If
    If rs_acn_tran_spc_lgr!fin_acnt_trn_time <> "" Then
        grid_lgr_rep.TextMatrix(b, 8) = rs_acn_tran_spc_lgr!fin_acnt_trn_time
    End If
    If rs_acn_tran_spc_lgr!fin_acnt_trn_user <> "" Then
    grid_lgr_rep.TextMatrix(b, 9) = rs_acn_tran_spc_lgr!fin_acnt_trn_user
    End If
    
    b = b + 1
Else
b = 1
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") Then
        ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") Then
        ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    
    grid_lgr_rep.AddItem ""
    If ledger_dr_total > ledger_cr_total Then
        temp_opening_balance = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        grid_lgr_rep.TextMatrix(b, 4) = "Opeining Balance is ..., " 'Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        grid_lgr_rep.TextMatrix(b, 5) = Format(ledger_dr_total - ledger_cr_total, "0.00") '& " Dr."
        grid_lgr_rep.TextMatrix(b, 6) = ""
        grid_lgr_rep.TextMatrix(b, 7) = Format(ledger_dr_total - ledger_cr_total, "0.00") & " Dr."
        
    ElseIf ledger_cr_total > ledger_dr_total Then
        temp_opening_balance = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
        grid_lgr_rep.TextMatrix(b, 4) = "Opeining Balance is ..., "
        grid_lgr_rep.TextMatrix(b, 5) = ""
        grid_lgr_rep.TextMatrix(b, 6) = Format(ledger_cr_total - ledger_dr_total, "0.00") '& " Cr."
        grid_lgr_rep.TextMatrix(b, 7) = Format(ledger_cr_total - ledger_dr_total, "0.00") & " Cr."
    End If
    b = b + 1
    
End If
rs_acn_tran_spc_lgr.MoveNext
Next

If b >= 2 Then
ledger_dr_total = 0
ledger_cr_total = 0
Dim grid_row_counter
For grid_row_counter = 1 To (b - 1)
If grid_lgr_rep.TextMatrix(grid_row_counter, 5) <> "" Then ledger_dr_total = ledger_dr_total + Val(grid_lgr_rep.TextMatrix(grid_row_counter, 5))
If grid_lgr_rep.TextMatrix(grid_row_counter, 6) <> "" Then ledger_cr_total = ledger_cr_total + Val(grid_lgr_rep.TextMatrix(grid_row_counter, 6))
Next

    If ledger_dr_total < ledger_cr_total Then
    
        lbl_customer_marq.Caption = selected_voucher_ledger & " Closing Balance is..... Dr. " & Format(ledger_cr_total - ledger_dr_total, "0.00")
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 4) = "Closing Balance is.....Dr."
        grid_lgr_rep.TextMatrix(b, 5) = Format(ledger_cr_total - ledger_dr_total, "0.00")
        b = b + 1
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 5) = "================="
        grid_lgr_rep.TextMatrix(b, 6) = "================="
        b = b + 1
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 5) = Format(ledger_cr_total, "0.00")
        grid_lgr_rep.TextMatrix(b, 6) = Format(ledger_cr_total, "0.00")
        b = b + 1
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 5) = "================="
        grid_lgr_rep.TextMatrix(b, 6) = "================="

    End If
    'grid_lgr_rep.TextMatrix(b, 6) = Format(ledger_cr_total, "0.00")
    If ledger_cr_total < ledger_dr_total Then
        lbl_customer_marq.Caption = selected_voucher_ledger & " Closing Balance is..... Cr. " & Format(ledger_dr_total - ledger_cr_total, "0.00")
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 4) = " Closing Balance is.....Cr. "
        grid_lgr_rep.TextMatrix(b, 6) = Format(ledger_dr_total - ledger_cr_total, "0.00")
        b = b + 1
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 5) = "=========="
        grid_lgr_rep.TextMatrix(b, 6) = "=========="
        b = b + 1
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 4) = "TOTAL AMOUNT..."
        grid_lgr_rep.TextMatrix(b, 5) = Format(ledger_dr_total, "0.00")
        grid_lgr_rep.TextMatrix(b, 6) = Format(ledger_dr_total, "0.00")
        b = b + 1
        grid_lgr_rep.AddItem ""
        grid_lgr_rep.TextMatrix(b, 5) = "=========="
        grid_lgr_rep.TextMatrix(b, 6) = "=========="
    End If
End If
Dim temp_entry_no
Dim temp_v_no
Dim temp_v_tp
For temp_entry_no = 1 To grid_lgr_rep.Rows - 3
If grid_lgr_rep.TextMatrix(temp_entry_no, 2) <> "" Then
    temp_v_tp = grid_lgr_rep.TextMatrix(temp_entry_no, 2)
    temp_v_no = Val(grid_lgr_rep.TextMatrix(temp_entry_no, 3))
    Call open_rs_acn_tran_all
    If rs_acn_tran_all.RecordCount > 0 Then rs_acn_tran_all.MoveFirst
    Do Until rs_acn_tran_all.EOF
            If LCase(selected_ledger) <> LCase(rs_acn_tran_all!fin_acnt_trn_ldgr) And LCase(temp_v_tp) = LCase(rs_acn_tran_all!fin_acnt_trn_vchr) And temp_v_no = Val(rs_acn_tran_all!fin_acnt_trn_vcno) Then
                grid_lgr_rep.TextMatrix(temp_entry_no, 4) = LCase(rs_acn_tran_all!fin_acnt_trn_ldgr)
                Exit Do
            End If
        
        rs_acn_tran_all.MoveNext
    Loop
End If
Next

End Sub
Public Sub set_grid_stk_dtl_data()
    grid_stk_dtl.RowHeightMin = 400
    grid_stk_dtl.Clear
    grid_stk_dtl.Rows = 2
    grid_stk_dtl.Cols = 12
    
    grid_stk_dtl.TextMatrix(0, 0) = "No."
    grid_stk_dtl.TextMatrix(0, 1) = "card"
    grid_stk_dtl.TextMatrix(0, 2) = "Start-No."
    grid_stk_dtl.TextMatrix(0, 3) = "End-No."
    grid_stk_dtl.TextMatrix(0, 4) = "Qty."
    grid_stk_dtl.TextMatrix(0, 5) = "Rate"
    grid_stk_dtl.TextMatrix(0, 6) = "Amount"
    grid_stk_dtl.TextMatrix(0, 7) = "F.Val"
    grid_stk_dtl.TextMatrix(0, 8) = "Dis(%)"
    grid_stk_dtl.TextMatrix(0, 9) = "Co."
    'grid_stk_dtl.TextMatrix(0, 10) = "VAT"
    grid_stk_dtl.TextMatrix(0, 10) = "Sup."
    
    grid_stk_dtl.ColWidth(0) = 300
    grid_stk_dtl.ColWidth(1) = 1500
    grid_stk_dtl.ColWidth(2) = 1800
    grid_stk_dtl.ColWidth(3) = 1800
    grid_stk_dtl.ColWidth(4) = 800
    grid_stk_dtl.ColWidth(5) = 800
    grid_stk_dtl.ColWidth(6) = 1000
    grid_stk_dtl.ColWidth(7) = 700
    grid_stk_dtl.ColWidth(8) = 600
    grid_stk_dtl.ColWidth(9) = 700
    grid_stk_dtl.ColWidth(10) = 700
    'grid_stk_dtl.ColWidth(11) = 800
    
    cmb_inv_vat_type.Visible = False
    Text1.Visible = False
    List_card.Visible = False
    
    grid_stk_dtl.Row = 0
    grid_stk_dtl.TextMatrix(1, 0) = grid_stk_dtl.Row + 1
    Call open_and_set_selected_voucher_data
End Sub
Public Sub open_and_set_selected_voucher_data()

End Sub

Private Sub grid_lgr_rep_Click()
Dim jjj
For jjj = 1 To grid_lgr_rep.Cols
'grid_lgr_rep.
'grid_lgr_rep.CellBackColor = RGB(100, 100, 100)
'grid_lgr_rep.CellBackColor(grid_lgr_rep.Row, jjj) = RGB(100, 100, 100)
Next
'grid_lgr_rep.CellBackColor
'grid_lgr_rep.BackColor
End Sub
Private Sub Timer1_Timer()
If lbl_customer_marq.Left + lbl_customer_marq.Width <= 0 Then lbl_customer_marq.Left = 13500 + lbl_customer_marq.Width
lbl_customer_marq.Left = lbl_customer_marq.Left - 500
End Sub

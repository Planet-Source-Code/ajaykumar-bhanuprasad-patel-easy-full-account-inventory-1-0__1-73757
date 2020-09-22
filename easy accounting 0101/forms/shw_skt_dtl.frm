VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form shw_skt_dtl 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "shw_skt_dtl.frx":0000
   LinkTopic       =   "Stock Detail"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   3201
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.ListBox List_card 
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
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   4455
   End
   Begin MSFlexGridLib.MSFlexGrid grid_stk_dtl 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6588
      _Version        =   393216
   End
End
Attribute VB_Name = "shw_skt_dtl"
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

Call set_stock_summary_grid
Call refresh_closing_stock
Call open_database
Call open_rs_stk_item_lgr
Do Until rs_stk_item_lgr.EOF
    List_card.AddItem rs_stk_item_lgr!stk_item_lgr_name
    If rs_stk_item_lgr!stk_item_lgr_alis <> "" Then List_card.AddItem rs_stk_item_lgr!stk_item_lgr_alis
rs_stk_item_lgr.MoveNext
Loop
End Sub
Public Sub set_stock_summary_grid()
    grid_stk_dtl.RowHeightMin = 400
    grid_stk_dtl.Clear
    grid_stk_dtl.Rows = 2
    grid_stk_dtl.Cols = 14
    
    grid_stk_dtl.TextMatrix(0, 0) = "No."
    grid_stk_dtl.TextMatrix(0, 1) = "Date"
    grid_stk_dtl.TextMatrix(0, 2) = "V.No"
    grid_stk_dtl.TextMatrix(0, 3) = "Tran."
    grid_stk_dtl.TextMatrix(0, 4) = "ledger"
    grid_stk_dtl.TextMatrix(0, 5) = "Starting Seril No."
    grid_stk_dtl.TextMatrix(0, 6) = "Ending Seril No."
    grid_stk_dtl.TextMatrix(0, 7) = "Inward Qty"
    grid_stk_dtl.TextMatrix(0, 8) = "Outward Qty"
    grid_stk_dtl.TextMatrix(0, 9) = "Balance Qty"
    grid_stk_dtl.TextMatrix(0, 10) = "F.Val"
    'grid_stk_dtl.TextMatrix(0, 11) = "Dis.Rt"
    grid_stk_dtl.TextMatrix(0, 11) = "Company name"
    grid_stk_dtl.TextMatrix(0, 12) = "VAT"
    grid_stk_dtl.TextMatrix(0, 13) = "Suplier"
    grid_stk_dtl.ColWidth(0) = 500
    grid_stk_dtl.ColWidth(1) = 1500
    grid_stk_dtl.ColWidth(2) = 800
    grid_stk_dtl.ColWidth(3) = 800
    grid_stk_dtl.ColWidth(4) = 2500
    grid_stk_dtl.ColWidth(5) = 2500
    grid_stk_dtl.ColWidth(6) = 2500
    grid_stk_dtl.ColWidth(7) = 1000
    grid_stk_dtl.ColWidth(8) = 1000
    grid_stk_dtl.ColWidth(9) = 1000
    grid_stk_dtl.ColWidth(10) = 800
    
    grid_stk_dtl.ColWidth(11) = 2000
    grid_stk_dtl.ColWidth(12) = 800
    grid_stk_dtl.ColWidth(13) = 2000
End Sub

Private Sub Text1_GotFocus()
    List_card.Visible = True
    List_card.Height = 2400
    List_card.SetFocus

End Sub

Public Sub enter_the_card_from_list()
selected_stock_item_name = Text1.Text

Call open_database
Call open_rs_stk_item_lgr
Do Until rs_stk_item_lgr.EOF
    If rs_stk_item_lgr!stk_item_lgr_alis = Text1.Text Then
    Text1.Text = rs_stk_item_lgr!stk_item_lgr_name
    selected_stock_item_name = Text1.Text
    Exit Do
    End If
    rs_stk_item_lgr.MoveNext
Loop

Call set_stock_summary_grid
Call open_database
Call open_rs_inv_tran_all
'If rs_inv_tran_all.State = 1 Then rs_inv_tran_all.Close
'rs_inv_tran_all.CursorLocation = adUseClient
'rs_inv_tran_all.Open "Select * From inv_tran_all", db_co, adOpenDynamic, adLockPessimistic
'rs_inv_tran_all.Open "SELECT * FROM inv_tran_all WHERE 'stk_invt_trn_card = " & selected_stock_item_name & "'", db_co, adOpenDynamic, adLockOptimistic
Dim rs_inv_tran_all_counter
Dim grid_stk_row_no
grid_stk_row_no = 1
Dim total_inward
Dim total_outward
Dim temp_stock_balance
For rs_inv_tran_all_counter = 1 To rs_inv_tran_all.RecordCount
If LCase(rs_inv_tran_all!stk_invt_trn_card) = LCase(selected_stock_item_name) Then

With rs_inv_tran_all
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 0) = grid_stk_row_no
        If !stk_invt_trn_date <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 1) = !stk_invt_trn_date
        If !stk_invt_trn_vcno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = !stk_invt_trn_vcno
        If !stk_invt_trn_vchr <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 3) = !stk_invt_trn_vchr
        If !stk_invt_trn_ldgr <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = !stk_invt_trn_ldgr
        If !stk_invt_trn_stno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 5) = !stk_invt_trn_stno
        If !stk_invt_trn_edno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 6) = !stk_invt_trn_edno
        If !stk_invt_trn_vchr = "purchase" Or !stk_invt_trn_vchr = "sale return" Or !stk_invt_trn_vchr = "opening stock" Then
            If !stk_invt_trn_fval <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = !stk_invt_trn_qnty
        total_inward = total_inward + Val(!stk_invt_trn_qnty)
        End If
        If !stk_invt_trn_vchr = "sale" Or !stk_invt_trn_vchr = "purchase return" Then
            If !stk_invt_trn_comp <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = !stk_invt_trn_qnty
            total_outward = total_outward + Val(!stk_invt_trn_qnty)
        End If
        temp_stock_balance = total_inward - total_outward
        If !stk_invt_trn_vtyp <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = temp_stock_balance
        If !stk_invt_trn_fval <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 10) = Format(!stk_invt_trn_fval, "0.00")
        'grid_stk_dtl.TextMatrix(grid_stk_row_no, 12) = !stk_invt_trn_splr
        If !stk_invt_trn_comp <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 11) = !stk_invt_trn_comp
        If !stk_invt_trn_vtyp <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 12) = !stk_invt_trn_vtyp
        If !stk_invt_trn_splr <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 13) = !stk_invt_trn_splr
            grid_stk_row_no = grid_stk_row_no + 1
            grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
End With
End If
rs_inv_tran_all.MoveNext
Next
End Sub
Private Sub List_card_Click()
'list_card.Text =
End Sub
Private Sub List_card_GotFocus()
List_card.Height = 2400

End Sub
Private Sub List_card_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text1.Text = List_card.Text
    List_card.Visible = False
    Call enter_the_card_from_list
    Text2.SetFocus
End If
End Sub
Private Sub List_card_LostFocus()
List_card.Height = 600
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form shw_clg_stk 
   Caption         =   "closing stock"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13095
   ScaleWidth      =   21480
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   -500
      Width           =   2655
   End
   Begin VB.ListBox List_card 
      Height          =   3375
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid grid_stk_dtl 
      Height          =   9615
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   16960
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Closing Stock"
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
      Left            =   10560
      TabIndex        =   3
      Top             =   1080
      Width           =   2955
   End
End
Attribute VB_Name = "shw_clg_stk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List_card_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text1.Text = List_card.Text
    List_card.Visible = False
    Call enter_the_card_from_list
    Text2.SetFocus
End If
Label1.Caption = "Closing Stock as on the " & Date & " of " & selected_stock_item_name
End Sub
Private Sub Text1_GotFocus()
    List_card.Visible = True
    List_card.Height = 2400
    List_card.SetFocus
End Sub
Private Sub Form_Load()
Call open_database
Call open_rs_stk_item_lgr
Do Until rs_stk_item_lgr.EOF
    List_card.AddItem rs_stk_item_lgr!stk_item_lgr_name
    If rs_stk_item_lgr!stk_item_lgr_alis <> "" Then List_card.AddItem rs_stk_item_lgr!stk_item_lgr_alis
rs_stk_item_lgr.MoveNext
Loop

Call set_stock_summary_grid
Call separation_of_all_inventory_to_inward_and_outward
Call search_closing_stock
List_card.Visible = False
End Sub
Public Sub set_stock_summary_grid()

    grid_stk_dtl.RowHeightMin = 400
    grid_stk_dtl.Clear
    grid_stk_dtl.Rows = 2
    grid_stk_dtl.Cols = 11
    

    grid_stk_dtl.TextMatrix(0, 1) = "Date"
    grid_stk_dtl.TextMatrix(0, 2) = "Starting Serial No."
    grid_stk_dtl.TextMatrix(0, 3) = "Ending Serial No."
    grid_stk_dtl.TextMatrix(0, 4) = "Balance Qty"
    grid_stk_dtl.TextMatrix(0, 5) = "Rate"
    grid_stk_dtl.TextMatrix(0, 6) = "Amount"
    grid_stk_dtl.TextMatrix(0, 7) = "F.Val"
    grid_stk_dtl.TextMatrix(0, 8) = "Company name"
    grid_stk_dtl.TextMatrix(0, 9) = "VAT"
    grid_stk_dtl.TextMatrix(0, 10) = "Suplier"
    
    grid_stk_dtl.ColWidth(0) = 500
    grid_stk_dtl.ColWidth(1) = 1500
    grid_stk_dtl.ColWidth(2) = 2500
    grid_stk_dtl.ColWidth(3) = 2500
    grid_stk_dtl.ColWidth(4) = 1000
    grid_stk_dtl.ColWidth(5) = 1000
    grid_stk_dtl.ColWidth(6) = 1500
    grid_stk_dtl.ColWidth(7) = 1000
    grid_stk_dtl.ColWidth(8) = 2500
    grid_stk_dtl.ColWidth(9) = 800
    grid_stk_dtl.ColWidth(10) = 2500

Dim temp_grid_col_no
Dim temp_grid_width
temp_grid_width = 0
For temp_grid_col_no = 0 To grid_stk_dtl.Cols - 1
temp_grid_width = temp_grid_width + grid_stk_dtl.ColWidth(temp_grid_col_no)
Next
grid_stk_dtl.Width = temp_grid_width + 800

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
Call open_rs_stk_clsg_srl
If rs_stk_clsg_srl.State = 1 Then rs_stk_clsg_srl.Close
Call open_rs_stk_clsg_srl

Dim rs_stk_clsg_srl_counter
Dim grid_stk_row_no

grid_stk_row_no = 1
Dim total_inward
Dim total_outward
Dim temp_stock_balance
grid_stk_dtl.Font.Size = 12
For rs_stk_clsg_srl_counter = 1 To rs_stk_clsg_srl.RecordCount
If rs_stk_clsg_srl!stk_invt_clg_card = selected_stock_item_name Then
With rs_stk_clsg_srl
       grid_stk_dtl.TextMatrix(grid_stk_row_no, 0) = grid_stk_row_no
       grid_stk_dtl.TextMatrix(grid_stk_row_no, 1) = Date
        If !stk_invt_clg_stno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = !stk_invt_clg_stno
        If !stk_invt_clg_edno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 3) = !stk_invt_clg_edno
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = (Val(grid_stk_dtl.TextMatrix(grid_stk_row_no, 3)) - Val(grid_stk_dtl.TextMatrix(grid_stk_row_no, 2))) + 1
        temp_stock_balance = temp_stock_balance + Val(grid_stk_dtl.TextMatrix(grid_stk_row_no, 4))
        If !stk_invt_clg_rate <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 5) = Format(!stk_invt_clg_rate, "0.00")
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 6) = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_row_no, 4)) * Val(!stk_invt_clg_rate), "0.00")
        If !stk_invt_clg_fval <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = Format(!stk_invt_clg_fval, "0.00")
        If !stk_invt_clg_comp <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = !stk_invt_clg_comp
        If !stk_invt_clg_vtyp <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = !stk_invt_clg_vtyp
        If !stk_invt_clg_splr <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 10) = !stk_invt_clg_splr
            grid_stk_row_no = grid_stk_row_no + 1
            grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
End With
End If
rs_stk_clsg_srl.MoveNext
Next
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = "==========="
    grid_stk_row_no = grid_stk_row_no + 1
    grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = "Balance Quantity On"
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 3) = Date
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = temp_stock_balance
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 5) = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_row_no - 2, 5)), "0.00")
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 6) = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_row_no, 4)) * Val(grid_stk_dtl.TextMatrix(grid_stk_row_no - 2, 5)), "0.00")

    grid_stk_row_no = grid_stk_row_no + 1
    grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = "==========="
End Sub

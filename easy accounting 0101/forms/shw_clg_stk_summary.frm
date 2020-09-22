VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form shw_clg_stk_summary 
   Caption         =   "closing stock"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "shw_clg_stk_summary.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to Exit"
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
      Left            =   0
      TabIndex        =   3
      Top             =   6240
      Width           =   11775
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   10080
      Top             =   10440
   End
   Begin MSFlexGridLib.MSFlexGrid grid_stk_dtl 
      Height          =   4575
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8070
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   10200
      TabIndex        =   7
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   51838977
      CurrentDate     =   40194
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
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
      Left            =   9480
      TabIndex        =   8
      Top             =   960
      Width           =   615
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
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   12540
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
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   12540
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12420
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
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   240
      TabIndex        =   2
      Top             =   6720
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Stock"
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
      Left            =   5400
      TabIndex        =   0
      Top             =   1080
      Width           =   1725
   End
End
Attribute VB_Name = "shw_clg_stk_summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================
    r_clr = 250
    g_clr = 50
    b_clr = 50

m_label.Caption = ""
Call open_database
Call open_rs_stk_item_lgr
Call set_stock_summary_grid
Call separation_of_all_inventory_to_inward_and_outward
Call search_closing_stock
Call enter_the_card_from_list
End Sub
Public Sub set_stock_summary_grid()

    grid_stk_dtl.RowHeightMin = 400
    grid_stk_dtl.Clear
    grid_stk_dtl.Rows = 2
    grid_stk_dtl.Cols = 7

    grid_stk_dtl.TextMatrix(0, 0) = "No"
    grid_stk_dtl.TextMatrix(0, 1) = "Item"
    grid_stk_dtl.TextMatrix(0, 2) = "Quantity"
    grid_stk_dtl.TextMatrix(0, 3) = "Rate"
    grid_stk_dtl.TextMatrix(0, 4) = "Amount"
    grid_stk_dtl.TextMatrix(0, 5) = "F.Val"
    grid_stk_dtl.TextMatrix(0, 6) = "Company"
    
    grid_stk_dtl.ColWidth(0) = 500
    grid_stk_dtl.ColWidth(1) = 4500
    grid_stk_dtl.ColWidth(2) = 1500
    grid_stk_dtl.ColWidth(3) = 1500
    grid_stk_dtl.ColWidth(4) = 1500
    grid_stk_dtl.ColWidth(5) = 1500
    grid_stk_dtl.ColWidth(6) = 2500

'Dim temp_grid_col_no
'Dim temp_grid_width
'temp_grid_width = 0

'For temp_grid_col_no = 0 To grid_stk_dtl.Cols - 1
'temp_grid_width = temp_grid_width + grid_stk_dtl.ColWidth(temp_grid_col_no)
'Next

'grid_stk_dtl.Width = temp_grid_width + 800

'grid_stk_dtl.Left = (Me.Width - grid_stk_dtl.Width) / 2
'grid_stk_dtl.Top = 2000
End Sub
Public Sub enter_the_card_from_list()
Call set_stock_summary_grid
    Dim rs_stk_clsg_srl_counter
    Dim grid_stk_row_no
    Dim total_inward
    Dim total_outward
    Dim temp_stock_balance
    grid_stk_row_no = 1
    grid_stk_dtl.Font.Size = 12

Call open_database
Call open_rs_stk_item_lgr
rs_stk_item_lgr.Sort = "stk_item_lgr_name"
Do Until rs_stk_item_lgr.EOF
    selected_stock_item_name = rs_stk_item_lgr!stk_item_lgr_name
    'Call open_database
    If rs_stk_clsg_srl.State = 1 Then rs_stk_clsg_srl.Close
    Call open_rs_stk_clsg_srl
    For rs_stk_clsg_srl_counter = 1 To rs_stk_clsg_srl.RecordCount
        If rs_stk_clsg_srl!stk_invt_clg_card = selected_stock_item_name Then
                temp_stock_balance = temp_stock_balance + (Val(rs_stk_clsg_srl!stk_invt_clg_edno) - Val(rs_stk_clsg_srl!stk_invt_clg_stno)) + 1
        End If
    rs_stk_clsg_srl.MoveNext
    Next
                    
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 0) = grid_stk_row_no
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 1) = rs_stk_item_lgr!stk_item_lgr_name
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = temp_stock_balance
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 3) = rs_stk_item_lgr!stk_item_lgr_rat1
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = Format(temp_stock_balance * Val(rs_stk_item_lgr!stk_item_lgr_rat1), "0.00")
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 5) = Format(rs_stk_item_lgr!stk_item_lgr_fcvl, "0.00")
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 6) = rs_stk_item_lgr!stk_item_lgr_comp
    
    m_label.Caption = m_label.Caption & "  " & rs_stk_item_lgr!stk_item_lgr_name & "   (" & temp_stock_balance & ") " & Format(temp_stock_balance * Val(rs_stk_item_lgr!stk_item_lgr_rat1), "0.00") & "£      "
    temp_stock_balance = 0
    grid_stk_row_no = grid_stk_row_no + 1
    grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
    rs_stk_item_lgr.MoveNext
Loop

Dim all_item_total_stock_balance
Dim all_item_total_stock_balance_amount
all_item_total_stock_balance = 0
Dim grid_stk_dtl_counter
For grid_stk_dtl_counter = 1 To grid_stk_dtl.Rows - 1
    all_item_total_stock_balance = all_item_total_stock_balance + Val(grid_stk_dtl.TextMatrix(grid_stk_dtl_counter, 2))
    all_item_total_stock_balance_amount = all_item_total_stock_balance_amount + Val(grid_stk_dtl.TextMatrix(grid_stk_dtl_counter, 4))
Next
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = "==========="
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = "==========="
        grid_stk_row_no = grid_stk_row_no + 1
        grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 1) = "Balance Quantity On" & Date & " "
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = all_item_total_stock_balance
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 3) = " Amount.."
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = Format(all_item_total_stock_balance_amount, "0.00")
        grid_stk_row_no = grid_stk_row_no + 1
        grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = "==========="
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = "==========="
End Sub
Private Sub grid_stk_dtl_DblClick()
show_stock_item_by_click = 1
selected_stock_item_name = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 1)
selected_procedure = "serial wise closing stock"
shw_item_wise_clg_stk.Show

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

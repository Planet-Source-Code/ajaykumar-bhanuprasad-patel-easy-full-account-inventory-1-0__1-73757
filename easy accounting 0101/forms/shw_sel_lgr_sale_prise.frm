VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form shw_sel_lgr_sale_prise 
   BackColor       =   &H00FFC0C0&
   Caption         =   "show_selected_ledger_sale_prise"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "shw_sel_lgr_sale_prise.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to exit."
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   6000
      Width           =   10935
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4800
      Top             =   11880
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
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   360
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
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid grid_report 
      Height          =   4695
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8281
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   705
   End
   Begin VB.Label m_label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please select ledger...,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   555
      Left            =   360
      TabIndex        =   5
      Top             =   6600
      Width           =   5175
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
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   11085
   End
End
Attribute VB_Name = "shw_sel_lgr_sale_prise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private report_temp_closing_balance
Public Sub set_ledger_account_detail()
m_label.Caption = ""
Call set_report_grid
selected_ledger = selected_from_list.Text
selected_voucher_ledger = selected_ledger
b = 0
                grid_report.TextMatrix(b, 1) = "Item"
                grid_report.TextMatrix(b, 2) = "Face Value"
                grid_report.TextMatrix(b, 3) = "Item Value"
                
b = 1
grid_report.Rows = grid_report.Rows + 1
Dim rs_inv_tran_sal_counter
Call open_database
Call open_rs_stk_item_lgr
Do Until rs_stk_item_lgr.EOF
    selected_stock_item_name = rs_stk_item_lgr!stk_item_lgr_name
    Call open_rs_inv_tran_sal
    If rs_inv_tran_sal.RecordCount > 0 Then rs_inv_tran_sal.MoveLast
    For rs_inv_tran_sal_counter = rs_inv_tran_sal.RecordCount To 1 Step -1
        If LCase(selected_stock_item_name) = LCase(rs_inv_tran_sal!stk_invt_trn_card) And LCase(selected_ledger) = LCase(rs_inv_tran_sal!stk_invt_trn_ldgr) Then
                grid_report.TextMatrix(b, 1) = rs_inv_tran_sal!stk_invt_trn_card
                grid_report.TextMatrix(b, 2) = Format(rs_inv_tran_sal!stk_invt_trn_fval, "0.00")
                grid_report.TextMatrix(b, 3) = Format(rs_inv_tran_sal!stk_invt_trn_rate, "0.00")
                m_label.Caption = m_label.Caption & "  " & rs_inv_tran_sal!stk_invt_trn_card & "   Rate :" & Format(rs_inv_tran_sal!stk_invt_trn_rate, "0.00") & "£      "
                grid_report.Rows = grid_report.Rows + 1
                b = b + 1
                Exit For
        End If
        rs_inv_tran_sal.MovePrevious
    Next
rs_stk_item_lgr.MoveNext
Loop
End Sub
Private Sub combo_list_LostFocus()
combo_list.Visible = False
End Sub
Private Sub combo_list_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    selected_from_list.Text = combo_list.Text
    combo_list.Visible = False
    Dim today_day As Integer
    Dim today_weekday As Integer
    today_weekday = Weekday(Now)
    today_day = Day(Now) - 1
    Call set_ledger_account_detail
    Text2.SetFocus
    Label1.Caption = selected_ledger & " Sale Prise on " & Date
    'Label1.Left = (Me.Width - Label1.Width) / 2
    'Label5.Left = (Me.Width - Label5.Width) / 2
    'selected_from_list.Left = (Me.Width - selected_from_list.Width) / 2
    'combo_list.Left = selected_from_list.Left
    'grid_report.Left = (Me.Width - grid_report.Width) / 2
End If
End Sub

Private Sub Command1_Click()
Unload Me
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
'Call set_ledger_account_detail
End Sub
Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================
    'Label1.Left = (Me.Width - Label1.Width) / 2
    'Label5.Left = (Me.Width - Label5.Width) / 2
    'selected_from_list.Left = (Me.Width - selected_from_list.Width) / 2
    'combo_list.Left = selected_from_list.Left
    'grid_report.Left = (Me.Width - grid_report.Width) / 2

    selected_from_list.Text = ""
    Call add_combo_list
    combo_list.Visible = False
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
    grid_report.Cols = 4
    grid_report.Font.Size = 12
    b = 0
    grid_report.TextMatrix(b, 0) = ""
    grid_report.TextMatrix(b, 1) = "card"
    grid_report.TextMatrix(b, 2) = "Face Value"
    grid_report.TextMatrix(b, 3) = "Rate"
    
    grid_report.ColWidth(0) = 1
    grid_report.ColWidth(1) = 6000
    grid_report.ColWidth(2) = 2000
    grid_report.ColWidth(3) = 2000
    
    'Dim x_grid_col
    'Dim total_grid_width
    'total_grid_width = 500
    'For x_grid_col = 0 To grid_report.Cols - 1
    '    total_grid_width = total_grid_width + grid_report.ColWidth(x_grid_col)
    'Next
    'grid_report.Width = total_grid_width
End Sub

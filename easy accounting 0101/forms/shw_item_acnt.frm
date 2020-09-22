VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form shw_item_acnt 
   Caption         =   "closing stock"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "shw_item_acnt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Click here exit"
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   6360
      Width           =   11295
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
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
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
      Left            =   9240
      TabIndex        =   13
      Text            =   "Select Option"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Option"
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   5400
      TabIndex        =   9
      Top             =   240
      Width           =   2295
      Begin VB.OptionButton Option3 
         Caption         =   "Find Serial No"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Specific Period"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normal Account "
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.TextBox Text3 
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
      Left            =   9240
      TabIndex        =   8
      Top             =   240
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   9840
      Top             =   6240
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   -615
      Width           =   2655
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
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid grid_stk_dtl 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7435
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   9960
      TabIndex        =   14
      Top             =   1320
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
      Format          =   111673345
      CurrentDate     =   40126
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9960
      TabIndex        =   15
      Top             =   840
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
      Format          =   111673345
      CurrentDate     =   40126
   End
   Begin VB.Label Label6 
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
      Left            =   9120
      TabIndex        =   18
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label5 
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
      Left            =   9120
      TabIndex        =   17
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label4 
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
      Height          =   495
      Left            =   8040
      TabIndex        =   16
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Serial No."
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
      Left            =   8040
      TabIndex        =   7
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Stock Item"
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
      Top             =   240
      Width           =   1755
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
      Left            =   120
      TabIndex        =   5
      Top             =   6960
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Closing Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   11235
   End
End
Attribute VB_Name = "shw_item_acnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Call enter_the_card_from_list
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
Dim today_day As Integer
Dim today_weekday As Integer
today_weekday = Weekday(Now)
today_day = Day(Now) - 1
Call enter_the_card_from_list
End Sub
Private Sub DTPicker2_Change()
Dim today_day As Integer
Dim today_weekday As Integer
today_weekday = Weekday(Now)
today_day = Day(Now) - 1
Call enter_the_card_from_list
End Sub
Private Sub List_card_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text1.Text = List_card.Text
    selected_stock_item_name = Text1.Text
    List_card.Visible = False
    Call enter_the_card_from_list
    Text2.SetFocus
End If
Label1.Caption = selected_stock_item_name & " Account "
End Sub

Private Sub Option1_Click()
Label3.Visible = False
Text3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Combo2.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Call set_stock_item_account_grid
End Sub

Private Sub Option2_Click()

Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Combo2.Visible = True

DTPicker1.Visible = True
DTPicker2.Visible = True

Label3.Visible = False
Text3.Visible = False
Call set_stock_item_account_grid
End Sub

Private Sub Option3_Click()

Label3.Visible = True
Text3.Visible = True

Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Combo2.Visible = False

DTPicker1.Visible = False
DTPicker2.Visible = False
Call set_stock_item_account_grid
End Sub

Private Sub Text1_GotFocus()
    List_card.Visible = True
    List_card.Height = 2400
    List_card.SetFocus
End Sub
Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================

Label3.Visible = False
Text3.Visible = False

Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Combo2.Visible = False

DTPicker1.Visible = False
DTPicker2.Visible = False

Combo2.AddItem "This Year"
Combo2.AddItem "This Month"
    
Combo2.AddItem "This Week"
Combo2.AddItem "Last Month"
Combo2.AddItem "Last Week"

Call open_database
Call open_rs_stk_item_lgr

Do Until rs_stk_item_lgr.EOF
    List_card.AddItem rs_stk_item_lgr!stk_item_lgr_name
    If rs_stk_item_lgr!stk_item_lgr_alis <> "" Then List_card.AddItem rs_stk_item_lgr!stk_item_lgr_alis
    rs_stk_item_lgr.MoveNext
Loop
If show_stock_item_by_click = 1 Then
    Text1.Text = selected_stock_item_name
    Call set_stock_item_account_grid
    Call separation_of_all_inventory_to_inward_and_outward
    Call search_closing_stock
    Call enter_the_card_from_list
    Label1.Caption = "Closing Stock as on the " & Date & " of " & selected_stock_item_name
Else
    Call set_stock_item_account_grid
    Call separation_of_all_inventory_to_inward_and_outward
    Call search_closing_stock
End If
End Sub
Public Sub set_stock_item_account_grid()
    grid_stk_dtl.RowHeightMin = 400
    grid_stk_dtl.Clear
    grid_stk_dtl.Rows = 2
    grid_stk_dtl.Cols = 10
    grid_stk_dtl.TextMatrix(0, 1) = "Date"
    grid_stk_dtl.TextMatrix(0, 2) = "Transaction"
    
    grid_stk_dtl.TextMatrix(0, 3) = "V. No"
    grid_stk_dtl.TextMatrix(0, 4) = "Ledger"

    grid_stk_dtl.TextMatrix(0, 5) = "Starting Serial No."
    grid_stk_dtl.TextMatrix(0, 6) = "Ending Serial No."
    grid_stk_dtl.TextMatrix(0, 7) = "Inward"
    grid_stk_dtl.TextMatrix(0, 8) = "Outward"
    grid_stk_dtl.TextMatrix(0, 9) = "Balance"
    
    'grid_stk_dtl.TextMatrix(0, 9) = "F.Val"
    'grid_stk_dtl.TextMatrix(0, 8) = "Company name"
    'grid_stk_dtl.TextMatrix(0, 9) = "VAT"
    'grid_stk_dtl.TextMatrix(0, 10) = "Suplier"
    
    grid_stk_dtl.ColWidth(0) = 500
    grid_stk_dtl.ColWidth(1) = 1500
    grid_stk_dtl.ColWidth(2) = 2500
    
    grid_stk_dtl.ColWidth(3) = 1000
    grid_stk_dtl.ColWidth(4) = 2500
    
    grid_stk_dtl.ColWidth(5) = 2500
    grid_stk_dtl.ColWidth(6) = 2500
    grid_stk_dtl.ColWidth(7) = 1200
    grid_stk_dtl.ColWidth(8) = 1200
    grid_stk_dtl.ColWidth(9) = 1200
    'grid_stk_dtl.ColWidth(8) = 2500
    'grid_stk_dtl.ColWidth(9) = 800
    'grid_stk_dtl.ColWidth(10) = 2500

'Dim temp_grid_col_no
'Dim temp_grid_width
'temp_grid_width = 0
'For temp_grid_col_no = 0 To grid_stk_dtl.Cols - 1
'temp_grid_width = temp_grid_width + grid_stk_dtl.ColWidth(temp_grid_col_no)
'Next
'grid_stk_dtl.Width = temp_grid_width + 200

End Sub
Public Sub enter_the_card_from_list()

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
Call set_stock_item_account_grid
Call open_database
Call open_rs_inv_tran_all
If rs_inv_tran_all.State = 1 Then rs_inv_tran_all.Close
Call open_rs_inv_tran_all
Dim rs_inv_tran_all_counter
Dim grid_stk_row_no
Dim total_inward
Dim total_outward
Dim temp_stock_balance
grid_stk_row_no = 1
total_inward = 0
total_outward = 0
grid_stk_dtl.Font.Size = 12
If Option1.Value = True Then ' for normal item account
        rs_inv_tran_all.Sort = "stk_invt_trn_stno,stk_invt_trn_date,stk_invt_trn_time"
        For rs_inv_tran_all_counter = 1 To rs_inv_tran_all.RecordCount
        If rs_inv_tran_all!stk_invt_trn_card = selected_stock_item_name Then
        With rs_inv_tran_all
               grid_stk_dtl.TextMatrix(grid_stk_row_no, 0) = grid_stk_row_no
               If !stk_invt_trn_date <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 1) = !stk_invt_trn_date
               If !stk_invt_trn_vchr <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = !stk_invt_trn_vchr
               If !stk_invt_trn_vcno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 3) = !stk_invt_trn_vcno
               If !stk_invt_trn_ldgr <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = !stk_invt_trn_ldgr
               If !stk_invt_trn_stno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 5) = !stk_invt_trn_stno
               If !stk_invt_trn_edno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 6) = !stk_invt_trn_edno
               If !stk_invt_trn_vchr = "purchase" Or !stk_invt_trn_vchr = "sale return" Or !stk_invt_trn_vchr = "opening stock" Then
                    grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                    total_inward = total_inward + (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
               ElseIf !stk_invt_trn_vchr = "sale" Or !stk_invt_trn_vchr = "purchase return" Then
                    grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                    total_outward = total_outward + (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
               End If
               temp_stock_balance = total_inward - total_outward
               grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = temp_stock_balance
               grid_stk_row_no = grid_stk_row_no + 1
               grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
        End With
        End If
        rs_inv_tran_all.MoveNext
        Next
ElseIf Option2.Value = True Then ' item account from x to y date period
        rep_starting_date = DTPicker1.Value
        rep_ending_date = DTPicker2.Value
        rs_inv_tran_all.Sort = "stk_invt_trn_date,stk_invt_trn_time,stk_invt_trn_stno"
        For rs_inv_tran_all_counter = 1 To rs_inv_tran_all.RecordCount
            If rs_inv_tran_all!stk_invt_trn_card = selected_stock_item_name Then
            If rep_starting_date <= rs_inv_tran_all!stk_invt_trn_date And rep_ending_date >= rs_inv_tran_all!stk_invt_trn_date Then
                If grid_stk_row_no = 1 Then
                        If Val(total_inward - total_outward) < 0 Then
                            temp_stock_balance = total_inward - total_outward
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = "Opening Stock"
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = 0
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = temp_stock_balance
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = temp_stock_balance
                            total_outward = temp_stock_balance
                            total_inward = 0
                        ElseIf Val(total_inward - total_outward) > 0 Then
                            temp_stock_balance = total_inward - total_outward
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = "Opening Stock"
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = temp_stock_balance
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = 0
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = temp_stock_balance
                            total_outward = 0
                            total_inward = temp_stock_balance
                        End If
                grid_stk_row_no = grid_stk_row_no + 1
                grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
                End If
                With rs_inv_tran_all
                           grid_stk_dtl.TextMatrix(grid_stk_row_no, 0) = grid_stk_row_no
                           If !stk_invt_trn_date <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 1) = !stk_invt_trn_date
                           If !stk_invt_trn_vchr <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = !stk_invt_trn_vchr
                           If !stk_invt_trn_vcno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 3) = !stk_invt_trn_vcno
                           If !stk_invt_trn_ldgr <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = !stk_invt_trn_ldgr
                           If !stk_invt_trn_stno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 5) = !stk_invt_trn_stno
                           If !stk_invt_trn_edno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 6) = !stk_invt_trn_edno
                           If !stk_invt_trn_vchr = "purchase" Or !stk_invt_trn_vchr = "sale return" Or !stk_invt_trn_vchr = "opening stock" Then
                                grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                                total_inward = total_inward + (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                           ElseIf !stk_invt_trn_vchr = "sale" Or !stk_invt_trn_vchr = "purchase return" Then
                                grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                                total_outward = total_outward + (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                           End If
                           temp_stock_balance = total_inward - total_outward
                           grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = temp_stock_balance
                           grid_stk_row_no = grid_stk_row_no + 1
                           grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
                End With
            Else
                        With rs_inv_tran_all
                           If !stk_invt_trn_vchr = "purchase" Or !stk_invt_trn_vchr = "sale return" Or !stk_invt_trn_vchr = "opening stock" Then
                                grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                                total_inward = total_inward + (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                           ElseIf !stk_invt_trn_vchr = "sale" Or !stk_invt_trn_vchr = "purchase return" Then
                                grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                                total_outward = total_outward + (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                           End If
                        End With
            End If
            End If
            rs_inv_tran_all.MoveNext
        Next
                If grid_stk_row_no = 1 Then
                        If Val(total_inward - total_outward) < 0 Then
                            temp_stock_balance = total_inward - total_outward
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = "Opening Stock"
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = 0
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = temp_stock_balance
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = temp_stock_balance
                            total_outward = temp_stock_balance
                            total_inward = 0
                        ElseIf Val(total_inward - total_outward) > 0 Then
                            temp_stock_balance = total_inward - total_outward
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = "Opening Stock"
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = temp_stock_balance
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = 0
                            grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = temp_stock_balance
                            total_outward = 0
                            total_inward = temp_stock_balance
                        End If
                grid_stk_row_no = grid_stk_row_no + 1
                grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
                End If
ElseIf Option3.Value = True Then ' find a card of specific serial no.
        Dim selected_item_serial_no
        selected_item_serial_no = Val(Text3.Text)
        For rs_inv_tran_all_counter = 1 To rs_inv_tran_all.RecordCount
        
        If rs_inv_tran_all!stk_invt_trn_card = selected_stock_item_name Then
        If selected_item_serial_no >= Val(rs_inv_tran_all!stk_invt_trn_stno) And selected_item_serial_no <= Val(rs_inv_tran_all!stk_invt_trn_edno) Then
            With rs_inv_tran_all
                        grid_stk_dtl.TextMatrix(grid_stk_row_no, 0) = grid_stk_row_no
                        If !stk_invt_trn_date <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 1) = !stk_invt_trn_date
                        If !stk_invt_trn_vchr <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = !stk_invt_trn_vchr
                        If !stk_invt_trn_vcno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 3) = !stk_invt_trn_vcno
                        If !stk_invt_trn_ldgr <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = !stk_invt_trn_ldgr
                        If !stk_invt_trn_stno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 5) = !stk_invt_trn_stno
                        If !stk_invt_trn_edno <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 6) = !stk_invt_trn_edno
                        If !stk_invt_trn_vchr = "purchase" Or !stk_invt_trn_vchr = "sale return" Or !stk_invt_trn_vchr = "opening stock" Then
                        grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                        total_inward = total_inward + (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                        ElseIf !stk_invt_trn_vchr = "sale" Or !stk_invt_trn_vchr = "purchase return" Then
                        grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                        total_outward = total_outward + (Val(!stk_invt_trn_edno) - Val(!stk_invt_trn_stno)) + 1
                        End If
                        temp_stock_balance = total_inward - total_outward
                        grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = temp_stock_balance
                       'If !stk_invt_trn_rate <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = Format(!stk_invt_trn_rate, "0.00")
                       'If !stk_invt_trn_fval <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = Format(!stk_invt_trn_fval, "0.00")
                       'If !stk_invt_trn_vtyp <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = !stk_invt_trn_vtyp
                       'If !stk_invt_trn_splr <> "" Then grid_stk_dtl.TextMatrix(grid_stk_row_no, 10) = !stk_invt_trn_splr
                        grid_stk_row_no = grid_stk_row_no + 1
                        grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
            End With
        End If
        End If
        rs_inv_tran_all.MoveNext
        Next
End If
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = "==========="
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = "==========="
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = "==========="
    grid_stk_row_no = grid_stk_row_no + 1
    grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 6) = "Total"
    'grid_stk_dtl.TextMatrix(grid_stk_row_no, 5) = Date
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = total_inward
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = total_outward
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = temp_stock_balance
    'grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_row_no - 2, 5)), "0.00")
    'Dim total_stock_balance_amount
    'total_stock_balance_amount = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_row_no, 4)) * Val(grid_stk_dtl.TextMatrix(grid_stk_row_no - 2, 5)), "0.00")
    'grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = Format(Val(grid_stk_dtl.TextMatrix(grid_stk_row_no, 4)) * Val(grid_stk_dtl.TextMatrix(grid_stk_row_no - 2, 5)), "0.00")
    grid_stk_row_no = grid_stk_row_no + 1
    grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 7) = "==========="
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 8) = "==========="
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 9) = "==========="
    If grid_stk_dtl.Rows > 15 Then grid_stk_dtl.Width = grid_stk_dtl.Width + 400
    'm_label.Caption = "Balance of Stock On " & Date & " of " & Text1.Text & " are " & temp_stock_balance & " The arpox value are .." & Format(total_stock_balance_amount, "0.00") & "£"
    m_label.Caption = Text1.Text & " (" & temp_stock_balance & ").."
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If Val(Text3.Text) >= 0 And KeyCode = 13 Then
Call enter_the_card_from_list
End If
End Sub

Private Sub Timer1_Timer()
If m_label.Left + m_label.Width <= 0 Then m_label.Left = Me.Width ' + m_label.Width
m_label.Left = m_label.Left - 500
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form creat_stock_lgr_detail 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Opening Stock Detail"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "creat_stock_lgr_detail.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   9720
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   8040
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7646
      _Version        =   393216
      ScrollBars      =   2
      AllowUserResizing=   3
   End
   Begin VB.Label lbl_3 
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   12
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lbl_2 
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   11
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lbl_1 
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lbl_card 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   5535
      TabIndex        =   5
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of company"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4155
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4815
      TabIndex        =   3
      Top             =   360
      Width           =   915
   End
   Begin VB.Label lbl_Heading 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Stock Detail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3570
      TabIndex        =   2
      Top             =   840
      Width           =   3435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..........................."
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   11085
   End
End
Attribute VB_Name = "creat_stock_lgr_detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private row_to_enter As Integer
Private record_is_available As Integer
Private row_no As Integer
Private temp_x As Integer
Private keycode_now  As Integer
Private pressed_key As Integer
Private selected_row
Private selected_col
Private pre_row
Private pre_col
Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================
Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)

If selected_procedure = "Stock_item_ledger_display" Then
Text1.Enabled = False
Combo1.Enabled = False
End If

record_is_available = 0
    
    Call set_form_values
    Call set_combo1
    Call set_grid1_data
    Call add_new_row
    Call refresh_total_lbl
    Call opening_stock_val
End Sub
Public Sub set_form_values()
    Text1.Text = ""
    lbl_card.Caption = selected_stock_item
    lbl_name.Caption = co_name
    lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
    'Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")
    If selected_path = "" Or selected_path = Null Then
    selected_path = App.Path & "\data\1000\co.mdb;"
End If
End Sub
Public Sub refresh_total_lbl()
    
    total_quantity = 0
    average_rate = 0
    total_amount = 0
    For temp_x = 2 To Grid1.Rows
    total_quantity = total_quantity + Val(Grid1.TextMatrix(temp_x - 1, 4))
    total_amount = total_amount + Val(Grid1.TextMatrix(temp_x - 1, 6))
    Next
    If total_amount > 0 And total_quantity > 0 Then average_rate = total_amount / total_quantity
    lbl_1.Caption = total_quantity
    lbl_3.Caption = Format(total_amount, "0.00")
    lbl_2.Caption = Format(average_rate, "0.00")
End Sub
Public Sub set_combo1()
    Combo1.FontSize = 12
    Call open_database
    Call open_rs_lgr_main_dtl
    Do Until rs_lgr_main_dtl.EOF
    selected_ledgers_group = rs_lgr_main_dtl!lgr_main_dtl_grup
    Call open_rs_lgr_main_grp
            Do Until rs_lgr_main_grp.EOF
                If rs_lgr_main_grp!lgr_main_grp_name = selected_ledgers_group Then
                    If rs_lgr_main_grp!lgr_main_grp_pgrp = "supplier" Or rs_lgr_main_grp!lgr_main_grp_pgrp = "Dealer" Then
                        Combo1.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
                    End If
                End If
            rs_lgr_main_grp.MoveNext
            Loop
    rs_lgr_main_dtl.MoveNext
    Loop
    Combo1.Text = "select a dealer.."
End Sub
Public Sub set_grid1_data()
    Grid1.RowHeightMin = 400
    Grid1.Clear
    Grid1.Rows = 2
    Grid1.Cols = 11
    Grid1.TextMatrix(0, 1) = "Item"
    Grid1.TextMatrix(0, 2) = "Start-No."
    Grid1.TextMatrix(0, 3) = "End-No."
    Grid1.TextMatrix(0, 4) = "Qty."
    Grid1.TextMatrix(0, 5) = "Rate"
    Grid1.TextMatrix(0, 6) = "Amount"
    Grid1.TextMatrix(0, 7) = "Face Val"
    Grid1.TextMatrix(0, 8) = "Dis.Rt."
    Grid1.TextMatrix(0, 9) = "Dealer"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 2000
    Grid1.ColWidth(3) = 2000
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 1200
    Grid1.ColWidth(10) = 1
End Sub
Public Sub opening_stock_val()

temp_x = 1
Call open_rs_stk_open_srl
Do Until rs_stk_open_srl.EOF
If rs_stk_open_srl!stk_open_srl_name = selected_stock_item_name And rs_stk_open_srl!stk_open_srl_type = selected_stock_item_type Then
    Grid1.Rows = temp_x + 1
    Grid1.TextMatrix(temp_x, 0) = temp_x 'rs_stk_open_srl!stk_open_srl_stid
    Grid1.TextMatrix(temp_x, 1) = rs_stk_open_srl!stk_open_srl_name
    Grid1.TextMatrix(temp_x, 2) = rs_stk_open_srl!stk_open_srl_stno
    Grid1.TextMatrix(temp_x, 3) = rs_stk_open_srl!stk_open_srl_edno
    Grid1.TextMatrix(temp_x, 4) = rs_stk_open_srl!stk_open_srl_qnty
    Grid1.TextMatrix(temp_x, 5) = Format(rs_stk_open_srl!stk_open_srl_rate, ".00")
    Grid1.TextMatrix(temp_x, 6) = Format(rs_stk_open_srl!stk_open_srl_amnt, ".00")
    Grid1.TextMatrix(temp_x, 7) = Format(rs_stk_open_srl!stk_open_srl_fcvl, ".00")
    Grid1.TextMatrix(temp_x, 8) = rs_stk_open_srl!stk_open_srl_disc
    Grid1.TextMatrix(temp_x, 9) = rs_stk_open_srl!stk_open_srl_splr
    temp_x = temp_x + 1
End If
rs_stk_open_srl.MoveNext

Loop
Call refresh_total_lbl

End Sub
Private Sub Command1_Click()

Call open_rs_stk_open_srl

If selected_procedure = "Stock_item_ledger_creat" Then
        For temp_x = 2 To Grid1.Rows - 1
        
        If Val(Grid1.TextMatrix(temp_x - 1, 4)) > 0 Then
        rs_stk_open_srl.AddNew
        rs_stk_open_srl!stk_open_srl_stid = Val(Grid1.TextMatrix(temp_x - 1, 0))
        rs_stk_open_srl!stk_open_srl_name = Grid1.TextMatrix(temp_x - 1, 1)
        rs_stk_open_srl!stk_open_srl_stno = Val(Grid1.TextMatrix(temp_x - 1, 2))
        rs_stk_open_srl!stk_open_srl_edno = Val(Grid1.TextMatrix(temp_x - 1, 3))
        rs_stk_open_srl!stk_open_srl_qnty = Val(Grid1.TextMatrix(temp_x - 1, 4))
        rs_stk_open_srl!stk_open_srl_rate = Val(Grid1.TextMatrix(temp_x - 1, 5))
        rs_stk_open_srl!stk_open_srl_amnt = Val(Grid1.TextMatrix(temp_x - 1, 6))
        rs_stk_open_srl!stk_open_srl_fcvl = Val(Grid1.TextMatrix(temp_x - 1, 7))
        rs_stk_open_srl!stk_open_srl_disc = Val(Grid1.TextMatrix(temp_x - 1, 8))
        rs_stk_open_srl!stk_open_srl_splr = Grid1.TextMatrix(temp_x - 1, 9)
        rs_stk_open_srl!stk_open_srl_type = selected_stock_item_type
        rs_stk_open_srl.UpdateBatch
        End If
        Next
ElseIf selected_procedure = "Stock_item_ledger_edit" Then 'write here code for the edit stock item detail
    
    
    If Val(Grid1.TextMatrix(Grid1.Rows - 1, 4)) <= 0 Then
    row_to_enter = Grid1.Rows - 2
    Else
    row_to_enter = Grid1.Rows - 1
    End If
    
    For row_no = 1 To row_to_enter
    
        'MsgBox "This is a record no....," & row_no
        
        record_is_available = 0
        Call open_rs_stk_open_srl
        
        Do Until rs_stk_open_srl.EOF
        If rs_stk_open_srl!stk_open_srl_name = selected_stock_item_name And _
            rs_stk_open_srl!stk_open_srl_stid = row_no And _
            rs_stk_open_srl!stk_open_srl_type = selected_stock_item_type Then
            rs_stk_open_srl!stk_open_srl_stid = Val(Grid1.TextMatrix(row_no, 0))
            rs_stk_open_srl!stk_open_srl_name = Grid1.TextMatrix(row_no, 1)
            rs_stk_open_srl!stk_open_srl_stno = Val(Grid1.TextMatrix(row_no, 2))
            rs_stk_open_srl!stk_open_srl_edno = Val(Grid1.TextMatrix(row_no, 3))
            rs_stk_open_srl!stk_open_srl_qnty = Val(Grid1.TextMatrix(row_no, 4))
            rs_stk_open_srl!stk_open_srl_rate = Val(Grid1.TextMatrix(row_no, 5))
            rs_stk_open_srl!stk_open_srl_amnt = Val(Grid1.TextMatrix(row_no, 6))
            rs_stk_open_srl!stk_open_srl_fcvl = Val(Grid1.TextMatrix(row_no, 7))
            rs_stk_open_srl!stk_open_srl_disc = Val(Grid1.TextMatrix(row_no, 8))
            rs_stk_open_srl!stk_open_srl_splr = Grid1.TextMatrix(row_no, 9)
            rs_stk_open_srl!stk_open_srl_type = selected_stock_item_type
            rs_stk_open_srl.UpdateBatch
                    
            record_is_available = 1
            Exit Do
        
        End If
        rs_stk_open_srl.MoveNext
        Loop
        
        
        Call open_rs_stk_open_srl
        If record_is_available = 0 Then
                If Val(Grid1.TextMatrix(row_no, 4)) = Null Or Val(Grid1.TextMatrix(row_no, 4)) <= 0 Then
                    MsgBox "One or more values is empty...!!!"
                    Exit Sub
                End If
                rs_stk_open_srl.AddNew
                rs_stk_open_srl!stk_open_srl_stid = Val(Grid1.TextMatrix(row_no, 0))
                rs_stk_open_srl!stk_open_srl_name = Grid1.TextMatrix(row_no, 1)
                rs_stk_open_srl!stk_open_srl_stno = Val(Grid1.TextMatrix(row_no, 2))
                rs_stk_open_srl!stk_open_srl_edno = Val(Grid1.TextMatrix(row_no, 3))
                rs_stk_open_srl!stk_open_srl_qnty = Val(Grid1.TextMatrix(row_no, 4))
                rs_stk_open_srl!stk_open_srl_rate = Val(Grid1.TextMatrix(row_no, 5))
                rs_stk_open_srl!stk_open_srl_amnt = Val(Grid1.TextMatrix(row_no, 6))
                rs_stk_open_srl!stk_open_srl_fcvl = Val(Grid1.TextMatrix(row_no, 7))
                rs_stk_open_srl!stk_open_srl_disc = Val(Grid1.TextMatrix(row_no, 8))
                rs_stk_open_srl!stk_open_srl_splr = Grid1.TextMatrix(row_no, 9)
                rs_stk_open_srl!stk_open_srl_type = selected_stock_item_type
                rs_stk_open_srl.UpdateBatch
        End If
    Next
End If
            lbl_1.Caption = total_quantity
            lbl_3.Caption = Format(total_amount, "0.00")
            lbl_2.Caption = Format(average_rate, "0.00")

If opg_stock_1_detail = 1 Then
        opg_stock_1_detail = 2
        opg_stock_2_detail = 0
        selected_stock_item_qnt1 = Val(lbl_1.Caption)
        selected_stock_item_rat1 = Val(lbl_2.Caption)
        selected_stock_item_amt1 = Val(lbl_3.Caption)
ElseIf opg_stock_2_detail = 1 Then
        opg_stock_1_detail = 0
        opg_stock_2_detail = 2
        selected_stock_item_qnt2 = Val(lbl_1.Caption)
        selected_stock_item_rat2 = Val(lbl_2.Caption)
        selected_stock_item_amt2 = Val(lbl_3.Caption)
End If

Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Combo1_Click()
Grid1.Text = Combo1.Text
If Grid1.Row = Grid1.Rows - 1 Then
                If Grid1.TextMatrix(Grid1.Row, 1) = "" Or Val(Grid1.TextMatrix(Grid1.Row, 2)) <= 0 Or Val(Grid1.TextMatrix(Grid1.Row, 3)) <= 0 Or Val(Grid1.TextMatrix(Grid1.Row, 4)) <= 0 Or Val(Grid1.TextMatrix(Grid1.Row, 5)) < 0 Or Val(Grid1.TextMatrix(Grid1.Row, 6)) < 0 Then
                MsgBox "Some Values are empty...!!!"
                Exit Sub
                End If
        Grid1.Text = Combo1.Text
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Col = 2
        Grid1.Row = Grid1.Row + 1
End If
Call add_new_row
End Sub
Public Sub add_new_row()
       Grid1.TextMatrix(Grid1.Row, 0) = Grid1.Row
       Grid1.TextMatrix(Grid1.Row, 1) = selected_stock_item_name
       Grid1.TextMatrix(Grid1.Row, 7) = Format(selected_stock_item_fval, ".00")
       Grid1.TextMatrix(Grid1.Row, 8) = Format(selected_stock_item_disc, ".00")
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
keycode_now = KeyCode
                If Grid1.TextMatrix(Grid1.Row, 1) = "" Or Val(Grid1.TextMatrix(Grid1.Row, 2)) <= 0 Or Val(Grid1.TextMatrix(Grid1.Row, 3)) <= 0 Or Val(Grid1.TextMatrix(Grid1.Row, 4)) <= 0 Or Val(Grid1.TextMatrix(Grid1.Row, 5)) < 0 Or Val(Grid1.TextMatrix(Grid1.Row, 6)) < 0 Then
                MsgBox "Some Values are empty...!!!"
                Exit Sub
                End If
If keycode_now = 37 Then
    Grid1.Text = Combo1.Text
    Grid1.Col = Grid1.Col - 1
ElseIf keycode_now = 39 Or keycode_now = 13 Then
    If Grid1.Col = 9 And Grid1.Row = Grid1.Rows - 1 Then
    Grid1.Text = Combo1.Text
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Col = 2
    Grid1.Row = Grid1.Row + 1
    Call add_new_row
    Exit Sub
    End If
    Grid1.Col = 2
    Grid1.Row = Grid1.Row + 1
End If
End Sub
Private Sub Grid1_Click()
If selected_procedure = "Stock_item_ledger_display" Then
Exit Sub
End If
selected_row = Grid1.Row
selected_col = Grid1.Col
If Grid1.Row = (pre_row + 1) And Grid1.Row >= 2 Then
    Grid1.TextMatrix(pre_row, 4) = (Val(Grid1.TextMatrix(pre_row, 3)) - Val(Grid1.TextMatrix(pre_row, 2))) + 1
End If
If selected_row = pre_row And selected_col = pre_col Then
    If selected_row = (pre_row + 1) Then
    Grid1.TextMatrix(pre_row, 5) = Val(Grid1.TextMatrix(pre_row, 4)) - Val(Grid1.TextMatrix(pre_row, 3))
    End If
'Else
    'Text1.Text = ""
End If
If Grid1.Col = 2 Or Grid1.Col = 3 Or Grid1.Col = 5 Then    ' Position and size the ComboBox, then show it.
Text1.Height = Grid1.CellHeight
Text1.Width = Grid1.CellWidth
Text1.Left = Grid1.CellLeft + Grid1.Left
Text1.Top = Grid1.CellTop + Grid1.Top
End If
If Grid1.Col = 9 Then
Combo1.Width = Grid1.CellWidth
Combo1.Left = Grid1.CellLeft + Grid1.Left
Combo1.Top = Grid1.CellTop + Grid1.Top
End If
pre_row = selected_row
pre_col = selected_col
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
keycode_now = KeyCode
If keycode_now = 37 Then
ElseIf keycode_now = 39 Or keycode_now = 13 Then
    If Grid1.Col < 9 Then
        Grid1.Col = Grid1.Col + 1
'    ElseIf Grid1.Col = 9 And Grid1.Row = Grid1.Rows - 1 Then
'        Grid1.Rows = Grid1.Rows + 1
'        Grid1.Col = 1
'        Grid1.Row = Grid1.Row + 1
'    ElseIf Grid1.Col = 5 And Grid1.Row <> Grid1.Rows - 1 Then
'        Grid1.Row = Grid1.Row + 1
'        Grid1.Col = 1
'    ElseIf Grid1.Col > 5 Then
'        Grid1.Row = Grid1.Row + 1
'        Grid1.Col = 1
    End If
End If
End Sub
Private Sub Grid1_RowColChange()
If selected_procedure = "Stock_item_ledger_display" Then
Exit Sub
End If
If Grid1.Col = 2 Or Grid1.Col = 3 Or Grid1.Col = 5 Then    ' Position and size the textbox, then show it.
    Combo1.Visible = False
    Text1.Visible = True
    If Grid1.TextMatrix(Grid1.Row, Grid1.Col) = "" Or Grid1.TextMatrix(Grid1.Row, Grid1.Col) = " " Or Grid1.TextMatrix(Grid1.Row, Grid1.Col) = Null Then
    Text1.Text = ""
    Else
    Text1.Text = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
    End If
    Text1.Height = Grid1.CellHeight
    Text1.Width = Grid1.CellWidth
    Text1.Left = Grid1.CellLeft + Grid1.Left
    Text1.Top = Grid1.CellTop + Grid1.Top
    Text1.SetFocus
ElseIf Grid1.Col = 9 Then
    Text1.Visible = False
    Combo1.Visible = True
    Combo1.Width = Grid1.CellWidth
    Combo1.Left = Grid1.CellLeft + Grid1.Left
    Combo1.Top = Grid1.CellTop + Grid1.Top
    Combo1.SetFocus
Else
    Combo1.Visible = False
    Text1.Visible = False
End If
End Sub

Private Sub Text1_Change()
    '   If Grid1.Col = 2 Then
    '   If Grid1.TextMatrix(Grid1.Row, 2) = "" Then
        '   Grid1.TextMatrix(Grid1.Row, 1) = selected_stock_item_name
        '   Grid1.TextMatrix(Grid1.Row, 7) = selected_stock_item_fval
        '   Grid1.TextMatrix(Grid1.Row, 8) = selected_stock_item_disc
        '   Grid1.TextMatrix(Grid1.Row, 0) = Grid1.Row
    '   End If
    '   End If
If Grid1.Col = 2 Or Grid1.Col = 3 Or Grid1.Col = 5 Then
    If Val(Text1.Text) < 0 Then
        Exit Sub
    ElseIf Grid1.Col = 5 Then
        Grid1.Text = Format(Text1.Text, ".00")
    Else
        Grid1.Text = Text1.Text
    End If
    If keycode_now = 13 And Grid1.Col + 1 = 3 Then
        If Val(Grid1.TextMatrix(Grid1.Row, 2)) < 0 And Val(Grid1.TextMatrix(Grid1.Row, 3)) < 0 Then Exit Sub
    End If
End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
keycode_now = KeyCode
If Grid1.Col = 2 Or Grid1.Col = 3 Or Grid1.Col = 5 Then    ' Position and size the ComboBox, then show it.
    Text1.Visible = True
    Text1.Height = Grid1.CellHeight
    Text1.Width = Grid1.CellWidth
    Text1.Left = Grid1.CellLeft + Grid1.Left
    Text1.Top = Grid1.CellTop + Grid1.Top
Else
    Text1.Visible = False
End If

If keycode_now = 37 Then
    If Grid1.Col = 1 And Grid1.Row = 1 Then
    MsgBox "Not Valid key.....!!!"
    Exit Sub
    End If
    If Grid1.Col = 2 Then
    Grid1.Row = Grid1.Row - 1
    Grid1.Col = 9
    Exit Sub
    End If
    Grid1.Col = Grid1.Col - 1
ElseIf keycode_now = 38 And Grid1.Row >= 1 And Grid1.Col <> 9 Then
    If Grid1.Row = 1 Then
    MsgBox "Not Valid key.....!!!"
    Exit Sub
    End If
    Grid1.Row = Grid1.Row - 1
    
ElseIf keycode_now = 39 Or keycode_now = 13 Then
    If Val(Text1.Text) < 0 Then
    MsgBox "You have entered invalid value...!!! please enter correct value...!!!"
    Text1.Text = ""
    Exit Sub
    End If
        
        If Grid1.Col = 3 Then
            Grid1.TextMatrix(Grid1.Row, 4) = (Val(Grid1.TextMatrix(Grid1.Row, 3)) - Val(Grid1.TextMatrix(Grid1.Row, 2))) + 1
        ElseIf Grid1.Col = 2 And Val(Grid1.TextMatrix(Grid1.Row, 3)) > 1 Then
            Grid1.TextMatrix(Grid1.Row, 4) = (Val(Grid1.TextMatrix(Grid1.Row, 3)) - Val(Grid1.TextMatrix(Grid1.Row, 2))) + 1
        ElseIf Grid1.Col = 5 Then
            Grid1.TextMatrix(Grid1.Row, 6) = Format(Val(Grid1.TextMatrix(Grid1.Row, 4)) * Val(Grid1.TextMatrix(Grid1.Row, 5)), ".00")
        End If
        
        Grid1.Col = Grid1.Col + 1
        Call refresh_total_lbl
ElseIf keycode_now = 40 And Grid1.Col <> 9 Then
    If Grid1.Rows = Grid1.Row + 1 Then
    MsgBox "Not Valid key.....!!!"
    Exit Sub
    Else
    Grid1.Row = Grid1.Row + 1
    End If
End If
'If selected_procedure = "Stock_item_ledger_edit" Then
Grid1.TextMatrix(Grid1.Row, 6) = Format(Val(Grid1.TextMatrix(Grid1.Row, 4)) * Val(Grid1.TextMatrix(Grid1.Row, 5)), ".00")
'End If
End Sub

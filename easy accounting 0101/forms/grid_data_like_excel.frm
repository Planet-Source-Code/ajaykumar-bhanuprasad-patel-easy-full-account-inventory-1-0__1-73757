VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form grid_data 
   Caption         =   "Opening Stock Detail"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   8880
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   8760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   615
      Left            =   7200
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8760
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   -500
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5415
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   9551
      _Version        =   393216
      ScrollBars      =   2
   End
   Begin VB.Label lbl_card 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   10695
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
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   10695
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
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   10935
   End
   Begin VB.Label lbl_Heading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Stock Detail"
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
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   10695
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   10695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   10695
   End
End
Attribute VB_Name = "grid_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private keycode_now  As Integer
Private pressed_key As Integer
Private selected_row
Private selected_col
Private pre_row
Private pre_col
Private Sub Form_Load()
lbl_card.Caption = selected_stock_item
lbl_name.Caption = co_name
lbl_add.Caption = co_add1 & ", " & co_add2 & ", " & co_pincode & ", " & co_city & ", " & co_contry
Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")
If selected_path = "" Or selected_path = Null Then
    selected_path = App.Path & "\data\1000\co.mdb;"
End If
Text1.Text = ""
Call set_combo1
Call set_grid1_data
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
'Grid1.Text = Combo1.Text
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
keycode_now = KeyCode
If keycode_now = 37 Then
    Grid1.Text = Combo1.Text
    Grid1.Col = Grid1.Col - 1
ElseIf keycode_now = 39 Or keycode_now = 13 Then
    If Grid1.Col = 9 And Grid1.Row = Grid1.Rows - 1 Then
        Grid1.Text = Combo1.Text
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Col = 2
        Grid1.Row = Grid1.Row + 1
    ElseIf Grid1.Col = 9 And Grid1.Row <> Grid1.Rows - 1 Then
        Grid1.Text = Combo1.Text
        Grid1.Row = Grid1.Row + 1
        Grid1.Col = 2
    End If
End If
End Sub
Private Sub Grid1_Click()
selected_row = Grid1.Row
selected_col = Grid1.Col
If Grid1.Row = (pre_row + 1) And Grid1.Row >= 2 Then
Grid1.TextMatrix(pre_row, 4) = Val(Grid1.TextMatrix(pre_row, 3)) - Val(Grid1.TextMatrix(pre_row, 2))
End If
If selected_row = pre_row And selected_col = pre_col Then
    If selected_row = (pre_row + 1) Then
        Grid1.TextMatrix(pre_row, 5) = Val(Grid1.TextMatrix(pre_row, 4)) - Val(Grid1.TextMatrix(pre_row, 3))
    End If
Else
Text1.Text = ""
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

'    If Grid1.Col > 1 Then
'        Grid1.Col = 9
'        Grid1.Row = Grid1.Row - 1
'    Else
'        Grid1.Col = Grid1.Col - 1
'    End If
ElseIf keycode_now = 39 Or keycode_now = 13 Then
'MsgBox "hello"
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
'        Combo1.Height = Grid1.CellHeight
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
    If Grid1.Col = 2 Or Grid1.Col = 3 Or Grid1.Col = 5 Then
                Grid1.Text = Text1.Text
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
        If Grid1.Col > 2 Then
            Grid1.Col = Grid1.Col - 1
        'ElseIf Grid1.Col = 2 Then
        'MsgBox "hello"
        '    Grid1.Col = 9
        '    Grid1.Row = Grid1.Row - 1
        End If
ElseIf keycode_now = 38 And Grid1.Row >= 1 And Grid1.Col <> 9 Then

If Grid1.Row = 1 Then
MsgBox "Not Valid key.....!!!"
Exit Sub
End If
        Grid1.Row = Grid1.Row - 1
ElseIf keycode_now = 39 Or keycode_now = 13 Then
    If Grid1.Col < 9 Then
        Grid1.Col = Grid1.Col + 1

'    ElseIf Grid1.Col = 5 And Grid1.Row = Grid1.Rows - 1 Then
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
ElseIf keycode_now = 40 And Grid1.Col <> 9 Then
'MsgBox Grid1.Rows
'MsgBox Grid1.Row
    If Grid1.Rows = Grid1.Row + 1 Then
    MsgBox "Not Valid key.....!!!"
    Exit Sub
    Else
    Grid1.Row = Grid1.Row + 1
    End If
End If
End Sub
Public Sub set_grid1_data()
'set data grid
Grid1.RowHeightMin = 400
Grid1.Clear
Grid1.Rows = 2
Grid1.Cols = 11
Grid1.TextMatrix(0, 1) = "Opening stock no"
Grid1.TextMatrix(0, 2) = "Starting Seril No."
Grid1.TextMatrix(0, 3) = "Ending Seril No."
Grid1.TextMatrix(0, 4) = "Quantity"
Grid1.TextMatrix(0, 5) = "Rate"
Grid1.TextMatrix(0, 6) = "Amount"

Grid1.TextMatrix(0, 7) = "Face Value"
Grid1.TextMatrix(0, 8) = "Dis. Rate"
Grid1.TextMatrix(0, 9) = "Dealer name"

Grid1.ColWidth(0) = 500
Grid1.ColWidth(1) = 1700
Grid1.ColWidth(2) = 2500
Grid1.ColWidth(3) = 2500
Grid1.ColWidth(4) = 800
Grid1.ColWidth(5) = 800
Grid1.ColWidth(6) = 1200

Grid1.ColWidth(7) = 800
Grid1.ColWidth(8) = 800
Grid1.ColWidth(9) = 2000

Grid1.ColWidth(10) = 1
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


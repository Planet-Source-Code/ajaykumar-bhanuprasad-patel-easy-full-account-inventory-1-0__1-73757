VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form disp_created_stock_item 
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "dsip_created_stock_item.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_detail_2 
      Caption         =   "Stock 2"
      Height          =   495
      Left            =   10680
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmd_detail_1 
      Caption         =   "Detail"
      Height          =   495
      Left            =   10560
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to Exit."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   10215
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   3585
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6324
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   -2147483637
      ForeColorSel    =   -2147483635
      SelectionMode   =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1320
      Width           =   2505
   End
   Begin VB.Label Label3 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   2505
   End
   Begin VB.Label Label4 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   2505
   End
   Begin VB.Label Label5 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   1320
      Width           =   2505
   End
   Begin VB.Label Label6 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   1680
      Width           =   2505
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
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
      Left            =   3000
      TabIndex        =   13
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
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
      Left            =   3000
      TabIndex        =   12
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
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
      Left            =   3000
      TabIndex        =   11
      Top             =   2040
      Width           =   825
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
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
      Left            =   8040
      TabIndex        =   10
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
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
      Left            =   8040
      TabIndex        =   9
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..............................................................."
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   9000
   End
   Begin VB.Label lbl_Heading 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_heading"
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
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   9000
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   9000
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9000
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
      Left            =   10215
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
End
Attribute VB_Name = "disp_created_stock_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_detail_1_Click()

show_opg_stk_srl_dtl_from_disp_list = 1

selected_stock_item_name = Grid1.TextMatrix(Grid1.Row, 1)
selected_stock_item_type = 1
disp_stock_lgr_detail.Show
End Sub

Private Sub cmd_detail_2_Click()
show_opg_stk_srl_dtl_from_disp_list = 1
selected_stock_item_name = Grid1.TextMatrix(Grid1.Row, 1)
selected_stock_item_type = 2
disp_stock_lgr_detail.Show
End Sub

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
cmd_detail_2.Visible = False

Label2.Caption = "Stock Item"
Label3.Caption = "Group"
Label4.Caption = "Company"
Label5.Caption = "F.Value"
Label6.Caption = "Rate"

Label8.Caption = " "
Label9.Caption = " "
Label10.Caption = " "
Label11.Caption = " "
Label12.Caption = " "

If selected_path = "" Or selected_path = Null Then
    selected_path = App.Path & "\data\1000\co.mdb;"
End If
Call arrange_form
Call arrange_grid1
Call open_database
Call open_rs_lgr_main_grp
Call open_grid1
End Sub
Public Sub arrange_form()
Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)
lbl_Heading.Caption = selected_procedure

lbl_Heading.Caption = "List of Created Stock Item....,"

lbl_name.Caption = co_name
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
'Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")

End Sub
Public Sub arrange_grid1()
    Grid1.RowHeightMin = 400
    Grid1.Clear
    Grid1.Rows = 2
    Grid1.Cols = 14
    
    Grid1.TextMatrix(0, 1) = "Item"
    Grid1.TextMatrix(0, 2) = "Alias"
    Grid1.TextMatrix(0, 3) = "Company"
    Grid1.TextMatrix(0, 4) = "Unit"
    Grid1.TextMatrix(0, 5) = "Gruop"
    Grid1.TextMatrix(0, 6) = "F.Val"
    Grid1.TextMatrix(0, 7) = "Dis."
    Grid1.TextMatrix(0, 8) = "Qty.1"
    Grid1.TextMatrix(0, 9) = "Rate 1"
    Grid1.TextMatrix(0, 10) = "Amount 1"
    
    Grid1.TextMatrix(0, 11) = "Qty.2"
    Grid1.TextMatrix(0, 12) = "Rate 2"
    Grid1.TextMatrix(0, 13) = "Amount 2"
    
    Grid1.ColWidth(0) = 300
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 2200
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 2000
    Grid1.ColWidth(6) = 1000
    
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 1500
    
    Grid1.ColWidth(11) = 1000
    Grid1.ColWidth(12) = 1000
    Grid1.ColWidth(13) = 1500
'    Grid1.ColWidth(14) = 2000
 '   Grid1.ColWidth(15) = 2000
    
    Grid1.Font.Size = 12
    
    'Grid1.Width = Grid1.ColWidth(0) + Grid1.ColWidth(1) + Grid1.ColWidth(2) + Grid1.ColWidth(3) + Grid1.ColWidth(4)

End Sub
Public Sub open_grid1()
Call open_database
Call open_rs_stk_item_lgr
Dim data_no As Integer
data_no = 1
Do Until rs_stk_item_lgr.EOF

Grid1.TextMatrix(data_no, 0) = data_no
Grid1.TextMatrix(data_no, 1) = rs_stk_item_lgr!stk_item_lgr_name
Grid1.TextMatrix(data_no, 2) = rs_stk_item_lgr!stk_item_lgr_alis
Grid1.TextMatrix(data_no, 3) = rs_stk_item_lgr!stk_item_lgr_comp
Grid1.TextMatrix(data_no, 4) = rs_stk_item_lgr!stk_item_lgr_unit
Grid1.TextMatrix(data_no, 5) = rs_stk_item_lgr!stk_item_lgr_grup
Grid1.TextMatrix(data_no, 6) = Format(rs_stk_item_lgr!stk_item_lgr_fcvl, "0.00")
Grid1.TextMatrix(data_no, 7) = rs_stk_item_lgr!stk_item_lgr_disc
Grid1.TextMatrix(data_no, 8) = rs_stk_item_lgr!stk_item_lgr_qnt1
Grid1.TextMatrix(data_no, 9) = Format(rs_stk_item_lgr!stk_item_lgr_rat1, "0.00")
Grid1.TextMatrix(data_no, 10) = Format(rs_stk_item_lgr!stk_item_lgr_amt1, "0.00")
Grid1.TextMatrix(data_no, 11) = rs_stk_item_lgr!stk_item_lgr_qnt2
Grid1.TextMatrix(data_no, 12) = rs_stk_item_lgr!stk_item_lgr_rat2
Grid1.TextMatrix(data_no, 13) = Format(rs_stk_item_lgr!stk_item_lgr_amt2, "0.00")


data_no = data_no + 1
If rs_stk_item_lgr.RecordCount < Grid1.Rows Then
Exit Sub
End If
Grid1.Rows = Grid1.Rows + 1
rs_stk_item_lgr.MoveNext
Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
show_opg_stk_srl_dtl_from_disp_list = 0
End Sub

Private Sub Grid1_Click()

Label8.Caption = Grid1.TextMatrix(Grid1.Row, 1) & " (Short Code Used:" & Grid1.TextMatrix(Grid1.Row, 2) & ")"
Label9.Caption = Grid1.TextMatrix(Grid1.Row, 5)
Label10.Caption = Grid1.TextMatrix(Grid1.Row, 3)
Label11.Caption = Format(Grid1.TextMatrix(Grid1.Row, 6), "0.00")
Label12.Caption = Format(Grid1.TextMatrix(Grid1.Row, 10), "0.00")

cmd_detail_1.Visible = True
'cmd_detail_2.Visible = True
cmd_detail_1.Height = 400
cmd_detail_1.Width = 1000
'cmd_detail_2.Height = 400
'cmd_detail_2.Width = 1000
cmd_detail_1.Top = Grid1.CellTop + Grid1.Top
'cmd_detail_2.Top = Grid1.CellTop + Grid1.Top
End Sub


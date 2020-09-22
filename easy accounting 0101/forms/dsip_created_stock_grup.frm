VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form disp_created_stock_group 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "dsip_created_stock_grup.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Click hete to Exit"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   6360
      Width           =   11055
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4725
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   8334
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "............................................................"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   9000
   End
   Begin VB.Label lbl_Heading 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   9000
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   9000
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   9000
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
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   10695
   End
End
Attribute VB_Name = "disp_created_stock_group"
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
'lbl_Heading.Caption = selected_procedure
lbl_Heading.Caption = "List of Stock Group....,"

lbl_name.Caption = co_name
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
'Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")
End Sub
Public Sub arrange_grid1()
    Grid1.RowHeightMin = 400
    Grid1.Clear
    Grid1.Rows = 2
    Grid1.Cols = 4
    Grid1.TextMatrix(0, 1) = "Stock Group Name"
    Grid1.TextMatrix(0, 2) = "Stock Group Alias"
    Grid1.TextMatrix(0, 3) = "Stock main Group"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 2500
    Grid1.ColWidth(3) = 3800
    'Grid1.Width = Grid1.ColWidth(0) + Grid1.ColWidth(1) + Grid1.ColWidth(2) + Grid1.ColWidth(3) + 500 ' + grid1.ColWidth(4)
    Grid1.Font.Size = 12
End Sub
Public Sub open_grid1()
Call open_database
Call open_rs_stk_item_grp
Dim data_no As Integer
data_no = 1
Do Until rs_stk_item_grp.EOF
Grid1.TextMatrix(data_no, 0) = data_no
Grid1.TextMatrix(data_no, 1) = rs_stk_item_grp!stk_item_grp_name
Grid1.TextMatrix(data_no, 2) = rs_stk_item_grp!stk_item_grp_alis
Grid1.TextMatrix(data_no, 3) = rs_stk_item_grp!stk_item_grp_mgrp
data_no = data_no + 1
If rs_stk_item_grp.RecordCount < Grid1.Rows Then
    Exit Sub
End If
Grid1.Rows = Grid1.Rows + 1
rs_stk_item_grp.MoveNext
Loop
End Sub


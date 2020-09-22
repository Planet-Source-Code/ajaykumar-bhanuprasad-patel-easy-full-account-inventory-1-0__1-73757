VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form disp_created_grup 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "disp_created_grup.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to Exit"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   6600
      Width           =   11175
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4935
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8705
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   12632256
      BackColorSel    =   -2147483645
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..............................."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   11535
   End
   Begin VB.Label lbl_Heading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_heading"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   405
      TabIndex        =   5
      Top             =   720
      Width           =   10995
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   4
      Top             =   360
      Width           =   10995
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name of company"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   400
      TabIndex        =   3
      Top             =   0
      Width           =   11000
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
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   10695
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "disp_created_grup"
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
'Label1.Caption = "Accounting Ledger Group Summery"
Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)

'lbl_Heading.Caption = selected_procedure
lbl_Heading.Caption = "Accounting Ledger Group Summery"

lbl_name.Caption = co_name
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
'Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")
End Sub
Public Sub arrange_grid1()
    Grid1.RowHeightMin = 400
    Grid1.Clear
    Grid1.Rows = 2
    Grid1.Cols = 5
    Grid1.TextMatrix(0, 1) = "Group Name"
    Grid1.TextMatrix(0, 2) = "Alias"
    Grid1.TextMatrix(0, 3) = "main Group"
    Grid1.TextMatrix(0, 4) = "Primary Group"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 3200
    Grid1.ColWidth(2) = 1500
    Grid1.ColWidth(3) = 2700
    Grid1.ColWidth(4) = 2700
    Grid1.Font.Size = 12
End Sub
Public Sub open_grid1()
Dim data_no As Integer
data_no = 1

Call open_database
Call open_rs_lgr_prim_grp

Do Until rs_lgr_prim_grp.EOF
Grid1.TextMatrix(data_no, 0) = data_no
Grid1.TextMatrix(data_no, 1) = rs_lgr_prim_grp!lgr_prim_grp_name
'grid1.TextMatrix(data_no, 2) = rs_lgr_prim_grp!lgr_main_grp_alis
Grid1.TextMatrix(data_no, 3) = "Primary Group"
Grid1.TextMatrix(data_no, 4) = rs_lgr_prim_grp!lgr_prim_grp_bgrp
data_no = data_no + 1
If rs_lgr_prim_grp.RecordCount < Grid1.Rows Then
Exit Sub
End If
Grid1.Rows = Grid1.Rows + 1
rs_lgr_prim_grp.MoveNext
Loop

Call open_database
Call open_rs_lgr_main_grp

Do Until rs_lgr_main_grp.EOF
Grid1.TextMatrix(data_no, 0) = data_no
Grid1.TextMatrix(data_no, 1) = rs_lgr_main_grp!lgr_main_grp_name
Grid1.TextMatrix(data_no, 2) = rs_lgr_main_grp!lgr_main_grp_alis
Grid1.TextMatrix(data_no, 3) = rs_lgr_main_grp!lgr_main_grp_sgrp
Grid1.TextMatrix(data_no, 4) = rs_lgr_main_grp!lgr_main_grp_pgrp
data_no = data_no + 1
If rs_lgr_main_grp.RecordCount < Grid1.Rows Then
Exit Sub
End If
Grid1.Rows = Grid1.Rows + 1
rs_lgr_main_grp.MoveNext
Loop
End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form disp_created 
   Caption         =   "Form1"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   13635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   12000
      TabIndex        =   1
      Top             =   9960
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   11033
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   12615
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
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   12615
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
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Width           =   12375
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
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   12615
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
      Height          =   2055
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "disp_created"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
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
Me.Caption = selected_procedure
lbl_Heading.Caption = selected_procedure
lbl_name.Caption = co_name
lbl_add.Caption = co_add1 & ", " & co_add2 & ", " & co_pincode & ", " & co_city & ", " & co_contry
Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")

End Sub
Public Sub arrange_grid1()
    grid1.RowHeightMin = 400
    grid1.Clear
    grid1.Rows = 2
    grid1.Cols = 5
    grid1.TextMatrix(0, 1) = "Group Name"
    grid1.TextMatrix(0, 2) = "Group Alias"
    grid1.TextMatrix(0, 3) = "main Group"
    grid1.TextMatrix(0, 4) = "Primary Group"
'    Grid1.TextMatrix(0, 5) = "Rate"
'    Grid1.TextMatrix(0, 6) = "Amount"
'    Grid1.TextMatrix(0, 7) = "Face Value"
'    Grid1.TextMatrix(0, 8) = "Dis. Rate"
'    Grid1.TextMatrix(0, 9) = "Dealer name"
    grid1.ColWidth(0) = 500
    grid1.ColWidth(1) = 3000
    grid1.ColWidth(2) = 3000
    grid1.ColWidth(3) = 3000
    grid1.ColWidth(4) = 3000
    grid1.Font.Size = 12
    
    'grid1.Width = grid1.ColWidth(0) + grid1.ColWidth(1) + grid1.ColWidth(2) + grid1.ColWidth(3) + grid1.ColWidth(4)

'    Grid1.ColWidth(5) = 800
'    Grid1.ColWidth(6) = 1200
'    Grid1.ColWidth(7) = 800
'    Grid1.ColWidth(8) = 800
'    Grid1.ColWidth(9) = 2000
'    Grid1.ColWidth(10) = 1
End Sub
Public Sub open_grid1()

'grid1.Row = 1
Call open_database
Call open_rs_lgr_main_grp
Dim data_no As Integer
data_no = 1
'For data_no = 1 To rs_lgr_main_grp.EOF
Do Until rs_lgr_main_grp.EOF
'MsgBox rs_lgr_main_grp!lgr_main_grp_name

grid1.TextMatrix(data_no, 0) = data_no
grid1.TextMatrix(data_no, 1) = rs_lgr_main_grp!lgr_main_grp_name
grid1.TextMatrix(data_no, 2) = rs_lgr_main_grp!lgr_main_grp_alis
grid1.TextMatrix(data_no, 3) = rs_lgr_main_grp!lgr_main_grp_sgrp
grid1.TextMatrix(data_no, 4) = rs_lgr_main_grp!lgr_main_grp_pgrp
data_no = data_no + 1
If rs_lgr_main_grp.RecordCount < grid1.Rows Then
Exit Sub
End If
grid1.Rows = grid1.Rows + 1
'MsgBox rs_lgr_main_grp.RecordCount
rs_lgr_main_grp.MoveNext
Loop
'Next
End Sub


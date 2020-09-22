VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form disp_created_ledger 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "disp_created_ledger.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to Exit"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   6720
      Width           =   11535
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   3975
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7011
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   -2147483637
      ForeColorSel    =   -2147483632
      SelectionMode   =   1
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   16
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   15
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   14
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   13
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   12
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2280
      Width           =   2505
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   2505
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Width           =   2505
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1200
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   2505
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   -120
      Width           =   10200
   End
   Begin VB.Label lbl_Heading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_heading"
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
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   10440
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
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   10560
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name of company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   10440
   End
   Begin VB.Label lbl_card 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      TabIndex        =   2
      Top             =   2280
      Width           =   2505
   End
End
Attribute VB_Name = "disp_created_ledger"
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
Label2.Caption = "Ledger Name"
Label3.Caption = "Address"
Label4.Caption = "Travel info"
Label5.Caption = "Contact No"
Label6.Caption = "Cr-Limit"
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
'lbl_Heading.Caption = selected_procedure
lbl_Heading.Caption = "List of Ledger Account...,"
lbl_name.Caption = co_name
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
'Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")
End Sub
Public Sub arrange_grid11()
    Grid1.RowHeightMin = 400
    Grid1.Clear
    Grid1.Rows = 2
    Grid1.Cols = 18 '20
    Grid1.Font.Size = 10
    
    Grid1.TextMatrix(0, 1) = "Name"
    Grid1.TextMatrix(0, 2) = "Alias"
    Grid1.TextMatrix(0, 3) = "Address"
    Grid1.TextMatrix(0, 4) = "Area"
    Grid1.TextMatrix(0, 5) = "city"
    Grid1.TextMatrix(0, 6) = "pin code"
    Grid1.TextMatrix(0, 7) = "transport"
    Grid1.TextMatrix(0, 8) = "Tel.1"
    Grid1.TextMatrix(0, 9) = "Tel.2"
    Grid1.TextMatrix(0, 10) = "Mobile"
    Grid1.TextMatrix(0, 11) = "E-mail"
    Grid1.TextMatrix(0, 12) = "Gruop"
    Grid1.TextMatrix(0, 13) = "Cr.Period"
    Grid1.TextMatrix(0, 14) = "Cr.Amount"
    Grid1.TextMatrix(0, 15) = "Opg.Bal."
    Grid1.TextMatrix(0, 16) = "Cr/Dr"
    'Grid1.TextMatrix(0, 17) = "Opg.Bal.2"
    'Grid1.TextMatrix(0, 18) = "Cr/Dr"
    Grid1.TextMatrix(0, 17) = "Sale.Under"
    
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 2500
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 2000
    Grid1.ColWidth(4) = 2000
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 700
    Grid1.ColWidth(7) = 3000
    Grid1.ColWidth(8) = 1500
    Grid1.ColWidth(9) = 1500
    Grid1.ColWidth(10) = 1500
    Grid1.ColWidth(11) = 1500
    Grid1.ColWidth(12) = 2000
    Grid1.ColWidth(13) = 1000
    Grid1.ColWidth(14) = 1000
    Grid1.ColWidth(15) = 1000
'    Grid1.ColWidth(16) = 500
'    Grid1.ColWidth(17) = 1000
    Grid1.ColWidth(16) = 500
'    grid1.ColWidth(19) = 1000
    
    Grid1.Font.Size = 10
    'grid1.Width = grid1.ColWidth(0) + grid1.ColWidth(1) + grid1.ColWidth(2) + grid1.ColWidth(3) + grid1.ColWidth(4)
End Sub
Public Sub arrange_grid1()
    Grid1.RowHeightMin = 400
    Grid1.Clear
    Grid1.Rows = 2
    Grid1.Cols = 18 '20
    Grid1.Font.Size = 10
    
    Grid1.TextMatrix(0, 1) = "Name"
    Grid1.TextMatrix(0, 2) = "Alias"
    Grid1.TextMatrix(0, 3) = "city"
    Grid1.TextMatrix(0, 4) = "Gruop"
    Grid1.TextMatrix(0, 5) = "Opg.Bal."
    Grid1.TextMatrix(0, 6) = "Cr/Dr"
    Grid1.TextMatrix(0, 7) = "Address"
    Grid1.TextMatrix(0, 8) = "Area"
    
    Grid1.TextMatrix(0, 9) = "pin code"
    Grid1.TextMatrix(0, 10) = "transport"
    Grid1.TextMatrix(0, 11) = "Tel.1"
    Grid1.TextMatrix(0, 12) = "Tel.2"
    Grid1.TextMatrix(0, 13) = "Mobile"
    Grid1.TextMatrix(0, 14) = "E-mail"
    
    Grid1.TextMatrix(0, 15) = "Cr.Period"
    Grid1.TextMatrix(0, 16) = "Cr.Amount"
    'Grid1.TextMatrix(0, 17) = "Opg.Bal.2"
    'Grid1.TextMatrix(0, 18) = "Cr/Dr"
    Grid1.TextMatrix(0, 17) = "Sale.Under"
    
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 2500
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1500
    Grid1.ColWidth(4) = 2000
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 700
    Grid1.ColWidth(7) = 3000
    Grid1.ColWidth(8) = 1500
    Grid1.ColWidth(9) = 1500
    Grid1.ColWidth(10) = 1500
    Grid1.ColWidth(11) = 1500
    Grid1.ColWidth(12) = 2000
    Grid1.ColWidth(13) = 1000
    Grid1.ColWidth(14) = 1000
    Grid1.ColWidth(15) = 1000
'    Grid1.ColWidth(16) = 500
'    Grid1.ColWidth(17) = 1000
    Grid1.ColWidth(16) = 500
'    grid1.ColWidth(19) = 1000
    
    Grid1.Font.Size = 10
    'grid1.Width = grid1.ColWidth(0) + grid1.ColWidth(1) + grid1.ColWidth(2) + grid1.ColWidth(3) + grid1.ColWidth(4)
End Sub
Public Sub open_grid1()
Call open_database
Call open_rs_lgr_main_dtl
Dim data_no As Integer
data_no = 1
Do Until rs_lgr_main_dtl.EOF
Grid1.TextMatrix(data_no, 0) = data_no
Grid1.TextMatrix(data_no, 1) = rs_lgr_main_dtl!lgr_main_dtl_name
If rs_lgr_main_dtl!lgr_main_dtl_alis <> "" Then Grid1.TextMatrix(data_no, 2) = rs_lgr_main_dtl!lgr_main_dtl_alis
If rs_lgr_main_dtl!lgr_main_dtl_add1 <> "" Then Grid1.TextMatrix(data_no, 7) = rs_lgr_main_dtl!lgr_main_dtl_add1
If rs_lgr_main_dtl!lgr_main_dtl_add2 <> "" Then Grid1.TextMatrix(data_no, 8) = rs_lgr_main_dtl!lgr_main_dtl_add2
If rs_lgr_main_dtl!lgr_main_dtl_city <> "" Then Grid1.TextMatrix(data_no, 3) = rs_lgr_main_dtl!lgr_main_dtl_city
If rs_lgr_main_dtl!lgr_main_dtl_pncd <> "" Then Grid1.TextMatrix(data_no, 9) = rs_lgr_main_dtl!lgr_main_dtl_pncd
If rs_lgr_main_dtl!lgr_main_dtl_trnp <> "" Then Grid1.TextMatrix(data_no, 10) = rs_lgr_main_dtl!lgr_main_dtl_trnp
If rs_lgr_main_dtl!lgr_main_dtl_tel1 <> "" Then Grid1.TextMatrix(data_no, 11) = rs_lgr_main_dtl!lgr_main_dtl_tel1
If rs_lgr_main_dtl!lgr_main_dtl_tel2 <> "" Then Grid1.TextMatrix(data_no, 12) = rs_lgr_main_dtl!lgr_main_dtl_tel2
If rs_lgr_main_dtl!lgr_main_dtl_mobl <> "" Then Grid1.TextMatrix(data_no, 13) = rs_lgr_main_dtl!lgr_main_dtl_mobl
If rs_lgr_main_dtl!lgr_main_dtl_emal <> "" Then Grid1.TextMatrix(data_no, 14) = rs_lgr_main_dtl!lgr_main_dtl_emal
If rs_lgr_main_dtl!lgr_main_dtl_grup <> "" Then Grid1.TextMatrix(data_no, 4) = rs_lgr_main_dtl!lgr_main_dtl_grup
If rs_lgr_main_dtl!lgr_main_dtl_crpd <> "" Then Grid1.TextMatrix(data_no, 15) = rs_lgr_main_dtl!lgr_main_dtl_crpd
If rs_lgr_main_dtl!lgr_main_dtl_cram <> "" Then Grid1.TextMatrix(data_no, 16) = rs_lgr_main_dtl!lgr_main_dtl_cram
If rs_lgr_main_dtl!lgr_main_dtl_obl1 <> "" Then Grid1.TextMatrix(data_no, 5) = rs_lgr_main_dtl!lgr_main_dtl_obl1
If rs_lgr_main_dtl!lgr_main_dtl_osd1 <> "" Then Grid1.TextMatrix(data_no, 6) = rs_lgr_main_dtl!lgr_main_dtl_osd1
'MsgBox rs_lgr_main_dtl!lgr_main_dtl_obl2
'Grid1.TextMatrix(data_no, 17) = rs_lgr_main_dtl!lgr_main_dtl_obl2
If rs_lgr_main_dtl!lgr_main_dtl_slun <> "" Then
'Grid1.TextMatrix(data_no, 19) = rs_lgr_main_dtl!lgr_main_dtl_slun
End If
'grid1.TextMatrix(data_no, 18) = rs_lgr_main_dtl!lgr_main_dtl_osd2
'If Val(rs_lgr_main_dtl!lgr_main_dtl_cram) = 0 Then grid1.TextMatrix(data_no, 14) = rs_lgr_main_dtl!lgr_main_dtl_cram
'If Val(rs_lgr_main_dtl!lgr_main_dtl_obl1) = 0 Then grid1.TextMatrix(data_no, 15) = rs_lgr_main_dtl!lgr_main_dtl_obl1
'If Val(rs_lgr_main_dtl!lgr_main_dtl_osd1) = 0 Then grid1.TextMatrix(data_no, 16) = rs_lgr_main_dtl!lgr_main_dtl_osd1
'If rs_lgr_main_dtl!lgr_main_dtl_obl2 <> Null Then grid1.TextMatrix(data_no, 17) = rs_lgr_main_dtl!lgr_main_dtl_obl2
'If rs_lgr_main_dtl!lgr_main_dtl_osd2 <> "" Then Grid1.TextMatrix(data_no, 18) = rs_lgr_main_dtl!lgr_main_dtl_osd2
data_no = data_no + 1
If rs_lgr_main_dtl.RecordCount < Grid1.Rows Then
Exit Sub
End If
Grid1.Rows = Grid1.Rows + 1
rs_lgr_main_dtl.MoveNext
Loop
End Sub

Public Sub open_grid11()
Call open_database
Call open_rs_lgr_main_dtl
Dim data_no As Integer
data_no = 1
Do Until rs_lgr_main_dtl.EOF
Grid1.TextMatrix(data_no, 0) = data_no
Grid1.TextMatrix(data_no, 1) = rs_lgr_main_dtl!lgr_main_dtl_name
If rs_lgr_main_dtl!lgr_main_dtl_alis <> "" Then Grid1.TextMatrix(data_no, 2) = rs_lgr_main_dtl!lgr_main_dtl_alis
If rs_lgr_main_dtl!lgr_main_dtl_add1 <> "" Then Grid1.TextMatrix(data_no, 3) = rs_lgr_main_dtl!lgr_main_dtl_add1
If rs_lgr_main_dtl!lgr_main_dtl_add2 <> "" Then Grid1.TextMatrix(data_no, 4) = rs_lgr_main_dtl!lgr_main_dtl_add2
If rs_lgr_main_dtl!lgr_main_dtl_city <> "" Then Grid1.TextMatrix(data_no, 5) = rs_lgr_main_dtl!lgr_main_dtl_city
If rs_lgr_main_dtl!lgr_main_dtl_pncd <> "" Then Grid1.TextMatrix(data_no, 6) = rs_lgr_main_dtl!lgr_main_dtl_pncd
If rs_lgr_main_dtl!lgr_main_dtl_trnp <> "" Then Grid1.TextMatrix(data_no, 7) = rs_lgr_main_dtl!lgr_main_dtl_trnp
If rs_lgr_main_dtl!lgr_main_dtl_tel1 <> "" Then Grid1.TextMatrix(data_no, 8) = rs_lgr_main_dtl!lgr_main_dtl_tel1
If rs_lgr_main_dtl!lgr_main_dtl_tel2 <> "" Then Grid1.TextMatrix(data_no, 9) = rs_lgr_main_dtl!lgr_main_dtl_tel2
If rs_lgr_main_dtl!lgr_main_dtl_mobl <> "" Then Grid1.TextMatrix(data_no, 10) = rs_lgr_main_dtl!lgr_main_dtl_mobl
If rs_lgr_main_dtl!lgr_main_dtl_emal <> "" Then Grid1.TextMatrix(data_no, 11) = rs_lgr_main_dtl!lgr_main_dtl_emal
If rs_lgr_main_dtl!lgr_main_dtl_grup <> "" Then Grid1.TextMatrix(data_no, 12) = rs_lgr_main_dtl!lgr_main_dtl_grup
If rs_lgr_main_dtl!lgr_main_dtl_crpd <> "" Then Grid1.TextMatrix(data_no, 13) = rs_lgr_main_dtl!lgr_main_dtl_crpd
If rs_lgr_main_dtl!lgr_main_dtl_cram <> "" Then Grid1.TextMatrix(data_no, 14) = rs_lgr_main_dtl!lgr_main_dtl_cram
If rs_lgr_main_dtl!lgr_main_dtl_obl1 <> "" Then Grid1.TextMatrix(data_no, 15) = rs_lgr_main_dtl!lgr_main_dtl_obl1
If rs_lgr_main_dtl!lgr_main_dtl_osd1 <> "" Then Grid1.TextMatrix(data_no, 16) = rs_lgr_main_dtl!lgr_main_dtl_osd1
'MsgBox rs_lgr_main_dtl!lgr_main_dtl_obl2
'Grid1.TextMatrix(data_no, 17) = rs_lgr_main_dtl!lgr_main_dtl_obl2
If rs_lgr_main_dtl!lgr_main_dtl_slun <> "" Then
'Grid1.TextMatrix(data_no, 19) = rs_lgr_main_dtl!lgr_main_dtl_slun
End If
'grid1.TextMatrix(data_no, 18) = rs_lgr_main_dtl!lgr_main_dtl_osd2
'If Val(rs_lgr_main_dtl!lgr_main_dtl_cram) = 0 Then grid1.TextMatrix(data_no, 14) = rs_lgr_main_dtl!lgr_main_dtl_cram
'If Val(rs_lgr_main_dtl!lgr_main_dtl_obl1) = 0 Then grid1.TextMatrix(data_no, 15) = rs_lgr_main_dtl!lgr_main_dtl_obl1
'If Val(rs_lgr_main_dtl!lgr_main_dtl_osd1) = 0 Then grid1.TextMatrix(data_no, 16) = rs_lgr_main_dtl!lgr_main_dtl_osd1
'If rs_lgr_main_dtl!lgr_main_dtl_obl2 <> Null Then grid1.TextMatrix(data_no, 17) = rs_lgr_main_dtl!lgr_main_dtl_obl2
'If rs_lgr_main_dtl!lgr_main_dtl_osd2 <> "" Then Grid1.TextMatrix(data_no, 18) = rs_lgr_main_dtl!lgr_main_dtl_osd2
data_no = data_no + 1
If rs_lgr_main_dtl.RecordCount < Grid1.Rows Then
Exit Sub
End If
Grid1.Rows = Grid1.Rows + 1
rs_lgr_main_dtl.MoveNext
Loop
End Sub

Private Sub Grid1_Click()
Label8.Caption = ": " & Grid1.TextMatrix(Grid1.Row, 1) & " (" & Grid1.TextMatrix(Grid1.Row, 4) & ")"
'If Grid1.TextMatrix(Grid1.Row, 19) <> "" Then Label8.Caption = Label8.Caption & " (Sale Under : " & Grid1.TextMatrix(Grid1.Row, 19) & " )"
Label9.Caption = ": " & Grid1.TextMatrix(Grid1.Row, 7) & "  " & Grid1.TextMatrix(Grid1.Row, 8) & "  " & Grid1.TextMatrix(Grid1.Row, 3) & "  " & Grid1.TextMatrix(Grid1.Row, 9)
Label10.Caption = ": " & Grid1.TextMatrix(Grid1.Row, 10)
Label11.Caption = ": " & Grid1.TextMatrix(Grid1.Row, 11) & "  " & Grid1.TextMatrix(Grid1.Row, 12) & "  " & Grid1.TextMatrix(Grid1.Row, 13)
Label12.Caption = ": " & Format(Grid1.TextMatrix(Grid1.Row, 16), "0.00")

'Label8.Caption = grid1.TextMatrix(grid1.Row, 1) & " (" & grid1.TextMatrix(grid1.Row, 12) & ")"
'If Grid1.TextMatrix(Grid1.Row, 19) <> "" Then Label8.Caption = Label8.Caption & " (Sale Under : " & Grid1.TextMatrix(Grid1.Row, 19) & " )"
'Label9.Caption = grid1.TextMatrix(grid1.Row, 3) & "  " & grid1.TextMatrix(grid1.Row, 4) & "  " & grid1.TextMatrix(grid1.Row, 5) & "  " & grid1.TextMatrix(grid1.Row, 6)
'Label10.Caption = grid1.TextMatrix(grid1.Row, 7)
'Label11.Caption = grid1.TextMatrix(grid1.Row, 8) & "  " & grid1.TextMatrix(grid1.Row, 9) & "  " & grid1.TextMatrix(grid1.Row, 10)
'Label12.Caption = Format(grid1.TextMatrix(grid1.Row, 14), "0.00")
End Sub


VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form emp_detail_repo2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "employee personal information"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "emp_detail_repo2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Name and Telephone print"
      Height          =   495
      Left            =   9240
      TabIndex        =   5
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "All Detail print"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   6480
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4935
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8705
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   720
      X2              =   9120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   9120
      Y1              =   1200
      Y2              =   1200
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
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   7335
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
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   7095
   End
   Begin VB.Label lbl_head 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   8415
   End
End
Attribute VB_Name = "emp_detail_repo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set DataReport3.DataSource = rs_emp_main_dtl.DataSource
DataReports3.Show
End Sub

Private Sub Command2_Click()
Set DataReport2.DataSource = rs_emp_main_dtl.DataSource
DataReport2.Show
End Sub

Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================
Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)
lbl_name.Caption = co_name
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country

'Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")
If selected_path = "" Or selected_path = Null Then
    selected_path = App.Path & "\data\1000\co.mdb;"
End If
Call open_database
Call open_rs_emp_main_dtl

Set DataGrid1.DataSource = rs_emp_main_dtl.DataSource

End Sub


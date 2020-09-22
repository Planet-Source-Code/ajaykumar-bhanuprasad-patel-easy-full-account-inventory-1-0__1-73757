VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form da_find_card_detail 
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "da_find_card_detail.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
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
      Height          =   375
      Left            =   6360
      TabIndex        =   20
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   6360
      TabIndex        =   18
      Text            =   "Select a Ref. No."
      Top             =   165
      Width           =   2535
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   6390
      TabIndex        =   16
      Text            =   "Stock Item"
      Top             =   765
      Width           =   2535
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   1560
      TabIndex        =   15
      Text            =   "Period"
      Top             =   1800
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
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
      CurrentDate     =   40149
   End
   Begin VB.ComboBox Combo_field 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1560
      TabIndex        =   9
      Text            =   "what ?"
      Top             =   90
      Width           =   3015
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
      ItemData        =   "da_find_card_detail.frx":1D2A
      Left            =   1560
      List            =   "da_find_card_detail.frx":1D2C
      TabIndex        =   7
      Text            =   "Select a Name"
      Top             =   1305
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4095
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7223
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
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
      Left            =   9240
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Report"
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
      Left            =   9240
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   12600
      Width           =   20655
      _ExtentX        =   36433
      _ExtentY        =   450
      _Version        =   393216
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
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
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
      Left            =   9240
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
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
      CurrentDate     =   40141
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1560
      TabIndex        =   0
      Text            =   "Select a Name"
      Top             =   690
      Width           =   3015
   End
   Begin VB.Label Label08 
      Caption         =   "Serial No"
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
      Left            =   4800
      TabIndex        =   21
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label04 
      Caption         =   "Reference No"
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
      Left            =   4800
      TabIndex        =   19
      Top             =   165
      Width           =   1935
   End
   Begin VB.Label Label05 
      Caption         =   "Name"
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
      Left            =   4800
      TabIndex        =   17
      Top             =   765
      Width           =   1455
   End
   Begin VB.Label Label07 
      Alignment       =   2  'Center
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
      Left            =   6600
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label06 
      Caption         =   "Duration"
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
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label03 
      Caption         =   "Supplier"
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
      Left            =   150
      TabIndex        =   11
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Label Label01 
      Caption         =   "Search By"
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
      Left            =   150
      TabIndex        =   10
      Top             =   165
      Width           =   1575
   End
   Begin VB.Label Label02 
      Caption         =   "Customer"
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
      Left            =   150
      TabIndex        =   3
      Top             =   765
      Width           =   1215
   End
End
Attribute VB_Name = "da_find_card_detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub xxx()
If Combo_field.Text = "Date" Then
    rs_dap_main_dtl_all.Open "SELECT * FROM dpdb1_main_all WHERE dap_main_dtl_date>= " & selected_str_date & " and dap_main_dtl_date <=#" & selected_end_date & "#", dpdb, adOpenDynamic, adLockOptimistic
    Set find_card_report.DataSource = rs_dap_main_dtl_all
    find_card_report.Show
ElseIf Combo_field.Text = "Refrense No." Then
    MsgBox "hello"
    rs_dap_main_dtl_all.Open "SELECT * FROM dpdb1_main_all WHERE ' dap_main_dtl_csrf = " & selected_ref_no & "'", dpdb, adOpenDynamic, adLockOptimistic
    Set find_card_report.DataSource = rs_dap_main_dtl_all
    find_card_report.Show
ElseIf Combo_field.Text = "Customer" Then
    rs_dap_main_dtl_all.Open "SELECT * FROM dpdb1_main_all WHERE 'dap_main_dtl_date>= " & selected_str_date & " and dap_main_dtl_date <= " & selected_end_date & " dap_main_dtl_name =" & Combo1.Text & "'", dpdb, adOpenDynamic, adLockOptimistic
    Set find_card_report.DataSource = rs_dap_main_dtl_all
    find_card_report.Show
End If
End Sub
Public Sub non_visible_all()

If Combo1.Visible = True Then Combo1.Visible = False
If Combo2.Visible = True Then Combo2.Visible = False
If Combo3.Visible = True Then Combo3.Visible = False
If Combo4.Visible = True Then Combo4.Visible = False
If Combo5.Visible = True Then Combo5.Visible = False

If Label08.Visible = True Then Label08.Visible = False
If Text1.Visible = True Then Text1.Visible = False

If Label02.Visible = True Then Label02.Visible = False
If Label03.Visible = True Then Label03.Visible = False
If Label04.Visible = True Then Label04.Visible = False
If Label05.Visible = True Then Label05.Visible = False
If Label06.Visible = True Then Label06.Visible = False
If Label07.Visible = True Then Label07.Visible = False

If DTPicker1.Visible = True Then DTPicker1.Visible = False
If DTPicker2.Visible = True Then DTPicker2.Visible = False

If Command1.Enabled = True Then Command1.Enabled = False
If Command2.Enabled = True Then Command2.Enabled = False

End Sub
Public Sub search_by_date()
'non visible all
Call non_visible_all
Label06.Visible = True
Label07.Visible = True
Combo5.Visible = True
DTPicker1.Visible = True
DTPicker2.Visible = True
Command1.Enabled = True
Command2.Enabled = True

DTPicker1.Value = 0
DTPicker2.Value = 0
'visible period & date combo only
End Sub

Public Sub search_by_ref()
Call non_visible_all
Label04.Visible = True
Combo3.Visible = True
Command1.Enabled = True
Command2.Enabled = True

End Sub
Public Sub search_by_customer()
Call non_visible_all
Call search_by_date
Label02.Visible = True
Combo1.Visible = True
Command1.Enabled = True
Command2.Enabled = True

End Sub
Public Sub search_by_card()
Call non_visible_all
Label08.Visible = True
Text1.Visible = True
Label05.Visible = True
Combo4.Visible = True
Command1.Enabled = True
Command2.Enabled = True

End Sub
Public Sub search_by_suplier()
Call non_visible_all
Call search_by_date
Label03.Visible = True
Combo2.Visible = True
Command1.Enabled = True
Command2.Enabled = True

End Sub


Private Sub Combo_field_Click()
If Combo_field.Text = "Date" Then
    Call search_by_date
ElseIf Combo_field.Text = "Refrense No." Then
    Call search_by_ref
ElseIf Combo_field.Text = "Customer" Then
    Call search_by_customer
ElseIf Combo_field.Text = "card" Then
    Call search_by_card
ElseIf Combo_field.Text = "Supplier" Then
    Call search_by_suplier
End If
End Sub


Private Sub Combo3_Click()
selected_ref_no = Combo3.Text
End Sub

Private Sub Combo5_Click()
Call search_a_period
End Sub

Private Sub Command1_Click()
Call delete_all_temp_data

Call set_grid1

selected_supl_name = Combo2.Text
selected_serial_no = Val(Text1.Text)
selected_card_name = Combo4.Text

selected_str_date = DTPicker1.Value
selected_end_date = DTPicker2.Value

selected_ref_no = Combo3.Text
selected_cust_nm = Combo1.Text

If rs_dap_main_dtl_all.State = 1 Then rs_dap_main_dtl_all.Close

Call open_rs_dap_main_dtl_all
Call open_rs_dap_main_dtl_temp

Do Until rs_dap_main_dtl_all.EOF
If Combo_field.Text = "Date" Then
    If rs_dap_main_dtl_all!dap_main_dtl_date >= selected_str_date And rs_dap_main_dtl_all!dap_main_dtl_date <= selected_end_date Then
        rs_dap_main_dtl_temp.AddNew
        rs_dap_main_dtl_temp!dap_main_dtl_date = rs_dap_main_dtl_all!dap_main_dtl_date
        rs_dap_main_dtl_temp!dap_main_dtl_name = rs_dap_main_dtl_all!dap_main_dtl_name
        rs_dap_main_dtl_temp!dap_main_dtl_card = rs_dap_main_dtl_all!dap_main_dtl_card
        rs_dap_main_dtl_temp!dap_main_dtl_stsl = rs_dap_main_dtl_all!dap_main_dtl_stsl
        rs_dap_main_dtl_temp!dap_main_dtl_edsl = rs_dap_main_dtl_all!dap_main_dtl_edsl
        rs_dap_main_dtl_temp!dap_main_dtl_qnty = rs_dap_main_dtl_all!dap_main_dtl_qnty
        rs_dap_main_dtl_temp!dap_main_dtl_rate = rs_dap_main_dtl_all!dap_main_dtl_rate
        rs_dap_main_dtl_temp!dap_main_dtl_amnt = rs_dap_main_dtl_all!dap_main_dtl_amnt
        rs_dap_main_dtl_temp!dap_main_dtl_comp = rs_dap_main_dtl_all!dap_main_dtl_comp
        rs_dap_main_dtl_temp!dap_main_dtl_splr = rs_dap_main_dtl_all!dap_main_dtl_splr
        rs_dap_main_dtl_temp!dap_main_dtl_csrf = rs_dap_main_dtl_all!dap_main_dtl_csrf
        rs_dap_main_dtl_temp!dap_main_dtl_user = user_name
        rs_dap_main_dtl_temp.UpdateBatch
    End If
ElseIf Combo_field.Text = "Refrense No." Then
    If rs_dap_main_dtl_all!dap_main_dtl_csrf = selected_ref_no Then
        rs_dap_main_dtl_temp.AddNew
        rs_dap_main_dtl_temp!dap_main_dtl_date = rs_dap_main_dtl_all!dap_main_dtl_date
        rs_dap_main_dtl_temp!dap_main_dtl_name = rs_dap_main_dtl_all!dap_main_dtl_name
        rs_dap_main_dtl_temp!dap_main_dtl_card = rs_dap_main_dtl_all!dap_main_dtl_card
        rs_dap_main_dtl_temp!dap_main_dtl_stsl = rs_dap_main_dtl_all!dap_main_dtl_stsl
        rs_dap_main_dtl_temp!dap_main_dtl_edsl = rs_dap_main_dtl_all!dap_main_dtl_edsl
        rs_dap_main_dtl_temp!dap_main_dtl_qnty = rs_dap_main_dtl_all!dap_main_dtl_qnty
        rs_dap_main_dtl_temp!dap_main_dtl_rate = rs_dap_main_dtl_all!dap_main_dtl_rate
        rs_dap_main_dtl_temp!dap_main_dtl_amnt = rs_dap_main_dtl_all!dap_main_dtl_amnt
        rs_dap_main_dtl_temp!dap_main_dtl_comp = rs_dap_main_dtl_all!dap_main_dtl_comp
        rs_dap_main_dtl_temp!dap_main_dtl_splr = rs_dap_main_dtl_all!dap_main_dtl_splr
        rs_dap_main_dtl_temp!dap_main_dtl_csrf = rs_dap_main_dtl_all!dap_main_dtl_csrf
        rs_dap_main_dtl_temp!dap_main_dtl_user = user_name
        rs_dap_main_dtl_temp.UpdateBatch
    End If

ElseIf Combo_field.Text = "Customer" Then
    If rs_dap_main_dtl_all!dap_main_dtl_name = selected_cust_nm And rs_dap_main_dtl_all!dap_main_dtl_date >= selected_str_date And rs_dap_main_dtl_all!dap_main_dtl_date <= selected_end_date Then
        rs_dap_main_dtl_temp.AddNew
        rs_dap_main_dtl_temp!dap_main_dtl_date = rs_dap_main_dtl_all!dap_main_dtl_date
        rs_dap_main_dtl_temp!dap_main_dtl_name = rs_dap_main_dtl_all!dap_main_dtl_name
        rs_dap_main_dtl_temp!dap_main_dtl_card = rs_dap_main_dtl_all!dap_main_dtl_card
        rs_dap_main_dtl_temp!dap_main_dtl_stsl = rs_dap_main_dtl_all!dap_main_dtl_stsl
        rs_dap_main_dtl_temp!dap_main_dtl_edsl = rs_dap_main_dtl_all!dap_main_dtl_edsl
        rs_dap_main_dtl_temp!dap_main_dtl_qnty = rs_dap_main_dtl_all!dap_main_dtl_qnty
        rs_dap_main_dtl_temp!dap_main_dtl_rate = rs_dap_main_dtl_all!dap_main_dtl_rate
        rs_dap_main_dtl_temp!dap_main_dtl_amnt = rs_dap_main_dtl_all!dap_main_dtl_amnt
        rs_dap_main_dtl_temp!dap_main_dtl_comp = rs_dap_main_dtl_all!dap_main_dtl_comp
        rs_dap_main_dtl_temp!dap_main_dtl_splr = rs_dap_main_dtl_all!dap_main_dtl_splr
        rs_dap_main_dtl_temp!dap_main_dtl_csrf = rs_dap_main_dtl_all!dap_main_dtl_csrf
        rs_dap_main_dtl_temp!dap_main_dtl_user = user_name
        rs_dap_main_dtl_temp.UpdateBatch
    
    End If
ElseIf Combo_field.Text = "Supplier" Then
    If rs_dap_main_dtl_all!dap_main_dtl_splr = selected_supl_name And rs_dap_main_dtl_all!dap_main_dtl_date >= selected_str_date And rs_dap_main_dtl_all!dap_main_dtl_date <= selected_end_date Then
        rs_dap_main_dtl_temp.AddNew
        rs_dap_main_dtl_temp!dap_main_dtl_date = rs_dap_main_dtl_all!dap_main_dtl_date
        rs_dap_main_dtl_temp!dap_main_dtl_name = rs_dap_main_dtl_all!dap_main_dtl_name
        rs_dap_main_dtl_temp!dap_main_dtl_card = rs_dap_main_dtl_all!dap_main_dtl_card
        rs_dap_main_dtl_temp!dap_main_dtl_stsl = rs_dap_main_dtl_all!dap_main_dtl_stsl
        rs_dap_main_dtl_temp!dap_main_dtl_edsl = rs_dap_main_dtl_all!dap_main_dtl_edsl
        rs_dap_main_dtl_temp!dap_main_dtl_qnty = rs_dap_main_dtl_all!dap_main_dtl_qnty
        rs_dap_main_dtl_temp!dap_main_dtl_rate = rs_dap_main_dtl_all!dap_main_dtl_rate
        rs_dap_main_dtl_temp!dap_main_dtl_amnt = rs_dap_main_dtl_all!dap_main_dtl_amnt
        rs_dap_main_dtl_temp!dap_main_dtl_comp = rs_dap_main_dtl_all!dap_main_dtl_comp
        rs_dap_main_dtl_temp!dap_main_dtl_splr = rs_dap_main_dtl_all!dap_main_dtl_splr
        rs_dap_main_dtl_temp!dap_main_dtl_csrf = rs_dap_main_dtl_all!dap_main_dtl_csrf
        rs_dap_main_dtl_temp!dap_main_dtl_user = user_name
        rs_dap_main_dtl_temp.UpdateBatch
    
    End If


ElseIf Combo_field.Text = "card" Then
    If rs_dap_main_dtl_all!dap_main_dtl_card = selected_card_name And Val(rs_dap_main_dtl_all!dap_main_dtl_stsl) <= selected_serial_no And Val(rs_dap_main_dtl_all!dap_main_dtl_edsl) >= selected_serial_no Then
        rs_dap_main_dtl_temp.AddNew
        rs_dap_main_dtl_temp!dap_main_dtl_date = rs_dap_main_dtl_all!dap_main_dtl_date
        rs_dap_main_dtl_temp!dap_main_dtl_name = rs_dap_main_dtl_all!dap_main_dtl_name
        rs_dap_main_dtl_temp!dap_main_dtl_card = rs_dap_main_dtl_all!dap_main_dtl_card
        rs_dap_main_dtl_temp!dap_main_dtl_stsl = rs_dap_main_dtl_all!dap_main_dtl_stsl
        rs_dap_main_dtl_temp!dap_main_dtl_edsl = rs_dap_main_dtl_all!dap_main_dtl_edsl
        rs_dap_main_dtl_temp!dap_main_dtl_qnty = rs_dap_main_dtl_all!dap_main_dtl_qnty
        rs_dap_main_dtl_temp!dap_main_dtl_rate = rs_dap_main_dtl_all!dap_main_dtl_rate
        rs_dap_main_dtl_temp!dap_main_dtl_amnt = rs_dap_main_dtl_all!dap_main_dtl_amnt
        rs_dap_main_dtl_temp!dap_main_dtl_comp = rs_dap_main_dtl_all!dap_main_dtl_comp
        rs_dap_main_dtl_temp!dap_main_dtl_splr = rs_dap_main_dtl_all!dap_main_dtl_splr
        rs_dap_main_dtl_temp!dap_main_dtl_csrf = rs_dap_main_dtl_all!dap_main_dtl_csrf
        rs_dap_main_dtl_temp!dap_main_dtl_user = user_name
        rs_dap_main_dtl_temp.UpdateBatch
    
    End If

End If
    rs_dap_main_dtl_all.MoveNext

Loop

Dim aa As Integer
aa = 1

Call open_rs_dap_main_dtl_all
Call open_rs_dap_main_dtl_temp
If rs_dap_main_dtl_temp.State <> 1 Then rs_dap_main_dtl_temp.MoveFirst
Do Until rs_dap_main_dtl_temp.EOF
        Grid1.AddItem aa
        Grid1.TextMatrix(aa, 1) = rs_dap_main_dtl_temp!dap_main_dtl_date
        Grid1.TextMatrix(aa, 2) = rs_dap_main_dtl_temp!dap_main_dtl_name
        Grid1.TextMatrix(aa, 3) = rs_dap_main_dtl_temp!dap_main_dtl_card
        Grid1.TextMatrix(aa, 4) = rs_dap_main_dtl_temp!dap_main_dtl_stsl
        Grid1.TextMatrix(aa, 5) = rs_dap_main_dtl_temp!dap_main_dtl_edsl
        Grid1.TextMatrix(aa, 6) = rs_dap_main_dtl_temp!dap_main_dtl_qnty
        Grid1.TextMatrix(aa, 7) = rs_dap_main_dtl_temp!dap_main_dtl_rate
        Grid1.TextMatrix(aa, 8) = rs_dap_main_dtl_temp!dap_main_dtl_amnt
        Grid1.TextMatrix(aa, 9) = rs_dap_main_dtl_temp!dap_main_dtl_comp
        Grid1.TextMatrix(aa, 10) = rs_dap_main_dtl_temp!dap_main_dtl_splr
        Grid1.TextMatrix(aa, 11) = rs_dap_main_dtl_temp!dap_main_dtl_csrf
        Grid1.TextMatrix(aa, 12) = rs_dap_main_dtl_temp!dap_main_dtl_user
    aa = aa + 1
rs_dap_main_dtl_temp.MoveNext
Loop
End Sub

Public Sub search_a_period()
Dim today_day As Integer
Dim today_weekday As Integer

today_weekday = Weekday(Now)
today_day = Day(Now) - 1

If Combo5.Text = "This Week" Then
    DTPicker1.Value = Date - (today_weekday + 1)
    DTPicker2.Value = Date
ElseIf Combo5.Text = "This Month" Then
    DTPicker1.Value = Date - today_day
    DTPicker2.Value = Date
ElseIf Combo5.Text = "Last Month" Then
    If Month(Now) = 1 Then
    DTPicker1.Value = Day(Now) - today_day & "/12/" & Year(Now)
    Else
        DTPicker1.Value = Day(Now) - today_day & "/" & Month(Now) - 1 & "/" & Year(Now)
    End If
    DTPicker2.Value = Date - (today_day + 1)
ElseIf Combo5.Text = "Last Week" Then
    DTPicker1.Value = Date - (today_weekday + 5)
    DTPicker2.Value = Date - (today_weekday - 1)
End If

End Sub
Public Sub delete_all_temp_data()
'deleting all the data from temp db file
    
Call open_rs_dap_main_dtl_temp
Do Until rs_dap_main_dtl_temp.EOF
    rs_dap_main_dtl_temp.Delete
    rs_dap_main_dtl_temp.MoveNext
Loop

End Sub

Private Sub Command2_Click()

Call open_rs_dap_main_dtl_temp
temparary_report_heading = ""

selected_supl_name = Combo2.Text
selected_serial_no = Val(Text1.Text)
selected_card_name = Combo4.Text

selected_str_date = DTPicker1.Value
selected_end_date = DTPicker2.Value

selected_ref_no = Combo3.Text
selected_cust_nm = Combo1.Text

If Combo_field.Text = "Date" Then
temparary_report_heading = "De-Activation Report from :  " & selected_str_date & "   to   " & selected_end_date
ElseIf Combo_field.Text = "Refrense No." Then
temparary_report_heading = "De-Activation Report of Ref.No.:  " & selected_ref_no
ElseIf Combo_field.Text = "Customer" Then
temparary_report_heading = "De-Activation Report of Customter:  " & selected_cust_nm & "    Period : " & selected_str_date & "   to   " & selected_end_date
ElseIf Combo_field.Text = "Supplier" Then
temparary_report_heading = "De-Activation Report of Supplier: " & selected_supl_name & "    Period : " & selected_str_date & "   to   " & selected_end_date
ElseIf Combo_field.Text = "card" Then
temparary_report_heading = "De-Activation Report of card:  " & selected_card_name '& "    Period :" & selected_str_date & "   to   " & selected_end_date
End If

With da_find_card_report.Sections("section4").Controls
    .item("label14").Caption = temparary_report_heading
End With

Set da_find_card_report.DataSource = rs_dap_main_dtl_temp
da_find_card_report.Show

End Sub

Private Sub temp_x()

Combo_field.Width = X
Grid1.Width = X
Combo1.Visible = False
Combo2.Visible = False
Combo3.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Label02.Visible = False
Label03.Visible = False
Label04.Visible = False
Label05.Visible = False
Label06.Visible = False
Label07.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Command1.Enabled = False
Command2.Enabled = False
Label08.Visible = False
Text1.Visible = False

End Sub
Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================
'Me.Icon = LoadPicture(App.Path & "\L.ico")
Me.Caption = "Ajay patel's card Deactivation...!!!  " & user_name
Call delete_all_temp_data
Call set_grid1
Call set_combo_field

Combo1.Visible = False
Combo2.Visible = False
Combo3.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Label02.Visible = False
Label03.Visible = False
Label04.Visible = False
Label05.Visible = False
Label06.Visible = False
Label07.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Command1.Enabled = False
Command2.Enabled = False
Label08.Visible = False
Text1.Visible = False

Combo5.AddItem "This Month"
Combo5.AddItem "This Week"
Combo5.AddItem "Last Month"
Combo5.AddItem "Last Week"
'add refrense no to the combo
Call open_database
Call open_rs_dap_main_dtl_all

Do Until rs_dap_main_dtl_all.EOF
    Combo3.AddItem rs_dap_main_dtl_all!dap_main_dtl_csrf
    rs_dap_main_dtl_all.MoveNext
Loop
Close All

'add customer from ledger to the combo
Combo1.Clear
Call open_database
Call open_rs_lgr_main_dtl

Do Until rs_lgr_main_dtl.EOF
selected_group = rs_lgr_main_dtl!lgr_main_dtl_grup 'Combo1.Text
selected_primary_group = ""
        Call open_rs_lgr_main_grp
        Do Until rs_lgr_main_grp.EOF
            If selected_group = rs_lgr_main_grp!lgr_main_grp_name Or selected_group = rs_lgr_main_grp!lgr_main_grp_alis Then
            selected_primary_group = rs_lgr_main_grp!lgr_main_grp_pgrp
            End If
            rs_lgr_main_grp.MoveNext
        Loop
        
        If selected_primary_group = "" Then
            Call open_rs_lgr_prim_grp
            If rs_lgr_prim_grp.RecordCount > 0 Then rs_lgr_prim_grp.MoveFirst
            Do Until rs_lgr_prim_grp.EOF
            If selected_group = rs_lgr_prim_grp!lgr_prim_grp_name Then
            selected_primary_group = rs_lgr_prim_grp!lgr_prim_grp_name
            End If
            rs_lgr_prim_grp.MoveNext
            Loop
        End If
        
        If LCase(selected_primary_group) = LCase("Sundry Debtors") Then ' if the created ledger is a debtor then
            Combo1.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
        End If
rs_lgr_main_dtl.MoveNext
Loop
Combo1.Text = "Select Customer..,"

'add suppliers

Combo2.Clear
Call open_database
Call open_rs_lgr_main_dtl

Do Until rs_lgr_main_dtl.EOF
selected_group = rs_lgr_main_dtl!lgr_main_dtl_grup 'Combo1.Text
selected_primary_group = ""
        Call open_rs_lgr_main_grp
        Do Until rs_lgr_main_grp.EOF
            If selected_group = rs_lgr_main_grp!lgr_main_grp_name Or selected_group = rs_lgr_main_grp!lgr_main_grp_alis Then
            selected_primary_group = rs_lgr_main_grp!lgr_main_grp_pgrp
            End If
            rs_lgr_main_grp.MoveNext
        Loop
        
        If selected_primary_group = "" Then
            Call open_rs_lgr_prim_grp
            If rs_lgr_prim_grp.RecordCount > 0 Then rs_lgr_prim_grp.MoveFirst
            Do Until rs_lgr_prim_grp.EOF
            If selected_group = rs_lgr_prim_grp!lgr_prim_grp_name Then
            selected_primary_group = rs_lgr_prim_grp!lgr_prim_grp_name
            End If
            rs_lgr_prim_grp.MoveNext
            Loop
        End If
        
        If LCase(selected_primary_group) = LCase("Sundry creditors") Then ' if the created ledger is a debtor then
            Combo2.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
        End If
rs_lgr_main_dtl.MoveNext
Loop
Combo2.Text = "Select supplier..,"

'add stock items

Call open_database
Call open_rs_stk_item_lgr
Do Until rs_stk_item_lgr.EOF
    Combo4.AddItem rs_stk_item_lgr!stk_item_lgr_name
    If rs_stk_item_lgr!stk_item_lgr_alis <> "" Then Combo4.AddItem rs_stk_item_lgr!stk_item_lgr_alis
rs_stk_item_lgr.MoveNext
Loop

End Sub

Private Sub Command3_Click()
Unload Me
End Sub
Public Sub set_grid1()
        Grid1.Width = Me.Width * 0.9
        Grid1.Left = Me.Width / 25
        Grid1.Height = Me.Height * 0.6
        Grid1.Top = Me.Height * 0.3
        
        Grid1.Clear
        Grid1.Rows = 1
        Grid1.Cols = 13
        
        Grid1.TextMatrix(0, 1) = "Date"
        Grid1.TextMatrix(0, 2) = "Customer Name"
       
        Grid1.TextMatrix(0, 3) = "card Name"
        Grid1.TextMatrix(0, 4) = "Starting Serial No"
        Grid1.TextMatrix(0, 5) = "Ending serial no"
        Grid1.TextMatrix(0, 6) = "Quantity"
        Grid1.TextMatrix(0, 7) = "Rate"
        Grid1.TextMatrix(0, 8) = "Amount"
        Grid1.TextMatrix(0, 9) = "Company"
        Grid1.TextMatrix(0, 10) = "Supplier"
        
        Grid1.TextMatrix(0, 11) = "Customer Ref No"
        'Grid1.TextMatrix(0, 12) = "response Ref No"
        Grid1.TextMatrix(0, 12) = "Enered by"
        
Grid1.CellAlignment = Center

Grid1.ColWidth(0) = 500

Grid1.ColWidth(1) = 1000
Grid1.ColWidth(2) = 3000

Grid1.ColWidth(3) = 1500
Grid1.ColWidth(4) = 2000
Grid1.ColWidth(5) = 2000
Grid1.ColWidth(6) = 1000
Grid1.ColWidth(7) = 500

Grid1.ColWidth(8) = 2500
Grid1.ColWidth(9) = 2500

Grid1.ColWidth(10) = 2500
Grid1.ColWidth(11) = 2000
Grid1.ColWidth(12) = 2500

End Sub
Public Sub set_combo_field()
Combo_field.AddItem "Date"
Combo_field.AddItem "Refrense No."
Combo_field.AddItem "Customer"
Combo_field.AddItem "card"
Combo_field.AddItem "Supplier"
End Sub


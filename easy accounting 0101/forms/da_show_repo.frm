VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form da_show_repo 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   8760
   Icon            =   "da_show_repo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5415
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   9551
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
      Height          =   975
      Left            =   9960
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print or Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16800
      TabIndex        =   8
      Top             =   11880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to Find records"
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
      TabIndex        =   4
      Top             =   720
      Width           =   9735
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
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   120
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
      CurrentDate     =   40141
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   120
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
      Left            =   1320
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Left            =   7920
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "da_show_repo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub delete_all_in_rs_dap_main_dtl_temp()

Close All
Call open_rs_dap_main_dtl_temp
Do Until rs_dap_main_dtl_temp.EOF
rs_dap_main_dtl_temp.Delete
rs_dap_main_dtl_temp.MoveNext
Loop
Close All

End Sub
Private Sub Command1_Click()

Grid1.Clear
If Grid1.Rows > 2 Then
set_grid1_data
End If

Call delete_all_in_rs_dap_main_dtl_temp
Call open_rs_dap_main_dtl_all

If rs_dap_main_dtl_all.RecordCount < 1 Then
    MsgBox "you have no records for..., first you have to enter record & then select the option & find again...!!! Thank's"
    Exit Sub
End If
Close All
selected_cust_name = Combo1.Text
selected_start_date = DTPicker1.Value
selected_end_date = DTPicker2.Value

Call open_rs_dap_main_dtl_all
Call open_rs_dap_main_dtl_temp

'copy main data to temp data of selected customer

'rs_dap_main_dtl_all.MoveFirst
Dim aa As Integer
aa = 1
Call set_grid1_data
Do Until rs_dap_main_dtl_all.EOF
    If LCase(rs_dap_main_dtl_all!dap_main_dtl_name) = LCase(selected_cust_name) And rs_dap_main_dtl_all!dap_main_dtl_date >= selected_start_date And rs_dap_main_dtl_all!dap_main_dtl_date <= selected_end_date Then
        rs_dap_main_dtl_temp.AddNew
        rs_dap_main_dtl_temp!dap_main_dtl_id = aa
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
        'Grid1.TextMatrix(aa, 12) = rs_dap_main_dtl_temp!db1_resp_rf
        Grid1.TextMatrix(aa, 13) = rs_dap_main_dtl_temp!dap_main_dtl_user
'        rs_dap_main_dtl_temp.UpdateBatch
    aa = aa + 1
    End If
    rs_dap_main_dtl_all.MoveNext
Loop

End Sub
Private Sub grid1_fill()
If rs_dap_main_dtl_temp.RecordCount > 0 Then rs_dap_main_dtl_temp.MoveFirst
Dim aa As Integer
aa = 1


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
'        Grid1.TextMatrix(aa, 11) = rs_dap_main_dtl_temp!dap_main_dtl_csrf
'        Grid1.TextMatrix(aa, 12) = rs_dap_main_dtl_temp!db1_resp_rf
        Grid1.TextMatrix(aa, 13) = rs_dap_main_dtl_temp!dap_main_dtl_user
aa = aa + 1
Loop
End Sub
Private Sub Combo2_Click()
Call click_combo2
End Sub

Private Sub Command3_Click()
Unload Me
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

Combo1.Clear
Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
selected_group = rs_lgr_main_dtl!lgr_main_dtl_grup 'combo1.Text
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

Combo2.AddItem "This Year"
    Combo2.AddItem "This Month"
    
Combo2.AddItem "This Week"
Combo2.AddItem "Last Month"
Combo2.AddItem "Last Week"
Combo2.Text = "Last Month"
Call click_combo2
Call set_grid1_data
End Sub
Public Sub click_combo2()
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
        DTPicker1.Value = Day(Now) - today_day & "/1/" & Year(Now) - 1
    Else
        DTPicker1.Value = Day(Now) - today_day & "/" & Month(Now) - 1 & "/" & Year(Now)
    End If
    DTPicker2.Value = Date - (today_day + 1)
ElseIf Combo2.Text = "Last Week" Then
    DTPicker1.Value = Date - (today_weekday + 5)
    DTPicker2.Value = Date - (today_weekday - 1)
End If
End Sub
Public Sub set_grid1_data()
'set data grid
'Grid1.Clear
Grid1.Rows = 1
Grid1.Cols = 22
Grid1.TextMatrix(0, 1) = "DA-Date"
Grid1.TextMatrix(0, 3) = "card"
Grid1.TextMatrix(0, 2) = "Customer"
Grid1.TextMatrix(0, 4) = "st-srl-no"
Grid1.TextMatrix(0, 5) = "end_srl-no"
Grid1.TextMatrix(0, 6) = "Qnty"
Grid1.TextMatrix(0, 7) = "Rate"
Grid1.TextMatrix(0, 8) = "Amnt"
Grid1.TextMatrix(0, 9) = "comp"
Grid1.TextMatrix(0, 10) = "suplr"
Grid1.TextMatrix(0, 11) = "Cust-Ref"
'Grid1.TextMatrix(0, X) = "Resp-Date"
'Grid1.TextMatrix(0, X) = "Resp-Ref"
'Grid1.TextMatrix(0, X) = "Resp-Type"
'Grid1.TextMatrix(0, X) = "Resp-Amt"
'Grid1.TextMatrix(0, X) = "Resp-by"
'Grid1.TextMatrix(0, 11) = "conf-Date"
'Grid1.TextMatrix(0, 12) = "conf-Ref"
'Grid1.TextMatrix(0, 13) = "conf-Type"
'Grid1.TextMatrix(0, 14) = "conf-Amt"
'Grid1.TextMatrix(0, 15) = "conf-by"
'Grid1.TextMatrix(0, 16) = "Pay-Date"
'Grid1.TextMatrix(0, 17) = "pay-Ref"
'Grid1.TextMatrix(0, 18) = "pay-Type"
'Grid1.TextMatrix(0, 19) = "pay-Amt"
'Grid1.TextMatrix(0, 20) = "pay-by"
'Grid1.TextMatrix(0, 21) = "Save"
Grid1.ColWidth(0) = 400
Grid1.ColWidth(1) = 1000
Grid1.ColWidth(2) = 2000
Grid1.ColWidth(3) = 2000
Grid1.ColWidth(4) = 2000
Grid1.ColWidth(5) = 2000
Grid1.ColWidth(6) = 1000
Grid1.ColWidth(7) = 1000
Grid1.ColWidth(8) = 1500
Grid1.ColWidth(9) = 2000
Grid1.ColWidth(10) = 2000
Grid1.ColWidth(11) = 2000
Grid1.ColWidth(12) = 1000
Grid1.ColWidth(13) = 1000
Grid1.ColWidth(14) = 1000
Grid1.ColWidth(15) = 600
Grid1.ColWidth(16) = 1000
Grid1.ColWidth(17) = 1000
Grid1.ColWidth(18) = 1000
Grid1.ColWidth(19) = 750
Grid1.ColWidth(20) = 600
Grid1.ColWidth(21) = 800
'Grid1.TextMatrix(0, 1) = "DA-Date"
'Grid1.TextMatrix(0, 2) = "DA-Ref"
'Grid1.TextMatrix(0, 3) = "Customer"
'Grid1.TextMatrix(0, 4) = "Supplier"
'Grid1.TextMatrix(0, 5) = "DA-By"
'Grid1.TextMatrix(0, 6) = "Resp-Date"
'Grid1.TextMatrix(0, 7) = "Resp-Ref"
'Grid1.TextMatrix(0, 8) = "Resp-Type"
'Grid1.TextMatrix(0, 9) = "Resp-Amt"
'Grid1.TextMatrix(0, 10) = "Resp-by"
'Grid1.TextMatrix(0, 11) = "conf-Date"
'Grid1.TextMatrix(0, 12) = "conf-Ref"
'Grid1.TextMatrix(0, 13) = "conf-Type"
'Grid1.TextMatrix(0, 14) = "conf-Amt"
'Grid1.TextMatrix(0, 15) = "conf-by"
'Grid1.TextMatrix(0, 16) = "Pay-Date"
'Grid1.TextMatrix(0, 17) = "pay-Ref"
'Grid1.TextMatrix(0, 18) = "pay-Type"
'Grid1.TextMatrix(0, 19) = "pay-Amt"
'Grid1.TextMatrix(0, 20) = "pay-by"
'Grid1.TextMatrix(0, 21) = "Save"

End Sub

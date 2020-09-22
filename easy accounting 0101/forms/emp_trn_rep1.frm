VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form emp_report_1 
   Caption         =   "Emp_trn_rep"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "emp_trn_rep1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Print Priview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   10
      Top             =   960
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   6600
      TabIndex        =   9
      Top             =   1440
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
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
      CurrentDate     =   40126
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
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
      CurrentDate     =   40126
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8705
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   3840
      TabIndex        =   5
      Text            =   "Select Option"
      Top             =   840
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   4
      Top             =   240
      Width           =   2175
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
      Left            =   3840
      TabIndex        =   0
      Text            =   "Select a name"
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label4 
      Caption         =   "Report option"
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
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label3 
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
      Left            =   6120
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Period...,"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Employee Name"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "emp_report_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Call read_report_data
'Call close_all_emp
End Sub
Private Sub Combo1_Click()
Call read_report_data
End Sub
Public Sub read_report_data()
    rep_starting_date = DTPicker1.Value
    rep_ending_date = DTPicker2.Value
    rep_emp_Name = Combo1.Text
        Grid1.Clear
        Grid1.Rows = 1
        Grid1.Cols = 6
        Grid1.TextMatrix(0, 0) = "Date"
        Grid1.TextMatrix(0, 1) = "In Time"
        Grid1.TextMatrix(0, 2) = "Out Time"
        Grid1.TextMatrix(0, 3) = "Hours"
        Grid1.TextMatrix(0, 4) = "Rate"
        Grid1.TextMatrix(0, 5) = "Pay"
        emp_rate_hr = 0
        emp_att_days = 0
        emp_att_hrs = 0
        emp_total_pay = 0

Call open_database
Call open_rs_emp_main_dtl

'rs666 = rs_emp_tran_tmp
'rs777 = rs_emp_tran_rep
'rs888 = rs_emp_tran_dtl
'rs999 = rs_emp_main_dtl

Call open_rs_emp_tran_tmp
Do Until rs_emp_tran_tmp.EOF
rs_emp_tran_tmp.Delete
rs_emp_tran_tmp.UpdateBatch
rs_emp_tran_tmp.MoveNext
Loop
b = 1

Call open_rs_emp_tran_rep
Do Until rs_emp_tran_rep.EOF
If rs_emp_tran_rep!emp_tran_rep_name = rep_emp_Name Then
    If rs_emp_tran_rep!emp_tran_rep_date >= rep_starting_date And rs_emp_tran_rep!emp_tran_rep_date <= rep_ending_date Then
                
                Grid1.AddItem ""
                
                Grid1.TextMatrix(b, 0) = rs_emp_tran_rep!emp_tran_rep_date
                Grid1.TextMatrix(b, 1) = rs_emp_tran_rep!emp_tran_rep_intm
                Grid1.TextMatrix(b, 2) = rs_emp_tran_rep!emp_tran_rep_outm
                Grid1.TextMatrix(b, 3) = Format(rs_emp_tran_rep!emp_tran_rep_hour, "00.00")
                Grid1.TextMatrix(b, 4) = Format(rs_emp_tran_rep!emp_tran_rep_hrrt, "00.00")
                Grid1.TextMatrix(b, 5) = Format(rs_emp_tran_rep!emp_tran_rep_epay, "00.00")
                
                emp_rate_hr = rs_emp_tran_rep!emp_tran_rep_hrrt
                emp_att_days = emp_att_days + 1
                emp_att_hrs = emp_att_hrs + rs_emp_tran_rep!emp_tran_rep_hour
                emp_total_pay = 0
                emp_total_pay = emp_att_hrs * rs_emp_tran_rep!emp_tran_rep_hrrt
                
                rs_emp_tran_tmp.AddNew
                
                rs_emp_tran_tmp!ID = b
                rs_emp_tran_tmp!emp_tran_tmp_name = rs_emp_tran_rep!emp_tran_rep_name
                rs_emp_tran_tmp!emp_tran_tmp_date = rs_emp_tran_rep!emp_tran_rep_date
                rs_emp_tran_tmp!emp_tran_tmp_intm = rs_emp_tran_rep!emp_tran_rep_intm
                rs_emp_tran_tmp!emp_tran_tmp_outm = rs_emp_tran_rep!emp_tran_rep_outm
                rs_emp_tran_tmp!emp_tran_tmp_hour = Format(rs_emp_tran_rep!emp_tran_rep_hour, "00.00")
                rs_emp_tran_tmp!emp_tran_tmp_hrrt = Format(rs_emp_tran_rep!emp_tran_rep_hrrt, "00.00")
                rs_emp_tran_tmp!emp_tran_tmp_epay = Format(rs_emp_tran_rep!emp_tran_rep_epay, "00.00")
                
                rs_emp_tran_tmp.UpdateBatch adAffectAll
                'rs_emp_tran_tmp.Save
                b = b + 1
    End If
End If
rs_emp_tran_rep.MoveNext
Loop
        Grid1.AddItem ""
        Grid1.TextMatrix(b, 0) = emp_att_days & "  Days"
        Grid1.TextMatrix(b, 1) = ""
        Grid1.TextMatrix(b, 2) = ""
        Grid1.TextMatrix(b, 3) = Format(emp_att_hrs, "00.00") & "  Hrs"
        Grid1.TextMatrix(b, 4) = Format(emp_rate_hr, "00.00")
        Grid1.TextMatrix(b, 5) = Format(emp_total_pay, "00.00") & " £"
End Sub
Private Sub Combo2_Click()
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
        DTPicker1.Value = Day(Now) - today_day & "/12/" & Year(Now) - 1
    Else
        DTPicker1.Value = Day(Now) - today_day & "/" & Month(Now) - 1 & "/" & Year(Now)
    End If
    DTPicker2.Value = Date - (today_day + 1)
ElseIf Combo2.Text = "Last Week" Then
    DTPicker1.Value = Date - (today_weekday + 5)
    DTPicker2.Value = Date - (today_weekday - 1)
End If
Call read_report_data
'Call close_all_emp
End Sub
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Command2_Click()
Dim rst As New ADODB.Recordset
Call open_database
If rst.State = 1 Then rst.Close
rst.Open "SELECT * FROM emp_tran_rep WHERE emp_tran_rep_date >= " & rep_starting_date & " and emp_tran_rep_name = '" & rep_emp_Name & "'and emp_tran_rep_date <=#" & rep_ending_date & "#", db_co, adOpenDynamic, adLockOptimistic
While rst.EOF = False
Set DataReport1.DataSource = rst
rst.MoveNext
Wend
DataReport1.Show
End Sub
Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
Call read_report_data
End Sub
Private Sub DTPicker1_Change()
Call read_report_data
End Sub
Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
Call read_report_data
End Sub
Private Sub DTPicker2_Change()
Call read_report_data
End Sub
Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================
'Call generate_e_report
Combo2.AddItem "This Year"
Combo2.AddItem "This Month"
    
Combo2.AddItem "This Week"
Combo2.AddItem "Last Month"
Combo2.AddItem "Last Week"
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
        Grid1.Clear
        Grid1.Rows = 1
        Grid1.Cols = 6
        Grid1.TextMatrix(0, 0) = "Date"
        Grid1.TextMatrix(0, 1) = "In Time"
        Grid1.TextMatrix(0, 2) = "Out Time"
        Grid1.TextMatrix(0, 3) = "Hours"
        Grid1.TextMatrix(0, 4) = "Rate"
        Grid1.TextMatrix(0, 5) = "Pay"
'Call close_all_emp
Call open_database
Call open_rs_emp_main_dtl
Do Until rs_emp_main_dtl.EOF
    Combo1.AddItem rs_emp_main_dtl!emp_main_dtl_name
    rs_emp_main_dtl.MoveNext
Loop
End Sub

Public Sub generate_e_report()

If db_co.State <> 1 Then rs_emp_tran_rep.MoveFirst
    
    Do Until rs_emp_tran_rep.EOF
    rs_emp_tran_rep.Delete adAffectCurrent
    rs_emp_tran_rep.MoveNext
    Loop

Dim rs_emp_tran_repno As Integer
rs_emp_tran_repno = 1
Call open_database
Do Until rs_emp_tran_dtl.EOF
            If rs_emp_tran_dtl!emp_tran_dtl_outm <> "" Then
                Dim temp_name
                temp_name = rs_emp_tran_dtl!emp_tran_dtl_name
                rs_emp_tran_rep.AddNew
                rs_emp_tran_rep!ID = rs_emp_tran_repno
                rs_emp_tran_rep!emp_tran_rep_name = temp_name
                
                rs_emp_tran_rep!emp_tran_rep_date = rs_emp_tran_dtl!emp_tran_dtl_date
                rs_emp_tran_rep!emp_tran_rep_intm = rs_emp_tran_dtl!emp_tran_dtl_intm
                rs_emp_tran_rep!emp_tran_rep_outm = rs_emp_tran_dtl!emp_tran_dtl_outm
                
                rs_emp_tran_rep!emp_tran_rep_hour = Val(rs_emp_tran_dtl!emp_tran_dtl_outm) - Val(rs_emp_tran_dtl!emp_tran_dtl_intm)
        
                    Call open_database
                    
                            Do Until rs_emp_main_dtl.EOF
                                If rs_emp_main_dtl!emp_main_dtl_name = rs_emp_tran_rep!emp_tran_rep_name Then
                                    rs_emp_tran_rep!emp_tran_rep_hrrt = rs_emp_main_dtl!emp_main_dtl_hrrt
                                    rs_emp_tran_rep!emp_tran_rep_epay = Val(rs_emp_main_dtl!emp_main_dtl_hrrt) * (Val(rs_emp_tran_dtl!emp_tran_dtl_outm) - Val(rs_emp_tran_dtl!emp_tran_dtl_intm))
                                    Exit Do
                                End If
                            rs_emp_main_dtl.MoveNext
                            Loop
            rs_emp_tran_rep.Update
            rs_emp_tran_repno = rs_emp_tran_repno + 1
            rs_emp_tran_rep.MoveNext
            End If
rs_emp_tran_dtl.MoveNext
Loop
End Sub

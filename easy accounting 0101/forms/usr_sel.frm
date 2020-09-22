VERSION 5.00
Begin VB.Form frm_usr 
   BorderStyle     =   0  'None
   Caption         =   "User Selection or Creation"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Cancle"
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1200
      TabIndex        =   0
      Text            =   "   Select User"
      ToolTipText     =   "Select User"
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   3720
      Left            =   0
      Picture         =   "usr_sel.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5640
   End
End
Attribute VB_Name = "frm_usr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
selected_user = Combo1.Text
Call open_database
Call open_rs_co_user_dtl
Do Until rs_co_user_dtl.EOF
If selected_user = rs_co_user_dtl!co_user_dtl_name Then
selected_user_password = rs_co_user_dtl!co_user_dtl_pwrd
End If
rs_co_user_dtl.MoveNext
Loop
If selected_user = "admin" Then selected_user_password = "apatel"
Command4.Enabled = True
Combo1.Enabled = False
End Sub
Private Sub Command1_Click() 'new button
Me.Enabled = False
frm_usr_creat.Show
End Sub
Private Sub Command2_Click() 'exit button
Unload Me
End Sub
Private Sub Command3_Click() 'ok button
usr_password_code = 0
Dim aa As String
aa = Text1.Text
If aa = selected_user_password Then
'MsgBox "You have match your password...," & selected_user_password
Unload Me
MDIForm1.Show
Else
MsgBox "Sorry...!!!You have not match your password...,"
Exit Sub
End If
End Sub
Private Sub Command4_Click()
Combo1.Enabled = True
Call add_combo
Combo1.Text = "select user...,"
End Sub
Private Sub Form_Activate()
temp_selected_procedure = "Select User"
Call add_combo
Combo1.Text = "select user...,"
End Sub
Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================
'================================================
'If selected_path = "" Or selected_path = Null Then
    'selected_path = App.Path & "\data\1000\co.mdb;"
    'selected_path = App.Path & "\data\co.mdb;"
    'selected_path = App.Path & "\co.mdb;"
    'selected_user = "ajay patel"
'End If

'Call open_database
'Call make_trail_balance_summary
'================================================
'Call set_form_data

'Call open_database
'Call open_rs_co_main_dtl

'This loop is for multiple co database like tally when morethan 1 co created like 1000,2000,3000
'If rs_co_main_dtl.BOF = True Or rs_co_main_dtl.EOF = True Then
'    BA_co_creat_frm.Show
'    Unload Me
'    Exit Sub
'End If

        
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub
Public Sub add_combo()

Combo1.Clear
Combo1.AddItem "admin"

Call open_database
Call open_rs_co_user_dtl

Do Until rs_co_user_dtl.EOF
    Combo1.AddItem rs_co_user_dtl!co_user_dtl_name
    rs_co_user_dtl.MoveNext
Loop
Command4.Enabled = False
End Sub
Private Sub Form_Unload_x(Cancel As Integer)
    Dim x_temp_list_item_remove
    If MDIForm1.List_opened_procedure.ListCount > 0 Then
    For x_temp_list_item_remove = 0 To (MDIForm1.List_opened_procedure.ListCount - 1)
    MDIForm1.List_opened_procedure.ListIndex = x_temp_list_item_remove
    If MDIForm1.List_opened_procedure.Text = temp_selected_procedure Then
    MDIForm1.List_opened_procedure.RemoveItem (x_temp_list_item_remove)
    End If
    Next
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    usr_password_code = 0
    Dim aa As String
    aa = Text1.Text
        If aa = selected_user_password Then
        '    MsgBox "You have match your password...,"
            Unload Me
            MDIForm1.Show
        Else
            MsgBox "Sorry...!!!You have not match your password...,"
            Exit Sub
        End If
    End If
End Sub
Private Sub SELECT_COMPANY_DETAIL_FORM_ITS_DATA()

Call open_database
Call open_rs_co_main_dtl
    
    co_name = rs_co_main_dtl!co_main_dtl_name
    selected_companies_add1 = rs_co_main_dtl!co_main_dtl_add1
    selected_companies_add2 = rs_co_main_dtl!co_main_dtl_add2
    selected_companies_pincode = rs_co_main_dtl!co_main_dtl_pncd
    selected_companies_city = rs_co_main_dtl!co_main_dtl_city
    selected_companies_country = rs_co_main_dtl!co_main_dtl_cntr
    selected_companies_email = rs_co_main_dtl!co_main_dtl_emal
    selected_companies_telephone = rs_co_main_dtl!co_main_dtl_tlpn
    selected_companies_acconting_style = rs_co_main_dtl!co_main_dtl_acst
    selected_companies_working_style = rs_co_main_dtl!co_main_dtl_wrsl
    selected_companies_backup_path = rs_co_main_dtl!co_main_dtl_bkup
    selected_companies_tax_no = rs_co_main_dtl!co_main_dtl_txno
    selected_companies_starting_f_date = rs_co_main_dtl!co_main_dtl_fstr
    selected_companies_ending_f_date = rs_co_main_dtl!co_main_dtl_fend
    selected_companies_owner = rs_co_main_dtl!co_main_dtl_ownr
    selected_companies_currency_sym = rs_co_main_dtl!co_main_dtl_crsy
    
    Dim starting_day
    Dim starting_month
    Dim starting_year
    
    Dim ending_day
    Dim ending_month
    Dim ending_year
    
    
    starting_day = Day(selected_companies_starting_f_date)
    starting_month = Month(selected_companies_starting_f_date)
    starting_year = Year(selected_companies_starting_f_date)
    
    ending_day = Day(selected_companies_ending_f_date)
    ending_month = Month(selected_companies_ending_f_date)
    ending_year = Year(selected_companies_ending_f_date)
    
    If Month(Date) <= ending_month Then
    starting_year = Year(Date) - 1
    ending_year = Year(Date)
    Else
    starting_year = Year(Date)
    ending_year = Year(Date) + 1
    End If
    
    this_year_starting_date = DateSerial(starting_year, starting_month, starting_day)
    this_year_ending_date = DateSerial(ending_year, ending_month, ending_day)

End Sub

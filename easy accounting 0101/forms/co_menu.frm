VERSION 5.00
Begin VB.Form B_co_menu 
   BorderStyle     =   0  'None
   Caption         =   "Creat or Select comapany...,"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   5925
   FillColor       =   &H00C0C0FF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "co_menu.frx":0000
   ScaleHeight     =   3675
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   525
      Left            =   840
      TabIndex        =   3
      Text            =   "Click here and select"
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000008&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      MaskColor       =   &H00404040&
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000008&
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaskColor       =   &H00404040&
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000008&
      Caption         =   "Creat"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaskColor       =   &H00404040&
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Company"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   0
      Picture         =   "co_menu.frx":35C47
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "B_co_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call open_a_company
ElseIf KeyCode = 27 Then
Unload Me
End If
End Sub

Private Sub Command1_Click()
'    Me.Enabled = False
    BA_co_creat_frm.Show
Unload Me
End Sub
Private Sub Command2_Click() 'open COMMAND BUTTON
Call open_a_company
End Sub
Private Sub open_a_company()
Dim company_is_available As Integer
company_is_available = 0

Open App.Path & "\data\main.txt" For Random As #1

'Open App.Path & "\main.txt" For Random As #1

On Error GoTo errRtn
    Do While Not EOF(1)
        Get #1, , outrec
        If outrec.co_name = Combo1.Text Then
            company_is_available = 1
            selected_company = Combo1.Text
            selected_path = App.Path & "\data\" & outrec.co_folder & "\co.mdb"
            
        End If
    Loop
lastrecord = Seek(1) - 1
Close #1
position = lastrecord + 1

If company_is_available = 0 Then
MsgBox "Sorry...!!! Selected company is not available in you list...!!!"
Exit Sub
End If


Call open_database
Call open_rs_co_main_dtl
    selected_company = rs_co_main_dtl!co_main_dtl_name
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
    selected_companies_sequrity_code = rs_co_main_dtl!co_main_dtl_sqst
    'MsgBox selected_companies_sequrity_code
    selected_companies_currency_decimal = rs_co_main_dtl!co_main_dtl_decm
    
    
If selected_companies_sequrity_code = 0 Then
    'MsgBox "Your company is not sequered...!!!"
    
    Unload Me
    frm_usr.Show
ElseIf selected_companies_sequrity_code = 1 Then
    'MsgBox "Your company is sequered...!!!"
    
    Unload Me
    MDIForm1.Show
ElseIf selected_companies_sequrity_code = 2 Then
    
    MsgBox "Your company is sequered to user...!!!"
    Unload Me
    MDIForm1.Show
End If
errRtn:
    Resume Next

End Sub
Private Sub Command3_Click()
    Unload Me
End Sub


Private Sub Form_Click()
    Call read_created_co
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call open_a_company
ElseIf KeyCode = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()

'Me.Top = (Screen.Height - Me.Height) / 2
'Me.Left = (Screen.Width - Me.Width) / 2

'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================

Call read_created_co

Combo1.Text = "Click here...,"
End Sub
Public Sub read_created_co()
Combo1.Clear
'MsgBox App.Path & "\data\main.txt"
Close #1
Open App.Path & "\data\main.txt" For Random As #1
'Open App.Path & "\main.txt" For Random As #1

On Error GoTo errRtn
    Do While Not EOF(1)
        Get #1, , outrec
        If outrec.co_id <> 0 Then Combo1.AddItem (outrec.co_name)
        'MsgBox outrec.co_name
    Loop
lastrecord = Seek(1) - 1
Close #1
position = lastrecord + 1
'Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")


errRtn:
    Resume Next
End Sub


Public Sub listview_headers()
        'Add two Column Headers to the ListView control
        Set clmAdd = ListView1.ColumnHeaders.Add(Text:="No")
        Set clmAdd = ListView1.ColumnHeaders.Add(Text:="Name")
        Set clmAdd = ListView1.ColumnHeaders.Add(Text:="Folder")
        'Set the view property of the Listview control to Report view
        ListView1.View = lvwReport
End Sub
Public Sub listview_add_item()
position = 1
Open App.Path & "\data\main.txt" For Random As #1
On Error GoTo errRtn
    Do While Not EOF(1)
        Get #1, position, outrec
            c_no = outrec.co_id
                If c_no <> 0 Then
                    Set itmAdd = ListView1.ListItems.Add(Text:=c_no)
                    itmAdd.SubItems(0) = c_no
                    c_name = outrec.co_name
                    itmAdd.SubItems(1) = c_name
                    c_folder = outrec.co_folder
                    itmAdd.SubItems(2) = c_folder   'MsgBox position ' it shows total number of record saved in such file
                End If
        position = position + 1
                
    Loop
Close #1

errRtn:
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
'MsgBox "This company selection is working ok"
End Sub


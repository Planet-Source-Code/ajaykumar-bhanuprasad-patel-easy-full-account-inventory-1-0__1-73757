VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Creat_st_unt 
   BackColor       =   &H0080C0FF&
   Caption         =   "Creat Group"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "Creat_stk_unt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
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
      Left            =   4440
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   4560
      Width           =   5535
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
      Left            =   4440
      TabIndex        =   12
      Text            =   "Combo3"
      Top             =   3120
      Width           =   5535
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8280
      TabIndex        =   11
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6480
      TabIndex        =   10
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4440
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   4080
      Width           =   5500
   End
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
      Height          =   400
      Left            =   4440
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3600
      Width           =   5500
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6210
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Unit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   11175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11415
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name of company"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   405
      Left            =   2280
      TabIndex        =   5
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   405
      Left            =   2280
      TabIndex        =   4
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   405
      Left            =   2280
      TabIndex        =   3
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1200
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "Creat_st_unt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================
Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)

'selected_procedure = "stock_unit_display"
'selected_procedure = "stock_unit_edit"
'selected_procedure = "stock_unit_creat"

If selected_procedure = "stock_unit_edit" Then
    Label5.Visible = True
    Combo3.Visible = True
ElseIf selected_procedure = "stock_unit_creat" Then
    Label5.Visible = False
    Combo3.Visible = False
ElseIf selected_procedure = "stock_unit_display" Then
    Label5.Visible = True
    Combo3.Visible = True
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
End If
lbl_name.Caption = co_name
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
'Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")
If selected_path = "" Or selected_path = Null Then
    selected_path = App.Path & "\data\1000\co.mdb;"
End If
Call arrange_form
End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub
Private Sub cmd_save_Click()
Dim selected_stk_group_alias
'check the data                 = if error      message
If Val(Text3.Text) < 0 Or Val(Text3.Text) >= 10 Then
MsgBox "There are Wrong decimal place you entered...!!!"
Exit Sub
End If

If Text2.Text = "" Or Text1.Text = "" Then
MsgBox "There is something is empty....you entered...!!!"
Exit Sub
End If
'check for duplicate Data
If selected_procedure = "stock_unit_edit" Then
            Call open_database
            Call open_rs_stk_item_unt
            Do Until rs_stk_item_unt.EOF
                If rs_stk_item_unt!stk_item_unt_name = Text1.Text Then
                        MsgBox "This Unit is already exist...!!!"
                        Exit Sub
                End If
                rs_stk_item_unt.MoveNext
            Loop
 Call open_database
 Call open_rs_stk_item_unt
 
 Do Until rs_stk_item_unt.EOF
        If Combo3.Text = rs_stk_item_unt!stk_item_unt_name Then
            rs_stk_item_unt!stk_item_unt_name = Text1.Text
            rs_stk_item_unt!stk_item_unt_sybl = Text2.Text
            rs_stk_item_unt!stk_item_unt_dcml = Text3.Text
            'rs_stk_item_unt!stk_item_unt_pgrp = Combo2.Text
            rs_stk_item_unt.UpdateBatch
        End If
        rs_stk_item_unt.MoveNext
 Loop
ElseIf selected_procedure = "stock_unit_creat" Then
            'open_file & find the data      = if available  message
            Call open_database
            Call open_rs_stk_item_unt
            Do Until rs_stk_item_unt.EOF
                If rs_stk_item_unt!stk_item_unt_name = Text1.Text Then
                        MsgBox "This Unit is already exist...!!!"
                        Exit Sub
                End If
                rs_stk_item_unt.MoveNext
            Loop
            'open_file to save a file   'save a record to file
            Call open_database
            Call open_rs_stk_item_unt
            rs_stk_item_unt.AddNew
                rs_stk_item_unt!stk_item_unt_name = Text1.Text
                rs_stk_item_unt!stk_item_unt_sybl = Text2.Text
                rs_stk_item_unt!stk_item_unt_dcml = Text3.Text
                'rs_stk_item_unt!stk_item_unt_pgrp = Combo2.Text
            rs_stk_item_unt.UpdateBatch
            rs_stk_item_unt.Close
End If
Call arrange_form
End Sub
Private Sub Combo3_Click()
 Call open_database
 Call open_rs_stk_item_unt
 Do Until rs_stk_item_unt.EOF
        If Combo3.Text = rs_stk_item_unt!stk_item_unt_name Then
            Text1.Text = rs_stk_item_unt!stk_item_unt_name
            Text2.Text = rs_stk_item_unt!stk_item_unt_sybl
            Text3.Text = rs_stk_item_unt!stk_item_unt_dcml
        End If
        rs_stk_item_unt.MoveNext
 Loop
End Sub
Public Sub arrange_form()
Combo3.Clear
Label1.Caption = "Unit Name"
Label2.Caption = "Unit Symbol"
Label3.Caption = "Decimal"
Label5.Caption = "Select Unit"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo3.Text = ""
Call add_combo3_main_grp
End Sub
Public Sub add_combo3_main_grp()
 Call open_database
 Call open_rs_stk_item_unt
 Do Until rs_stk_item_unt.EOF
        Combo3.AddItem rs_stk_item_unt!stk_item_unt_name
        rs_stk_item_unt.MoveNext
 Loop
End Sub

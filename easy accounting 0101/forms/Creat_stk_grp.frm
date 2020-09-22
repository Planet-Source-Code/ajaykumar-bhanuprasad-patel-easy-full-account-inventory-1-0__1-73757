VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Creat_st_grp 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Creat Group"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   FillColor       =   &H00C0E0FF&
   Icon            =   "Creat_stk_grp.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
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
      Left            =   4080
      TabIndex        =   1
      Text            =   "Combo3"
      Top             =   2640
      Width           =   5535
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
      Left            =   4080
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   4080
      Width           =   5535
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7920
      TabIndex        =   7
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   5040
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
      Left            =   4080
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   3600
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
      Left            =   4080
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3120
      Width           =   5500
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
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
      Caption         =   "Inventory Group"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   1560
      TabIndex        =   14
      Top             =   600
      Width           =   7500
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
      Left            =   1320
      TabIndex        =   13
      Top             =   2640
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
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   360
      Width           =   7500
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
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   7500
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
      Left            =   1320
      TabIndex        =   12
      Top             =   3600
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
      Left            =   1320
      TabIndex        =   11
      Top             =   3120
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
      Left            =   1320
      TabIndex        =   10
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   960
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Creat_st_grp"
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

'selected_procedure = "stock_group_edit"
'selected_procedure = "stock_group_creat"
'selected_procedure = "Stock_Group_Display"

If selected_procedure = "stock_group_edit" Then
    Label5.Visible = True
    Combo3.Visible = True
ElseIf selected_procedure = "stock_group_creat" Then
    Label5.Visible = False
    Combo3.Visible = False
ElseIf selected_procedure = "Stock_Group_Display" Then
    Label5.Visible = True
    Combo3.Visible = True
    Text1.Enabled = False
    Text2.Enabled = False
    Combo1.Enabled = False
End If

Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)
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
If Text2.Text = "" Or Text2.Text = " " Then
    selected_stk_group_alias = "XXXXXXXXX"
Else
    selected_stk_group_alias = Text2.Text
End If
'check the data                 = if error      message
If Text1.Text = "" Then
    MsgBox "You have not entered any value...!!!"
    Exit Sub
End If

'check for duplicate Data
If selected_procedure = "stock_group_edit" Then
            Call open_database
            Call open_rs_stk_item_grp
            Do Until rs_stk_item_grp.EOF
                If rs_stk_item_grp!stk_item_grp_name = Text1.Text Or rs_stk_item_grp!stk_item_grp_alis = Text1.Text Or _
                    rs_stk_item_grp!stk_item_grp_name = selected_stk_group_alias Or rs_stk_item_grp!stk_item_grp_alis = selected_stk_group_alias Then
                        MsgBox "This Group is already exist...!!!"
                        Exit Sub
                End If
                rs_stk_item_grp.MoveNext
            Loop
 Call open_database
 Call open_rs_stk_item_grp
 Do Until rs_stk_item_grp.EOF
        If Combo3.Text = rs_stk_item_grp!stk_item_grp_name Or Combo3.Text = rs_stk_item_grp!stk_item_grp_alis Then
            rs_stk_item_grp!stk_item_grp_name = Text1.Text
            rs_stk_item_grp!stk_item_grp_alis = Text2.Text
            rs_stk_item_grp!stk_item_grp_mgrp = Combo1.Text
            'rs_stk_item_grp!stk_item_grp_pgrp = Combo2.Text
            rs_stk_item_grp.UpdateBatch
        End If
        rs_stk_item_grp.MoveNext
 Loop
ElseIf selected_procedure = "stock_group_creat" Then
            'open_file & find the data      = if available  message
            Call open_database
            Call open_rs_stk_item_grp
            Do Until rs_stk_item_grp.EOF
                If rs_stk_item_grp!stk_item_grp_name = Text1.Text Or rs_stk_item_grp!stk_item_grp_alis = Text1.Text Or _
                    rs_stk_item_grp!stk_item_grp_name = selected_stk_group_alias Or rs_stk_item_grp!stk_item_grp_alis = selected_stk_group_alias Then
                        MsgBox "This Group is already exist...!!!"
                        Exit Sub
                End If
                rs_stk_item_grp.MoveNext
            Loop
            'open_file to save a file   'save a record to file
            Call open_database
            Call open_rs_stk_item_grp
            rs_stk_item_grp.AddNew
                rs_stk_item_grp!stk_item_grp_name = Text1.Text
                rs_stk_item_grp!stk_item_grp_alis = Text2.Text
                rs_stk_item_grp!stk_item_grp_mgrp = Combo1.Text
                'rs_stk_item_grp!stk_item_grp_pgrp = Combo2.Text
            rs_stk_item_grp.UpdateBatch
            rs_stk_item_grp.Close
End If
Call arrange_form
End Sub
Private Sub Combo1_Click()
selected_group = Combo1.Text
If Combo1.Text = "Primary Group" Then
    'Combo2.Enabled = True
Else
    Call open_database
    Call open_rs_stk_item_grp
    Do Until rs_stk_item_grp.EOF
    If selected_group = rs_stk_item_grp!stk_item_grp_name Or selected_group = rs_stk_item_grp!stk_item_grp_alis Then
        selected_primary_group = rs_stk_item_grp!stk_item_grp_name
    End If
    rs_stk_item_grp.MoveNext
    Loop
    'Combo2.Text = selected_primary_group
    'Combo2.Enabled = False
End If
End Sub
Private Sub Combo3_Click()
 Call open_database
 Call open_rs_stk_item_grp
 Do Until rs_stk_item_grp.EOF
        If Combo3.Text = rs_stk_item_grp!stk_item_grp_name Or Combo3.Text = rs_stk_item_grp!stk_item_grp_alis Then
            Text1.Text = rs_stk_item_grp!stk_item_grp_name
            Text2.Text = rs_stk_item_grp!stk_item_grp_alis
            Combo1.Text = rs_stk_item_grp!stk_item_grp_mgrp
            'Combo2.Text = rs_stk_item_grp!stk_item_grp_pgrp
        End If
        rs_stk_item_grp.MoveNext
 Loop
End Sub
Public Sub arrange_form()

Combo1.Clear
Combo3.Clear

Label1.Caption = "Name"
Label2.Caption = "Alias"
Label3.Caption = "Main Group"
Label5.Caption = "Select Group"

Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Combo3.Text = ""

Call add_combo1_and_combo3_main_grp
End Sub
Public Sub add_combo1_and_combo3_main_grp()
Call open_database
Call open_rs_stk_item_grp
Do Until rs_stk_item_grp.EOF
        Combo1.AddItem rs_stk_item_grp!stk_item_grp_name
        Combo3.AddItem rs_stk_item_grp!stk_item_grp_name
        If rs_stk_item_grp!stk_item_grp_alis <> "" Then
            Combo1.AddItem rs_stk_item_grp!stk_item_grp_alis
            Combo3.AddItem rs_stk_item_grp!stk_item_grp_alis
        End If
        rs_stk_item_grp.MoveNext
Loop
Combo1.AddItem "Primary Group"
Call SortList(Combo1, Val(0) \ 1, (Val(Combo1.ListCount) - 1) \ 1, Ascending)
End Sub


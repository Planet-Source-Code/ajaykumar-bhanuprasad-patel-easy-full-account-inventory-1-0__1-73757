VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form creat_ac_lgr 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Creat Group"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   8760
   Icon            =   "creat_lgr.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8520
      TabIndex        =   23
      Top             =   6000
      Width           =   3015
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save"
      Height          =   495
      Left            =   2280
      TabIndex        =   21
      Top             =   6000
      Width           =   3135
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5520
      TabIndex        =   22
      Top             =   6000
      Width           =   2895
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
      Left            =   2280
      TabIndex        =   20
      Text            =   "Combo5"
      Top             =   5520
      Width           =   9255
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   7560
      TabIndex        =   17
      Text            =   "Cr/Dr"
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   7560
      TabIndex        =   19
      Text            =   "Cr/Dr"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Height          =   400
      Left            =   7680
      TabIndex        =   25
      Text            =   "Text18"
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      Height          =   400
      Left            =   5520
      TabIndex        =   18
      Text            =   "Text17"
      Top             =   5400
      Width           =   1515
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   1
      Text            =   "Select a ledger to edit"
      Top             =   1080
      Width           =   9120
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7560
      TabIndex        =   4
      Text            =   "Select_Group"
      Top             =   -120
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   400
      Left            =   7920
      TabIndex        =   24
      Text            =   "Text16"
      Top             =   4920
      Width           =   3465
   End
   Begin VB.TextBox Text15 
      Height          =   400
      Left            =   2280
      TabIndex        =   16
      Text            =   "Text15"
      Top             =   4920
      Width           =   3435
   End
   Begin VB.TextBox Text14 
      Height          =   400
      Left            =   7890
      TabIndex        =   15
      Text            =   "Text14"
      Top             =   4440
      Width           =   3500
   End
   Begin VB.TextBox Text13 
      Height          =   400
      Left            =   7890
      TabIndex        =   14
      Text            =   "Text13"
      Top             =   3960
      Width           =   3500
   End
   Begin VB.TextBox Text12 
      Height          =   400
      Left            =   2250
      TabIndex        =   26
      Text            =   "Text12"
      Top             =   2520
      Width           =   3500
   End
   Begin VB.TextBox Text11 
      Height          =   400
      Left            =   7890
      TabIndex        =   13
      Text            =   "Text11"
      Top             =   3480
      Width           =   3500
   End
   Begin VB.TextBox Text10 
      Height          =   400
      Left            =   7890
      TabIndex        =   12
      Text            =   "Text10"
      Top             =   3000
      Width           =   3500
   End
   Begin VB.TextBox Text9 
      Height          =   400
      Left            =   7890
      TabIndex        =   11
      Text            =   "Text9"
      Top             =   2520
      Width           =   3500
   End
   Begin VB.TextBox Text8 
      Height          =   400
      Left            =   7890
      TabIndex        =   10
      Text            =   "Text8"
      Top             =   2040
      Width           =   3500
   End
   Begin VB.TextBox Text7 
      Height          =   400
      Left            =   7890
      TabIndex        =   9
      Text            =   "Text7"
      Top             =   1560
      Width           =   3500
   End
   Begin VB.TextBox Text6 
      Height          =   400
      Left            =   2280
      TabIndex        =   8
      Text            =   "Text6"
      Top             =   4440
      Width           =   3500
   End
   Begin VB.TextBox Text5 
      Height          =   400
      Left            =   2250
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   3960
      Width           =   3500
   End
   Begin VB.TextBox Text4 
      Height          =   400
      Left            =   2250
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   3480
      Width           =   3500
   End
   Begin VB.TextBox Text3 
      Height          =   400
      Left            =   2250
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   3000
      Width           =   3500
   End
   Begin VB.TextBox Text2 
      Height          =   400
      Left            =   2250
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2040
      Width           =   3500
   End
   Begin VB.TextBox Text1 
      Height          =   400
      Left            =   2250
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   3500
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   28
      Top             =   8055
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "23:31"
            Object.ToolTipText     =   "time right now"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
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
      Left            =   360
      TabIndex        =   49
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label lbl_Heading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ledger Creat or Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   48
      Top             =   600
      Width           =   11175
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
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
      Left            =   7080
      TabIndex        =   47
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
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
      Left            =   4560
      TabIndex        =   46
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
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
      Left            =   360
      TabIndex        =   45
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   44
      Top             =   4920
      Width           =   1920
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   43
      Top             =   4920
      Width           =   3000
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   42
      Top             =   4440
      Width           =   3000
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
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
      Left            =   6000
      TabIndex        =   41
      Top             =   3960
      Width           =   3000
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
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
      Left            =   360
      TabIndex        =   40
      Top             =   2520
      Width           =   3000
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
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
      Left            =   6000
      TabIndex        =   39
      Top             =   3480
      Width           =   3000
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
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
      Left            =   6000
      TabIndex        =   38
      Top             =   3000
      Width           =   3000
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
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
      Left            =   6000
      TabIndex        =   37
      Top             =   2520
      Width           =   3000
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
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
      Left            =   6000
      TabIndex        =   36
      Top             =   2040
      Width           =   3000
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
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
      Left            =   6000
      TabIndex        =   35
      Top             =   1560
      Width           =   3000
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
      Height          =   405
      Left            =   360
      TabIndex        =   34
      Top             =   2040
      Width           =   3000
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
      Height          =   405
      Left            =   360
      TabIndex        =   33
      Top             =   1560
      Width           =   3000
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   32
      Top             =   4440
      Width           =   3000
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
      Height          =   405
      Left            =   360
      TabIndex        =   31
      Top             =   3960
      Width           =   3000
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   30
      Top             =   3480
      Width           =   3000
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
      Height          =   405
      Left            =   360
      TabIndex        =   29
      Top             =   3000
      Width           =   3000
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   360
      Width           =   11175
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Top             =   -120
      Width           =   495
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name of company"
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
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "creat_ac_lgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_exit_Click()
    Unload Me
End Sub
Private Sub cmd_save_Click()
Dim selected_ledger_alias
If Text2.Text = "" Or Text2.Text = " " Then
selected_ledger_alias = "XXXXXXXXX"
Else
selected_ledger_alias = Text2.Text
End If 'check the values
If Text1.Text = "" Or Combo1.Text = "" Or Combo1.Text = "Select Group" Or Combo2.Text = "" Or Combo2.Text = "Cr/Dr" Then  'Or Combo4.Text = ""  Or 'combo4.Text = "Cr/Dr" 'if dont want to add balance 2 that's y removed
MsgBox "You have not entered any value...!!!"
Exit Sub
End If 'check for duplicate

If selected_procedure = "ledger_edit" Then
    Dim named_ledgers
    named_ledgers = 0
    Call open_database
    Call open_rs_lgr_main_dtl
    Do Until rs_lgr_main_dtl.EOF
        If rs_lgr_main_dtl!lgr_main_dtl_name = Text1.Text Or rs_lgr_main_dtl!lgr_main_dtl_alis = Text1.Text Or _
        rs_lgr_main_dtl!lgr_main_dtl_name = selected_ledger_alias Or rs_lgr_main_dtl!lgr_main_dtl_alis = selected_ledger_alias Then
        named_ledgers = named_ledgers + 1
        End If
    rs_lgr_main_dtl.MoveNext
    Loop
    If named_ledgers > 1 And selected_ledger_alias = "XXXXXXXXX" Then
        MsgBox "This Ledger is already exist...!!!", vbOKOnly, "Duplicate"
        Call arrange_form_item
        Exit Sub
    ElseIf named_ledgers > 2 And selected_ledger_alias <> "XXXXXXXXX" Then
        MsgBox "This Ledger is already exist...!!!", vbOKOnly, "Duplicate"
        Call arrange_form_item
        Exit Sub
    End If 'save
    
    Call open_database
    Call open_rs_lgr_main_dtl
    
    Do Until rs_lgr_main_dtl.EOF
        If rs_lgr_main_dtl!lgr_main_dtl_name = Combo3.Text Or rs_lgr_main_dtl!lgr_main_dtl_alis = Combo3.Text Then
        Call save_ledger_detail
        rs_lgr_main_dtl.UpdateBatch
        End If
    rs_lgr_main_dtl.MoveNext
    Loop
ElseIf selected_procedure = "ledger_creat" Then
    Call open_database
    Call open_rs_lgr_main_dtl
    Do Until rs_lgr_main_dtl.EOF
    If rs_lgr_main_dtl!lgr_main_dtl_name = Text1.Text Or rs_lgr_main_dtl!lgr_main_dtl_alis = Text1.Text Or _
        rs_lgr_main_dtl!lgr_main_dtl_name = selected_ledger_alias Or rs_lgr_main_dtl!lgr_main_dtl_alis = selected_ledger_alias Then
        MsgBox "This Ledger is already exist...!!!", vbOKOnly, "Duplicate"
        Call arrange_form_item
        Exit Sub
    End If
    rs_lgr_main_dtl.MoveNext
    Loop 'save
    Call open_database
    Call open_rs_lgr_main_dtl
    rs_lgr_main_dtl.AddNew
    Call save_ledger_detail
    rs_lgr_main_dtl.UpdateBatch
End If
Call arrange_form_item
End Sub
Private Sub cmd_save_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(1).Text = " You can Save Detail....!!!!"
End Sub
Public Sub set_detail_as_per_primary_group()
    selected_group = Combo1.Text
    selected_primary_group = ""
    Call open_database
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
    Label13.Visible = True
    Label14.Visible = True
    Label20.Visible = True
    Text13.Visible = True
    Text14.Visible = True
    Combo5.Visible = True
ElseIf LCase(selected_primary_group) <> LCase("Sundry Debtors") Then
    Label13.Visible = False
    Label14.Visible = False
    Label20.Visible = False
    Text13.Visible = False
    Text14.Visible = False
    Combo5.Visible = False
End If
If LCase(selected_primary_group) = LCase("Sundry Debtors") Or LCase(selected_primary_group) = LCase("Sundry Creditors") Or LCase(selected_primary_group) = LCase("Deposits(Liabilities)") Or LCase(selected_primary_group) = LCase("Loan and Advances(Assets)") Or LCase(selected_primary_group) = LCase("Capital Account") Then
    Call unlock_address_label_and_text
ElseIf LCase(selected_primary_group) <> LCase("Sundry Debtors") Or LCase(selected_primary_group) <> LCase("Sundry Creditors") Or LCase(selected_primary_group) <> LCase("Deposits(Liabilities)") Or LCase(selected_primary_group) <> LCase("Loan and Advances(Assets)") Or LCase(selected_primary_group) <> LCase("Capital Account") Then
    Call lock_address_label_and_text
End If
End Sub
Public Sub lock_address_label_and_text()

Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Text7.Visible = False
Text8.Visible = False
Text9.Visible = False
Text10.Visible = False
Text11.Visible = False

Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
Label6.Enabled = False
Label7.Enabled = False
Label8.Enabled = False
Label9.Enabled = False
Label10.Enabled = False
Label11.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False

End Sub
Public Sub unlock_address_label_and_text()

Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text8.Visible = True
Text9.Visible = True
Text10.Visible = True
Text11.Visible = True

Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Label6.Enabled = True
Label7.Enabled = True
Label8.Enabled = True
Label9.Enabled = True
Label10.Enabled = True
Label11.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True

End Sub
Private Sub Combo1_Click() 'WHEN SELECT GROUP

Call set_detail_as_per_primary_group

Call open_database
Call open_rs_lgr_main_grp
Call open_rs_lgr_prim_grp
    
    Do Until rs_lgr_main_grp.EOF
        If rs_lgr_main_grp!lgr_main_grp_name = Combo1.Text Then
        Do Until rs_lgr_prim_grp.EOF
            If rs_lgr_main_grp!lgr_main_grp_pgrp = rs_lgr_prim_grp!lgr_prim_grp_name Then
            Combo2.Text = rs_lgr_prim_grp!lgr_prim_grp_side
            'combo4.Text = rs_lgr_prim_grp!lgr_prim_grp_side
            Exit Sub
            End If
            rs_lgr_prim_grp.MoveNext
        Loop
        End If
    rs_lgr_main_grp.MoveNext
    Loop
    
    Call open_rs_lgr_prim_grp
    Do Until rs_lgr_prim_grp.EOF
    If rs_lgr_prim_grp!lgr_prim_grp_name = Combo1.Text Then
    Combo2.Text = rs_lgr_prim_grp!lgr_prim_grp_side
    'combo4.Text = rs_lgr_prim_grp!lgr_prim_grp_side
    Exit Sub
    End If
    rs_lgr_prim_grp.MoveNext
    Loop

End Sub
Private Sub Combo3_Click()
Call blank_all_text
Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
    If rs_lgr_main_dtl!lgr_main_dtl_name = Combo3.Text Or rs_lgr_main_dtl!lgr_main_dtl_alis = Combo3.Text Then
        Call set_selected_ledger_detail
    Exit Sub
    End If
rs_lgr_main_dtl.MoveNext
Loop
End Sub

Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================

Label18.Visible = False
Label19.Visible = False
Text17.Visible = False
Text18.Visible = False
Combo4.Visible = False
'disabled all label18,label19,text17,text18 & combo4 becuase don't want balance 2(black balance)

Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)
'selected_procedure = "ledger_edit"
'selected_procedure = "ledger_creat"
lbl_name.Caption = co_name
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
'Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")
If selected_path = "" Or selected_path = Null Then
    selected_path = App.Path & "\data\1000\co.mdb;"
End If
Call arrange_form_item
Call lock_address_label_and_text
End Sub
Public Sub arrange_form_item()
Call clear_all_combos_and_labels
Call blank_all_text
Call set_label_caption
Call add_combo1_main_grp
Call set_combo1_2_4_5
If selected_procedure = "ledger_edit" Then
Combo3.Visible = True
Label17.Visible = True
ElseIf selected_procedure = "ledger_creat" Then
Combo3.Visible = False
Label17.Visible = False
ElseIf selected_procedure = "ledger_display" Then
Combo3.Visible = True
Label17.Visible = True
Call enable_texts_and_combos
cmd_save.Enabled = False
End If
Combo3.Text = "Select a ledger..!!!"
End Sub
Public Sub add_combo1_main_grp()
Call open_database
Call open_rs_lgr_prim_grp
Do Until rs_lgr_prim_grp.EOF
    Combo1.AddItem rs_lgr_prim_grp!lgr_prim_grp_name
    rs_lgr_prim_grp.MoveNext
Loop
Call open_rs_lgr_main_grp
Do Until rs_lgr_main_grp.EOF
    Combo1.AddItem rs_lgr_main_grp!lgr_main_grp_name
        If rs_lgr_main_grp!lgr_main_grp_alis <> "" Then
        Combo1.AddItem rs_lgr_main_grp!lgr_main_grp_alis
        End If
    rs_lgr_main_grp.MoveNext
Loop

Call SortList(Combo1, Val(0) \ 1, (Val(Combo1.ListCount) - 1) \ 1, Ascending)
Combo1.Text = "Select Group"

Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
    Combo3.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
    If rs_lgr_main_dtl!lgr_main_dtl_alis <> "" Then Combo3.AddItem rs_lgr_main_dtl!lgr_main_dtl_alis
rs_lgr_main_dtl.MoveNext
Loop
Call SortList(Combo1, Val(0) \ 1, (Val(Combo1.ListCount) - 1) \ 1, Ascending)
Call SortList(Combo3, Val(0) \ 1, (Val(Combo3.ListCount) - 1) \ 1, Ascending)
End Sub
Public Sub clear_all_combos_and_labels()
Combo1.Clear
Combo2.Clear
Combo3.Clear
'combo4.Clear
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""
Label13.Caption = ""
Label14.Caption = ""
Label15.Caption = ""
Label16.Caption = ""
'label17.Caption = ""
'label18.Caption = ""
'label19.Caption = ""
Label20.Caption = ""

Text1.FontSize = 12
Text2.FontSize = 12
Text3.FontSize = 12
Text4.FontSize = 12
Text5.FontSize = 12
Text6.FontSize = 12
Text7.FontSize = 12
Text8.FontSize = 12
Text9.FontSize = 12
Text10.FontSize = 12
Text11.FontSize = 12
Text12.FontSize = 12
Text13.FontSize = 12
Text14.FontSize = 12
Text15.FontSize = 12
Text16.FontSize = 12
'Text17.FontSize = 12
'Text18.FontSize = 12
End Sub
Public Sub blank_all_text()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
'Text17.Text = ""
'Text18.Text = ""
End Sub
Public Sub enable_texts_and_combos()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
'Text17.Enabled = False
'Text18.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
'Combo3.Enabled = False
'combo4.Enabled = False
End Sub
Public Sub set_label_caption()
Label1.Caption = "Name"
Label2.Caption = "Alias"
Label3.Caption = "Address"
Label4.Caption = "Street"
Label5.Caption = "City"
Label6.Caption = "Pin-Code"
Label7.Caption = "Travel Detail"
Label8.Caption = "Telephone 1"
Label9.Caption = "Telephone 2"
Label10.Caption = "Mobile"
Label11.Caption = "E-mail"
Label12.Caption = "Group"
Label13.Caption = "Max.Cr. Period"
Label14.Caption = "Max.Cr. Limit"
Label15.Caption = "Opening Balance"
Label16.Caption = "Cr/Dr"
'label17.Caption = "Select a ledger"
'label18.Caption = "Opening Balance 2"
'label19.Caption = "Cr/Dr"
Label20.Caption = "Sales man"

'hide the customers options
Label13.Visible = False
Label14.Visible = False
Label20.Visible = False

Text13.Visible = False
Text14.Visible = False
Combo5.Visible = False
End Sub

Public Sub set_combo1_2_4_5()
Combo1.Left = Text12.Left
Combo1.Top = Text12.Top
Combo1.Width = Text12.Width
Combo1.FontSize = 12
Combo2.Left = Text16.Left
Combo2.Top = Text16.Top
Combo2.Width = Text16.Width
Combo2.FontSize = 12
'combo4.Left = Text18.Left
'combo4.Top = Text18.Top
'combo4.Width = Text18.Width
'combo4.FontSize = 12
Combo2.AddItem "Dr"
Combo2.AddItem "Cr"
Combo2.Text = "Cr/Dr"
'combo4.AddItem "Dr"
'combo4.AddItem "Cr"
'combo4.Text = "Cr/Dr"
Call open_rs_emp_main_dtl
Do Until rs_emp_main_dtl.EOF
Combo5.AddItem rs_emp_main_dtl!emp_main_dtl_name
rs_emp_main_dtl.MoveNext
Loop
Combo5.Text = "Select Sales Under"

End Sub

Public Sub save_ledger_detail()
    rs_lgr_main_dtl!lgr_main_dtl_name = Text1.Text
    rs_lgr_main_dtl!lgr_main_dtl_alis = Text2.Text
    rs_lgr_main_dtl!lgr_main_dtl_add1 = Text3.Text
    rs_lgr_main_dtl!lgr_main_dtl_add2 = Text4.Text
    rs_lgr_main_dtl!lgr_main_dtl_city = Text5.Text
    rs_lgr_main_dtl!lgr_main_dtl_pncd = Text6.Text
    rs_lgr_main_dtl!lgr_main_dtl_trnp = Text7.Text
    rs_lgr_main_dtl!lgr_main_dtl_tel1 = Text8.Text
    rs_lgr_main_dtl!lgr_main_dtl_tel2 = Text9.Text
    rs_lgr_main_dtl!lgr_main_dtl_mobl = Text10.Text
    rs_lgr_main_dtl!lgr_main_dtl_emal = Text11.Text
    rs_lgr_main_dtl!lgr_main_dtl_grup = Combo1.Text
    If Text13.Text = "" Then Text13.Text = 0
        rs_lgr_main_dtl!lgr_main_dtl_crpd = Text13.Text
    If Text14.Text = "" Then Text14.Text = 0
        rs_lgr_main_dtl!lgr_main_dtl_cram = Text14.Text
    If Text15.Text = "" Then Text15.Text = 0
        rs_lgr_main_dtl!lgr_main_dtl_obl1 = Text15.Text
    'If Text17.Text = "" Then Text17.Text = 0
    'rs_lgr_main_dtl!lgr_main_dtl_obl2 = Text17.Text
    rs_lgr_main_dtl!lgr_main_dtl_osd1 = Combo2.Text
    'rs_lgr_main_dtl!lgr_main_dtl_osd2 = Combo4.Text
    If LCase(selected_primary_group) = LCase("Sundry Debtors") Then rs_lgr_main_dtl!lgr_main_dtl_slun = Combo5.Text
End Sub
Public Sub set_selected_ledger_detail()
       Text1.Text = rs_lgr_main_dtl!lgr_main_dtl_name
       Text2.Text = rs_lgr_main_dtl!lgr_main_dtl_alis
       Text3.Text = rs_lgr_main_dtl!lgr_main_dtl_add1
       Text4.Text = rs_lgr_main_dtl!lgr_main_dtl_add2
       Text5.Text = rs_lgr_main_dtl!lgr_main_dtl_city
       Text6.Text = rs_lgr_main_dtl!lgr_main_dtl_pncd
       Text7.Text = rs_lgr_main_dtl!lgr_main_dtl_trnp
       Text8.Text = rs_lgr_main_dtl!lgr_main_dtl_tel1
       Text9.Text = rs_lgr_main_dtl!lgr_main_dtl_tel2
       Text10.Text = rs_lgr_main_dtl!lgr_main_dtl_mobl
       Text11.Text = rs_lgr_main_dtl!lgr_main_dtl_emal
       Combo1.Text = rs_lgr_main_dtl!lgr_main_dtl_grup
       If rs_lgr_main_dtl!lgr_main_dtl_crpd <> 0 Then Text13.Text = rs_lgr_main_dtl!lgr_main_dtl_crpd
       If rs_lgr_main_dtl!lgr_main_dtl_cram <> 0 Then Text14.Text = rs_lgr_main_dtl!lgr_main_dtl_cram
       If rs_lgr_main_dtl!lgr_main_dtl_obl1 <> 0 Then Text15.Text = rs_lgr_main_dtl!lgr_main_dtl_obl1
       'If rs_lgr_main_dtl!lgr_main_dtl_obl2 <> 0 Then Text17.Text = rs_lgr_main_dtl!lgr_main_dtl_obl2
       Combo2.Text = rs_lgr_main_dtl!lgr_main_dtl_osd1
       'combo4.Text = rs_lgr_main_dtl!lgr_main_dtl_osd2
    'If rs_lgr_main_dtl!lgr_main_dtl_osd1 <> Null Then Combo2.Text = rs_lgr_main_dtl!lgr_main_dtl_osd1
    'If rs_lgr_main_dtl!lgr_main_dtl_osd2 <> Null Then Combo4.Text = rs_lgr_main_dtl!lgr_main_dtl_osd2
    Call set_detail_as_per_primary_group
End Sub

Private Sub Form_Unload(Cancel As Integer)
'selected_procedure = "ledger_edit"
'selected_procedure = "ledger_creat"

End Sub

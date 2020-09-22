VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form vchr_payment 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Payment Voucher"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   Icon            =   "voucher_payment.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   11895
      Begin VB.ListBox list_lgr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   375
         Left            =   10320
         TabIndex        =   37
         Top             =   3050
         Width           =   1335
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
         Height          =   420
         Left            =   8640
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox text_amt 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   5160
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox combo_lgr 
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
         Index           =   2
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   18
         Text            =   "Combo2"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox text_amt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   6480
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text4 
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
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "voucher_payment.frx":1D2A
         Top             =   2280
         Width           =   6375
      End
      Begin VB.TextBox Text5 
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
         Left            =   8640
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton cmd_sv_n_new 
         Caption         =   "&Save and New"
         Height          =   450
         Left            =   10320
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "Edit"
         Height          =   450
         Left            =   10320
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "Cancel"
         Height          =   450
         Left            =   10320
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmd_save_n_exit 
         Caption         =   "Save and exit"
         Height          =   450
         Left            =   10320
         TabIndex        =   11
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Pirnt"
         Height          =   450
         Left            =   10320
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox Combo0 
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
         Left            =   8640
         TabIndex        =   9
         Text            =   "Combo0"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox combo_sub_entry_no 
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
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Text            =   "Paid by"
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox combo_sub_entry_no 
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
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Text            =   "Paid to"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox combo_lgr 
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
         Index           =   1
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   10320
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   8640
         TabIndex        =   21
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
         CurrentDate     =   40166
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
         ForeColor       =   &H00004040&
         Height          =   495
         Left            =   7920
         TabIndex        =   33
         Top             =   240
         Width           =   975
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
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   7920
         TabIndex        =   32
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   8280
         TabIndex        =   31
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   8160
         TabIndex        =   30
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Dr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   5400
         TabIndex        =   29
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   6840
         TabIndex        =   28
         Top             =   1800
         Width           =   855
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
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   1215
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
         ForeColor       =   &H00004040&
         Height          =   495
         Left            =   7920
         TabIndex        =   26
         Top             =   2760
         Width           =   975
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
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   5400
         TabIndex        =   25
         Top             =   120
         Width           =   855
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
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   6720
         TabIndex        =   24
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label0 
         BackStyle       =   0  'Transparent
         Caption         =   "Label0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   7920
         TabIndex        =   23
         Top             =   1920
         Width           =   735
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
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Frame2"
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   11895
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   4471
         _Version        =   393216
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   10440
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lbl_head 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   495
      Left            =   195
      TabIndex        =   36
      Top             =   360
      Width           =   12000
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   375
      Left            =   240
      TabIndex        =   35
      Top             =   240
      Width           =   12000
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name of company"
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
      Left            =   195
      TabIndex        =   34
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "vchr_payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public combo_temp_index
Public index_x

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub text_amt_Click(index As Integer)
If index <> 1 Then
If text_amt(index - 1) = "" Then text_amt(index - 1) = "0.00"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
temp_selected_procedure = selected_procedure
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
Private Sub list_lgr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 39 Then
list_lgr.Visible = False
combo_temp_index = list_lgr.ListIndex
text_amt(index_x).SetFocus
ElseIf KeyCode = 27 Then
combo_lgr(index_x).ListIndex = combo_temp_index
text_amt(index_x).SetFocus
End If
End Sub
Private Sub combo_lgr_GotFocus(index As Integer)
list_lgr.TabIndex = combo_lgr(index).TabIndex
combo_temp_index = combo_lgr(index).ListIndex
index_x = index
list_lgr.Visible = True
list_lgr.Left = combo_lgr(index).Left
list_lgr.Top = combo_lgr(index).Top
list_lgr.Width = combo_lgr(index).Width
For list_lgr_temp = 0 To combo_lgr(index).ListCount - 1
combo_lgr(index).ListIndex = list_lgr_temp
list_lgr.AddItem combo_lgr(index).Text
Next
list_lgr.Height = 2000
list_lgr.SetFocus
End Sub
Private Sub list_lgr_LostFocus()
If combo_temp_index < 0 Then combo_temp_index = 0
    combo_lgr(index_x).ListIndex = combo_temp_index
    Call open_database
    Call open_rs_lgr_main_dtl
    Do Until rs_lgr_main_dtl.EOF
    If combo_lgr(index_x).Text = rs_lgr_main_dtl!lgr_main_dtl_alis Then
        combo_lgr(index_x).Text = rs_lgr_main_dtl!lgr_main_dtl_name
    End If
    rs_lgr_main_dtl.MoveNext
    Loop
    Dim lgr_is_available_or_not As Integer
    lgr_is_available_or_not = 0
    Call open_database
    Call open_rs_lgr_main_dtl
    Do Until rs_lgr_main_dtl.EOF
        If combo_lgr(index_x).Text = rs_lgr_main_dtl!lgr_main_dtl_name Then
        lgr_is_available_or_not = 1
        Else
        End If
        rs_lgr_main_dtl.MoveNext
    Loop
    If lgr_is_available_or_not = 0 Then
        MsgBox "You are Entered invalid ledger...!!! select proper account...!!!"
        combo_lgr(index_x).Text = "select a ledger"
        combo_lgr(index_x).SetFocus
        Exit Sub
    End If
    list_lgr.Visible = False
    list_lgr.Clear
End Sub
Private Sub cmd_delete_Click()
Dim delete_sure
delete_sure = MsgBox("You want to delete voucher....?", vbQuestion + vbYesNo, "Are You Sure !!!!")
If delete_sure = 6 Then
Call delete_transaction
cmd_delete.Enabled = False
cmd_edit.Enabled = False
selected_procedure = "Payment voucher"
xsub_entry_no = 0
DTPicker1.Value = Date
voucher_total_cr_amt = 0
voucher_total_dr_amt = 0
current_sub_entry_no = 0
sub_entry_no = 1
dr_sub_entry_no = 1
cr_sub_entry_no = 1
Text5.Text = selected_user
Call set_form_headings
Call set_form_labels
Call set_vourcher_detail
Call unlock_all_combo_text
End If
End Sub
Public Sub delete_transaction()
Call open_rs_acn_tran_pmt
Do Until rs_acn_tran_pmt.EOF
If rs_acn_tran_pmt!fin_acnt_trn_vcno = Text1.Text Then
rs_acn_tran_pmt.Delete
rs_acn_tran_pmt.UpdateBatch
this_entry_is_saved = 0
End If
rs_acn_tran_pmt.MoveNext
Loop
Call open_rs_acn_tran_all
Do Until rs_acn_tran_all.EOF
If rs_acn_tran_all!fin_acnt_trn_vcno = Text1.Text And rs_acn_tran_all!fin_acnt_trn_vchr = "Payment" Then
rs_acn_tran_all.Delete
rs_acn_tran_all.UpdateBatch
this_entry_is_saved = 0
End If
rs_acn_tran_all.MoveNext
Loop
Call open_rs_acn_tran_all_temp
Do Until rs_acn_tran_all_temp.EOF
If rs_acn_tran_all_temp!fin_acnt_trn_vcno = Text1.Text And rs_acn_tran_all_temp!fin_acnt_trn_vchr = "Payment" Then
rs_acn_tran_all_temp.Delete
rs_acn_tran_all_temp.UpdateBatch
this_entry_is_saved = 0
End If
rs_acn_tran_all_temp.MoveNext
Loop
End Sub
Private Sub cmd_edit_Click()
change_the_old_voucher = 1
Call unlock_all_combo_text
End Sub
Private Sub cmd_save_n_exit_Click()
If Text1.Text = "" Or _
Val(text_amt(sub_entry_no).Text) < 0 Or _
Val(text_amt(sub_entry_no).Text) < 0 Or _
combo_lgr(sub_entry_no).Text = "select a ledger" Or _
combo_lgr(sub_entry_no).Text = "select a ledger" Or _
DTPicker1.Value > Date Or _
DTPicker1.Value < this_year_starting_date Then 'Combo0.Text > 2 Or Combo0.Text < 1 Or
MsgBox "You have not entered proper or sufficient detail...!!!"
Exit Sub
End If
Call save_new_transaction
Unload Me
End Sub
Private Sub combo_lgr_Click(index As Integer)
If index <> 1 Then
If text_amt(index - 1) = "" Then text_amt(index - 1) = "0.00"
End If
End Sub
Private Sub Form_Load()
'this is a code for sizing===================================
'    RePosForm = True   ' Flag for positioning Form
'    DoResize = False   ' Flag for Resize Event
'Call set_screen_resolution
'Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================

Label0.Visible = False
Combo0.Visible = False

list_lgr.Top = -340
selected_procedure = "Payment voucher"
cmd_delete.Enabled = False
cmd_edit.Enabled = False
xsub_entry_no = 0
DTPicker1.Value = Date
voucher_total_cr_amt = 0
voucher_total_dr_amt = 0
current_sub_entry_no = 0
sub_entry_no = 1
dr_sub_entry_no = 1
cr_sub_entry_no = 1
Text5.Text = selected_user
Call set_form_headings
Call set_form_labels
Call set_vourcher_detail
Call voucher_type_1_tab_indexing
If show_ledger_detail = 1 Then
Text1.Text = selected_voucher_no
Call search_voucher_and_show_detail
End If
End Sub
Public Sub set_vourcher_detail()
transaction_type = "Payment"
Text1.Text = "" 'voucher no
text_amt(sub_entry_no).Text = "" 'amount
Text4.Text = "" 'narration
Text1.Enabled = False
Text5.Enabled = False
'combo0.Text = "" 'transaction type 1/2
Frame1.Caption = "Current Transaction Detail"
selected_date = DTPicker1.Value
Frame2.Caption = selected_date & "s Transactions Detail"
Call add_account_combo0
Call start_new_voucher
End Sub
Public Sub start_new_voucher()
sub_entry_no = 1
total_sub_entry_no = 2
dr_sub_entry_no = 1
this_entry = "cr"
Call add_combo_sub_entry_no(sub_entry_no)
Call add_account_combo_lgr(sub_entry_no)
Call add_text_amt(sub_entry_no)
this_entry = "dr"
sub_entry_no = sub_entry_no + 1
cr_sub_entry_no = 1
Call add_combo_sub_entry_no(sub_entry_no)
Call add_account_combo_lgr(sub_entry_no)
Call add_text_amt(sub_entry_no)
Call reset_voucher_detail
End Sub
Public Sub reset_voucher_detail()
Call remove_controls
Call open_database
Call open_rs_acn_tran_pmt
Call find_last_voucher_no
If rs_acn_tran_pmt.RecordCount = 0 Then Text1.Text = 1
Text4.Text = ""
Label3.Caption = WeekdayName(Weekday(DTPicker1.Value - 1)) ' Day(Weekday(Now))
Label4.Caption = Time
Call arrange_grid1
Call open_grid1
End Sub
Private Sub DTPicker1_Change()
selected_date = DTPicker1.Value
Label3.Caption = WeekdayName(Weekday(DTPicker1.Value - 1)) ' Day(Weekday(Now))
Label4.Caption = Time
Frame2.Caption = selected_date & "s Transactions Detail"
If change_the_old_voucher <> 1 Then Call read_all_dated_transaction
End Sub
Public Sub read_all_dated_transaction()
'Call remove_controls
'Call start_new_voucher
'Call reset_voucher_detail
Call refresh_grid1
'Call refresh_dr_cr_total_amt
'Call move_all_command_to_bottom
End Sub
Private Sub MSFlexGrid1_Click()
'Call read_current_transaction
End Sub
Public Sub add_account_combo0()
'Combo0.AddItem "1"
'Combo0.AddItem "2"
'Combo0.Text = "2"
End Sub
Public Sub add_combo_sub_entry_no(sub_entry_no)
combo_sub_entry_no(sub_entry_no).Left = 299
combo_sub_entry_no(sub_entry_no).Top = (sub_entry_no * 600)
combo_sub_entry_no(sub_entry_no).Visible = True
combo_sub_entry_no(sub_entry_no).AddItem "Cr / To"
combo_sub_entry_no(sub_entry_no).AddItem "Dr / By"
If LCase(this_entry) = "cr" Then
combo_sub_entry_no(sub_entry_no).Text = "Cr / To"
ElseIf LCase(this_entry) = "dr" Then
combo_sub_entry_no(sub_entry_no).Text = "Dr / By"
End If
End Sub
Public Sub add_account_combo_lgr(sub_entry_no)
combo_lgr(sub_entry_no).Text = "select a ledger"
combo_lgr(sub_entry_no).Left = Frame1.Left + 1499
combo_lgr(sub_entry_no).Top = (sub_entry_no * 600)
combo_lgr(sub_entry_no).Visible = True
If LCase(this_entry) = "dr" Then
Call add_cr_ledger(sub_entry_no)
ElseIf LCase(this_entry) = "cr" Then
Call add_dr_ledger(sub_entry_no)
End If
End Sub
Public Sub add_cr_ledger(sub_entry_no)
combo_lgr(sub_entry_no).Clear
Call open_database
Call add_other_then_cash_or_bank_ledgers
combo_lgr(sub_entry_no).Text = "select a ledger"
End Sub
Public Sub add_dr_ledger(sub_entry_no)
combo_lgr(sub_entry_no).Clear
Call add_cash_or_bank_ledgers
combo_lgr(sub_entry_no).Text = "select a ledger"
End Sub
Public Sub add_text_amt(sub_entry_no)
text_amt(sub_entry_no).Top = (sub_entry_no * 600)
text_amt(sub_entry_no).Visible = True
text_amt(sub_entry_no).Text = "" 'amount
If LCase(this_entry) = "dr" Then
Call cr_amt_text_adjust(sub_entry_no)
ElseIf LCase(this_entry) = "cr" Then
dr_amt_text_adjust (sub_entry_no)
End If
End Sub
Public Sub cr_amt_text_adjust(sub_entry_no)
text_amt(sub_entry_no).Left = 6399
End Sub
Public Sub dr_amt_text_adjust(sub_entry_no)
text_amt(sub_entry_no).Left = 5199
End Sub
Public Sub set_form_headings()
lbl_name.Width = Me.Width
lbl_name.Left = 0
lbl_name.Caption = co_name
lbl_add.Width = Me.Width
lbl_add.Top = -1000
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
lbl_head.Width = Me.Width
lbl_head.Left = 0
lbl_head.Caption = UCase(selected_procedure)
Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)
End Sub
Public Sub set_form_labels()
'Label0.Caption = "Type"
Label1.Caption = "No"
Label2.Caption = "Date"
'Label3.Caption = "Day:" & Day(DTPicker1.Value)
'Label4.Caption = "Time:" & Time
Label7.Caption = "Narration"
Label8.Caption = "User"
Label9.Caption = "Amount"
Label10.Caption = "Amount"
Label11.Caption = "Paid"
End Sub
Private Sub text_amt_cr_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
text_amt(sub_entry_no).Text = Format(text_amt(sub_entry_no).Text, "0.00")
End If
End Sub
Private Sub Grid1_Click()
If show_ledger_detail = 0 Then
If Grid1.TextMatrix(Grid1.Row, 0) = "" Then
MsgBox "This is a invalid entry....,"
Exit Sub
End If
change_the_old_voucher = 0
current_sub_entry_no = 0
total_sub_entry_no = 0
sub_entry_no = 1
Dim entry_no
For entry_no = 1 To 12
If Val(Grid1.TextMatrix(Grid1.Row, 2)) <> 0 Then
selected_voucher_no = Grid1.TextMatrix(Grid1.Row, 2)
Exit For
ElseIf Val(Grid1.TextMatrix(Grid1.Row - entry_no, 2)) <> 0 Then
selected_voucher_no = Grid1.TextMatrix(Grid1.Row - entry_no, 2)
Exit For
End If
Next
Dim control_counter
For control_counter = 3 To xsub_entry_no ' To 2 Step -1
Unload combo_sub_entry_no(control_counter)
Unload combo_lgr(control_counter)
Unload text_amt(control_counter)
Next
Call search_voucher_and_show_detail
End If
End Sub
Public Sub search_voucher_and_show_detail()
change_the_old_voucher = 0
current_sub_entry_no = 0
total_sub_entry_no = 0
sub_entry_no = 1
Call open_database
Call open_rs_acn_tran_pmt
Do Until rs_acn_tran_pmt.EOF
If rs_acn_tran_pmt!fin_acnt_trn_vcno = selected_voucher_no Then
With rs_acn_tran_pmt
If sub_entry_no = 1 Then
Text1.Text = !fin_acnt_trn_vcno
'Combo0.Text = !fin_acnt_trn_vtyp
DTPicker1.Value = !fin_acnt_trn_date
Label4.Caption = !fin_acnt_trn_time
Label3.Caption = !fin_acnt_trn_wday
Text4.Text = !fin_acnt_trn_nrtn
Text5.Text = !fin_acnt_trn_user
If !fin_acnt_trn_side = "dr" Then
Call add_other_then_cash_or_bank_ledgers
text_amt(sub_entry_no).Left = 6399
combo_sub_entry_no(sub_entry_no).Text = "Dr / By"
ElseIf !fin_acnt_trn_side = "cr" Then
Call add_cash_or_bank_ledgers
text_amt(sub_entry_no).Left = 5199
combo_sub_entry_no(sub_entry_no).Text = "Cr / To"
End If
combo_lgr(sub_entry_no).Text = !fin_acnt_trn_ldgr
text_amt(sub_entry_no).Text = Format(!fin_acnt_trn_amnt, "0.00")
sub_entry_no = sub_entry_no + 1
total_sub_entry_no = 1
ElseIf sub_entry_no = 2 Then
If !fin_acnt_trn_side = "dr" Then
Call add_other_then_cash_or_bank_ledgers
text_amt(sub_entry_no).Left = 6399
combo_sub_entry_no(sub_entry_no).Text = "Dr / By"
ElseIf !fin_acnt_trn_side = "cr" Then
Call add_cash_or_bank_ledgers
text_amt(sub_entry_no).Left = 5199
combo_sub_entry_no(sub_entry_no).Text = "Cr / To"
End If
combo_lgr(sub_entry_no).Text = !fin_acnt_trn_ldgr
text_amt(sub_entry_no).Text = Format(!fin_acnt_trn_amnt, "0.00")
Call move_all_command_to_bottom
sub_entry_no = sub_entry_no + 1
total_sub_entry_no = total_sub_entry_no + 1
ElseIf sub_entry_no > 2 Then
Load combo_sub_entry_no(sub_entry_no)
Load combo_lgr(sub_entry_no)
Load text_amt(sub_entry_no)
If !fin_acnt_trn_side = "cr" Then
this_entry = "cr"
cr_sub_entry_no = 1
combo_sub_entry_no(sub_entry_no).Text = "Cr / To"
combo_lgr(sub_entry_no).Text = "select a ledger"
combo_lgr(sub_entry_no).Left = Frame1.Left + 1499
combo_lgr(sub_entry_no).Top = (sub_entry_no * 600)
combo_lgr(sub_entry_no).Visible = True
combo_sub_entry_no(sub_entry_no).Left = 299
combo_sub_entry_no(sub_entry_no).Top = (sub_entry_no * 600)
combo_sub_entry_no(sub_entry_no).Visible = True
text_amt(sub_entry_no).Top = (sub_entry_no * 600)
text_amt(sub_entry_no).Visible = True
text_amt(sub_entry_no).Left = 5199
combo_sub_entry_no(sub_entry_no).Text = "Cr / To"
Call move_all_command_to_bottom
If !fin_acnt_trn_side = "dr" Then
Call add_other_then_cash_or_bank_ledgers
text_amt(sub_entry_no).Left = 6399
combo_sub_entry_no(sub_entry_no).Text = "Dr / By"
ElseIf !fin_acnt_trn_side = "cr" Then
Call add_cash_or_bank_ledgers
text_amt(sub_entry_no).Left = 5199
combo_sub_entry_no(sub_entry_no).Text = "Cr / To"
End If
combo_lgr(sub_entry_no).Text = !fin_acnt_trn_ldgr
text_amt(sub_entry_no).Text = Format(!fin_acnt_trn_amnt, "0.00")
sub_entry_no = sub_entry_no + 1
ElseIf !fin_acnt_trn_side = "dr" Then
this_entry = "dr"
combo_sub_entry_no(sub_entry_no).Text = "Dr / By"
this_entry = "cr"
cr_sub_entry_no = 1
combo_sub_entry_no(sub_entry_no).Text = "Cr / To"
combo_lgr(sub_entry_no).Text = "select a ledger"
combo_lgr(sub_entry_no).Left = Frame1.Left + 1499
combo_lgr(sub_entry_no).Top = (sub_entry_no * 600)
combo_lgr(sub_entry_no).Visible = True
combo_sub_entry_no(sub_entry_no).Left = 299
combo_sub_entry_no(sub_entry_no).Top = (sub_entry_no * 600)
combo_sub_entry_no(sub_entry_no).Visible = True
text_amt(sub_entry_no).Top = (sub_entry_no * 600)
text_amt(sub_entry_no).Visible = True
text_amt(sub_entry_no).Left = 6399
Call move_all_command_to_bottom
combo_sub_entry_no(sub_entry_no).Text = "Dr / By"
If !fin_acnt_trn_side = "dr" Then
Call add_other_then_cash_or_bank_ledgers
text_amt(sub_entry_no).Left = 6399
combo_sub_entry_no(sub_entry_no).Text = "Dr / By"
ElseIf !fin_acnt_trn_side = "cr" Then
Call add_cash_or_bank_ledgers
text_amt(sub_entry_no).Left = 5199
combo_sub_entry_no(sub_entry_no).Text = "Cr / To"
End If
combo_lgr(sub_entry_no).Text = !fin_acnt_trn_ldgr
text_amt(sub_entry_no).Text = Format(!fin_acnt_trn_amnt, "0.00")
sub_entry_no = sub_entry_no + 1
End If
total_sub_entry_no = total_sub_entry_no + 1
End If
End With
End If
rs_acn_tran_pmt.MoveNext
Loop
sub_entry_no = sub_entry_no - 1
xsub_entry_no = sub_entry_no
Call refresh_dr_cr_total_amt
Call lock_all_combo_text
If show_ledger_detail = 1 Then
selected_date = DTPicker1.Value
Frame2.Caption = selected_date & "s Transactions Detail"
Call refresh_grid1
Call refresh_dr_cr_total_amt
Call move_all_command_to_bottom
End If
End Sub
Public Sub lock_all_combo_text()
Dim ix
For ix = 1 To sub_entry_no
combo_sub_entry_no(ix).Enabled = False
combo_lgr(ix).Enabled = False
combo_sub_entry_no(ix).Enabled = False
text_amt(ix).Enabled = False
Text4.Enabled = False
Next
cmd_delete.Enabled = True
cmd_edit.Enabled = True
End Sub
Public Sub unlock_all_combo_text()
Dim ix
For ix = 1 To sub_entry_no
combo_sub_entry_no(ix).Enabled = True
combo_lgr(ix).Enabled = True
combo_sub_entry_no(ix).Enabled = True
text_amt(ix).Enabled = True
Text4.Enabled = True
Next
End Sub
Private Sub Label7_Click()
MsgBox Label7.Top
End Sub
Private Sub cmd_sv_n_new_Click()
If Text1.Text = "" Or _
Val(text_amt(sub_entry_no).Text) < 0 Or _
Val(text_amt(sub_entry_no).Text) < 0 Or _
combo_lgr(sub_entry_no).Text = "select a ledger" Or _
combo_lgr(sub_entry_no).Text = "select a ledger" Or _
DTPicker1.Value > Date Or _
DTPicker1.Value < this_year_starting_date Then 'Combo0.Text > 2 Or Combo0.Text < 1 Or
MsgBox "You have not entered proper or sufficient detail...!!!"
Exit Sub
End If
Call save_new_transaction
Call find_last_voucher_no
Call remove_controls
selected_procedure = "Payment voucher"
xsub_entry_no = 0
DTPicker1.Value = Date
voucher_total_cr_amt = 0
voucher_total_dr_amt = 0
current_sub_entry_no = 0
sub_entry_no = 1
dr_sub_entry_no = 1
cr_sub_entry_no = 1
Text5.Text = selected_user
Label5.Caption = Format(voucher_total_dr_amt, "0.00")
Label6.Caption = Format(voucher_total_cr_amt, "0.00")
Call set_form_headings
Call set_form_labels
Call set_vourcher_detail
Call move_all_command_to_bottom
End Sub
Public Sub find_last_voucher_no()
Call open_database
Call open_rs_acn_tran_pmt
Dim iflvn
Dim this_voucher_no
Dim biggest_voucher_no
If rs_acn_tran_pmt.RecordCount > 0 Then
For iflvn = 1 To rs_acn_tran_pmt.RecordCount
this_voucher_no = rs_acn_tran_pmt!fin_acnt_trn_vcno
If this_voucher_no > biggest_voucher_no Then
   biggest_voucher_no = this_voucher_no
End If
rs_acn_tran_pmt.MoveNext
Next
End If
Text1.Text = biggest_voucher_no + 1
End Sub
Public Sub save_new_transaction()
Dim transaction_counter
Dim this_entry_is_saved
For transaction_counter = 1 To sub_entry_no
this_entry_is_saved = 0
Call open_database
Call open_rs_acn_tran_pmt 'MsgBox rs_acn_tran_pmt.RecordCount
For available_tran_no = 1 To rs_acn_tran_pmt.RecordCount
With rs_acn_tran_pmt
If .RecordCount > 0 Then
If .EOF = True Or .BOF = True Then Exit Sub
End If
If !fin_acnt_trn_vcno = Text1.Text And !fin_acnt_trn_seno = transaction_counter Then
'.EditMode
!fin_acnt_trn_vcno = Text1.Text
!fin_acnt_trn_seno = transaction_counter
'!fin_acnt_trn_vtyp = Combo0.Text
!fin_acnt_trn_date = DTPicker1.Value
!fin_acnt_trn_time = Label4.Caption
!fin_acnt_trn_wday = Label3.Caption
!fin_acnt_trn_ldgr = combo_lgr(transaction_counter).Text
!fin_acnt_trn_amnt = text_amt(transaction_counter).Text
If LCase(combo_sub_entry_no(transaction_counter).Text) = LCase("Cr / To") Then
!fin_acnt_trn_side = "cr"
ElseIf LCase(combo_sub_entry_no(transaction_counter).Text) = LCase("Dr / By") Then
!fin_acnt_trn_side = "dr"
End If
!fin_acnt_trn_nrtn = Text4.Text
!fin_acnt_trn_user = Text5.Text
!fin_acnt_trn_vchr = "Payment"
this_entry_is_saved = 1
End If
End With
rs_acn_tran_pmt.MoveNext
Next
If this_entry_is_saved <> 1 Then
Call open_database
Call open_rs_acn_tran_pmt
rs_acn_tran_pmt.AddNew
With rs_acn_tran_pmt
!fin_acnt_trn_vcno = Text1.Text
!fin_acnt_trn_seno = transaction_counter
'!fin_acnt_trn_vtyp = Combo0.Text
!fin_acnt_trn_date = DTPicker1.Value
!fin_acnt_trn_time = Label4.Caption
!fin_acnt_trn_wday = Label3.Caption
!fin_acnt_trn_ldgr = combo_lgr(transaction_counter).Text
!fin_acnt_trn_amnt = text_amt(transaction_counter).Text
If LCase(combo_sub_entry_no(transaction_counter).Text) = LCase("Cr / To") Then
!fin_acnt_trn_side = "cr"
ElseIf LCase(combo_sub_entry_no(transaction_counter).Text) = LCase("Dr / By") Then
!fin_acnt_trn_side = "dr"
End If
!fin_acnt_trn_nrtn = Text4.Text
!fin_acnt_trn_user = Text5.Text
!fin_acnt_trn_vchr = "Payment"
End With
End If
rs_acn_tran_pmt.UpdateBatch
Next
If db_co.State = 1 Then db_co.Close
'FileCopy selected_path, selected_backup_path
End Sub
Public Sub find_voucher_entry_is_available()
End Sub
Public Sub remove_controls()
Dim control_counter
For control_counter = sub_entry_no To 3 Step -1
Unload combo_sub_entry_no(control_counter)
Unload combo_lgr(control_counter)
Unload text_amt(control_counter)
Next
End Sub
Public Sub refresh_grid1()
Call arrange_grid1
Call open_grid1
End Sub
Public Sub arrange_grid1()
Grid1.RowHeightMin = 250
Grid1.Clear
Grid1.Rows = 2
Grid1.Cols = 12

Grid1.TextMatrix(0, 1) = "Type"
Grid1.TextMatrix(0, 2) = "V.No"
Grid1.TextMatrix(0, 3) = "Date"
Grid1.TextMatrix(0, 4) = "Day"
Grid1.TextMatrix(0, 5) = "Time"
Grid1.TextMatrix(0, 6) = "Dr / By"
Grid1.TextMatrix(0, 7) = "Amount"
Grid1.TextMatrix(0, 8) = "Cr / To"
Grid1.TextMatrix(0, 9) = "Amount"
Grid1.TextMatrix(0, 10) = "Nar."
Grid1.TextMatrix(0, 11) = "User"

    
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 200
    Grid1.ColWidth(2) = 500
    Grid1.ColWidth(3) = 1100
    Grid1.ColWidth(4) = 900
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 2000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 2000
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 900
    Grid1.ColWidth(11) = 600

Grid1.Font.Size = 10
End Sub
Public Sub open_grid1()
Dim saw_voucher_no
Call open_database
Call open_rs_acn_tran_pmt
rs_acn_tran_pmt.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno"
Dim data_no As Integer
data_no = 1
Do Until rs_acn_tran_pmt.EOF
With rs_acn_tran_pmt
If selected_date = !fin_acnt_trn_date Then
Grid1.TextMatrix(data_no, 0) = data_no
'Grid1.TextMatrix(data_no, 1) = !fin_acnt_trn_vtyp
If saw_voucher_no = !fin_acnt_trn_vcno Then
Else
saw_voucher_no = !fin_acnt_trn_vcno
Grid1.TextMatrix(data_no, 2) = !fin_acnt_trn_vcno
Grid1.TextMatrix(data_no, 3) = !fin_acnt_trn_date
Grid1.TextMatrix(data_no, 4) = !fin_acnt_trn_wday
Grid1.TextMatrix(data_no, 5) = !fin_acnt_trn_time
Grid1.TextMatrix(data_no, 10) = !fin_acnt_trn_nrtn
Grid1.TextMatrix(data_no, 11) = !fin_acnt_trn_user
End If
If LCase(!fin_acnt_trn_side) = "dr" Then
Grid1.TextMatrix(data_no, 7) = Format(!fin_acnt_trn_amnt, "0.00")
Grid1.TextMatrix(data_no, 6) = !fin_acnt_trn_ldgr
ElseIf LCase(!fin_acnt_trn_side) = "cr" Then
Grid1.TextMatrix(data_no, 9) = Format(!fin_acnt_trn_amnt, "0.00")
Grid1.TextMatrix(data_no, 8) = !fin_acnt_trn_ldgr
End If
data_no = data_no + 1
If rs_acn_tran_pmt.RecordCount < Grid1.Rows Then
Exit Sub
End If
Grid1.Rows = Grid1.Rows + 1
End If
End With
rs_acn_tran_pmt.MoveNext
Loop
End Sub
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub text_amt_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
current_sub_entry_no = index
If KeyCode = 13 And index = 1 And change_the_old_voucher = 0 Then
text_amt(index).Text = Format(text_amt(index).Text, "0.00")
text_amt(index + 1).Text = Format(text_amt(index).Text, "0.00")
ElseIf KeyCode = 13 And index > 1 And change_the_old_voucher = 0 Then
text_amt(index).Text = Format(text_amt(index).Text, "0.00")
End If
If KeyCode = 13 Then
If Val(text_amt(index).Text) <= 0 Then
MsgBox "Enter Correct value and try agian...,"
Exit Sub
End If
Call refresh_dr_cr_total_amt
If voucher_total_dr_amt = voucher_total_cr_amt Then
If combo_lgr(sub_entry_no).Text = "select a ledger" Then
combo_lgr(sub_entry_no).SetFocus
Exit Sub
End If
Text4.SetFocus
Exit Sub
End If
If voucher_total_dr_amt < voucher_total_cr_amt Then
If sub_entry_no > 12 Then
MsgBox "Only 12 Entries allowed Once....!!!!"
Exit Sub
End If
If combo_lgr(sub_entry_no - 1).Text = "" Or LCase(combo_lgr(sub_entry_no - 1).Text) = " " Or LCase(combo_lgr(sub_entry_no - 1).Text) = "select a ledger" _
Or Val(text_amt(sub_entry_no - 1)) < 0 Or Val(text_amt(sub_entry_no - 1)) = Null Then
MsgBox "Please, Enter proper values....!!!! and click again...!!!"
Exit Sub
End If
'If LCase(combo_sub_entry_no(Index).Text) = LCase("Cr / To") And Index >= 1 Then 'And Index = sub_entry_no Then
this_entry = "cr"
sub_entry_no = sub_entry_no + 1
total_sub_entry_no = total_sub_entry_no + 1
Load combo_sub_entry_no(sub_entry_no)
Load combo_lgr(sub_entry_no)
Load text_amt(sub_entry_no)
Call add_combo_sub_entry_no(sub_entry_no)
Call add_text_amt(sub_entry_no)
Call add_account_combo_lgr(sub_entry_no)
Call move_all_command_to_bottom
text_amt(sub_entry_no) = Format(voucher_total_cr_amt - voucher_total_dr_amt, "0.00")
'End If
ElseIf voucher_total_dr_amt > voucher_total_cr_amt Then
If sub_entry_no > 12 Then
MsgBox "Only 12 Entries allowed Once....!!!!"
Exit Sub
End If
If combo_lgr(sub_entry_no - 1).Text = "" Or LCase(combo_lgr(sub_entry_no - 1).Text) = " " Or LCase(combo_lgr(sub_entry_no - 1).Text) = "select a ledger" _
Or Val(text_amt(sub_entry_no - 1)) < 0 Or Val(text_amt(sub_entry_no - 1)) = Null Then
MsgBox "Please, Enter proper values....!!!! and click again...!!!"
Exit Sub
End If
this_entry = "dr"
sub_entry_no = sub_entry_no + 1
total_sub_entry_no = total_sub_entry_no + 1
Load combo_sub_entry_no(sub_entry_no)
Load combo_lgr(sub_entry_no)
Load text_amt(sub_entry_no)
Call add_combo_sub_entry_no(sub_entry_no)
Call add_text_amt(sub_entry_no)
Call add_account_combo_lgr(sub_entry_no)
Call move_all_command_to_bottom
text_amt(sub_entry_no) = Format(voucher_total_dr_amt - voucher_total_cr_amt, "0.00")
ElseIf voucher_total_dr_amt = voucher_total_cr_amt Then
Call refresh_dr_cr_total_amt
Text4.SetFocus
Exit Sub
End If
'If change_the_old_voucher <> 1 Then
'MsgBox sub_entry_no
combo_lgr(sub_entry_no).SetFocus
Call refresh_dr_cr_total_amt
End If
End Sub
Public Sub refresh_dr_cr_total_amt()
voucher_total_cr_amt = 0
voucher_total_dr_amt = 0
Dim int_i
For int_i = 1 To total_sub_entry_no
If LCase(combo_sub_entry_no(int_i).Text) = LCase("Dr / By") Then
voucher_total_cr_amt = voucher_total_cr_amt + Val(text_amt(int_i).Text)
ElseIf LCase(combo_sub_entry_no(int_i).Text) = LCase("Cr / To") Then
voucher_total_dr_amt = voucher_total_dr_amt + Val(text_amt(int_i).Text)
End If
If voucher_total_cr_amt = voucher_total_dr_amt And change_the_old_voucher = 1 Then
Call remove_all_the_data_after_this_point
Exit For
End If
Next
If change_the_old_voucher = 2 Then total_sub_entry_no = int_i
change_the_old_voucher = 0
Label5.Caption = Format(voucher_total_dr_amt, "0.00")
Label6.Caption = Format(voucher_total_cr_amt, "0.00")
End Sub
Public Sub remove_all_the_data_after_this_point()
Dim control_counter
For control_counter = current_sub_entry_no + 1 To total_sub_entry_no
Unload combo_sub_entry_no(control_counter)
Unload combo_lgr(control_counter)
Unload text_amt(control_counter)
sub_entry_no = sub_entry_no - 1
Label7.Top = (sub_entry_no * 600) + 1000
Text4.Top = (sub_entry_no * 600) + 800
Frame1.Height = (sub_entry_no * 600) + 2300
Frame2.Top = (sub_entry_no * 600) + 2999
Label5.Top = (sub_entry_no * 600) + 500
Label6.Top = (sub_entry_no * 600) + 500
change_the_old_voucher = 2
Next
End Sub
Private Sub combo_sub_entry_no_Click(index As Integer)
If sub_entry_no > 12 Then
MsgBox "Only 12 Entries allowed Once....!!!!"
Exit Sub
End If
If combo_lgr(sub_entry_no - 1).Text = "" Or LCase(combo_lgr(sub_entry_no - 1).Text) = " " Or LCase(combo_lgr(sub_entry_no - 1).Text) = "select a ledger" _
Or Val(text_amt(sub_entry_no - 1)) < 0 Or Val(text_amt(sub_entry_no - 1)) = Null Then
MsgBox "Please, Enter proper values....!!!! and click again...!!!"
Exit Sub
End If
If index = 2 Then
If LCase(combo_sub_entry_no(index).Text) = LCase("Dr / By") Then
this_entry = "dr"
'write here text to insert row
Call add_account_combo_lgr(sub_entry_no)
Call cr_amt_text_adjust(sub_entry_no)
Call add_text_amt(sub_entry_no)
text_amt(sub_entry_no) = Format(voucher_total_dr_amt - voucher_total_cr_amt, "0.00")
Call refresh_dr_cr_total_amt
this_entry = "cr"
'write here text to insert row
sub_entry_no = sub_entry_no + 1
Load combo_sub_entry_no(sub_entry_no)
Load combo_lgr(sub_entry_no)
Load text_amt(sub_entry_no)
Call add_combo_sub_entry_no(sub_entry_no)
Call add_text_amt(sub_entry_no)
Call add_account_combo_lgr(sub_entry_no)
Call move_all_command_to_bottom
text_amt(sub_entry_no) = Format(voucher_total_cr_amt - voucher_total_dr_amt, "0.00")
End If
ElseIf index > 2 Then
If LCase(combo_sub_entry_no(index).Text) = LCase("Cr / To") Then
this_entry = "cr"
'write here text to insert row
Call add_account_combo_lgr(sub_entry_no)
Call dr_amt_text_adjust(sub_entry_no)
Call add_text_amt(sub_entry_no)
text_amt(sub_entry_no) = Format(voucher_total_cr_amt - voucher_total_dr_amt, "0.00")
ElseIf LCase(combo_sub_entry_no(index).Text) = LCase("Dr / By") Then
this_entry = "dr"
'write here text to insert row
Call add_account_combo_lgr(sub_entry_no)
Call cr_amt_text_adjust(sub_entry_no)
Call add_text_amt(sub_entry_no)
text_amt(sub_entry_no) = Format(voucher_total_dr_amt - voucher_total_cr_amt, "0.00")
End If
End If
End Sub
Public Sub move_all_command_to_bottom()
Label7.Top = (sub_entry_no * 600) + 1000
Text4.Top = (sub_entry_no * 600) + 800
Frame1.Height = (sub_entry_no * 600) + 2300
Frame2.Top = (sub_entry_no * 600) + 2999
Label5.Top = (sub_entry_no * 600) + 500
Label6.Top = (sub_entry_no * 600) + 500
End Sub
Public Sub add_cash_or_bank_ledgers()
combo_lgr(sub_entry_no).Clear
'Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
If rs_lgr_main_dtl!lgr_main_dtl_grup = "Cash-on-hand" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Balances" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Loans" Then
combo_lgr(sub_entry_no).AddItem rs_lgr_main_dtl!lgr_main_dtl_name
If rs_lgr_main_dtl!lgr_main_dtl_alis <> "" Then combo_lgr(sub_entry_no).AddItem rs_lgr_main_dtl!lgr_main_dtl_alis
End If
rs_lgr_main_dtl.MoveNext
Loop
End Sub
Public Sub add_other_then_cash_or_bank_ledgers()
combo_lgr(sub_entry_no).Clear
'Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
If rs_lgr_main_dtl!lgr_main_dtl_grup = "Cash-on-hand" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Balances" Or rs_lgr_main_dtl!lgr_main_dtl_grup = "Bank Loans" Then
Else
combo_lgr(sub_entry_no).AddItem rs_lgr_main_dtl!lgr_main_dtl_name
If rs_lgr_main_dtl!lgr_main_dtl_alis <> "" Then combo_lgr(sub_entry_no).AddItem rs_lgr_main_dtl!lgr_main_dtl_alis
End If
rs_lgr_main_dtl.MoveNext
Loop
End Sub
Private Sub text_amt_LostFocus(index As Integer)
Call voucher_type_1_tab_indexing
End Sub
Private Sub voucher_type_1_tab_indexing()
TabIndex_counter = 1
DTPicker1.TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
combo_sub_entry_no(1).TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
For i_TabIndex_counter = 1 To total_sub_entry_no
If i_TabIndex_counter > 1 Then
combo_sub_entry_no(i_TabIndex_counter).TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
End If
If index_x = i_TabIndex_counter Then
list_lgr.TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
End If
combo_lgr(i_TabIndex_counter).TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
text_amt(i_TabIndex_counter).TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
Next
'Combo0.TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
Text4.TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
cmd_sv_n_new.TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
cmd_edit.TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
cmd_print.TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
cmd_cancel.TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
cmd_save_n_exit.TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
cmd_delete.TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
Command1.TabIndex = TabIndex_counter
TabIndex_counter = TabIndex_counter + 1
End Sub

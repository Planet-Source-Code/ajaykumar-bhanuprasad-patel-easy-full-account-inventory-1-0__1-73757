VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00000000&
   ClientHeight    =   8130
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8760
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":1D2A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   7755
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   15875
            MinWidth        =   15875
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "21:13"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "30/01/2011"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      Height          =   6270
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   6210
      ScaleWidth      =   2445
      TabIndex        =   1
      Top             =   1485
      Visible         =   0   'False
      Width           =   2505
      Begin VB.CommandButton Command1 
         Caption         =   "<"
         Height          =   11655
         Left            =   1800
         TabIndex        =   3
         Top             =   120
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Height          =   11775
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2535
         Begin VB.ListBox List_opened_procedure 
            Height          =   10980
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1485
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   8700
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8760
      Begin VB.CommandButton Command2 
         Caption         =   "^"
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   840
         Width           =   20055
      End
   End
   Begin VB.Menu main 
      Caption         =   "Main Menu"
      Begin VB.Menu co_creat 
         Caption         =   "Creat company"
         Visible         =   0   'False
      End
      Begin VB.Menu co_select 
         Caption         =   "Select company"
         Visible         =   0   'False
      End
      Begin VB.Menu co_change 
         Caption         =   "Change company"
      End
      Begin VB.Menu co_cls 
         Caption         =   "Close company"
      End
      Begin VB.Menu chng_usr 
         Caption         =   "Change User"
      End
      Begin VB.Menu system_info 
         Caption         =   "system information"
         Index           =   100
         Visible         =   0   'False
      End
      Begin VB.Menu co_pref 
         Caption         =   "Preferense"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Ac 
      Caption         =   "Account"
      Begin VB.Menu grp 
         Caption         =   "Group"
         Begin VB.Menu grp_creat 
            Caption         =   "Creat"
         End
         Begin VB.Menu grp_display 
            Caption         =   "Display"
         End
         Begin VB.Menu grp_alter 
            Caption         =   "Alter"
         End
         Begin VB.Menu grp_list 
            Caption         =   "Display List"
         End
      End
      Begin VB.Menu lgr 
         Caption         =   "Ledger"
         Begin VB.Menu lgr_creat 
            Caption         =   "Creat"
            Shortcut        =   ^L
         End
         Begin VB.Menu lgr_display 
            Caption         =   "Display"
         End
         Begin VB.Menu lgr_list 
            Caption         =   "Display List"
         End
         Begin VB.Menu lgr_alter 
            Caption         =   "Alter"
         End
      End
   End
   Begin VB.Menu inventory 
      Caption         =   "Inventory"
      Begin VB.Menu igrp 
         Caption         =   "Group"
         Begin VB.Menu igrp_creat 
            Caption         =   "Creat"
         End
         Begin VB.Menu igrp_display 
            Caption         =   "Display"
         End
         Begin VB.Menu igrp_list 
            Caption         =   "Display list"
         End
         Begin VB.Menu igrp_alter 
            Caption         =   "Alter"
         End
      End
      Begin VB.Menu item 
         Caption         =   "Item"
         Begin VB.Menu itm_creat 
            Caption         =   "Creat"
            Shortcut        =   ^I
         End
         Begin VB.Menu itm_display 
            Caption         =   "Dispaly"
         End
         Begin VB.Menu itm_list 
            Caption         =   "Display list"
         End
         Begin VB.Menu itm_alter 
            Caption         =   "Alter"
         End
      End
      Begin VB.Menu unit 
         Caption         =   "Unit"
         Begin VB.Menu unit_creat 
            Caption         =   "Creat"
         End
         Begin VB.Menu unit_display 
            Caption         =   "Display"
         End
         Begin VB.Menu unit_display_list 
            Caption         =   "Display list"
         End
         Begin VB.Menu unit_alter 
            Caption         =   "Alter"
         End
      End
   End
   Begin VB.Menu trn 
      Caption         =   "Transaction"
      Begin VB.Menu sales_return 
         Caption         =   "Sales Return"
         Shortcut        =   {F12}
      End
      Begin VB.Menu purchase_return 
         Caption         =   "Purchaser Return"
         Shortcut        =   {F11}
      End
      Begin VB.Menu payment 
         Caption         =   "Payment"
         Shortcut        =   {F5}
      End
      Begin VB.Menu receipt 
         Caption         =   "Receipt"
         Shortcut        =   {F6}
      End
      Begin VB.Menu sales 
         Caption         =   "Sales"
         Shortcut        =   {F8}
      End
      Begin VB.Menu purchase 
         Caption         =   "Purchase"
         Shortcut        =   {F9}
      End
      Begin VB.Menu banking 
         Caption         =   "Banking"
         Shortcut        =   {F4}
      End
      Begin VB.Menu adjustment 
         Caption         =   "Adjustment"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu rpt 
      Caption         =   "Reports"
      Begin VB.Menu Acc_report 
         Caption         =   "Accouting Reports"
         Begin VB.Menu Trial_balance_sheet 
            Caption         =   "Trial Balance Sheet"
         End
         Begin VB.Menu smry_group 
            Caption         =   "Group summary"
         End
         Begin VB.Menu specific_ledger_detail 
            Caption         =   "Ledger Account"
            Shortcut        =   ^A
         End
         Begin VB.Menu sm_ledger 
            Caption         =   "Ladger"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Stock_reports 
         Caption         =   "Stock Reports"
         Begin VB.Menu stock_summary 
            Caption         =   "stock summary"
         End
         Begin VB.Menu stock_item_account 
            Caption         =   "stock item account"
         End
         Begin VB.Menu closing_stock 
            Caption         =   "Item closing stock"
         End
      End
      Begin VB.Menu smry 
         Caption         =   "Voucher_Summaries"
         Begin VB.Menu smry_payment 
            Caption         =   "payment summary"
         End
         Begin VB.Menu smry_receipt 
            Caption         =   "Receipt summary"
         End
         Begin VB.Menu smry_sales 
            Caption         =   "Sales summary"
            Shortcut        =   ^S
         End
         Begin VB.Menu smry_purchase 
            Caption         =   "Purchase summary"
         End
         Begin VB.Menu smry_banking 
            Caption         =   "Banking summary"
         End
         Begin VB.Menu summary_purchase_return 
            Caption         =   "Purchase Return Summary"
         End
         Begin VB.Menu summary_sale_return 
            Caption         =   "Sale Return summary"
         End
         Begin VB.Menu smry_adjustment 
            Caption         =   "Adjustment summary"
         End
      End
      Begin VB.Menu other_reports 
         Caption         =   "Other Reports"
         Begin VB.Menu sales_man_report 
            Caption         =   "sales man report"
         End
         Begin VB.Menu c_rate 
            Caption         =   "Customer Rates"
         End
      End
      Begin VB.Menu p_l_ac 
         Caption         =   "Profit & Loss account"
         Visible         =   0   'False
      End
      Begin VB.Menu b_s 
         Caption         =   "Balance sheet"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu emp 
      Caption         =   "employee"
      Begin VB.Menu emp_add 
         Caption         =   "Add employee "
      End
      Begin VB.Menu emp_in_entry 
         Caption         =   "In Entry"
      End
      Begin VB.Menu emp_out_entry 
         Caption         =   "Out Entry"
      End
      Begin VB.Menu emp_detail_repo 
         Caption         =   "Employee Detail Report"
      End
      Begin VB.Menu emp_salary_rep 
         Caption         =   "Employee Salary Report"
      End
   End
   Begin VB.Menu da_repo 
      Caption         =   "De Activation"
      Begin VB.Menu da_find_card 
         Caption         =   "DA Find a card"
      End
      Begin VB.Menu show_da_report 
         Caption         =   "Show DA Report"
      End
      Begin VB.Menu da_gene_repo 
         Caption         =   "Generate Report"
      End
      Begin VB.Menu resp_report 
         Caption         =   "Response Report"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu sft_help 
         Caption         =   "Software Help"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()

End Sub

Private Sub about_Click()
frm_about.Show
End Sub
Private Sub resp_report_Click()
selected_procedure = "Response Entry"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
da_resp_entry.Show

End Sub
Private Sub sales_Click()
show_ledger_detail = 0
selected_procedure = "Sales voucher"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
vchr_sales.Show
End Sub
Private Sub co_change_Click()
'selected_path = ""
Me.Caption = "Ajay patel's Accounting software....., You have to select a company to work"
B_co_menu.Show
End Sub
Private Sub emp_add_Click()
selected_procedure = "Employee creat"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
empl_creat.Show
End Sub
Private Sub emp_detail_repo_Click()
selected_procedure = "Employee In Entry"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
emp_detail_repo2.Show

End Sub

Private Sub closing_stock_Click()
show_stock_item_by_click = 0
selected_procedure = "serial wise closing stock"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
shw_item_wise_clg_stk.Show
End Sub
Private Sub co_pref_Click()
selected_procedure = "Co. prefrence"
path_sel.Show
End Sub
Private Sub da_find_card_Click()
selected_procedure = "Find card Deactivation Report"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
da_find_card_detail.Show
End Sub
Private Sub da_gene_repo_Click()
selected_procedure = "Generate Deactivation Report"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
da_Gene_cust_repo_main.Show
End Sub

Private Sub payment_Click()
show_ledger_detail = 0
selected_procedure = "payment voucher"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
vchr_payment.Show
End Sub
Private Sub purchase_Click()
show_ledger_detail = 0
selected_procedure = "purchase voucher"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
vchr_purchase.Show
End Sub
Private Sub purchase_return_Click()
show_ledger_detail = 0
selected_procedure = "purchase return voucher"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
vchr_purchase_return.Show
End Sub
Private Sub receipt_Click()
show_ledger_detail = 0
selected_procedure = "Receipt voucher"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
vchr_receipt.Show
End Sub
Private Sub adjustment_Click()
show_ledger_detail = 0
selected_procedure = "Adjustment/Journal voucher"
vchr_Journal.Show
End Sub
Private Sub banking_Click()
show_ledger_detail = 0
selected_procedure = "Banking voucher"
vchr_contra.Show
End Sub
Private Sub chng_usr_Click()
selected_procedure = "Select User"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
frm_usr.Show
End Sub
Private Sub List_opened_procedure_Clickx()
Dim selected_list_item
selected_list_item = MDIForm1.List_opened_procedure.Text

If LCase(selected_list_item) = LCase("group_Display") Or LCase(selected_list_item) = LCase("group_edit") Or LCase(selected_list_item) = LCase("group_creat") Then
    SetWindowPos Creat_ac_grp.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("Select User") Then
    SetWindowPos frm_usr.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("payment voucher") Then
    SetWindowPos vchr_payment.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("receipt voucher") Then
    SetWindowPos vchr_receipt.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("sales voucher") Then
    SetWindowPos vchr_sales.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("purchase voucher") Then
    SetWindowPos vchr_purchase.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("banking voucher") Then
    SetWindowPos vchr_contra.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("Adjustment/Journal voucher") Then
    SetWindowPos vchr_Journal.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("purchase return voucher") Then
    SetWindowPos vchr_purchase_return.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE

ElseIf LCase(selected_list_item) = LCase("serial wise closing stock") Then
    SetWindowPos shw_item_wise_clg_stk.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("Co. prefrence") Then
    SetWindowPos path_sel.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("Find card Deactivation Report") Then
    SetWindowPos da_find_card_detail.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    
ElseIf LCase(selected_list_item) = LCase("Response Entry") Then
    SetWindowPos da_resp_entry.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("Sales voucher") Then
    SetWindowPos vchr_sales.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("Employee creat") Then
    SetWindowPos empl_creat.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("Employee In Entry") Then
    SetWindowPos emp_detail_repo2.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    
ElseIf LCase(selected_list_item) = LCase("payment voucher") Then
    SetWindowPos vchr_payment.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("payment voucher") Then
    SetWindowPos vchr_payment.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("payment voucher") Then
    SetWindowPos vchr_payment.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("payment voucher") Then
    SetWindowPos vchr_payment.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("payment voucher") Then
    SetWindowPos vchr_payment.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("payment voucher") Then
    SetWindowPos vchr_payment.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("payment voucher") Then
    SetWindowPos vchr_payment.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
ElseIf LCase(selected_list_item) = LCase("payment voucher") Then
    SetWindowPos vchr_payment.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End If
End Sub
Private Sub emp_in_entry_Click()
selected_procedure = "Employee In Entry"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
emp_tran.Show
End Sub
Private Sub emp_out_entry_Click()
selected_procedure = "Employee Out Entry"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
emp_tran.Show
End Sub
Private Sub emp_salary_rep_Click()
selected_procedure = "Employee Salary Report"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
emp_report_1.Show
End Sub
Private Sub grp_alter_Click()
selected_procedure = "group_edit"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
Creat_ac_grp.Show
End Sub
Private Sub grp_creat_Click()
selected_procedure = "group_creat"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
'ListView1.ListItems.Add(1) = selected_procedure
Creat_ac_grp.Show
End Sub
Private Sub grp_display_Click()
selected_procedure = "group_Display"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
Creat_ac_grp.Show
End Sub
Private Sub grp_list_Click()
selected_procedure = "group_display_list"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
disp_created_grup.Show
End Sub
Private Sub igrp_alter_Click()
'selected_procedure = "stock_group_creat"
selected_procedure = "stock_group_edit"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
Creat_st_grp.Show
End Sub
Private Sub igrp_creat_Click()
selected_procedure = "stock_group_creat"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
'selected_procedure = "stock_group_edit"
Creat_st_grp.Show
End Sub
Private Sub igrp_display_Click()
selected_procedure = "Stock_Group_Display"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
Creat_st_grp.Show
End Sub
Private Sub igrp_list_Click()
selected_procedure = "Stock_Group_Display_List"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
disp_created_stock_group.Show
End Sub
Private Sub itm_display_Click()
'selected_procedure = "Stock_item_ledger_edit"
selected_procedure = "Stock_item_ledger_display"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
creat_stock_lgr.Show
End Sub
Private Sub itm_list_Click()
selected_procedure = "Stock_Item_Display_List"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
'selected_procedure = "stock_unit_creat"
disp_created_stock_item.Show
End Sub
Private Sub sales_man_report_Click()
selected_procedure = "sales man report"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
disp_sales_man_repo.Show
End Sub
Private Sub sales_return_Click()
show_ledger_detail = 0
selected_procedure = "sale return voucher"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
vchr_sale_return.Show
End Sub

Private Sub sft_help_Click()
frm_about_software.Show
End Sub

Private Sub show_da_report_Click()
selected_procedure = "Show Report"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
da_show_repo.Show
End Sub
Private Sub sm_ledger_Click()
selected_procedure = "Display Ledger Account"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
shw_lgr_dtl.Show
End Sub
Private Sub smry_adjustment_Click()
ledger_clicked_from_other = 0
selected_voucher_name = "journal"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
selected_procedure = "show Journal voucher summary"
show_sel_vchr.Show
End Sub
Private Sub smry_banking_Click()
ledger_clicked_from_other = 0
selected_voucher_name = "contra"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
ledger_clicked_from_other = 0
selected_procedure = "show Banking voucher summary"
show_sel_vchr.Show
End Sub
Private Sub smry_group_Click()
ledger_clicked_from_other = 0
selected_procedure = "show group summary"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
show_sel_grp_smry.Show
End Sub
Private Sub smry_payment_Click()
ledger_clicked_from_other = 0
selected_voucher_name = "payment"
selected_procedure = "show payment voucher summary"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
show_sel_vchr.Show
End Sub
Private Sub smry_purchase_Click()
ledger_clicked_from_other = 0
selected_voucher_name = "purchase"
selected_procedure = "show purchase voucher summary"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
show_sel_vchr.Show
End Sub
Private Sub smry_receipt_Click()
ledger_clicked_from_other = 0
selected_voucher_name = "receipt"
selected_procedure = "show Receipt voucher summary"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
show_sel_vchr.Show
End Sub
Private Sub smry_sales_Click()
ledger_clicked_from_other = 0
selected_voucher_name = "sale"
selected_procedure = "show Sale voucher summary"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
show_sel_vchr.Show
End Sub
Private Sub specific_ledger_detail_Click()
ledger_clicked_from_other = 0
selected_procedure = "show ledger account"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
shw_sel_lgr_dtl.Show
End Sub
Private Sub stock_item_account_Click()
show_ledger_detail = 0
selected_procedure = "stock item account"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
shw_item_acnt.Show
End Sub
Private Sub stock_summary_Click()
selected_procedure = "closing_stock_display"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
shw_clg_stk_summary.Show
End Sub
Private Sub summary_purchase_return_Click()
ledger_clicked_from_other = 0
selected_voucher_name = "purchase return"
selected_procedure = "show purchase Return voucher summary"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
show_sel_vchr.Show
End Sub
Private Sub summary_sale_return_Click()
ledger_clicked_from_other = 0
selected_voucher_name = "sale return"
selected_procedure = "show Sale Return voucher summary"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
show_sel_vchr.Show
End Sub

Private Sub Trial_balance_sheet_Click()
selected_procedure = "Trial Balance Sheet"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
show_trial_balance.Show
End Sub
Private Sub unit_display_Click()
selected_procedure = "stock_unit_display"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
'selected_procedure = "stock_unit_creat"
Creat_st_unt.Show
End Sub
Private Sub unit_alter_Click()
selected_procedure = "stock_unit_edit"
'selected_procedure = "stock_unit_creat"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
Creat_st_unt.Show
End Sub
Private Sub unit_creat_Click()
'selected_procedure = "stock_unit_edit"
'selected_procedure = "stock_unit_creat"
selected_procedure = "stock_unit_creat"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
Creat_st_unt.Show
End Sub
Private Sub itm_alter_Click()
selected_procedure = "Stock_item_ledger_edit"
'selected_procedure = "Stock_item_ledger_creat"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
creat_stock_lgr.Show
End Sub
Private Sub itm_creat_Click()
'selected_procedure = "Stock_item_ledger_edit"
selected_procedure = "Stock_item_ledger_creat"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
creat_stock_lgr.Show
End Sub
Private Sub lgr_alter_Click()
selected_procedure = "ledger_edit"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
creat_ac_lgr.Show
End Sub
Private Sub lgr_creat_Click()
selected_procedure = "ledger_creat"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
creat_ac_lgr.Show
End Sub
Private Sub lgr_display_Click()
selected_procedure = "ledger_display"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
creat_ac_lgr.Show
End Sub
Private Sub lgr_list_Click()
selected_procedure = "ledger_display_list"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
disp_created_ledger.Show
End Sub
Private Sub MDIForm_Activate()
Me.Caption = "Ajay patel's Telecome Accounting software.." & selected_company & "..." & selected_user
StatusBar1.Panels(1) = UCase(selected_company)
StatusBar1.Panels(6) = UCase(selected_user)
End Sub
Private Sub c_rate_Click()
selected_procedure = "Latest Sale Prise of Customer"
shw_sel_lgr_sale_prise.Show '    Shell App.Path & "\other exes\Customer rate.exe", vbNormalFocus
End Sub
Private Sub co_cls_Click()
B_co_menu.Show
selected_path = ""
Me.Caption = "Ajay patel's Telecome Accounting software....., You have to select a company to work"
Unload Me
End Sub

Private Sub co_creat_Click()
MDIForm1.Enabled = False
BA_co_creat_frm.Show
'ListView1.ListItems.Add = "Creat Company"
End Sub

Private Sub co_select_Click()
B_co_menu.Show
End Sub

Private Sub Command1_Click()
If Command1.Caption = "<" Then
    Command1.Caption = ">"
    Picture2.Width = 200
    Command1.Left = 0
'Frame2.Width = 600
'Frame2.Height = Me.Height - 600
'Frame2.Top = 600
'Frame2.Left = 0
    ElseIf Command1.Caption = ">" Then
    Command1.Caption = "<"
    Picture2.Width = 3000
    Command1.Left = 3000 - 255
'Frame2.Width = 3000 - 600
'Frame2.Height = 600
'Frame2.Top = 0
'Frame2.Left = 0
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "^" Then
    Command2.Caption = "V"
    Picture1.Height = 300
    Command2.Top = 0
ElseIf Command2.Caption = "V" Then
Command2.Caption = "^"
Picture1.Height = 1000
Command2.Top = 1000 - 300
End If
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub MDIForm_Load()
'================================================
If selected_path = "" Or selected_path = Null Then
    'selected_path = App.Path & "\data\1000\co.mdb;"
    selected_path = App.Path & "\co.mdb;"
    'selected_user = "Sukdev_AP"
    selected_user = "admin"
End If
'Call open_database
'Call make_trail_balance_summary

'================================================
Call set_form_data
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
        
    Command1.Caption = ">"
    Picture2.Width = 200
    Command1.Left = 0
    Command2.Caption = "V"
    Picture1.Height = 300
    Command2.Top = 0

End Sub

Public Sub set_form_data()
Command2.Left = 50
Command2.Height = 300
Command2.Width = Screen.Width
Frame1.Top = 0
End Sub
'Public Sub get_system_data()
'sysinfo_c_name = machinename()                          '"Computer Name"
'sysinfo_c_user = Environ$("username")                   '"User name"
'sysinfo_c_windir = Environ$("windir")                   '"Windows Dir"
'sysinfo_c_wintempdir = Environ$("temp")                 '"Windows temp dir"
'sysinfo_c_winsysdir = SystemDir                         '"win system dir"
'sysinfo_c_sysdrive = Environ$("systemdrive")            '"System Drive"
'sysinfo_c_osname = Environ$("os")                       'os version
'    processorsinfo = ReadKey("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0\Processornamestring")
'sysinfo_c_processor = processorsinfo                    'processor information
'    Call GlobalMemoryStatus(memInfo)
'sysinfo_c_totalram = memInfo.dwTotalPhys / 1024 & " KB" 'total ram
'sysinfo_c_noprocessor = Environ$("Number_Of_Processors")  'number of processor
'Call computerbiosinfo
'sysinfo_c_biosver = bios_name                           'bios version
'sysinfo_c_biosman = bios_manufacturer
'        sysdrv = Environ$("systemdrive")
'sysinfo_c_sysdrv_serial_no = GetSerialNumber(Environ$("systemdrive"))   'system drive serial no
'End Sub

Private Sub unit_display_list_Click()
'selected_procedure = "stock_unit_edit"
'selected_procedure = "stock_unit_creat"
selected_procedure = "stock_unit_display_list"
'MDIForm1.List_opened_procedure.AddItem selected_procedure
disp_created_unit.Show
End Sub

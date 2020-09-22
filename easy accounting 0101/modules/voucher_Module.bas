Attribute VB_Name = "voucher_module"
Public TabIndex_counter
Public i_TabIndex_counter

Public sub_entry_no As Integer
Public xsub_entry_no As Integer
Public dr_sub_entry_no As Integer
Public cr_sub_entry_no As Integer

Public selected_voucher_type As Integer
Public selected_voucher_no As Integer
Public selected_voucher_date As Date
Public selected_voucher_ledger As String
Public selected_voucher_amount As Double
Public selected_voucher_side As String
Public selected_voucher_narration As String
Public selected_voucher_user As String
Public selected_voucher_name As String
Public selected_date As Date
Public show_ledger_detail As Integer
Public selected_ledger As String

Public total_sub_entry_no As Integer
Public this_entry As String

Public voucher_total_dr_amt As Double
Public voucher_total_cr_amt As Double
Public transaction_type As String
Public current_sub_entry_no As Integer
Public change_the_old_voucher As Integer

Public selected_voucher_card As String


Attribute VB_Name = "ledger_merge_all_to_table"
Public Sub make_trail_balance_summary()

Call open_rs_lgr_clsg_smr

Do Until rs_lgr_clsg_smr.EOF
    rs_lgr_clsg_smr.Delete
    rs_lgr_clsg_smr.Update
    rs_lgr_clsg_smr.MoveNext
Loop

Call collect_all_transaction_in_one_file
'Call open_database
Call open_rs_lgr_main_dtl
Call open_rs_lgr_clsg_smr
'MsgBox rs_lgr_main_dtl.RecordCount


rs_lgr_main_dtl.MoveFirst
Do Until rs_lgr_main_dtl.EOF
selected_ledger = rs_lgr_main_dtl!lgr_main_dtl_name
'MsgBox rs_lgr_main_dtl!lgr_main_dtl_name
ledger_dr_total = 0
ledger_cr_total = 0
Call copy_specific_ledger_transaction_to_acn_tran_spc_lgr
Call open_rs_acn_tran_spc_lgr
rs_acn_tran_spc_lgr.Sort = "fin_acnt_trn_date,fin_acnt_trn_vcno" 'time"
Dim temp_starting_dt As Date
Dim temp_ending_dt As Date
b = 1
Do Until rs_acn_tran_spc_lgr.EOF
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("dr") And rs_acn_tran_spc_lgr!fin_acnt_trn_date <= selected_date Then
    ledger_dr_total = ledger_dr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    If LCase(rs_acn_tran_spc_lgr!fin_acnt_trn_side) = LCase("cr") And rs_acn_tran_spc_lgr!fin_acnt_trn_date <= selected_date Then
    ledger_cr_total = ledger_cr_total + Val(rs_acn_tran_spc_lgr!fin_acnt_trn_amnt)
    End If
    b = b + 1
        'If LCase(selected_ledger) = "cash" Then MsgBox ledger_dr_total - ledger_cr_total
rs_acn_tran_spc_lgr.MoveNext
Loop
        
        
        If ledger_dr_total < ledger_cr_total Then
                rs_lgr_clsg_smr.AddNew
                rs_lgr_clsg_smr!lgr_clsg_dtl_id = rs_lgr_main_dtl!lgr_main_dtl_id
                rs_lgr_clsg_smr!lgr_clsg_dtl_name = rs_lgr_main_dtl!lgr_main_dtl_name
                rs_lgr_clsg_smr!lgr_clsg_dtl_grup = rs_lgr_main_dtl!lgr_main_dtl_grup
                                selected_group = rs_lgr_main_dtl!lgr_main_dtl_grup
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
                rs_lgr_clsg_smr!lgr_clsg_dtl_pgrp = selected_primary_group
                rs_lgr_clsg_smr!lgr_clsg_dtl_crpd = rs_lgr_main_dtl!lgr_main_dtl_crpd
                rs_lgr_clsg_smr!lgr_clsg_dtl_cram = rs_lgr_main_dtl!lgr_main_dtl_cram
                'rs_lgr_clsg_smr!lgr_clsg_dtl_bal1 = rs_lgr_main_dtl!lgr_main_dtl_obl1
                'rs_lgr_clsg_smr!lgr_clsg_dtl_bal2 = rs_lgr_main_dtl!lgr_main_dtl_obl2
                'rs_lgr_clsg_smr!lgr_clsg_dtl_sid1 = rs_lgr_main_dtl!lgr_main_dtl_osd1
                'rs_lgr_clsg_smr!lgr_clsg_dtl_sid2 = rs_lgr_main_dtl!lgr_main_dtl_osd2
                rs_lgr_clsg_smr!lgr_clsg_dtl_slun = rs_lgr_main_dtl!lgr_main_dtl_slun
                rs_lgr_clsg_smr!lgr_clsg_dtl_tbal = ledger_cr_total - ledger_dr_total
                rs_lgr_clsg_smr!lgr_clsg_dtl_tsid = "cr"
                rs_lgr_clsg_smr.Update
        ElseIf ledger_cr_total < ledger_dr_total Then
        'MsgBox rs_lgr_main_dtl!lgr_main_dtl_name & "..." & ledger_dr_total - ledger_cr_total
                rs_lgr_clsg_smr.AddNew
                rs_lgr_clsg_smr!lgr_clsg_dtl_id = rs_lgr_main_dtl!lgr_main_dtl_id
                rs_lgr_clsg_smr!lgr_clsg_dtl_name = rs_lgr_main_dtl!lgr_main_dtl_name
                rs_lgr_clsg_smr!lgr_clsg_dtl_grup = rs_lgr_main_dtl!lgr_main_dtl_grup
                
                                selected_group = rs_lgr_main_dtl!lgr_main_dtl_grup
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
                rs_lgr_clsg_smr!lgr_clsg_dtl_pgrp = selected_primary_group
                rs_lgr_clsg_smr!lgr_clsg_dtl_crpd = rs_lgr_main_dtl!lgr_main_dtl_crpd
                rs_lgr_clsg_smr!lgr_clsg_dtl_cram = rs_lgr_main_dtl!lgr_main_dtl_cram
                'rs_lgr_clsg_smr!lgr_clsg_dtl_bal1 = rs_lgr_main_dtl!lgr_main_dtl_obl1
                'rs_lgr_clsg_smr!lgr_clsg_dtl_bal2 = rs_lgr_main_dtl!lgr_main_dtl_obl2
                'rs_lgr_clsg_smr!lgr_clsg_dtl_sid1 = rs_lgr_main_dtl!lgr_main_dtl_osd1
                'rs_lgr_clsg_smr!lgr_clsg_dtl_sid2 = rs_lgr_main_dtl!lgr_main_dtl_osd2
                rs_lgr_clsg_smr!lgr_clsg_dtl_slun = rs_lgr_main_dtl!lgr_main_dtl_slun
                rs_lgr_clsg_smr!lgr_clsg_dtl_tbal = ledger_dr_total - ledger_cr_total
                rs_lgr_clsg_smr!lgr_clsg_dtl_tsid = "dr"
                rs_lgr_clsg_smr.Update
      'grid_report.TextMatrix(b, 6) = Format(ledger_dr_total - ledger_cr_total, "0.00")
        ElseIf ledger_cr_total = ledger_dr_total And ledger_dr_total <> 0 And ledger_cr_total <> 0 Then
                rs_lgr_clsg_smr.AddNew
                rs_lgr_clsg_smr!lgr_clsg_dtl_id = rs_lgr_main_dtl!lgr_main_dtl_id
                rs_lgr_clsg_smr!lgr_clsg_dtl_name = rs_lgr_main_dtl!lgr_main_dtl_name
                rs_lgr_clsg_smr!lgr_clsg_dtl_grup = rs_lgr_main_dtl!lgr_main_dtl_grup
                
                                selected_group = rs_lgr_main_dtl!lgr_main_dtl_grup
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
    rs_lgr_clsg_smr!lgr_clsg_dtl_pgrp = selected_primary_group
    rs_lgr_clsg_smr!lgr_clsg_dtl_crpd = rs_lgr_main_dtl!lgr_main_dtl_crpd
    rs_lgr_clsg_smr!lgr_clsg_dtl_cram = rs_lgr_main_dtl!lgr_main_dtl_cram
    'rs_lgr_clsg_smr!lgr_clsg_dtl_bal1 = rs_lgr_main_dtl!lgr_main_dtl_obl1
    'rs_lgr_clsg_smr!lgr_clsg_dtl_bal2 = rs_lgr_main_dtl!lgr_main_dtl_obl2
    'rs_lgr_clsg_smr!lgr_clsg_dtl_sid1 = rs_lgr_main_dtl!lgr_main_dtl_osd1
    'rs_lgr_clsg_smr!lgr_clsg_dtl_sid2 = rs_lgr_main_dtl!lgr_main_dtl_osd2
    rs_lgr_clsg_smr!lgr_clsg_dtl_slun = rs_lgr_main_dtl!lgr_main_dtl_slun
    rs_lgr_clsg_smr!lgr_clsg_dtl_tbal = ledger_dr_total - ledger_cr_total
    rs_lgr_clsg_smr!lgr_clsg_dtl_tsid = "dr"
    rs_lgr_clsg_smr.Update
    End If
    rs_lgr_main_dtl.MoveNext
Loop
End Sub
Public Sub collect_all_transaction_in_one_file()
Dim i_trans
''Call open_database
Call open_rs_acn_tran_all
For i_trans = 1 To rs_acn_tran_all.RecordCount
rs_acn_tran_all.Delete
rs_acn_tran_all.MoveNext
Next
temp_entry_no = 1
Call copy_all_data_from_opening_to_all_temp
Call copy_all_data_from_payment_to_all
Call copy_all_data_from_receipt_to_all
Call copy_all_data_from_sales_to_all
Call copy_all_data_from_purchase_to_all
Call copy_all_data_from_salesreturn_to_all
Call copy_all_data_from_purchasereturn_to_all
Call copy_all_data_from_contra_to_all
Call copy_all_data_from_journal_to_all

End Sub
Public Sub copy_all_data_from_opening_to_all_temp()
Dim i_trans
'Call open_database
Call open_rs_lgr_main_dtl
Call open_rs_acn_tran_all
For i_trans = 1 To rs_lgr_main_dtl.RecordCount
    rs_acn_tran_all.AddNew
    rs_acn_tran_all!fin_acnt_trn_seno = 1
    rs_acn_tran_all!fin_acnt_trn_vtyp = 1
    rs_acn_tran_all!fin_acnt_trn_vchr = "opening balance 1"
    rs_acn_tran_all!fin_acnt_trn_date = selected_companies_starting_f_date
    rs_acn_tran_all!fin_acnt_trn_ldgr = rs_lgr_main_dtl!lgr_main_dtl_name
    rs_acn_tran_all!fin_acnt_trn_amnt = rs_lgr_main_dtl!lgr_main_dtl_obl1
    rs_acn_tran_all!fin_acnt_trn_side = rs_lgr_main_dtl!lgr_main_dtl_osd1
    rs_acn_tran_all!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    
    rs_acn_tran_all.UpdateBatch
    rs_acn_tran_all.AddNew
    rs_acn_tran_all!fin_acnt_trn_seno = 2
    rs_acn_tran_all!fin_acnt_trn_vtyp = 2
    rs_acn_tran_all!fin_acnt_trn_date = selected_companies_starting_f_date
    rs_acn_tran_all!fin_acnt_trn_vchr = "opening balance 2"
    rs_acn_tran_all!fin_acnt_trn_ldgr = rs_lgr_main_dtl!lgr_main_dtl_name
    'rs_acn_tran_all!fin_acnt_trn_amnt = rs_lgr_main_dtl!lgr_main_dtl_obl2
    'rs_acn_tran_all!fin_acnt_trn_side = rs_lgr_main_dtl!lgr_main_dtl_osd2
    rs_acn_tran_all!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    rs_acn_tran_all.UpdateBatch
rs_lgr_main_dtl.MoveNext
Next
End Sub
Public Sub copy_all_data_from_to_all_temp()
Dim i_trans
''Call open_database
Call open_rs_acn_tran_all
Call open_rs_acn_tran_all_temp
For i_trans = 1 To rs_acn_tran_all.RecordCount
    rs_acn_tran_all_temp.AddNew
    rs_acn_tran_all_temp!fin_acnt_trn_vcno = rs_acn_tran_all!fin_acnt_trn_vcno
    rs_acn_tran_all_temp!fin_acnt_trn_seno = rs_acn_tran_all!fin_acnt_trn_seno
    rs_acn_tran_all_temp!fin_acnt_trn_vtyp = rs_acn_tran_all!fin_acnt_trn_vtyp
    rs_acn_tran_all_temp!fin_acnt_trn_date = rs_acn_tran_all!fin_acnt_trn_date
    rs_acn_tran_all_temp!fin_acnt_trn_time = rs_acn_tran_all!fin_acnt_trn_time
    rs_acn_tran_all_temp!fin_acnt_trn_wday = rs_acn_tran_all!fin_acnt_trn_wday
    rs_acn_tran_all_temp!fin_acnt_trn_ldgr = rs_acn_tran_all!fin_acnt_trn_ldgr
    rs_acn_tran_all_temp!fin_acnt_trn_amnt = rs_acn_tran_all!fin_acnt_trn_amnt
    rs_acn_tran_all_temp!fin_acnt_trn_side = rs_acn_tran_all!fin_acnt_trn_side
    rs_acn_tran_all_temp!fin_acnt_trn_nrtn = rs_acn_tran_all!fin_acnt_trn_nrtn
    rs_acn_tran_all_temp!fin_acnt_trn_user = rs_acn_tran_all!fin_acnt_trn_user
    rs_acn_tran_all_temp!fin_acnt_trn_vchr = rs_acn_tran_all!fin_acnt_trn_vchr
    rs_acn_tran_all_temp!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    rs_acn_tran_all_temp.UpdateBatch
rs_acn_tran_all.MoveNext
Next
End Sub
Public Sub copy_all_data_from_payment_to_all()
Dim i_trans
''Call open_database
Call open_rs_acn_tran_pmt
Call open_rs_acn_tran_all
For i_trans = 1 To rs_acn_tran_pmt.RecordCount
    rs_acn_tran_all.AddNew
    rs_acn_tran_all!fin_acnt_trn_vcno = rs_acn_tran_pmt!fin_acnt_trn_vcno
    rs_acn_tran_all!fin_acnt_trn_seno = rs_acn_tran_pmt!fin_acnt_trn_seno
    rs_acn_tran_all!fin_acnt_trn_vtyp = rs_acn_tran_pmt!fin_acnt_trn_vtyp
    rs_acn_tran_all!fin_acnt_trn_date = rs_acn_tran_pmt!fin_acnt_trn_date
    rs_acn_tran_all!fin_acnt_trn_time = rs_acn_tran_pmt!fin_acnt_trn_time
    rs_acn_tran_all!fin_acnt_trn_wday = rs_acn_tran_pmt!fin_acnt_trn_wday
    rs_acn_tran_all!fin_acnt_trn_ldgr = rs_acn_tran_pmt!fin_acnt_trn_ldgr
    rs_acn_tran_all!fin_acnt_trn_amnt = rs_acn_tran_pmt!fin_acnt_trn_amnt
    rs_acn_tran_all!fin_acnt_trn_side = rs_acn_tran_pmt!fin_acnt_trn_side
    rs_acn_tran_all!fin_acnt_trn_nrtn = rs_acn_tran_pmt!fin_acnt_trn_nrtn
    rs_acn_tran_all!fin_acnt_trn_user = rs_acn_tran_pmt!fin_acnt_trn_user
    rs_acn_tran_all!fin_acnt_trn_vchr = rs_acn_tran_pmt!fin_acnt_trn_vchr
    rs_acn_tran_all!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    rs_acn_tran_all.UpdateBatch
rs_acn_tran_pmt.MoveNext
Next
End Sub
Public Sub copy_all_data_from_receipt_to_all()
Dim i_trans
''Call open_database
Call open_rs_acn_tran_rct
Call open_rs_acn_tran_all
For i_trans = 1 To rs_acn_tran_rct.RecordCount
    rs_acn_tran_all.AddNew
    rs_acn_tran_all!fin_acnt_trn_vcno = rs_acn_tran_rct!fin_acnt_trn_vcno
    rs_acn_tran_all!fin_acnt_trn_seno = rs_acn_tran_rct!fin_acnt_trn_seno
    rs_acn_tran_all!fin_acnt_trn_vtyp = rs_acn_tran_rct!fin_acnt_trn_vtyp
    rs_acn_tran_all!fin_acnt_trn_date = rs_acn_tran_rct!fin_acnt_trn_date
    rs_acn_tran_all!fin_acnt_trn_time = rs_acn_tran_rct!fin_acnt_trn_time
    rs_acn_tran_all!fin_acnt_trn_wday = rs_acn_tran_rct!fin_acnt_trn_wday
    rs_acn_tran_all!fin_acnt_trn_ldgr = rs_acn_tran_rct!fin_acnt_trn_ldgr
    rs_acn_tran_all!fin_acnt_trn_amnt = rs_acn_tran_rct!fin_acnt_trn_amnt
    rs_acn_tran_all!fin_acnt_trn_side = rs_acn_tran_rct!fin_acnt_trn_side
    rs_acn_tran_all!fin_acnt_trn_nrtn = rs_acn_tran_rct!fin_acnt_trn_nrtn
    rs_acn_tran_all!fin_acnt_trn_user = rs_acn_tran_rct!fin_acnt_trn_user
    rs_acn_tran_all!fin_acnt_trn_vchr = rs_acn_tran_rct!fin_acnt_trn_vchr
    rs_acn_tran_all!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    rs_acn_tran_all.UpdateBatch
rs_acn_tran_rct.MoveNext
Next
End Sub
Public Sub copy_all_data_from_journal_to_all()
Dim i_trans
'Call open_database
Call open_rs_acn_tran_jrn
Call open_rs_acn_tran_all
For i_trans = 1 To rs_acn_tran_jrn.RecordCount
    rs_acn_tran_all.AddNew
    rs_acn_tran_all!fin_acnt_trn_vcno = rs_acn_tran_jrn!fin_acnt_trn_vcno
    rs_acn_tran_all!fin_acnt_trn_seno = rs_acn_tran_jrn!fin_acnt_trn_seno
    rs_acn_tran_all!fin_acnt_trn_vtyp = rs_acn_tran_jrn!fin_acnt_trn_vtyp
    rs_acn_tran_all!fin_acnt_trn_date = rs_acn_tran_jrn!fin_acnt_trn_date
    rs_acn_tran_all!fin_acnt_trn_time = rs_acn_tran_jrn!fin_acnt_trn_time
    rs_acn_tran_all!fin_acnt_trn_wday = rs_acn_tran_jrn!fin_acnt_trn_wday
    rs_acn_tran_all!fin_acnt_trn_ldgr = rs_acn_tran_jrn!fin_acnt_trn_ldgr
    rs_acn_tran_all!fin_acnt_trn_amnt = rs_acn_tran_jrn!fin_acnt_trn_amnt
    rs_acn_tran_all!fin_acnt_trn_side = rs_acn_tran_jrn!fin_acnt_trn_side
    rs_acn_tran_all!fin_acnt_trn_nrtn = rs_acn_tran_jrn!fin_acnt_trn_nrtn
    rs_acn_tran_all!fin_acnt_trn_user = rs_acn_tran_jrn!fin_acnt_trn_user
    rs_acn_tran_all!fin_acnt_trn_vchr = rs_acn_tran_jrn!fin_acnt_trn_vchr
    rs_acn_tran_all!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    rs_acn_tran_all.UpdateBatch
rs_acn_tran_jrn.MoveNext
Next
End Sub
Public Sub copy_all_data_from_contra_to_all()
Dim i_trans
'Call open_database
Call open_rs_acn_tran_cnt
Call open_rs_acn_tran_all
For i_trans = 1 To rs_acn_tran_cnt.RecordCount
    rs_acn_tran_all.AddNew
    rs_acn_tran_all!fin_acnt_trn_vcno = rs_acn_tran_cnt!fin_acnt_trn_vcno
    rs_acn_tran_all!fin_acnt_trn_seno = rs_acn_tran_cnt!fin_acnt_trn_seno
    rs_acn_tran_all!fin_acnt_trn_vtyp = rs_acn_tran_cnt!fin_acnt_trn_vtyp
    rs_acn_tran_all!fin_acnt_trn_date = rs_acn_tran_cnt!fin_acnt_trn_date
    rs_acn_tran_all!fin_acnt_trn_time = rs_acn_tran_cnt!fin_acnt_trn_time
    rs_acn_tran_all!fin_acnt_trn_wday = rs_acn_tran_cnt!fin_acnt_trn_wday
    rs_acn_tran_all!fin_acnt_trn_ldgr = rs_acn_tran_cnt!fin_acnt_trn_ldgr
    rs_acn_tran_all!fin_acnt_trn_amnt = rs_acn_tran_cnt!fin_acnt_trn_amnt
    rs_acn_tran_all!fin_acnt_trn_side = rs_acn_tran_cnt!fin_acnt_trn_side
    rs_acn_tran_all!fin_acnt_trn_nrtn = rs_acn_tran_cnt!fin_acnt_trn_nrtn
    rs_acn_tran_all!fin_acnt_trn_user = rs_acn_tran_cnt!fin_acnt_trn_user
    rs_acn_tran_all!fin_acnt_trn_vchr = rs_acn_tran_cnt!fin_acnt_trn_vchr
    rs_acn_tran_all!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    rs_acn_tran_all.UpdateBatch
rs_acn_tran_cnt.MoveNext
Next
End Sub
Public Sub copy_all_data_from_sales_to_all()
Dim i_trans
'Call open_database
Call open_rs_acn_tran_sal
Call open_rs_acn_tran_all
For i_trans = 1 To rs_acn_tran_sal.RecordCount
    rs_acn_tran_all.AddNew
    rs_acn_tran_all!fin_acnt_trn_vcno = rs_acn_tran_sal!fin_acnt_trn_vcno
    rs_acn_tran_all!fin_acnt_trn_seno = rs_acn_tran_sal!fin_acnt_trn_seno
    rs_acn_tran_all!fin_acnt_trn_vtyp = rs_acn_tran_sal!fin_acnt_trn_vtyp
    rs_acn_tran_all!fin_acnt_trn_date = rs_acn_tran_sal!fin_acnt_trn_date
    rs_acn_tran_all!fin_acnt_trn_time = rs_acn_tran_sal!fin_acnt_trn_time
    rs_acn_tran_all!fin_acnt_trn_wday = rs_acn_tran_sal!fin_acnt_trn_wday
    rs_acn_tran_all!fin_acnt_trn_ldgr = rs_acn_tran_sal!fin_acnt_trn_ldgr
    rs_acn_tran_all!fin_acnt_trn_amnt = rs_acn_tran_sal!fin_acnt_trn_amnt
    rs_acn_tran_all!fin_acnt_trn_side = rs_acn_tran_sal!fin_acnt_trn_side
    rs_acn_tran_all!fin_acnt_trn_nrtn = rs_acn_tran_sal!fin_acnt_trn_nrtn
    rs_acn_tran_all!fin_acnt_trn_user = rs_acn_tran_sal!fin_acnt_trn_user
    rs_acn_tran_all!fin_acnt_trn_vchr = rs_acn_tran_sal!fin_acnt_trn_vchr
    rs_acn_tran_all!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    rs_acn_tran_all.UpdateBatch
rs_acn_tran_sal.MoveNext
Next
End Sub
Public Sub copy_all_data_from_purchase_to_all()
Dim i_trans
'Call open_database
Call open_rs_acn_tran_prs
Call open_rs_acn_tran_all
For i_trans = 1 To rs_acn_tran_prs.RecordCount
    rs_acn_tran_all.AddNew
    rs_acn_tran_all!fin_acnt_trn_vcno = rs_acn_tran_prs!fin_acnt_trn_vcno
    rs_acn_tran_all!fin_acnt_trn_seno = rs_acn_tran_prs!fin_acnt_trn_seno
    rs_acn_tran_all!fin_acnt_trn_vtyp = rs_acn_tran_prs!fin_acnt_trn_vtyp
    rs_acn_tran_all!fin_acnt_trn_date = rs_acn_tran_prs!fin_acnt_trn_date
    rs_acn_tran_all!fin_acnt_trn_time = rs_acn_tran_prs!fin_acnt_trn_time
    rs_acn_tran_all!fin_acnt_trn_wday = rs_acn_tran_prs!fin_acnt_trn_wday
    rs_acn_tran_all!fin_acnt_trn_ldgr = rs_acn_tran_prs!fin_acnt_trn_ldgr
    rs_acn_tran_all!fin_acnt_trn_amnt = rs_acn_tran_prs!fin_acnt_trn_amnt
    rs_acn_tran_all!fin_acnt_trn_side = rs_acn_tran_prs!fin_acnt_trn_side
    rs_acn_tran_all!fin_acnt_trn_nrtn = rs_acn_tran_prs!fin_acnt_trn_nrtn
    rs_acn_tran_all!fin_acnt_trn_user = rs_acn_tran_prs!fin_acnt_trn_user
    rs_acn_tran_all!fin_acnt_trn_vchr = rs_acn_tran_prs!fin_acnt_trn_vchr
    rs_acn_tran_all!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    rs_acn_tran_all.UpdateBatch
rs_acn_tran_prs.MoveNext
Next
End Sub
Public Sub copy_all_data_from_salesreturn_to_all()
Dim i_trans
'Call open_database
Call open_rs_acn_tran_srt
Call open_rs_acn_tran_all
For i_trans = 1 To rs_acn_tran_srt.RecordCount
    rs_acn_tran_all.AddNew
    rs_acn_tran_all!fin_acnt_trn_vcno = rs_acn_tran_srt!fin_acnt_trn_vcno
    rs_acn_tran_all!fin_acnt_trn_seno = rs_acn_tran_srt!fin_acnt_trn_seno
    rs_acn_tran_all!fin_acnt_trn_vtyp = rs_acn_tran_srt!fin_acnt_trn_vtyp
    rs_acn_tran_all!fin_acnt_trn_date = rs_acn_tran_srt!fin_acnt_trn_date
    rs_acn_tran_all!fin_acnt_trn_time = rs_acn_tran_srt!fin_acnt_trn_time
    rs_acn_tran_all!fin_acnt_trn_wday = rs_acn_tran_srt!fin_acnt_trn_wday
    rs_acn_tran_all!fin_acnt_trn_ldgr = rs_acn_tran_srt!fin_acnt_trn_ldgr
    rs_acn_tran_all!fin_acnt_trn_amnt = rs_acn_tran_srt!fin_acnt_trn_amnt
    rs_acn_tran_all!fin_acnt_trn_side = rs_acn_tran_srt!fin_acnt_trn_side
    rs_acn_tran_all!fin_acnt_trn_nrtn = rs_acn_tran_srt!fin_acnt_trn_nrtn
    rs_acn_tran_all!fin_acnt_trn_user = rs_acn_tran_srt!fin_acnt_trn_user
    rs_acn_tran_all!fin_acnt_trn_vchr = rs_acn_tran_srt!fin_acnt_trn_vchr
    rs_acn_tran_all!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    rs_acn_tran_all.UpdateBatch
rs_acn_tran_srt.MoveNext
Next
End Sub
Public Sub copy_all_data_from_purchasereturn_to_all()
Dim i_trans
'Call open_database
Call open_rs_acn_tran_prt
Call open_rs_acn_tran_all
For i_trans = 1 To rs_acn_tran_prt.RecordCount
    rs_acn_tran_all.AddNew
    rs_acn_tran_all!fin_acnt_trn_vcno = rs_acn_tran_prt!fin_acnt_trn_vcno
    rs_acn_tran_all!fin_acnt_trn_seno = rs_acn_tran_prt!fin_acnt_trn_seno
    rs_acn_tran_all!fin_acnt_trn_vtyp = rs_acn_tran_prt!fin_acnt_trn_vtyp
    rs_acn_tran_all!fin_acnt_trn_date = rs_acn_tran_prt!fin_acnt_trn_date
    rs_acn_tran_all!fin_acnt_trn_time = rs_acn_tran_prt!fin_acnt_trn_time
    rs_acn_tran_all!fin_acnt_trn_wday = rs_acn_tran_prt!fin_acnt_trn_wday
    rs_acn_tran_all!fin_acnt_trn_ldgr = rs_acn_tran_prt!fin_acnt_trn_ldgr
    rs_acn_tran_all!fin_acnt_trn_amnt = rs_acn_tran_prt!fin_acnt_trn_amnt
    rs_acn_tran_all!fin_acnt_trn_side = rs_acn_tran_prt!fin_acnt_trn_side
    rs_acn_tran_all!fin_acnt_trn_nrtn = rs_acn_tran_prt!fin_acnt_trn_nrtn
    rs_acn_tran_all!fin_acnt_trn_user = rs_acn_tran_prt!fin_acnt_trn_user
    rs_acn_tran_all!fin_acnt_trn_vchr = rs_acn_tran_prt!fin_acnt_trn_vchr
    rs_acn_tran_all!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    rs_acn_tran_all.UpdateBatch
rs_acn_tran_prt.MoveNext
Next
End Sub

Public Sub unknown_1_subprocedure()
Call copy_all_data_from_to_all_temp
'Call open_database
Call open_rs_acn_tran_all
Call open_rs_acn_tran_all_temp

For i_trans = 1 To rs_acn_tran_all.RecordCount
rs_acn_tran_all.Delete
rs_acn_tran_all.MoveNext
Next

Dim TipeSort As String
TipeSort = "ASC"
   Dim adoSort As ADODB.Recordset
   Set adoSort = New ADODB.Recordset
   adoSort.Open ("select * from rs_acn_tran_all_temp order by fin_acnt_trn_date")
   adoSort.Open "SHAPE " & _
     "{SELECT * FROM " & rs_acn_tran_all_temp.Source & " " & _
     "ORDER BY " & fin_acnt_trn_date & " " & TipeSort & "} " & _
     "AS ParentCMD APPEND " & _
     "({SELECT * FROM " & rs_acn_tran_all_temp.Source & " " & _
     "ORDER BY " & fin_acnt_trn_date & " " & TipeSort & "} " & _
     "AS ChildCMD RELATE " & fin_acnt_trn_time & " TO " & fin_acnt_trn_time & " ) " & _
     "AS ChildCMD", _
     db_co, adOpenStatic, adLockOptimistic
'rs_acn_tran_all.Sort = rs_acn_tran_all!fin_acnt_trn_date
'Call open_database
Call open_rs_acn_tran_all
rs_acn_tran_all.Source = adoSort.Source
rs_acn_tran_all.UpdateBatch

End Sub


Public Sub copy_specific_ledger_transaction_to_acn_tran_spc_lgr()
'Call open_database
'Call collect_all_transaction_in_one_file
Call open_rs_acn_tran_spc_lgr
Do Until rs_acn_tran_spc_lgr.EOF
rs_acn_tran_spc_lgr.Delete
rs_acn_tran_spc_lgr.MoveNext
Loop
Dim temp_entry_no
temp_entry_no = 1
Call open_rs_acn_tran_spc_lgr
Call open_rs_acn_tran_all
Do Until rs_acn_tran_all.EOF

If selected_ledger = rs_acn_tran_all!fin_acnt_trn_ldgr Then

    rs_acn_tran_spc_lgr.AddNew
    rs_acn_tran_spc_lgr!fin_acnt_trn_vcno = rs_acn_tran_all!fin_acnt_trn_vcno
    rs_acn_tran_spc_lgr!fin_acnt_trn_seno = rs_acn_tran_all!fin_acnt_trn_seno
    rs_acn_tran_spc_lgr!fin_acnt_trn_vtyp = rs_acn_tran_all!fin_acnt_trn_vtyp
    rs_acn_tran_spc_lgr!fin_acnt_trn_date = rs_acn_tran_all!fin_acnt_trn_date
    rs_acn_tran_spc_lgr!fin_acnt_trn_time = rs_acn_tran_all!fin_acnt_trn_time
    rs_acn_tran_spc_lgr!fin_acnt_trn_wday = rs_acn_tran_all!fin_acnt_trn_wday
    rs_acn_tran_spc_lgr!fin_acnt_trn_ldgr = rs_acn_tran_all!fin_acnt_trn_ldgr
    rs_acn_tran_spc_lgr!fin_acnt_trn_amnt = rs_acn_tran_all!fin_acnt_trn_amnt
    rs_acn_tran_spc_lgr!fin_acnt_trn_side = rs_acn_tran_all!fin_acnt_trn_side
    rs_acn_tran_spc_lgr!fin_acnt_trn_nrtn = rs_acn_tran_all!fin_acnt_trn_nrtn
    rs_acn_tran_spc_lgr!fin_acnt_trn_user = rs_acn_tran_all!fin_acnt_trn_user
    rs_acn_tran_spc_lgr!fin_acnt_trn_vchr = rs_acn_tran_all!fin_acnt_trn_vchr
    rs_acn_tran_spc_lgr!fin_acnt_trn_id = temp_entry_no
    temp_entry_no = temp_entry_no + 1
    rs_acn_tran_spc_lgr.UpdateBatch
End If
rs_acn_tran_all.MoveNext
Loop

End Sub

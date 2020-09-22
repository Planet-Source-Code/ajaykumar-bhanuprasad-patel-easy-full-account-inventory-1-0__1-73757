Attribute VB_Name = "file_management"
Public db_co As New ADODB.Connection
'=========== employee

Public rs_emp_main_dtl As New ADODB.Recordset
Public rs_emp_tran_dtl As New ADODB.Recordset
Public rs_emp_tran_rep As New ADODB.Recordset
Public rs_emp_tran_tmp As New ADODB.Recordset
'=========== company
Public rs_co_main_dtl As New ADODB.Recordset
Public rs_co_user_dtl As New ADODB.Recordset
'=========== accounting & ledger

Public rs_acn_tran_spc_lgr_print As New ADODB.Recordset
Public rs_lgr_main_dtl As New ADODB.Recordset
Public rs_lgr_main_grp As New ADODB.Recordset
Public rs_lgr_prim_grp As New ADODB.Recordset
Public rs_lgr_bsic_grp As New ADODB.Recordset
Public rs_lgr_clsg_smr As New ADODB.Recordset

Public rs_acn_tran_pmt As New ADODB.Recordset
Public rs_acn_tran_rct As New ADODB.Recordset
Public rs_acn_tran_cnt As New ADODB.Recordset
Public rs_acn_tran_jrn As New ADODB.Recordset
Public rs_acn_tran_sal As New ADODB.Recordset
Public rs_acn_tran_srt As New ADODB.Recordset
Public rs_acn_tran_prs As New ADODB.Recordset
Public rs_acn_tran_prt As New ADODB.Recordset


Public rs_acn_tran_all As New ADODB.Recordset
Public rs_acn_tran_all_temp As New ADODB.Recordset
Public rs_acn_tran_spc_lgr As New ADODB.Recordset

'=========== inventory
Public rs_stk_item_grp As New ADODB.Recordset
Public rs_stk_item_lgr As New ADODB.Recordset
Public rs_stk_item_unt As New ADODB.Recordset
Public rs_stk_open_srl As New ADODB.Recordset
Public rs_stk_clsg_srl As New ADODB.Recordset
Public rs_tmp_clsg_stk As New ADODB.Recordset
Public rs_tmp_spec_itm_clg_stk As New ADODB.Recordset


Public rs_inv_tran_sal As New ADODB.Recordset
Public rs_inv_tran_srt As New ADODB.Recordset
Public rs_inv_tran_prs As New ADODB.Recordset
Public rs_inv_tran_prt As New ADODB.Recordset
Public rs_inv_tran_all As New ADODB.Recordset
Public rs_inv_tran_inw As New ADODB.Recordset
Public rs_inv_tran_otw As New ADODB.Recordset

Public int_i As Integer
'=========== De-Activation project
Public rs_dap_main_dtl As New ADODB.Recordset
Public rs_dap_main_dtl_temp As New ADODB.Recordset
Public rs_dap_main_dtl_all As New ADODB.Recordset
Public rs_dap_rspn_dtl As New ADODB.Recordset

'=========== open De-Activation project
Public Sub open_rs_dap_main_dtl()
If rs_dap_main_dtl.State = 1 Then rs_dap_main_dtl.Close
rs_dap_main_dtl.CursorLocation = adUseClient
rs_dap_main_dtl.Open "Select * From [dap_main_dtl]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_dap_main_dtl_temp()
If rs_dap_main_dtl_temp.State = 1 Then rs_dap_main_dtl_temp.Close
rs_dap_main_dtl_temp.CursorLocation = adUseClient
rs_dap_main_dtl_temp.Open "Select * From [dap_main_dtl_temp]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_dap_main_dtl_all()
If rs_dap_main_dtl_all.State = 1 Then rs_dap_main_dtl_all.Close
rs_dap_main_dtl_all.CursorLocation = adUseClient
rs_dap_main_dtl_all.Open "Select * From [dap_main_dtl_all]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_dap_rspn_dtl()
If rs_dap_rspn_dtl.State = 1 Then rs_dap_rspn_dtl.Close
rs_dap_rspn_dtl.CursorLocation = adUseClient
rs_dap_rspn_dtl.Open "Select * From [dap_rspn_dtl]", db_co, adOpenDynamic, adLockPessimistic
End Sub

'===================== open main file
Public Sub open_database()
If db_co.State = 1 Then db_co.Close
db_co.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & selected_path
db_co.Open
End Sub
'===================== open company database tables
Public Sub open_rs_co_main_dtl()
If rs_co_main_dtl.State = 1 Then rs_co_main_dtl.Close
rs_co_main_dtl.CursorLocation = adUseClient
rs_co_main_dtl.Open "Select * From [co_main_dtl]", db_co, adOpenDynamic, adLockPessimistic
End Sub
'===================== open user database tables
Public Sub open_rs_co_user_dtl()
If rs_co_user_dtl.State = 1 Then rs_co_user_dtl.Close
rs_co_user_dtl.CursorLocation = adUseClient
rs_co_user_dtl.Open "Select * From [co_user_dtl]", db_co, adOpenDynamic, adLockPessimistic
End Sub
'===================== open financial database tables
Public Sub open_rs_lgr_main_dtl()
If rs_lgr_main_dtl.State = 1 Then rs_lgr_main_dtl.Close
rs_lgr_main_dtl.CursorLocation = adUseClient
rs_lgr_main_dtl.Open "Select * From [lgr_main_dtl]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_lgr_clsg_smr()
If rs_lgr_clsg_smr.State = 1 Then rs_lgr_clsg_smr.Close
rs_lgr_clsg_smr.CursorLocation = adUseClient
rs_lgr_clsg_smr.Open "Select * From [lgr_clsg_smr]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_lgr_main_grp()
If rs_lgr_main_grp.State = 1 Then rs_lgr_main_grp.Close
rs_lgr_main_grp.CursorLocation = adUseClient
rs_lgr_main_grp.Open "Select * From [lgr_main_grp]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_lgr_prim_grp()
If rs_lgr_prim_grp.State = 1 Then rs_lgr_prim_grp.Close
rs_lgr_prim_grp.CursorLocation = adUseClient
rs_lgr_prim_grp.Open "Select * From [lgr_prim_grp]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_lgr_bsic_grp()
If rs_lgr_bsic_grp.State = 1 Then rs_lgr_bsic_grp.Close
rs_lgr_bsic_grp.CursorLocation = adUseClient
rs_lgr_bsic_grp.Open "Select * From [lgr_bsic_grp]", db_co, adOpenDynamic, adLockPessimistic
End Sub
'===================== open inventory database tables
Public Sub open_rs_stk_item_lgr()
If rs_stk_item_lgr.State = 1 Then rs_stk_item_lgr.Close
rs_stk_item_lgr.CursorLocation = adUseClient
rs_stk_item_lgr.Open "Select * From [stk_item_lgr]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_stk_item_grp()
If rs_stk_item_grp.State = 1 Then rs_stk_item_grp.Close
rs_stk_item_grp.CursorLocation = adUseClient
rs_stk_item_grp.Open "Select * From [stk_item_grp]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_stk_item_unt()
If rs_stk_item_unt.State = 1 Then rs_stk_item_unt.Close
rs_stk_item_unt.CursorLocation = adUseClient
rs_stk_item_unt.Open "Select * From [stk_item_unt]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_stk_open_srl()
If rs_stk_open_srl.State = 1 Then rs_stk_open_srl.Close
rs_stk_open_srl.CursorLocation = adUseClient
rs_stk_open_srl.Open "Select * From [stk_open_srl]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_stk_clsg_srl()
If rs_stk_clsg_srl.State = 1 Then rs_stk_clsg_srl.Close
rs_stk_clsg_srl.CursorLocation = adUseClient
rs_stk_clsg_srl.Open "Select * From [stk_clsg_srl]", db_co, adOpenDynamic, adLockPessimistic
End Sub


'===================== open employee database tables
Public Sub open_rs_emp_main_dtl()
If rs_emp_main_dtl.State = 1 Then rs_emp_main_dtl.Close
rs_emp_main_dtl.CursorLocation = adUseClient
rs_emp_main_dtl.Open "Select * From [emp_main_dtl]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_emp_tran_dtl()
If rs_emp_tran_dtl.State = 1 Then rs_emp_tran_dtl.Close
rs_emp_tran_dtl.CursorLocation = adUseClient
rs_emp_tran_dtl.Open "Select * From [emp_tran_dtl]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_emp_tran_rep()
If rs_emp_tran_rep.State = 1 Then rs_emp_tran_rep.Close
rs_emp_tran_rep.CursorLocation = adUseClient
rs_emp_tran_rep.Open "Select * From [emp_tran_rep]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_emp_tran_tmp()
If rs_emp_tran_tmp.State = 1 Then rs_emp_tran_tmp.Close
rs_emp_tran_tmp.CursorLocation = adUseClient
rs_emp_tran_tmp.Open "Select * From [emp_tran_tmp]", db_co, adOpenDynamic, adLockPessimistic
End Sub
'===================== open acccounting voucher database tables
Public Sub open_rs_acn_tran_pmt()
If rs_acn_tran_pmt.State = 1 Then rs_acn_tran_pmt.Close
rs_acn_tran_pmt.CursorLocation = adUseClient
rs_acn_tran_pmt.Open "Select * From [acn_tran_pmt]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_acn_tran_rct()
If rs_acn_tran_rct.State = 1 Then rs_acn_tran_rct.Close
rs_acn_tran_rct.CursorLocation = adUseClient
rs_acn_tran_rct.Open "Select * From [acn_tran_rct]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_acn_tran_cnt()
If rs_acn_tran_cnt.State = 1 Then rs_acn_tran_cnt.Close
rs_acn_tran_cnt.CursorLocation = adUseClient
rs_acn_tran_cnt.Open "Select * From [acn_tran_cnt]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_acn_tran_jrn()
If rs_acn_tran_jrn.State = 1 Then rs_acn_tran_jrn.Close
rs_acn_tran_jrn.CursorLocation = adUseClient
rs_acn_tran_jrn.Open "Select * From [acn_tran_jrn]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_acn_tran_sal()
If rs_acn_tran_sal.State = 1 Then rs_acn_tran_sal.Close
rs_acn_tran_sal.CursorLocation = adUseClient
rs_acn_tran_sal.Open "Select * From [acn_tran_sal]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_acn_tran_prs()
If rs_acn_tran_prs.State = 1 Then rs_acn_tran_prs.Close
rs_acn_tran_prs.CursorLocation = adUseClient
rs_acn_tran_prs.Open "Select * From [acn_tran_prs]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_acn_tran_prt()
If rs_acn_tran_prt.State = 1 Then rs_acn_tran_prt.Close
rs_acn_tran_prt.CursorLocation = adUseClient
rs_acn_tran_prt.Open "Select * From [acn_tran_prt]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_acn_tran_srt()
If rs_acn_tran_srt.State = 1 Then rs_acn_tran_srt.Close
rs_acn_tran_srt.CursorLocation = adUseClient
rs_acn_tran_srt.Open "Select * From [acn_tran_srt]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_acn_tran_all()
If rs_acn_tran_all.State = 1 Then rs_acn_tran_all.Close
rs_acn_tran_all.CursorLocation = adUseClient
rs_acn_tran_all.Open "Select * From [acn_tran_all]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_acn_tran_all_temp()
If rs_acn_tran_all_temp.State = 1 Then rs_acn_tran_all_temp.Close
rs_acn_tran_all_temp.CursorLocation = adUseClient
rs_acn_tran_all_temp.Open "Select * From [acn_tran_all_temp]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_acn_tran_spc_lgr()
If rs_acn_tran_spc_lgr.State = 1 Then rs_acn_tran_spc_lgr.Close
rs_acn_tran_spc_lgr.CursorLocation = adUseClient
rs_acn_tran_spc_lgr.Open "Select * From [acn_tran_spc_lgr]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_acn_tran_spc_lgr_print()
If rs_acn_tran_spc_lgr_print.State = 1 Then rs_acn_tran_spc_lgr_print.Close
rs_acn_tran_spc_lgr_print.CursorLocation = adUseClient
rs_acn_tran_spc_lgr_print.Open "Select * From [acn_tran_spc_lgr_print]", db_co, adOpenDynamic, adLockPessimistic
End Sub

'===================== open inventory voucher database tables
Public Sub open_rs_inv_tran_sal()
If rs_inv_tran_sal.State = 1 Then rs_inv_tran_sal.Close
rs_inv_tran_sal.CursorLocation = adUseClient
rs_inv_tran_sal.Open "Select * From [inv_tran_sal]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_inv_tran_prs()
If rs_inv_tran_prs.State = 1 Then rs_inv_tran_prs.Close
rs_inv_tran_prs.CursorLocation = adUseClient
rs_inv_tran_prs.Open "Select * From [inv_tran_prs]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_inv_tran_srt()
If rs_inv_tran_srt.State = 1 Then rs_inv_tran_srt.Close
rs_inv_tran_srt.CursorLocation = adUseClient
rs_inv_tran_srt.Open "Select * From [inv_tran_srt]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_inv_tran_prt()
If rs_inv_tran_prt.State = 1 Then rs_inv_tran_prt.Close
rs_inv_tran_prt.CursorLocation = adUseClient
rs_inv_tran_prt.Open "Select * From [inv_tran_prt]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_inv_tran_all()
If rs_inv_tran_all.State = 1 Then rs_inv_tran_all.Close
rs_inv_tran_all.CursorLocation = adUseClient
rs_inv_tran_all.Open "Select * From [inv_tran_all]", db_co, adOpenDynamic, adLockPessimistic

End Sub
Public Sub open_rs_inv_tran_inw()
If rs_inv_tran_inw.State = 1 Then rs_inv_tran_inw.Close
rs_inv_tran_inw.CursorLocation = adUseClient
rs_inv_tran_inw.Open "Select * From [inv_tran_inw]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_inv_tran_otw()
If rs_inv_tran_otw.State = 1 Then rs_inv_tran_otw.Close
rs_inv_tran_otw.CursorLocation = adUseClient
rs_inv_tran_otw.Open "Select * From [inv_tran_otw]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_tmp_clsg_stk()
If rs_tmp_clsg_stk.State = 1 Then rs_tmp_clsg_stk.Close
rs_tmp_clsg_stk.CursorLocation = adUseClient
rs_tmp_clsg_stk.Open "Select * From [tmp_clsg_stk ]", db_co, adOpenDynamic, adLockPessimistic
End Sub
Public Sub open_rs_tmp_spec_itm_clg_stk()
If rs_tmp_spec_itm_clg_stk.State = 1 Then rs_tmp_spec_itm_clg_stk.Close
rs_tmp_spec_itm_clg_stk.CursorLocation = adUseClient
rs_tmp_spec_itm_clg_stk.Open "Select * From [tmp_spec_itm_clg_stk ]", db_co, adOpenDynamic, adLockPessimistic
End Sub


Attribute VB_Name = "closing_stock"
Public Sub remove_rs_stk_clsg_srl()
Call open_database
Call open_rs_stk_clsg_srl
'rs_stk_clsg_srl.MoveFirst
Do Until rs_stk_clsg_srl.EOF
    rs_stk_clsg_srl.Delete
    rs_stk_clsg_srl.MoveNext
Loop
End Sub
Public Sub copy_all_inward_to_stk()
Call open_database
Call open_rs_inv_tran_otw
Call open_rs_inv_tran_inw
Call open_rs_stk_clsg_srl
If rs_inv_tran_inw.RecordCount > 0 Then rs_inv_tran_inw.MoveFirst
Do Until rs_inv_tran_inw.EOF
With rs_inv_tran_inw
    rs_stk_clsg_srl.AddNew
    If !stk_invt_trn_date <> "" Then rs_stk_clsg_srl!stk_invt_clg_date = !stk_invt_trn_date
    If !stk_invt_trn_card <> "" Then rs_stk_clsg_srl!stk_invt_clg_card = !stk_invt_trn_card
    If !stk_invt_trn_stno <> "" Then rs_stk_clsg_srl!stk_invt_clg_stno = !stk_invt_trn_stno
    If !stk_invt_trn_edno <> "" Then rs_stk_clsg_srl!stk_invt_clg_edno = !stk_invt_trn_edno
    If !stk_invt_trn_qnty <> "" Then rs_stk_clsg_srl!stk_invt_clg_qnty = !stk_invt_trn_qnty
    If !stk_invt_trn_rate <> 0 Then rs_stk_clsg_srl!stk_invt_clg_rate = !stk_invt_trn_rate
    If !stk_invt_trn_amnt <> 0 Then rs_stk_clsg_srl!stk_invt_clg_amnt = !stk_invt_trn_amnt
    If !stk_invt_trn_fval <> 0 Then rs_stk_clsg_srl!stk_invt_clg_fval = !stk_invt_trn_fval
    If !stk_invt_trn_splr <> "" Then rs_stk_clsg_srl!stk_invt_clg_splr = !stk_invt_trn_splr
    If !stk_invt_trn_vtyp <> 0 Then rs_stk_clsg_srl!stk_invt_clg_vtyp = !stk_invt_trn_vtyp
    If !stk_invt_trn_vchr <> "" Then rs_stk_clsg_srl!stk_invt_clg_vchr = !stk_invt_trn_vchr
    rs_stk_clsg_srl.UpdateBatch
End With
rs_inv_tran_inw.MoveNext
Loop
End Sub
Public Sub search_closing_stock()

Call remove_rs_stk_clsg_srl
Call copy_all_inward_to_stk

Call open_database
Call open_rs_stk_clsg_srl
Call open_rs_inv_tran_otw
Call open_rs_tmp_clsg_stk

Dim stk_starting_srl_no
Dim stk_ending_srl_no

Dim temp1_stk_starting_srl_no
Dim temp1_stk_ending_srl_no

Dim temp2_stk_starting_srl_no
Dim temp2_stk_ending_srl_no

Dim out_starting_srl_no
Dim out_ending_srl_no

Do Until rs_inv_tran_otw.EOF

    Call open_rs_stk_clsg_srl
    If rs_stk_clsg_srl.RecordCount > 0 Then rs_stk_clsg_srl.MoveFirst
    Do Until rs_stk_clsg_srl.EOF

        stk_starting_srl_no = Val(rs_stk_clsg_srl!stk_invt_clg_stno)
        stk_ending_srl_no = Val(rs_stk_clsg_srl!stk_invt_clg_edno)
        
        out_starting_srl_no = Val(rs_inv_tran_otw!stk_invt_trn_stno)
        out_ending_srl_no = Val(rs_inv_tran_otw!stk_invt_trn_edno)
                    
'MsgBox "opg_stock.." & stk_starting_srl_no & "-" & stk_ending_srl_no
'MsgBox "out_stock.." & out_starting_srl_no & "-" & out_ending_srl_no
'MsgBox stk_ending_srl_no & "=" & out_ending_srl_no
        
        If stk_starting_srl_no <= out_starting_srl_no And stk_ending_srl_no >= out_ending_srl_no Then
            
            'clear all in temp stock entry
            'calculate and add to temp stock
                '1...if serial no is on begining
                '2...if serial no is on ending
                '3...if serial no is on middle
            'delete original stock
            'add temp stock entry to original stock
            'clear all in temp stock entry
            
            'clear all in temp stock entry
            'calculate and add to temp stock
                If stk_starting_srl_no = out_starting_srl_no Then       '1...if serial no is on begining
                    'stk =12500-13500
                    'sale=12500-12510
                    'temp=12511-13500
                    
                    temp1_stk_starting_srl_no = out_ending_srl_no + 1
                    temp1_stk_ending_srl_no = stk_ending_srl_no
                    'MsgBox "clg_stock.." & temp1_stk_starting_srl_no & "-" & temp1_stk_ending_srl_no
                    rs_tmp_clsg_stk.AddNew
                        If rs_stk_clsg_srl!stk_invt_clg_card <> "" Then rs_tmp_clsg_stk!stk_invt_clg_card = rs_stk_clsg_srl!stk_invt_clg_card
                        If rs_stk_clsg_srl!stk_invt_clg_stno <> "" Then rs_tmp_clsg_stk!stk_invt_clg_stno = temp1_stk_starting_srl_no
                        If rs_stk_clsg_srl!stk_invt_clg_edno <> "" Then rs_tmp_clsg_stk!stk_invt_clg_edno = temp1_stk_ending_srl_no
                        If rs_stk_clsg_srl!stk_invt_clg_qnty <> "" Then rs_tmp_clsg_stk!stk_invt_clg_qnty = temp1_stk_ending_srl_no - temp1_stk_starting_srl_no + 1
                        If rs_stk_clsg_srl!stk_invt_clg_rate <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_rate = rs_stk_clsg_srl!stk_invt_clg_rate
                        If rs_stk_clsg_srl!stk_invt_clg_amnt <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_amnt = rs_stk_clsg_srl!stk_invt_clg_amnt
                        If rs_stk_clsg_srl!stk_invt_clg_fval <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_fval = rs_stk_clsg_srl!stk_invt_clg_fval
                        If rs_stk_clsg_srl!stk_invt_clg_splr <> "" Then rs_tmp_clsg_stk!stk_invt_clg_splr = rs_stk_clsg_srl!stk_invt_clg_splr
                        If rs_stk_clsg_srl!stk_invt_clg_vtyp <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_vtyp = rs_stk_clsg_srl!stk_invt_clg_vtyp
                        If rs_stk_clsg_srl!stk_invt_clg_vchr <> "" Then rs_tmp_clsg_stk!stk_invt_clg_vchr = "closing stock"
                    rs_tmp_clsg_stk.UpdateBatch
                        'delete original stock
                        rs_stk_clsg_srl.Delete
                        rs_stk_clsg_srl.UpdateBatch

                ElseIf stk_ending_srl_no = out_ending_srl_no Then       '2...if serial no is on ending
                    'stk =12500-13500
                    'sale=13481-13500
                    'temp=12500-13480
                    temp1_stk_starting_srl_no = stk_starting_srl_no
                    temp1_stk_ending_srl_no = out_starting_srl_no - 1
                    'MsgBox "clg_stock.." & temp1_stk_starting_srl_no & "-" & temp1_stk_ending_srl_no
                    rs_tmp_clsg_stk.AddNew
                        If rs_stk_clsg_srl!stk_invt_clg_card <> "" Then rs_tmp_clsg_stk!stk_invt_clg_card = rs_stk_clsg_srl!stk_invt_clg_card
                        If rs_stk_clsg_srl!stk_invt_clg_stno <> "" Then rs_tmp_clsg_stk!stk_invt_clg_stno = temp1_stk_starting_srl_no
                        If rs_stk_clsg_srl!stk_invt_clg_edno <> "" Then rs_tmp_clsg_stk!stk_invt_clg_edno = temp1_stk_ending_srl_no
                        If rs_stk_clsg_srl!stk_invt_clg_qnty <> "" Then rs_tmp_clsg_stk!stk_invt_clg_qnty = temp1_stk_ending_srl_no - temp1_stk_starting_srl_no + 1
                        If rs_stk_clsg_srl!stk_invt_clg_rate <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_rate = rs_stk_clsg_srl!stk_invt_clg_rate
                        If rs_stk_clsg_srl!stk_invt_clg_amnt <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_amnt = rs_stk_clsg_srl!stk_invt_clg_amnt
                        If rs_stk_clsg_srl!stk_invt_clg_fval <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_fval = rs_stk_clsg_srl!stk_invt_clg_fval
                        If rs_stk_clsg_srl!stk_invt_clg_splr <> "" Then rs_tmp_clsg_stk!stk_invt_clg_splr = rs_stk_clsg_srl!stk_invt_clg_splr
                        If rs_stk_clsg_srl!stk_invt_clg_vtyp <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_vtyp = rs_stk_clsg_srl!stk_invt_clg_vtyp
                        If rs_stk_clsg_srl!stk_invt_clg_vchr <> "" Then rs_tmp_clsg_stk!stk_invt_clg_vchr = "closing stock"
                    rs_tmp_clsg_stk.UpdateBatch
                        'delete original stock
                        rs_stk_clsg_srl.Delete
                        rs_stk_clsg_srl.UpdateBatch
                
                Else                                '3...if serial no is on middle
                    '     st----end
                    'stk =12500-13500
                    'sale=12511-12520
                    
                    'temp1=12500-12510
                    'temp2=12521-13500
                    
                    temp1_stk_starting_srl_no = stk_starting_srl_no
                    temp1_stk_ending_srl_no = out_starting_srl_no - 1
                    'MsgBox "clg_stock.." & temp1_stk_starting_srl_no & "-" & temp1_stk_ending_srl_no
                    rs_tmp_clsg_stk.AddNew
                    If rs_stk_clsg_srl!stk_invt_clg_card <> "" Then rs_tmp_clsg_stk!stk_invt_clg_card = rs_stk_clsg_srl!stk_invt_clg_card
                    If rs_stk_clsg_srl!stk_invt_clg_stno <> "" Then rs_tmp_clsg_stk!stk_invt_clg_stno = temp1_stk_starting_srl_no
                    If rs_stk_clsg_srl!stk_invt_clg_edno <> "" Then rs_tmp_clsg_stk!stk_invt_clg_edno = temp1_stk_ending_srl_no
                    If rs_stk_clsg_srl!stk_invt_clg_qnty <> "" Then rs_tmp_clsg_stk!stk_invt_clg_qnty = temp1_stk_ending_srl_no - temp1_stk_starting_srl_no + 1
                    If rs_stk_clsg_srl!stk_invt_clg_rate <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_rate = rs_stk_clsg_srl!stk_invt_clg_rate
                    If rs_stk_clsg_srl!stk_invt_clg_amnt <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_amnt = rs_stk_clsg_srl!stk_invt_clg_amnt
                    If rs_stk_clsg_srl!stk_invt_clg_fval <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_fval = rs_stk_clsg_srl!stk_invt_clg_fval
                    If rs_stk_clsg_srl!stk_invt_clg_splr <> "" Then rs_tmp_clsg_stk!stk_invt_clg_splr = rs_stk_clsg_srl!stk_invt_clg_splr
                    If rs_stk_clsg_srl!stk_invt_clg_vtyp <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_vtyp = rs_stk_clsg_srl!stk_invt_clg_vtyp
                    If rs_stk_clsg_srl!stk_invt_clg_vchr <> "" Then rs_tmp_clsg_stk!stk_invt_clg_vchr = "closing stock"
                    rs_tmp_clsg_stk.UpdateBatch
                    
                    
                    temp2_stk_starting_srl_no = out_ending_srl_no + 1
                    temp2_stk_ending_srl_no = stk_ending_srl_no

                    
                    rs_tmp_clsg_stk.AddNew
                        If rs_stk_clsg_srl!stk_invt_clg_card <> "" Then rs_tmp_clsg_stk!stk_invt_clg_card = rs_stk_clsg_srl!stk_invt_clg_card
                        If rs_stk_clsg_srl!stk_invt_clg_stno <> "" Then rs_tmp_clsg_stk!stk_invt_clg_stno = temp2_stk_starting_srl_no
                        If rs_stk_clsg_srl!stk_invt_clg_edno <> "" Then rs_tmp_clsg_stk!stk_invt_clg_edno = temp2_stk_ending_srl_no
                        If rs_stk_clsg_srl!stk_invt_clg_qnty <> "" Then rs_tmp_clsg_stk!stk_invt_clg_qnty = temp2_stk_ending_srl_no - temp1_stk_starting_srl_no + 1
                        If rs_stk_clsg_srl!stk_invt_clg_rate <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_rate = rs_stk_clsg_srl!stk_invt_clg_rate
                        If rs_stk_clsg_srl!stk_invt_clg_amnt <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_amnt = rs_stk_clsg_srl!stk_invt_clg_amnt
                        If rs_stk_clsg_srl!stk_invt_clg_fval <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_fval = rs_stk_clsg_srl!stk_invt_clg_fval
                        If rs_stk_clsg_srl!stk_invt_clg_splr <> "" Then rs_tmp_clsg_stk!stk_invt_clg_splr = rs_stk_clsg_srl!stk_invt_clg_splr
                        If rs_stk_clsg_srl!stk_invt_clg_vtyp <> 0 Then rs_tmp_clsg_stk!stk_invt_clg_vtyp = rs_stk_clsg_srl!stk_invt_clg_vtyp
                        If rs_stk_clsg_srl!stk_invt_clg_vchr <> "" Then rs_tmp_clsg_stk!stk_invt_clg_vchr = "closing stock"
                    rs_tmp_clsg_stk.UpdateBatch
                        'delete original stock
                        rs_stk_clsg_srl.Delete
                        rs_stk_clsg_srl.UpdateBatch
                End If
            'add temp stock entry to original stock
            Call open_rs_tmp_clsg_stk
            If rs_tmp_clsg_stk.RecordCount > 0 Then rs_tmp_clsg_stk.MoveFirst
            Do Until rs_tmp_clsg_stk.EOF
                    rs_stk_clsg_srl.AddNew
                    If rs_tmp_clsg_stk!stk_invt_clg_card <> "" Then rs_stk_clsg_srl!stk_invt_clg_card = rs_tmp_clsg_stk!stk_invt_clg_card
                    If rs_tmp_clsg_stk!stk_invt_clg_stno <> "" Then rs_stk_clsg_srl!stk_invt_clg_stno = rs_tmp_clsg_stk!stk_invt_clg_stno
                    If rs_tmp_clsg_stk!stk_invt_clg_edno <> "" Then rs_stk_clsg_srl!stk_invt_clg_edno = rs_tmp_clsg_stk!stk_invt_clg_edno
                    If rs_tmp_clsg_stk!stk_invt_clg_qnty <> "" Then rs_stk_clsg_srl!stk_invt_clg_qnty = rs_tmp_clsg_stk!stk_invt_clg_qnty
                    If rs_tmp_clsg_stk!stk_invt_clg_rate <> 0 Then rs_stk_clsg_srl!stk_invt_clg_rate = rs_tmp_clsg_stk!stk_invt_clg_rate
                    If rs_tmp_clsg_stk!stk_invt_clg_amnt <> 0 Then rs_stk_clsg_srl!stk_invt_clg_amnt = rs_tmp_clsg_stk!stk_invt_clg_amnt
                    If rs_tmp_clsg_stk!stk_invt_clg_fval <> 0 Then rs_stk_clsg_srl!stk_invt_clg_fval = rs_tmp_clsg_stk!stk_invt_clg_fval
                    If rs_tmp_clsg_stk!stk_invt_clg_splr <> "" Then rs_stk_clsg_srl!stk_invt_clg_splr = rs_tmp_clsg_stk!stk_invt_clg_splr
                    If rs_tmp_clsg_stk!stk_invt_clg_vtyp <> 0 Then rs_stk_clsg_srl!stk_invt_clg_vtyp = rs_tmp_clsg_stk!stk_invt_clg_vtyp
                    If rs_tmp_clsg_stk!stk_invt_clg_vchr <> "" Then rs_stk_clsg_srl!stk_invt_clg_vchr = "closing stock"
                    rs_stk_clsg_srl.UpdateBatch
            rs_tmp_clsg_stk.MoveNext
            Loop
            
            'clear all in temp stock entry
            Call open_rs_tmp_clsg_stk
            If rs_tmp_clsg_stk.RecordCount > 0 Then rs_tmp_clsg_stk.MoveFirst
            Do Until rs_tmp_clsg_stk.EOF
                rs_tmp_clsg_stk.Delete
                rs_tmp_clsg_stk.MoveNext
            Loop
    End If
    rs_stk_clsg_srl.MoveNext
    Loop
    rs_inv_tran_otw.MoveNext
Loop
End Sub
Public Sub separation_of_all_inventory_to_inward_and_outward()
Dim rs_inv_tran_all_counter
Dim grid_stk_row_no
Dim total_inward
Dim total_outward
Dim temp_stock_balance
Dim temp_i

Call refresh_closing_stock
Call open_database
Call open_rs_inv_tran_inw

For temp_i = rs_inv_tran_inw.RecordCount To 1 Step -1
    rs_inv_tran_inw.Delete
    rs_inv_tran_inw.MoveNext
Next
Call open_rs_inv_tran_otw
For temp_i = rs_inv_tran_otw.RecordCount To 1 Step -1
    rs_inv_tran_otw.Delete
    rs_inv_tran_otw.MoveNext
Next
Dim out_entry_no
Dim in_entry_no
in_entry_no = 1
out_entry_no = 1
Call open_rs_stk_item_lgr
Call open_rs_stk_clsg_srl
Call open_database
Call open_rs_inv_tran_all
'MsgBox rs_inv_tran_all.RecordCount
Call open_rs_inv_tran_inw
Call open_rs_inv_tran_otw
'Do Until rs_stk_item_lgr.EOF
'selected_stock_item_name = rs_stk_item_lgr!stk_item_lgr_name
    If rs_inv_tran_all.RecordCount > 0 Then rs_inv_tran_all.MoveFirst
    Do Until rs_inv_tran_all.EOF
        'If rs_inv_tran_all!stk_invt_trn_card = selected_stock_item_name Then
        With rs_inv_tran_all
                    If LCase(!stk_invt_trn_vchr) = "purchase" Or LCase(!stk_invt_trn_vchr) = "sale return" Or LCase(!stk_invt_trn_vchr) = "opening stock" Then
                    
                                rs_inv_tran_inw.AddNew
                                rs_inv_tran_inw!stk_invt_trn_id = in_entry_no
                                
                                If !stk_invt_trn_date <> "" Then rs_inv_tran_inw!stk_invt_trn_date = !stk_invt_trn_date
                                If !stk_invt_trn_time <> "" Then rs_inv_tran_inw!stk_invt_trn_time = !stk_invt_trn_time
                                If !stk_invt_trn_wday <> "" Then rs_inv_tran_inw!stk_invt_trn_wday = !stk_invt_trn_wday
                                If !stk_invt_trn_vcno <> "" Then rs_inv_tran_inw!stk_invt_trn_vcno = !stk_invt_trn_vcno
                                If !stk_invt_trn_ldgr <> "" Then rs_inv_tran_inw!stk_invt_trn_ldgr = !stk_invt_trn_ldgr
                                If !stk_invt_trn_card <> "" Then rs_inv_tran_inw!stk_invt_trn_card = !stk_invt_trn_card
                                If !stk_invt_trn_stno <> "" Then rs_inv_tran_inw!stk_invt_trn_stno = !stk_invt_trn_stno
                                If !stk_invt_trn_edno <> "" Then rs_inv_tran_inw!stk_invt_trn_edno = !stk_invt_trn_edno
                                If !stk_invt_trn_qnty <> "" Then rs_inv_tran_inw!stk_invt_trn_qnty = !stk_invt_trn_qnty
                                If !stk_invt_trn_rate <> 0 Then rs_inv_tran_inw!stk_invt_trn_rate = !stk_invt_trn_rate
                                If !stk_invt_trn_amnt <> 0 Then rs_inv_tran_inw!stk_invt_trn_amnt = !stk_invt_trn_amnt
                                If !stk_invt_trn_fval <> 0 Then rs_inv_tran_inw!stk_invt_trn_fval = !stk_invt_trn_fval
                                If !stk_invt_trn_splr <> "" Then rs_inv_tran_inw!stk_invt_trn_splr = !stk_invt_trn_splr
                                If !stk_invt_trn_vtyp <> 0 Then rs_inv_tran_inw!stk_invt_trn_vtyp = !stk_invt_trn_vtyp
                                If !stk_invt_trn_vchr <> "" Then rs_inv_tran_inw!stk_invt_trn_vchr = !stk_invt_trn_vchr
                                in_entry_no = in_entry_no + 1
                                rs_inv_tran_inw.UpdateBatch
                    End If
                    If LCase(!stk_invt_trn_vchr) = "sale" Or LCase(!stk_invt_trn_vchr) = "purchase return" Then
                                rs_inv_tran_otw.AddNew
                                rs_inv_tran_otw!stk_invt_trn_id = out_entry_no
                                If !stk_invt_trn_date <> "" Then rs_inv_tran_otw!stk_invt_trn_date = !stk_invt_trn_date
                                If !stk_invt_trn_time <> "" Then rs_inv_tran_otw!stk_invt_trn_time = !stk_invt_trn_time
                                If !stk_invt_trn_wday <> "" Then rs_inv_tran_otw!stk_invt_trn_wday = !stk_invt_trn_wday
                                If !stk_invt_trn_vcno <> "" Then rs_inv_tran_otw!stk_invt_trn_vcno = !stk_invt_trn_vcno
                                
                                If !stk_invt_trn_ldgr <> "" Then rs_inv_tran_otw!stk_invt_trn_ldgr = !stk_invt_trn_ldgr
                                If !stk_invt_trn_card <> "" Then rs_inv_tran_otw!stk_invt_trn_card = !stk_invt_trn_card
                                If !stk_invt_trn_stno <> "" Then rs_inv_tran_otw!stk_invt_trn_stno = !stk_invt_trn_stno
                                If !stk_invt_trn_edno <> "" Then rs_inv_tran_otw!stk_invt_trn_edno = !stk_invt_trn_edno
                                If !stk_invt_trn_qnty <> "" Then rs_inv_tran_otw!stk_invt_trn_qnty = !stk_invt_trn_qnty
                                If !stk_invt_trn_rate <> 0 Then rs_inv_tran_otw!stk_invt_trn_rate = !stk_invt_trn_rate
                                If !stk_invt_trn_amnt <> 0 Then rs_inv_tran_otw!stk_invt_trn_amnt = !stk_invt_trn_amnt
                                If !stk_invt_trn_fval <> 0 Then rs_inv_tran_otw!stk_invt_trn_fval = !stk_invt_trn_fval
                                If !stk_invt_trn_splr <> "" Then rs_inv_tran_otw!stk_invt_trn_splr = !stk_invt_trn_splr
                                If !stk_invt_trn_vtyp <> 0 Then rs_inv_tran_otw!stk_invt_trn_vtyp = !stk_invt_trn_vtyp
                                If !stk_invt_trn_vchr <> "" Then rs_inv_tran_otw!stk_invt_trn_vchr = !stk_invt_trn_vchr
                                out_entry_no = out_entry_no + 1
                                rs_inv_tran_otw.UpdateBatch
                    End If
        End With
        'End If
    rs_inv_tran_all.MoveNext
    Loop
rs_inv_tran_inw.Sort = "stk_invt_trn_card ,stk_invt_trn_stno"
rs_inv_tran_otw.Sort = "stk_invt_trn_card ,stk_invt_trn_stno"
'rs_stk_item_lgr.MoveNext
'Loop
End Sub
Public Sub refresh_closing_stock()
Dim temp_rs_inv_tran_all
temp_rs_inv_tran_all = 1
Dim temp_counter
Call open_database
Call open_rs_inv_tran_all
For temp_counter = 1 To rs_inv_tran_all.RecordCount
    rs_inv_tran_all.Delete
    rs_inv_tran_all.MoveNext
Next

Call open_database
Call open_rs_stk_open_srl
Call open_rs_inv_tran_all
For temp_counter = 1 To rs_stk_open_srl.RecordCount

    rs_inv_tran_all.AddNew
    rs_inv_tran_all!stk_invt_trn_id = temp_rs_inv_tran_all
    rs_inv_tran_all!stk_invt_trn_vcno = rs_stk_open_srl!stk_open_srl_stid
    'rs_inv_tran_all!stk_invt_trn_seno = rs_inv_tran_sal!stk_invt_trn_seno
    rs_inv_tran_all!stk_invt_trn_date = selected_companies_starting_f_date
    'rs_inv_tran_all!stk_invt_trn_time = rs_inv_tran_sal!stk_invt_trn_time
    'rs_inv_tran_all!stk_invt_trn_wday = rs_inv_tran_sal!stk_invt_trn_wday
    'rs_inv_tran_all!stk_invt_trn_vtyp = rs_inv_tran_sal!stk_invt_trn_vtyp
    'rs_inv_tran_all!stk_invt_trn_ldgr = rs_inv_tran_sal!stk_invt_trn_ldgr
    rs_inv_tran_all!stk_invt_trn_card = rs_stk_open_srl!stk_open_srl_name
    rs_inv_tran_all!stk_invt_trn_stno = rs_stk_open_srl!stk_open_srl_stno
    rs_inv_tran_all!stk_invt_trn_edno = rs_stk_open_srl!stk_open_srl_edno
    rs_inv_tran_all!stk_invt_trn_qnty = rs_stk_open_srl!stk_open_srl_qnty
    rs_inv_tran_all!stk_invt_trn_rate = rs_stk_open_srl!stk_open_srl_rate
    rs_inv_tran_all!stk_invt_trn_amnt = rs_stk_open_srl!stk_open_srl_rate
    rs_inv_tran_all!stk_invt_trn_fval = rs_stk_open_srl!stk_open_srl_fcvl
    rs_inv_tran_all!stk_invt_trn_splr = rs_stk_open_srl!stk_open_srl_splr
    'rs_inv_tran_all!stk_invt_trn_comp = rs_inv_tran_sal!stk_invt_trn_comp
    'rs_inv_tran_all!stk_invt_trn_user = rs_inv_tran_sal!stk_invt_trn_user
    rs_inv_tran_all!stk_invt_trn_vchr = "opening stock"
    'rs_inv_tran_all!stk_invt_trn_slmn = rs_inv_tran_sal!stk_invt_trn_slmn
    rs_inv_tran_all!stk_invt_trn_vtyp = rs_stk_open_srl!stk_open_srl_type
    rs_inv_tran_all!stk_invt_trn_splr = rs_stk_open_srl!stk_open_srl_splr
    temp_rs_inv_tran_all = temp_rs_inv_tran_all + 1
    rs_inv_tran_all.UpdateBatch
    rs_stk_open_srl.MoveNext
Next

Call open_database
Call open_rs_inv_tran_sal
Call open_rs_inv_tran_all
For temp_counter = 1 To rs_inv_tran_sal.RecordCount
    rs_inv_tran_all.AddNew
    rs_inv_tran_all!stk_invt_trn_id = temp_rs_inv_tran_all
    rs_inv_tran_all!stk_invt_trn_vcno = rs_inv_tran_sal!stk_invt_trn_vcno
    rs_inv_tran_all!stk_invt_trn_seno = rs_inv_tran_sal!stk_invt_trn_seno
    rs_inv_tran_all!stk_invt_trn_date = rs_inv_tran_sal!stk_invt_trn_date
    rs_inv_tran_all!stk_invt_trn_time = rs_inv_tran_sal!stk_invt_trn_time
    rs_inv_tran_all!stk_invt_trn_wday = rs_inv_tran_sal!stk_invt_trn_wday
    rs_inv_tran_all!stk_invt_trn_ldgr = rs_inv_tran_sal!stk_invt_trn_ldgr
    rs_inv_tran_all!stk_invt_trn_card = rs_inv_tran_sal!stk_invt_trn_card
    rs_inv_tran_all!stk_invt_trn_stno = rs_inv_tran_sal!stk_invt_trn_stno
    rs_inv_tran_all!stk_invt_trn_edno = rs_inv_tran_sal!stk_invt_trn_edno
    rs_inv_tran_all!stk_invt_trn_qnty = rs_inv_tran_sal!stk_invt_trn_qnty
    rs_inv_tran_all!stk_invt_trn_rate = rs_inv_tran_sal!stk_invt_trn_rate
    rs_inv_tran_all!stk_invt_trn_amnt = rs_inv_tran_sal!stk_invt_trn_amnt
    rs_inv_tran_all!stk_invt_trn_fval = rs_inv_tran_sal!stk_invt_trn_fval
    rs_inv_tran_all!stk_invt_trn_splr = rs_inv_tran_sal!stk_invt_trn_splr
    rs_inv_tran_all!stk_invt_trn_comp = rs_inv_tran_sal!stk_invt_trn_comp
    rs_inv_tran_all!stk_invt_trn_user = rs_inv_tran_sal!stk_invt_trn_user
    rs_inv_tran_all!stk_invt_trn_vchr = rs_inv_tran_sal!stk_invt_trn_vchr
    rs_inv_tran_all!stk_invt_trn_slmn = rs_inv_tran_sal!stk_invt_trn_slmn
    rs_inv_tran_all!stk_invt_trn_vtyp = rs_inv_tran_sal!stk_invt_trn_vtyp
    rs_inv_tran_all!stk_invt_trn_splr = rs_inv_tran_sal!stk_invt_trn_splr
    temp_rs_inv_tran_all = temp_rs_inv_tran_all + 1
    rs_inv_tran_all.UpdateBatch
    rs_inv_tran_sal.MoveNext
Next

Call open_database
Call open_rs_inv_tran_prs
Call open_rs_inv_tran_all
For temp_counter = 1 To rs_inv_tran_prs.RecordCount
    rs_inv_tran_all.AddNew
    rs_inv_tran_all!stk_invt_trn_id = temp_rs_inv_tran_all
    rs_inv_tran_all!stk_invt_trn_vcno = rs_inv_tran_prs!stk_invt_trn_vcno
    rs_inv_tran_all!stk_invt_trn_seno = rs_inv_tran_prs!stk_invt_trn_seno
    rs_inv_tran_all!stk_invt_trn_date = rs_inv_tran_prs!stk_invt_trn_date
    rs_inv_tran_all!stk_invt_trn_time = rs_inv_tran_prs!stk_invt_trn_time
    rs_inv_tran_all!stk_invt_trn_wday = rs_inv_tran_prs!stk_invt_trn_wday
    rs_inv_tran_all!stk_invt_trn_ldgr = rs_inv_tran_prs!stk_invt_trn_ldgr
    rs_inv_tran_all!stk_invt_trn_card = rs_inv_tran_prs!stk_invt_trn_card
    rs_inv_tran_all!stk_invt_trn_stno = rs_inv_tran_prs!stk_invt_trn_stno
    rs_inv_tran_all!stk_invt_trn_edno = rs_inv_tran_prs!stk_invt_trn_edno
    rs_inv_tran_all!stk_invt_trn_qnty = rs_inv_tran_prs!stk_invt_trn_qnty
    rs_inv_tran_all!stk_invt_trn_rate = rs_inv_tran_prs!stk_invt_trn_rate
    rs_inv_tran_all!stk_invt_trn_amnt = rs_inv_tran_prs!stk_invt_trn_amnt
    rs_inv_tran_all!stk_invt_trn_fval = rs_inv_tran_prs!stk_invt_trn_fval
    rs_inv_tran_all!stk_invt_trn_splr = rs_inv_tran_prs!stk_invt_trn_splr
    rs_inv_tran_all!stk_invt_trn_comp = rs_inv_tran_prs!stk_invt_trn_comp
    rs_inv_tran_all!stk_invt_trn_user = rs_inv_tran_prs!stk_invt_trn_user
    rs_inv_tran_all!stk_invt_trn_vchr = rs_inv_tran_prs!stk_invt_trn_vchr
    rs_inv_tran_all!stk_invt_trn_slmn = rs_inv_tran_prs!stk_invt_trn_slmn
    rs_inv_tran_all!stk_invt_trn_vtyp = rs_inv_tran_prs!stk_invt_trn_vtyp
    rs_inv_tran_all!stk_invt_trn_splr = rs_inv_tran_prs!stk_invt_trn_splr
    temp_rs_inv_tran_all = temp_rs_inv_tran_all + 1
    rs_inv_tran_all.UpdateBatch
rs_inv_tran_prs.MoveNext
Next

Call open_database
Call open_rs_inv_tran_srt
Call open_rs_inv_tran_all
For temp_counter = 1 To rs_inv_tran_srt.RecordCount
    rs_inv_tran_all.AddNew
    rs_inv_tran_all!stk_invt_trn_id = temp_rs_inv_tran_all
    rs_inv_tran_all!stk_invt_trn_vcno = rs_inv_tran_srt!stk_invt_trn_vcno
    rs_inv_tran_all!stk_invt_trn_seno = rs_inv_tran_srt!stk_invt_trn_seno
    rs_inv_tran_all!stk_invt_trn_date = rs_inv_tran_srt!stk_invt_trn_date
    rs_inv_tran_all!stk_invt_trn_time = rs_inv_tran_srt!stk_invt_trn_time
    rs_inv_tran_all!stk_invt_trn_wday = rs_inv_tran_srt!stk_invt_trn_wday
    rs_inv_tran_all!stk_invt_trn_ldgr = rs_inv_tran_srt!stk_invt_trn_ldgr
    rs_inv_tran_all!stk_invt_trn_card = rs_inv_tran_srt!stk_invt_trn_card
    rs_inv_tran_all!stk_invt_trn_stno = rs_inv_tran_srt!stk_invt_trn_stno
    rs_inv_tran_all!stk_invt_trn_edno = rs_inv_tran_srt!stk_invt_trn_edno
    rs_inv_tran_all!stk_invt_trn_qnty = rs_inv_tran_srt!stk_invt_trn_qnty
    rs_inv_tran_all!stk_invt_trn_rate = rs_inv_tran_srt!stk_invt_trn_rate
    rs_inv_tran_all!stk_invt_trn_amnt = rs_inv_tran_srt!stk_invt_trn_amnt
    rs_inv_tran_all!stk_invt_trn_fval = rs_inv_tran_srt!stk_invt_trn_fval
    rs_inv_tran_all!stk_invt_trn_splr = rs_inv_tran_srt!stk_invt_trn_splr
    rs_inv_tran_all!stk_invt_trn_comp = rs_inv_tran_srt!stk_invt_trn_comp
    rs_inv_tran_all!stk_invt_trn_user = rs_inv_tran_srt!stk_invt_trn_user
    rs_inv_tran_all!stk_invt_trn_vchr = rs_inv_tran_srt!stk_invt_trn_vchr
    rs_inv_tran_all!stk_invt_trn_slmn = rs_inv_tran_srt!stk_invt_trn_slmn
    rs_inv_tran_all!stk_invt_trn_vtyp = rs_inv_tran_srt!stk_invt_trn_vtyp
    rs_inv_tran_all!stk_invt_trn_splr = rs_inv_tran_srt!stk_invt_trn_splr
    temp_rs_inv_tran_all = temp_rs_inv_tran_all + 1
    rs_inv_tran_all.UpdateBatch
rs_inv_tran_srt.MoveNext
Next

Call open_database
Call open_rs_inv_tran_prt
Call open_rs_inv_tran_all
For temp_counter = 1 To rs_inv_tran_prt.RecordCount
    rs_inv_tran_all.AddNew
    rs_inv_tran_all!stk_invt_trn_id = temp_rs_inv_tran_all
    rs_inv_tran_all!stk_invt_trn_vcno = rs_inv_tran_prt!stk_invt_trn_vcno
    rs_inv_tran_all!stk_invt_trn_seno = rs_inv_tran_prt!stk_invt_trn_seno
    rs_inv_tran_all!stk_invt_trn_date = rs_inv_tran_prt!stk_invt_trn_date
    rs_inv_tran_all!stk_invt_trn_time = rs_inv_tran_prt!stk_invt_trn_time
    rs_inv_tran_all!stk_invt_trn_wday = rs_inv_tran_prt!stk_invt_trn_wday
    rs_inv_tran_all!stk_invt_trn_ldgr = rs_inv_tran_prt!stk_invt_trn_ldgr
    rs_inv_tran_all!stk_invt_trn_card = rs_inv_tran_prt!stk_invt_trn_card
    rs_inv_tran_all!stk_invt_trn_stno = rs_inv_tran_prt!stk_invt_trn_stno
    rs_inv_tran_all!stk_invt_trn_edno = rs_inv_tran_prt!stk_invt_trn_edno
    rs_inv_tran_all!stk_invt_trn_qnty = rs_inv_tran_prt!stk_invt_trn_qnty
    rs_inv_tran_all!stk_invt_trn_rate = rs_inv_tran_prt!stk_invt_trn_rate
    rs_inv_tran_all!stk_invt_trn_amnt = rs_inv_tran_prt!stk_invt_trn_amnt
    rs_inv_tran_all!stk_invt_trn_fval = rs_inv_tran_prt!stk_invt_trn_fval
    rs_inv_tran_all!stk_invt_trn_splr = rs_inv_tran_prt!stk_invt_trn_splr
    rs_inv_tran_all!stk_invt_trn_comp = rs_inv_tran_prt!stk_invt_trn_comp
    rs_inv_tran_all!stk_invt_trn_user = rs_inv_tran_prt!stk_invt_trn_user
    rs_inv_tran_all!stk_invt_trn_vchr = rs_inv_tran_prt!stk_invt_trn_vchr
    rs_inv_tran_all!stk_invt_trn_slmn = rs_inv_tran_prt!stk_invt_trn_slmn
    rs_inv_tran_all!stk_invt_trn_vtyp = rs_inv_tran_prt!stk_invt_trn_vtyp
    rs_inv_tran_all!stk_invt_trn_splr = rs_inv_tran_prt!stk_invt_trn_splr
    temp_rs_inv_tran_all = temp_rs_inv_tran_all + 1
    rs_inv_tran_all.UpdateBatch
    rs_inv_tran_prt.MoveNext
Next
End Sub

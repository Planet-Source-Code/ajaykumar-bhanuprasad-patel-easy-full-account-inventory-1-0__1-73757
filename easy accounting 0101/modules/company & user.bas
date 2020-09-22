Attribute VB_Name = "Module2"
Option Explicit

Public position As Integer
Public lastrecord As Integer

Public intCnt As Integer
Public outrec As PhoneRec

Type PhoneRec
    co_id As Integer
    co_name As String * 20
    co_folder As Integer
End Type


Public my_path As String
Public back_up_path As String

Public selected_company As String
Public selected_path As String
Public selected_backup_path As String
Public selected_group As String
Public selected_primary_group As String


Public co_name As String
Public selected_companies_add1 As String
Public selected_companies_add2 As String
Public selected_companies_pincode As String
Public selected_companies_city As String
Public selected_companies_country As String
Public selected_companies_email As String
Public selected_companies_telephone As String
Public selected_companies_acconting_style As Integer
Public selected_companies_working_style As Integer
Public selected_companies_backup_path As String
Public selected_companies_tax_no As String
Public selected_companies_starting_f_date As Date
Public selected_companies_ending_f_date As Date

Public this_year_starting_date As Date
Public this_year_ending_date As Date
Public selected_companies_owner As String
Public selected_companies_currency_sym As String
Public selected_companies_currency_decimal As Integer
Public selected_companies_sequrity_code As Integer



Public selected_procedure As String

Public selected_stock_group As String

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


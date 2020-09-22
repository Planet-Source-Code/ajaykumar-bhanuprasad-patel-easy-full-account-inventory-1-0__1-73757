VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "key_code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function check() As Integer
 'To detect an existing file, use the function below:
' Add this code to the appropriate event:
Dim success%
success% = FileExists%(App.Path & "\activation.txt") 'A full path and filename
' FileExists% returns True if file exists
If success% = True Then
    Dim FieldContent
    Dim reg_code
    Open App.Path & "\activation.txt" For Input As #1
    Input #1, FieldContent
    reg_code = FieldContent
    Close #1
    
    If reg_code = "aum" Then
        check = 1
        MsgBox "you have activated your product...."
    Else
        check = 0
    End If
Else
   check = 0
End If

End Function
Function FileExists%(fname$)
On Local Error Resume Next

Dim ff%
        ff% = FreeFile
        Open fname$ For Input As ff%

        If Err Then
        FileExists% = False
        Else
        FileExists% = True
        End If

        Close ff%

End Function

Attribute VB_Name = "user_000"

Option Explicit
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public user_name As String
Public xyz
Public position_usr As Integer
Public lastrecord_usr As Integer
Public outrec_usr As PhoneRec
    Type PhoneRec
        user_id As Integer
        uname As String * 10
        upass As String * 10
    End Type

Public usr_password_code As Integer
Public selected_user As String
Public selected_user_password As String

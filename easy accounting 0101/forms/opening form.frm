VERSION 5.00
Begin VB.Form A_Opening_win 
   BorderStyle     =   0  'None
   Caption         =   "001_opening Window"
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "opening form.frx":0000
   ScaleHeight     =   4755
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   600
   End
   Begin VB.Image Image1 
      Height          =   4755
      Left            =   0
      Picture         =   "opening form.frx":35C47
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6195
   End
End
Attribute VB_Name = "A_Opening_win"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'set screen size
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

'set path null
If selected_path = "" Or selected_path = Null Then

'this is a direct path to co
'selected_path = App.Path & "\co.mdb;"
'selected_backup_path = App.Path & "\bk_up_co.mdb;"

'this is a direct path to 1st created company.
'selected_path = App.Path & "\data\1000\co.mdb"
'selected_backup_path = App.Path & "\data\1000\backup\co.mdb"

End If

'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'this is a code for sizing===================================
End Sub
Private Sub select_co_path()

Dim FieldContent
Open App.Path & "\data\main.txt" For Input As #1
        Input #1, FieldContent
        selected_path = FieldContent
        Input #1, FieldContent
        selected_backup_path = FieldContent
Close #1

Dim success%
Dim search_file As String
search_file = sysinfo_c_windir & "\Sw.txt"
success% = reg.FileExists%(selected_path)
If success% = True Then
Else
    Me.Visible = False
    MsgBox "There is error loading your programme....., contact ajay patel(M) 9998175413 "
    End
End If

Open App.Path & "\data\main.txt" For Random As #1
Do While Not EOF(1)
Get #1, , outrec
selected_path = App.Path & "\data\" & outrec.co_folder
Loop
Close #1

Open App.Path & "\main.txt" For Random As #1
On Error GoTo errRtn
        Get #1, , outrec
        selected_path = outrec.co_folder
Close #1

errRtn:
    Resume

End Sub


Private Sub Timer1_Timer()
Dim k As Integer
'this is a direct tunnel to any co...,
'k = 0
'If k = 0 Then
'Call open_database
'Call make_trail_balance_summary
'End If

For k = 1 To 10000
    If k > 9000 Then
            Unload Me
            'frm_usr.Show
            Set newfrm = B_co_menu
            newfrm.Show
    End If
Next k
End Sub

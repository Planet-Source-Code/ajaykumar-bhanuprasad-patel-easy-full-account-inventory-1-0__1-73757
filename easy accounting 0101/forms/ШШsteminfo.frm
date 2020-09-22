VERSION 5.00
Begin VB.Form Form000002 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "system information"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10095
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLog1 
         Caption         =   "Make a log file of ""Computer Profile"""
      End
      Begin VB.Menu mnuLog2 
         Caption         =   "Make a log file of ""System Info"""
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuMinimizeFoxInfo 
         Caption         =   "Minimize Fox-Info"
      End
      Begin VB.Menu mnuAlwaysOnTop 
         Caption         =   "Always On Top"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form000002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private processorsinfo
Private sys_drv As String
Private Sub Form_Load()
Me.Visible = False
Call get_system_data
sysinfo_path = sysinfo_c_windir & "\systeminfo.txt"
Call msgme
End Sub
Public Sub msgme()
MsgBox sysinfo_c_name
MsgBox sysinfo_c_user
MsgBox sysinfo_c_windir
MsgBox sysinfo_c_wintempdir
MsgBox sysinfo_c_winsysdir
MsgBox sysinfo_c_sysdrive
MsgBox sysinfo_c_osname
MsgBox sysinfo_c_processor
MsgBox sysinfo_c_totalram
MsgBox sysinfo_c_noprocessor
MsgBox sysinfo_c_sysdrv_serial_no
End Sub
Public Sub get_system_data()
sysinfo_c_name = machinename()                          '"Computer Name"
sysinfo_c_user = Environ$("username")                   '"User name"
sysinfo_c_windir = Environ$("windir")                   '"Windows Dir"
sysinfo_c_wintempdir = Environ$("temp")                 '"Windows temp dir"
sysinfo_c_winsysdir = SystemDir                         '"win system dir"
sysinfo_c_sysdrive = Environ$("systemdrive")            '"System Drive"
sysinfo_c_osname = Environ$("os")                       'os version
    processorsinfo = ReadKey("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0\Processornamestring")
sysinfo_c_processor = processorsinfo                    'processor information
    Call GlobalMemoryStatus(memInfo)
sysinfo_c_totalram = memInfo.dwTotalPhys / 1024 & " KB" 'total ram
sysinfo_c_noprocessor = Environ$("Number_Of_Processors")  'number of processor
'Call computerbiosinfo
sysinfo_c_biosver = bios_name                           'bios version
sysinfo_c_biosman = bios_manufacturer
        sysdrv = Environ$("systemdrive")
sysinfo_c_sysdrv_serial_no = GetSerialNumber(sys_drv)   'system drive serial no

End Sub

Public Sub open_and_read_system_data()

Dim FieldContent
Open sysinfo_path For Input As #1
        Input #1, FieldContent
         old_sysinfo_c_name = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_user = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_windir = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_wintempdir = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_winsysdir = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_sysdrive = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_osname = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_processor = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_totalram = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_noprocessor = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_biosver = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_biosman = FieldContent
        Input #1, FieldContent
         old_sysinfo_c_sysdrv_serial_no = FieldContent
Close #1

End Sub

Public Sub compare_data()

Dim score As Integer
score = 0
If Mid(old_sysinfo_c_name, 1, 4) = Mid(sysinfo_c_name, 1, 4) Then score = score + 1
If old_sysinfo_c_user = sysinfo_c_user Then score = score + 1
If old_sysinfo_c_windir = sysinfo_c_windir Then score = score + 1
If old_sysinfo_c_wintempdir = sysinfo_c_wintempdir Then score = score + 1
If Mid(old_sysinfo_c_winsysdir, 1, 10) = Mid(sysinfo_c_winsysdir, 1, 10) Then score = score + 1
If old_sysinfo_c_sysdrive = sysinfo_c_sysdrive Then score = score + 1
If old_sysinfo_c_osname = sysinfo_c_osname Then score = score + 1
If old_sysinfo_c_processor = sysinfo_c_processor Then score = score + 1
If old_sysinfo_c_totalram = sysinfo_c_totalram Then score = score + 1
If old_sysinfo_c_noprocessor = sysinfo_c_noprocessor Then score = score + 1
If old_sysinfo_c_biosver = sysinfo_c_biosver Then score = score + 1
If old_sysinfo_c_biosman = sysinfo_c_biosman Then score = score + 1
If old_sysinfo_c_sysdrv_serial_no = sysinfo_c_sysdrv_serial_no Then score = score + 1

If score < 11 Then
    MsgBox "There is error loading your programme....., contact ajay patel(M) 9998175413 "
    Unload Me
Else
    Kill sysinfo_path
    Call write_system_data
    Unload Me
End If

End Sub
Public Sub write_system_data()
Open sysinfo_path For Append As #1
    Write #1, sysinfo_c_name
    Write #1, sysinfo_c_user
    Write #1, sysinfo_c_windir
    Write #1, sysinfo_c_wintempdir
    Write #1, sysinfo_c_winsysdir
    Write #1, sysinfo_c_sysdrive
    Write #1, sysinfo_c_osname
    Write #1, sysinfo_c_processor
    Write #1, sysinfo_c_totalram
    Write #1, sysinfo_c_noprocessor
    Write #1, sysinfo_c_biosver
    Write #1, sysinfo_c_biosman
    Write #1, sysinfo_c_sysdrv_serial_no
Close #1
End Sub

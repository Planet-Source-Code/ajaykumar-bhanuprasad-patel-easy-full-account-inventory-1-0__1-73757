Attribute VB_Name = "Module5"
Option Explicit
Public newfrm As Form
Public activation As Integer
Public processorsinfo
Public sysdrv As String
Public my_customer_key As String
Public my_customer_no As Double
Public set_time1 As Integer
Public set_time2 As Integer
Public my_serial_key As Double
Public act_time1 As Integer
Public act_time2 As Integer
Public my_activation_key As Double

Public x1 As String
Public x2 As String
Public x3 As Integer
Public x4 As Integer
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long


Public sysinfo_c_name _
       , sysinfo_c_user _
       , sysinfo_c_windir _
       , sysinfo_c_wintempdir _
       , sysinfo_c_winsysdir _
       , sysinfo_c_sysdrive _
       , sysinfo_c_osname _
       , sysinfo_c_processor _
       , sysinfo_c_totalram _
       , sysinfo_c_noprocessor _
       , sysinfo_c_biosver _
       , sysinfo_c_biosman _
       , sysinfo_c_sysdrv_serial_no _
As String
    
Public old_sysinfo_c_name _
       , old_sysinfo_c_user _
       , old_sysinfo_c_windir _
       , old_sysinfo_c_wintempdir _
       , old_sysinfo_c_winsysdir _
       , old_sysinfo_c_sysdrive _
       , old_sysinfo_c_osname _
       , old_sysinfo_c_processor _
       , old_sysinfo_c_totalram _
       , old_sysinfo_c_noprocessor _
       , old_sysinfo_c_biosver _
       , old_sysinfo_c_biosman _
       , old_sysinfo_c_sysdrv_serial_no _
As String

Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public memInfo As MEMORYSTATUS
Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

'Public BiosSet As SWbemObjectSet
'Public bios As SWbemObject
Public sysinfo_path As String
Public cnt As Long
Public msg As String
Public bios_name As String
Public bios_manufacturer As String



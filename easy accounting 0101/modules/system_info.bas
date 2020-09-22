Attribute VB_Name = "Module1"
Public Function SystemDir() As String
    Dim result
    Dim SystemDirectory As String
    SystemDirectory = Space(144)
    result = GetSystemDirectory(SystemDirectory, 144)
    If result = 0 Then
        MsgBox "Cannot Get the Windows System Directory", vbCritical, "Warning"
    Else
        SystemDir = Trim(SystemDirectory)
    End If
End Function
Public Function GetSerialNumber(strDrive As String) As Long
    Dim SerialNum As Long
    Dim Res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    Res = GetVolumeInformation(strDrive, Temp1, _
    Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialNumber = SerialNum
End Function

Public Function GetCountry() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGCOUNTRY, Buffer, 99)
   GetCountry = LPSTRToVBString(Buffer)
End Function
Public Function machinename() As String
    Dim NameSize As Long
    Dim X As Long
    
    machinename = Space$(16)
    NameSize = Len(machinename)
    X = GetComputerName(machinename, NameSize)
    
End Function
Public Function LPSTRToVBString$(ByVal s$)
   Dim nullpos&
   nullpos& = InStr(s$, Chr$(0))
   If nullpos > 0 Then
      LPSTRToVBString = Left$(s$, nullpos - 1)
   Else
      LPSTRToVBString = ""
   End If
End Function
Public Function ReadKey(Value As String) As String
Dim b As Object
Dim R
On Error Resume Next
Set b = CreateObject("wscript.shell")
R = b.regread(Value)
ReadKey = R
End Function

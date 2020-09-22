Attribute VB_Name = "screen_size"
'Option Explicit
Public ScaleFactorX As Single, ScaleFactorY As Single
Public MyForm As FRMSIZE
Public DesignX As Integer
Public DesignY As Integer

Public Xtwips As Integer, Ytwips As Integer
Public Xpixels As Integer, Ypixels As Integer

Type FRMSIZE
   Height As Long
   Width As Long
End Type

Public RePosForm As Boolean
Public DoResize As Boolean
Public Sub set_screen_resolution()

  ' Scaling factors
    ' Size of Form in Pixels at design resolution
    DesignX = 800
    DesignY = 600
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
    ' Set up the screen values
    Xtwips = Screen.TwipsPerPixelX
    Ytwips = Screen.TwipsPerPixelY
    Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
    Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution
    ' Determine scaling factors
    ScaleFactorX = (Xpixels / DesignX)
    ScaleFactorY = (Ypixels / DesignY)
    ScaleMode = 1  ' twips 'Exit Sub
    ' uncomment to see how Form1 looks without resizing
    
End Sub
Public Sub Resize_For_Resolution(ByVal SFX As Single, ByVal SFY As Single, MyForm As Form)
    Dim I As Integer
    Dim SFFont As Single

    SFFont = (SFX + SFY) / 2  ' average scale
    ' Size the Controls for the new resolution
    On Error Resume Next  ' for read-only or nonexistent properties
    With MyForm
      For I = 0 To .Count - 1
       If TypeOf .Controls(I) Is ComboBox Then   ' cannot change Height
         .Controls(I).Left = .Controls(I).Left * SFX
         .Controls(I).Top = .Controls(I).Top * SFY
         .Controls(I).Width = .Controls(I).Width * SFX
       Else
         .Controls(I).Move .Controls(I).Left * SFX, _
          .Controls(I).Top * SFY, _
          .Controls(I).Width * SFX, _
          .Controls(I).Height * SFY
       End If
         ' Be sure to resize and reposition before changing the FontSize
         .Controls(I).FontSize = .Controls(I).FontSize * SFFont
      Next I
      If RePosForm Then
        ' Now size the Form
        .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
      End If
    End With
End Sub



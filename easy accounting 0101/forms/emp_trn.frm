VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form emp_tran 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   8430
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "emp_trn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Close1 
      Caption         =   "Close"
      Height          =   495
      Left            =   4680
      TabIndex        =   24
      Top             =   5640
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3960
      TabIndex        =   0
      Text            =   "Select Name"
      Top             =   720
      Width           =   6375
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   22
      Top             =   8160
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9498
            MinWidth        =   7408
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "(C) Masino Sinaga (masino_sinaga@yahoo.com)"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "It's up to you..."
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "23/01/2011"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Date today"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1464
            MinWidth        =   1464
            TextSave        =   "23:01"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Time right now"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_tran_dtl_name"
      Height          =   525
      Index           =   0
      Left            =   3960
      TabIndex        =   8
      Top             =   720
      Width           =   6375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_tran_dtl_date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   1
      Left            =   3960
      TabIndex        =   9
      Top             =   1320
      Width           =   6375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_tran_dtl_intm"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   4
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   2
      Left            =   3960
      TabIndex        =   3
      Top             =   2160
      Width           =   6375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_tran_dtl_outm"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   4
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   3
      Left            =   3960
      TabIndex        =   10
      Top             =   2160
      Width           =   6375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_tran_dtl_aupr"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   4
      Left            =   3960
      TabIndex        =   4
      Top             =   2880
      Width           =   6375
   End
   Begin VB.PictureBox picButtons 
      Height          =   1095
      Left            =   1800
      ScaleHeight     =   1035
      ScaleWidth      =   7395
      TabIndex        =   11
      Top             =   4320
      Width           =   7455
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&New"
         Height          =   585
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Sa&ve"
         Height          =   585
         Left            =   3720
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   585
         Left            =   5400
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdBookmark 
         Caption         =   "&Bookmark"
         Height          =   585
         Left            =   5400
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Exit to Main Menu "
         Height          =   585
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   585
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   585
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   585
         Left            =   3720
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   23
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Employee.... :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date..........:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Time....................:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   5
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Time......................:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   720
      TabIndex        =   6
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Authorized_Person......:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   720
      TabIndex        =   7
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label lblAngka 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   3360
      Width           =   2655
   End
End
Attribute VB_Name = "emp_tran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public int_i As Integer
Public display_tran  As Integer
Public WithEvents rsstrFindData As Recordset
Attribute rsstrFindData.VB_VarHelpID = -1

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim blnCancel As Boolean
Dim NumData As Integer
Dim intRecord As Integer
Dim intField As Integer

Private Sub Form_Load()
'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

'this is a code for sizing===================================
Combo1.Enabled = False
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
txtFields(4).Text = selected_user

If selected_procedure = "Employee In Entry" Then
    Call display_off
    Call in_transaction_entry 'Call click_add
    txtFields(1).Text = Date
    txtFields(3).Text = Time
    Close1.Visible = False
ElseIf selected_procedure = "Employee Out Entry" Then
    Call display_off '   Label1.Caption = "Click New to creat a transaction and save after...,"
    Call out_tranaction_entry 'Call click_add
    txtFields(1).Text = Date
    txtFields(2).Text = Time
    Close1.Visible = False
End If

Call open_database
Call open_rs_emp_main_dtl

Do Until rs_emp_main_dtl.EOF
    Combo1.AddItem rs_emp_main_dtl!emp_main_dtl_name
    rs_emp_main_dtl.MoveNext
Loop

cmdAdd.Enabled = True
cmdClose.Enabled = True
cmdUpdate.Enabled = False
cmdCancel.Enabled = False
End Sub

Private Sub Combo1_Change()
    txtFields(0).Text = Combo1.Text
End Sub
Private Sub Combo1_Click()
    txtFields(0).Text = Combo1.Text
End Sub
Private Sub edit_on()
Dim z As Integer
For z = 0 To 4
lblLabels(z).Visible = True
txtFields(z).Visible = True
Combo1.Visible = True
Next
Combo1.Text = txtFields(0).Text
Call display_off
cmdClose.Enabled = True
Label1.Caption = " You have to Edit Selected Entry and Click save button"
End Sub
Private Sub Close1_Click()
Unload Me
End Sub
Private Sub Command1_Click()
Me.Visible = False
MDIForm1.WindowState = 2
End Sub
Public Sub emp_tran_sub_procedure()
Call open_database
Call open_rs_emp_tran_dtl
   
Dim oText As TextBox 'Bind textbox to recordset
For Each oText In Me.txtFields
Set oText.DataSource = rs_emp_tran_dtl
Next 'Bind recordset to datagrid
  
  
  If rs_emp_tran_dtl.RecordCount < 1 Then
     MsgBox "Recordset is empty. Please click Add button to add new record!", vbExclamation, "Empty Recordset"
     Exit Sub
  End If
  
  LockTheForm 'Lock textbox, combobox, and optionbutton'Except Datagrid....
  SetButtons True
  Exit Sub
Message:
  MsgBox Err.Number & " - " & Err.Description
  End
End Sub
Private Sub Message(strMessage As String)
  StatusBar1.Panels(1).Text = strMessage
End Sub
Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
  Response = -1
  'DataError = -1
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If cmdUpdate.Enabled = True And cmdCancel.Enabled = True Then
     MsgBox "You have to save or cancel the changes " & vbCrLf & "that you have just made before quit!", vbExclamation, "Warning"
     cmdUpdate.SetFocus
     Cancel = -1
     Exit Sub
  End If
  If Not rs_emp_tran_dtl Is Nothing Then Set rs_emp_tran_dtl = Nothing     'Clear memory from recordset
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Screen.MousePointer = vbDefault 'Mouse pointer back to normal
End Sub
'Display the selected record in datagrid
'Public Sub rs_emp_tran_dtl_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  NumData = rs_emp_tran_dtl.AbsolutePosition
'  lblStatus.Caption = "Record number " & CStr(NumData) & " from " & rs_emp_tran_dtl.RecordCount
'  CheckNavigation
'End Sub

Private Sub click_add()
  On Error GoTo AddErr
  With rs_emp_tran_dtl
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    UnlockTheForm
    .AddNew
    
    mbAddNewFlag = True
    SetButtons False
  End With
  
  On Error Resume Next
  txtFields(0).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdAdd_Click()
Combo1.Enabled = True
Call emp_tran_sub_procedure
Call click_add
If display_tran = 3 Then
    txtFields(2).Text = Time
ElseIf display_tran = 0 Then
    txtFields(3).Text = Time
End If
txtFields(1).Text = Date
txtFields(4).Text = selected_user
End Sub
Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Add new record.")
End Sub
Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Delete the selected record.")
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next
  LockTheForm
  If blnCancel = True Then
     Exit Sub
  End If
  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  rs_emp_tran_dtl.CancelUpdate
  If mvBookMark > 0 Then
    rs_emp_tran_dtl.Bookmark = mvBookMark
  Else
    rs_emp_tran_dtl.MoveFirst
  End If
  LockTheForm    'Lock textbox
End Sub
Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Cancel the change or new record that have not been saved.")
End Sub
Private Sub cmdUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Save the change or new record.")
End Sub
Private Sub cmdUpdate_Click()
cmdAdd.Enabled = True
cmdClose.Enabled = True
cmdUpdate.Enabled = False
cmdCancel.Enabled = False
Combo1.Enabled = False
If display_tran = 3 Then ' its procedure for employee in transaction

Call open_database
Call open_rs_emp_tran_dtl
        
        Do Until rs_emp_tran_dtl.EOF
            If rs_emp_tran_dtl!emp_tran_dtl_name = Combo1.Text And rs_emp_tran_dtl!emp_tran_dtl_date = txtFields(1).Text Then
                MsgBox "you have already entered the entry....,"
                Exit Sub
            End If
            rs_emp_tran_dtl.MoveNext
        Loop
        
            On Error GoTo UpdateErr
          
          For int_i = 0 To 4
            If I <> 3 And txtFields(I).Text = "" Then
               MsgBox "You have to fill in all textbox!", vbExclamation, "Validation"
               txtFields(I).SetFocus
               Exit Sub
             End If
          Next int_i
        rs_emp_tran_dtl.UpdateBatch adAffectAll

    Exit Sub
ElseIf display_tran = 0 Then ' its procedure for employee out transaction

Call open_database
Call open_rs_emp_tran_dtl
        Do Until rs_emp_tran_dtl.EOF
         If rs_emp_tran_dtl!emp_tran_dtl_name = Combo1.Text And rs_emp_tran_dtl!emp_tran_dtl_date = txtFields(1).Text Then
            If rs_emp_tran_dtl!emp_tran_dtl_outm <> 0 Then
                MsgBox "You have already entered time...!!!"
                Exit Sub
            End If
            rs_emp_tran_dtl!emp_tran_dtl_outm = txtFields(3).Text
            rs_emp_tran_dtl.Update
              On Error GoTo UpdateErr
                      On Error Resume Next
              Exit Sub
        End If
    rs_emp_tran_dtl.MoveNext
    Loop
            Call click_add
            txtFields(1).Text = Date
            txtFields(2).Text = Time
            MsgBox "You have not entered In Entry of Employee...,"
    Exit Sub
End If
UpdateErr:
  MsgBox Err.Number & " - " & Err.Description, vbCritical, "Error Occured"
End Sub
Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Quit from this program now.")
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub SetButtons(bVal As Boolean)

  cmdAdd.Enabled = bVal
  cmdUpdate.Enabled = Not bVal
  cmdCancel.Enabled = Not bVal
  cmdEdit.Enabled = bVal
  cmdDelete.Enabled = bVal
  cmdRefresh.Enabled = bVal
  cmdClose.Enabled = bVal

  End Sub


Private Sub picButtons_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub
Private Sub LockTheForm()
  For int_i = 0 To 4
    txtFields(I).Locked = True
  Next int_i
  
End Sub
'Unlock textbox in order that we can edit data
Sub UnlockTheForm()
  For int_i = 0 To 4
    txtFields(I).Locked = False
  Next int_i
  
End Sub
Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Find record (find first and find next).")
End Sub
Private Sub cmdFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Filter recordset.")
End Sub
Private Sub display_off()
    cmdAdd.Visible = True
    cmdUpdate.Visible = True
    cmdCancel.Visible = True
    cmdEdit.Visible = False
    cmdDelete.Visible = False
    cmdRefresh.Visible = False
    
    cmdClose.Visible = True
    Label1.Caption = " Employee Daily Transaction Entry Form...,"
    Dim z As Integer
    For z = 0 To 4
    lblLabels(z).Visible = True
    txtFields(z).Visible = True
    Next
    picButtons.Height = 1200
End Sub
Private Sub display_on()
    cmdAdd.Visible = False
    cmdUpdate.Visible = False
    cmdCancel.Visible = False
    cmdEdit.Visible = True
    cmdDelete.Visible = True
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    Label1.Caption = " Employee Daily Transaction...,"
    
    Dim z As Integer
    For z = 0 To 4
    lblLabels(z).Visible = False
    txtFields(z).Visible = False
    Combo1.Visible = False
    Next
End Sub
Public Sub in_transaction_entry()
txtFields(4).Text = selected_user
lblLabels(3).Visible = False
txtFields(3).Visible = False
End Sub
Public Sub out_tranaction_entry()
txtFields(4).Text = selected_user
lblLabels(2).Visible = False
txtFields(2).Visible = False
End Sub

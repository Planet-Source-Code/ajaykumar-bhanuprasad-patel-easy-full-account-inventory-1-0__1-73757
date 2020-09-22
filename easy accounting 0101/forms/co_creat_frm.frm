VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form BA_co_creat_frm 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comapany Creation From"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8535
   Icon            =   "co_creat_frm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Currency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1935
      Left            =   240
      TabIndex        =   19
      Top             =   5400
      Width           =   7935
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   41
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   40
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text14 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   34
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   4680
         TabIndex        =   33
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   4680
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Select your regular backup path...,"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   39
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "How many Decimal place?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   22
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "What is a Curency symbol?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "You want to use Security Contry?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit without Save"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save and Exit"
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Accounting Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   7935
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         TabIndex        =   38
         Text            =   "Select"
         Top             =   2160
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   5520
         TabIndex        =   36
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109969409
         CurrentDate     =   40114
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3240
         TabIndex        =   35
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109969409
         CurrentDate     =   40114
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   3240
         TabIndex        =   31
         Top             =   1680
         Width           =   3735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3240
         TabIndex        =   30
         Text            =   "Select"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   3240
         TabIndex        =   29
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Comapny"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   37
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   15
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Accounting Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   14
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Financial year from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   5760
         TabIndex        =   28
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   2040
         TabIndex        =   27
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   5760
         TabIndex        =   26
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2040
         TabIndex        =   25
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   5760
         TabIndex        =   24
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   840
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Pincode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4725
         TabIndex        =   17
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Contry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Company Creation Form."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "BA_co_creat_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Click()
'MsgBox Combo2.ListIndex
End Sub

Private Sub Command1_Click() 'when save & exit button
Call write_company_data
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text12.Text = "" Then
    MsgBox "Sorry...!!! You are not filled information properly...!!!"
    Exit Sub
End If
Me.Enabled = False
path_sel.Show
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Activate()

Text1.TabIndex = 1
Text2.TabIndex = 2
Text3.TabIndex = 3
Text4.TabIndex = 4
Text5.TabIndex = 5
Text6.TabIndex = 6
Text7.TabIndex = 7
Text8.TabIndex = 8
Text9.TabIndex = 9
DTPicker1.TabIndex = 10
DTPicker2.TabIndex = 11
Combo1.TabIndex = 12
Text12.TabIndex = 13
Combo3.TabIndex = 14
Combo2.TabIndex = 15
Text13.TabIndex = 16
Text14.TabIndex = 17
Text15.TabIndex = 18
Command2.TabIndex = 19
Command1.TabIndex = 20
Command3.TabIndex = 21


    If back_up_path = "" Then
        back_up_path = App.Path & "\data\back_up\"
        Text15.Text = back_up_path
    Else
        Text15.Text = back_up_path
    End If
End Sub

Private Sub Form_Load()

'this is a code for sizing===================================
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
Call set_screen_resolution
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'this is a code for sizing===================================

    'Text15.Text = App.Path & "\data\folder no\back_up\co.mdb;"
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Combo1.AddItem ("Accounting olny")
    Combo1.AddItem ("Inventory olny")
    Combo1.AddItem ("Accounting and Inventory")
    
    Combo2.AddItem ("Yes")
    Combo2.AddItem ("No")
    
    Combo3.AddItem ("Individual")
    Combo3.AddItem ("partnership firm")
    Combo3.AddItem ("Limited company")

End Sub
Private Sub Form_Unload(Cancel As Integer)
'    MDIForm1.Enabled = True
End Sub

Public Sub write_company_data()
'=====================================================
'STEP 1 : put condition to check the syntex of the entered text
'STEP 2 :open files and read the last record of the file main.txt
'STEP 3 :save the co name in main.txt
'STEP 4 :creat folder named & copy common files in to such folder
'STEP 5 :save the companies detail in created folder co_main.mdb
'STEP 6 :check the user control detail & save the record in user table
'STEP 7 :save the detail of comapny at last position of main.txt
'=====================================================
'================================================================
'STEP 1 : put condition to check the syntex of the entered text
'================================================================

If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text12.Text = "" Then
    MsgBox "Sorry...!!! You are not filled information properly...!!!"
    Exit Sub
End If

'================================================================
'STEP 2 :open files and read the last record of the file main.txt
'================================================================

Open App.Path & "\data\main.txt" For Random As #1
On Error GoTo errRtn
    Do While Not EOF(1)
        Get #1, , outrec
    Loop
lastrecord = Seek(1) - 1
lastrecord = lastrecord + 1
Close #1

errRtn:
    Resume Next

'text1.Text to text12
'not a 10 & 11

'=====================================================
'STEP 3 :save the co name in main.txt
'=====================================================
'position

If lastrecord <= 0 Or lastrecord = Null Then lastrecord = 1
Open App.Path & "\data\main.txt" For Random As #1
On Error GoTo errRtn
    outrec.co_id = lastrecord
    outrec.co_name = Text1.Text
    outrec.co_folder = lastrecord * 1000
Put #1, lastrecord, outrec
Close #1

'errRtn:
'    Resume Next
'=====================================================
'STEP 4 :creat folder named & copy common files in to such folder
'=====================================================
MkDir App.Path & "\data\" & lastrecord * 1000
MkDir App.Path & "\data\back_up"
FileCopy App.Path & "\data\main\co.mdb", App.Path & "\data\" & lastrecord * 1000 & "\co.mdb"

'MkDir App.Path & "\data"
'MkDir App.Path & "\data\back_up"
'FileCopy App.Path & "\co.mdb", App.Path & "\data\co.mdb"

'=====================================================
'STEP 5 :save the companies detail in created folder co_main.mdb
'=====================================================
selected_path = App.Path & "\data\" & lastrecord * 1000 & "\co.mdb"

Call open_database
Call open_rs_co_main_dtl

        rs_co_main_dtl.AddNew
        rs_co_main_dtl!co_main_dtl_name = Text1.Text
        rs_co_main_dtl!co_main_dtl_add1 = Text2.Text
        rs_co_main_dtl!co_main_dtl_add2 = Text3.Text
        rs_co_main_dtl!co_main_dtl_pncd = Text4.Text
        rs_co_main_dtl!co_main_dtl_city = Text5.Text
        rs_co_main_dtl!co_main_dtl_cntr = Text6.Text
        rs_co_main_dtl!co_main_dtl_emal = Text8.Text
        rs_co_main_dtl!co_main_dtl_tlpn = Text7.Text
        rs_co_main_dtl!co_main_dtl_acst = Combo1.ListIndex
        rs_co_main_dtl!co_main_dtl_wrsl = Combo3.ListIndex
            Text15.Text = back_up_path
        rs_co_main_dtl!co_main_dtl_bkup = Text15.Text
        rs_co_main_dtl!co_main_dtl_txno = Text9.Text
        rs_co_main_dtl!co_main_dtl_fstr = DTPicker1.Value
        rs_co_main_dtl!co_main_dtl_fend = DTPicker2.Value
        rs_co_main_dtl!co_main_dtl_ownr = Text12.Text
        rs_co_main_dtl!co_main_dtl_sqst = Combo2.ListIndex
        rs_co_main_dtl!co_main_dtl_crsy = Text13.Text
        rs_co_main_dtl!co_main_dtl_decm = Text14.Text
        
        rs_co_main_dtl.UpdateBatch adAffectAllChapters

rs_co_main_dtl.Close
If db_co.State = 1 Then db_co.Close
'MDIForm1.Enabled = True
Close All
'frm_usr.Enabled = True
'frm_usr.Show

Set newfrm = B_co_menu
newfrm.Show
Unload Me
'=====================================================
'STEP 6 :check the user control detail & save the record in user table
'=====================================================
'=====================================================
'STEP 7 :save the detail of comapny at last position of main.txt
'=====================================================
End Sub
Public Sub save_co_detail()
'MsgBox lastrecord * 1000
End Sub

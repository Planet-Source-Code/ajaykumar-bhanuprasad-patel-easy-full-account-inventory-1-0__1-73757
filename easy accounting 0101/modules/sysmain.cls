VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mdi_dll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub close_all()
    Dim objApp
    Set objApp = CreateObject("Photoshop.Application")
    Dim dialogMode
    dialogMode = 3
    Do While objApp.Documents.Count
        objApp.ActiveDocument.Close
    Loop
End Sub


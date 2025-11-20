VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestClass 
   Caption         =   "Test class:"
   ClientHeight    =   2325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   OleObjectBlob   =   "frmTestClass.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents clsMSCh As clsMultiStateCheckBox
Attribute clsMSCh.VB_VarHelpID = -1

Private Sub clsMSCh_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
    lbValue.Caption = vbNewLine & vbTab & "Item: " & Item & vbNewLine & vbTab & "Code Icon: " & CodeIcon
End Sub

Private Sub cmbSetValue_Change()
    clsMSCh.Item = cmbSetValue.Value
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With
    With cmbSetValue
        .AddItem 0
        .AddItem 1
        .AddItem 2
        .AddItem 3
    End With

    Set clsMSCh = New clsMultiStateCheckBox
    Call clsMSCh.Initialize(Label1, 0, Array(59193, 59194, 59195, 59634, 59198, 60513), Array(rgbBlue, rgbGreen, rgbRed))
    lbVersion.Caption = clsMSCh.Version
End Sub
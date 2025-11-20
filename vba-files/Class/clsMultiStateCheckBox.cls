VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMultiStateCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================================================================================
' Class: clsMultiStateCheckBox
' Author: VBATools
' Version: 1.0.2
' Creation Date: 13.11.2025 15:41
' Last Update Date: 20.11.2025 11:47
' License: Apache License
' Description: A multi-state checkbox control for VBA UserForms with customizable icons and colors
'========================================================================================

Private WithEvents mLabelBtn As MSForms.Label
Attribute mLabelBtn.VB_VarHelpID = -1
Private WithEvents mParent As MSForms.UserForm
Attribute mParent.VB_VarHelpID = -1

Public Event Click(ByRef control As Object, ByRef Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)

Private mArrIcon    As Variant
Private mArrIconColor As Variant
Private mItem       As Byte
Private mCount      As Byte
Private mCountColor As Byte

' Property to get the current state text
Public Property Get StateText() As String
    If IsArray(mArrIcon) And mItem <= UBound(mArrIcon) Then
        StateText = VBA.ChrW$(mArrIcon(mItem))
    Else
        StateText = vbNullString
    End If
End Property

Public Property Get Version() As String
    Version = "Class: clsMultiStateCheckBox" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "Version: 1.0.2" & vbNewLine & _
            "Date of creation: 13.11.2025 15:41" & vbNewLine & _
            "Date of update: 20.11.2025 11:47" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Description: A multi-state checkbox control for VBA UserForms with customizable icons and colors"
End Property

Public Property Get Item() As Byte
    Item = mItem
End Property

Public Property Let Item(ByRef Item As Byte)
    If Item > mCount Then Exit Property
    mItem = Item
    With mLabelBtn
        .Caption = VBA.ChrW$(mArrIcon(mItem))
        If Item > mCountColor Then
            .ForeColor = mArrIconColor(mCountColor)
        Else
            .ForeColor = mArrIconColor(mItem)
        End If
    End With
End Property

Public Property Get Count() As Byte
    Count = UBound(mArrIcon)
End Property

Public Property Get ArrayIcon() As Variant
    ArrayIcon = mArrIcon
End Property

Public Property Let ArrayIcon(ByVal arr As Variant)
    mArrIcon = arr
    mCount = UBound(mArrIcon)
End Property

' Method to update icon colors after initialization
Public Sub UpdateIconColors(ArrIconColor As Variant)
    mArrIconColor = ArrIconColor
    If IsArray(mArrIconColor) Then
        mCountColor = UBound(mArrIconColor)
        ' Update the current color to match the new color array
        With mLabelBtn
            If mItem > mCountColor Then
                .ForeColor = mArrIconColor(mCountColor)
            Else
                .ForeColor = mArrIconColor(mItem)
            End If
        End With
    End If
End Sub

Public Sub Initialize(ByRef LabelBtn As MSForms.Label, _
        Optional Item As Byte, _
        Optional arrIcon As Variant, _
        Optional ArrIconColor As Variant)

    Set mLabelBtn = LabelBtn
    Set mParent = mLabelBtn.Parent
    
    ' Validate and set the icon array
    If Not IsArray(arrIcon) Then
        mArrIcon = Array(59193, 59194, 59195)
    Else
        mArrIcon = arrIcon
    End If
    
    ' Calculate count after setting the array
    mCount = UBound(mArrIcon)
    
    ' Set the initial item value, ensuring it's within bounds
    mItem = Item
    If mItem > mCount Then mItem = 0

    ' Validate and set the color array
    If Not IsArray(ArrIconColor) Then
        mArrIconColor = Array(vbBlack)
    Else
        mArrIconColor = ArrIconColor
    End If
    
    ' Calculate color count after setting the color array
    mCountColor = UBound(mArrIconColor)

    With mLabelBtn
        .Font.Name = "Segoe MDL2 Assets"
        .Font.Size = .Width * 0.9
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleNone
        .Height = .Width
        .Caption = VBA.ChrW$(mArrIcon(mItem))
        ' Apply the appropriate color based on current item and color array
        If mItem > mCountColor Then
            .ForeColor = mArrIconColor(mCountColor)
        Else
            .ForeColor = mArrIconColor(mItem)
        End If
    End With
End Sub

Private Sub mLabelBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call hoverColor(150)
End Sub

Private Sub mParent_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call hoverColor(0)
End Sub

' Function to blend colors for hover effect
Private Function BlendColor(ByVal color1 As Long, color2 As Long, factor As Double) As Long
    Dim r1 As Byte, g1 As Byte, b1 As Byte
    Dim r2 As Byte, g2 As Byte, b2 As Byte
    Dim r As Byte, g As Byte, b As Byte
    
    ' Extract RGB components from the first color
    r1 = color1 Mod 256
    g1 = (color1 \ 256) Mod 256
    b1 = (color1 \ 65536) Mod 256
    
    ' Extract RGB components from the second color
    r2 = color2 Mod 256
    g2 = (color2 \ 256) Mod 256
    b2 = (color2 \ 65536) Mod 256
    
    ' Blend the colors
    r = r1 + (r2 - r1) * factor
    g = g1 + (g2 - g1) * factor
    b = b1 + (b2 - b1) * factor
    
    ' Return the blended color
    BlendColor = RGB(r, g, b)
End Function

Private Sub hoverColor(ByVal shade As Byte)
    With mLabelBtn
        If mCountColor >= mItem Then
            ' Use a more sophisticated hover effect by blending colors
            .ForeColor = BlendColor(mArrIconColor(mItem), RGB(255, 255, 255), shade / 255)
        Else
            .ForeColor = BlendColor(mArrIconColor(mCountColor), RGB(255, 255, 255), shade / 255)
        End If
    End With
End Sub

Private Sub mLabelBtn_Click()
    mItem = mItem + 1
    If mCount < mItem Then mItem = 0
    mLabelBtn.Caption = VBA.ChrW$(mArrIcon(mItem))
    If IsArray(mArrIconColor) And mCountColor >= mItem Then mLabelBtn.ForeColor = mArrIconColor(mItem)
    RaiseEvent Click(mLabelBtn, mItem, mArrIcon(mItem), StateText)
End Sub

Private Sub Class_Terminate()
    Set mLabelBtn = Nothing
    Set mParent = Nothing
    Erase mArrIcon
    Erase mArrIconColor
End Sub
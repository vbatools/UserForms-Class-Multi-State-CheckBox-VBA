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
' Version: 1.0.4
' Creation Date: 13.11.2025 15:41
' Last Update Date: 20.11.2025 18:15
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
Private mCyclic     As Boolean
Private mFontName   As String
Private mFontSizeFactor As Double

Public Property Get Version() As String
    Version = "Class: clsMultiStateCheckBox" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "Version: 1.0.4" & vbNewLine & _
            "Date of creation: 13.11.2025 15:41" & vbNewLine & _
            "Date of update: 20.11.2025 18:15" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Description: A multi-state checkbox control for VBA UserForms with customizable icons and colors"
End Property

' Property to get the current state text
Public Property Get StateText() As String
    If IsArray(mArrIcon) And mItem <= UBound(mArrIcon) Then
        StateText = VBA.ChrW$(mArrIcon(mItem))
    Else
        StateText = vbNullString
    End If
End Property

' Property to set the current state text
Public Property Let StateText(ByVal value As String)
    ' Find the index of the character in the icon array
    Dim i As Byte
    For i = 0 To mCount
        If VBA.ChrW$(mArrIcon(i)) = value Then
            mItem = i
            With mLabelBtn
                .Caption = value
                If mItem > mCountColor Then
                    .ForeColor = mArrIconColor(mCountColor)
                Else
                    .ForeColor = mArrIconColor(mItem)
                End If
            End With
            Exit Property
        End If
    Next i
    
    ' If the character is not found in the array, raise an error
    Err.Raise vbObjectError + 1006, "clsMultiStateCheckBox", "StateText value not found in icon array"
End Property

' Property to get/set the current icon
Public Property Get CurrentIcon() As Long
    If IsArray(mArrIcon) And mItem <= UBound(mArrIcon) Then
        CurrentIcon = mArrIcon(mItem)
    Else
        CurrentIcon = 0
    End If
End Property

Public Property Let CurrentIcon(ByVal value As Long)
    ' Find the index of the icon in the array
    Dim i As Byte
    For i = 0 To mCount
        If mArrIcon(i) = value Then
            mItem = i
            With mLabelBtn
                .Caption = VBA.ChrW$(mArrIcon(mItem))
                If mItem > mCountColor Then
                    .ForeColor = mArrIconColor(mCountColor)
                Else
                    .ForeColor = mArrIconColor(mItem)
                End If
            End With
            Exit Property
        End If
    Next i
    
    ' If the icon is not found in the array, raise an error
        Err.Raise vbObjectError + 1007, "clsMultiStateCheckBox", "CurrentIcon value not found in icon array"
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

Public Property Get Cyclic() As Boolean
    Cyclic = mCyclic
End Property

Public Property Let Cyclic(ByVal value As Boolean)
    mCyclic = value
End Property

' Property to get/set the font name
Public Property Get FontName() As String
    If mFontName = "" Then
        FontName = "Segoe MDL2 Assets"
    Else
        FontName = mFontName
    End If
End Property

Public Property Let FontName(ByVal value As String)
    mFontName = value
    If Not mLabelBtn Is Nothing Then
        mLabelBtn.Font.Name = value
    End If
End Property

' Property to get/set the font size factor
Public Property Get FontSizeFactor() As Double
    If mFontSizeFactor = 0 Then
        FontSizeFactor = 0.9
    Else
        FontSizeFactor = mFontSizeFactor
    End If
End Property

Public Property Let FontSizeFactor(ByVal value As Double)
    mFontSizeFactor = value
    If Not mLabelBtn Is Nothing And mLabelBtn.Width > 0 Then
        mLabelBtn.Font.Size = mLabelBtn.Width * value
    End If
End Property

' Method to update icon colors after initialization
Public Sub UpdateIconColors(ArrIconColor As Variant)
    ' Check that ArrIconColor is an array
    If Not IsArray(ArrIconColor) Then
        Err.Raise vbObjectError + 1005, "clsMultiStateCheckBox", "ArrIconColor must be an array"
        Exit Sub
    End If
    
    mArrIconColor = ArrIconColor
    
    ' Check that the array is not empty
    If UBound(mArrIconColor) < 0 Then
        Err.Raise vbObjectError + 1004, "clsMultiStateCheckBox", "ArrIconColor cannot be empty"
        Exit Sub
    End If
    
    mCountColor = UBound(mArrIconColor)
    
    ' Update the current color to match the new color array
    With mLabelBtn
        If mItem > mCountColor Then
            .ForeColor = mArrIconColor(mCountColor)
        Else
            .ForeColor = mArrIconColor(mItem)
        End If
    End With
End Sub

' Method to set color for a specific state
Public Sub SetStateColor(StateIndex As Byte, Color As Long)
    If StateIndex > mCount Then
        Err.Raise vbObjectError + 107, "clsMultiStateCheckBox", "StateIndex exceeds the number of available states"
        Exit Sub
    End If
    
    If Not IsArray(mArrIconColor) Then
        ' If the color array doesn't exist, initialize it with the default color
        mArrIconColor = Array(vbBlack)
        mCountColor = 0
    End If
    
    ' Ensure the color array is large enough
    If StateIndex > mCountColor Then
        ' Resize the array to accommodate the new state
        Dim tempArray() As Long
        ReDim tempArray(0 To StateIndex)
        Dim i As Byte
        For i = 0 To mCountColor
            tempArray(i) = mArrIconColor(i)
        Next i
        For i = mCountColor + 1 To StateIndex
            ' Use the last color as default for new states
            tempArray(i) = mArrIconColor(mCountColor)
        Next i
        mArrIconColor = tempArray
        mCountColor = StateIndex
    End If
    
    ' Set the color for the specified state
    mArrIconColor(StateIndex) = Color
    
    ' Update the display if this is the current state
    If mItem = StateIndex Then
        mLabelBtn.ForeColor = Color
    End If
End Sub

' Method to reset to initial state
Public Sub ResetToInitialState(InitialItem As Byte)
    If InitialItem > mCount Then
        ' If the provided initial item is out of bounds, default to 0
        mItem = 0
    Else
        mItem = InitialItem
    End If
    
    With mLabelBtn
        .Caption = VBA.ChrW$(mArrIcon(mItem))
        If mItem > mCountColor Then
            .ForeColor = mArrIconColor(mCountColor)
        Else
            .ForeColor = mArrIconColor(mItem)
        End If
    End With
End Sub

' Method to get all available states
Public Function GetAllStates() As Variant
    Dim result() As String
    ReDim result(0 To mCount)
    
    Dim i As Byte
    For i = 0 To mCount
        result(i) = VBA.ChrW$(mArrIcon(i))
    Next i
    
    GetAllStates = result
End Function

Public Sub Initialize(ByRef LabelBtn As MSForms.Label, _
        Optional Item As Byte, _
        Optional arrIcon As Variant, _
        Optional ArrIconColor As Variant, _
        Optional fontName As String = "Segoe MDL2 Assets", _
        Optional fontSizeFactor As Double = 0.9)

    ' Check for Nothing on LabelBtn
    If LabelBtn Is Nothing Then
        Err.Raise vbObjectError + 100, "clsMultiStateCheckBox", "LabelBtn cannot be Nothing"
        Exit Sub
    End If
    
    Set mLabelBtn = LabelBtn
    Set mParent = mLabelBtn.Parent
    
    ' Validate and set the icon array
    If Not IsArray(arrIcon) Then
        mArrIcon = Array(59193, 59194, 59195)
    Else
        ' Check that arrIcon is actually an array
        If Not IsArray(arrIcon) Then
            Err.Raise vbObjectError + 1001, "clsMultiStateCheckBox", "arrIcon must be an array"
            Exit Sub
        End If
        mArrIcon = arrIcon
    End If
    
    ' Check that the array is not empty
    If UBound(mArrIcon) < 0 Then
        Err.Raise vbObjectError + 1002, "clsMultiStateCheckBox", "arrIcon cannot be empty"
        Exit Sub
    End If
    
    mCount = UBound(mArrIcon)
    
    ' Set the initial item value, ensuring it's within bounds
    mItem = IIf(Item <= mCount, Item, 0)

    ' Validate and set the color array
    If Not IsArray(ArrIconColor) Then
        mArrIconColor = Array(vbBlack)
    Else
        ' Check that ArrIconColor is actually an array
        If Not IsArray(ArrIconColor) Then
            Err.Raise vbObjectError + 1003, "clsMultiStateCheckBox", "ArrIconColor must be an array"
            Exit Sub
        End If
        mArrIconColor = ArrIconColor
    End If
    
    ' Check that the color array is not empty
    If UBound(mArrIconColor) < 0 Then
        Err.Raise vbObjectError + 1004, "clsMultiStateCheckBox", "ArrIconColor cannot be empty"
        Exit Sub
    End If
    
    mCountColor = UBound(mArrIconColor)

    ' Initialize the cyclic property to True by default
    mCyclic = True
    
    ' Set font properties
    mFontName = fontName
    mFontSizeFactor = fontSizeFactor

    With mLabelBtn
        .Font.Name = fontName
        .Font.Size = .Width * fontSizeFactor
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
    
    ' Extract BGR components from the first color (VBA uses BGR format)
    b1 = color1 Mod 256
    g1 = (color1 \ 256) Mod 256
    r1 = (color1 \ 6536) Mod 256
    
    ' Extract BGR components from the second color (VBA uses BGR format)
    b2 = color2 Mod 256
    g2 = (color2 \ 256) Mod 256
    r2 = (color2 \ 65536) Mod 256
    
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
    If mCyclic Then
        mItem = (mItem + 1) Mod (mCount + 1)
    Else
        If mItem < mCount Then
            mItem = mItem + 1
        Else
            ' If the last state is reached, we can either stay or reset to 0 depending on requirements
        End If
    End If
    
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
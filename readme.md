# Multi-State CheckBox VBA Class

A VBA class implementation that provides a multi-state checkbox control with three states: unchecked, checked, and indeterminate.

## Features

- Three-state checkbox functionality (unchecked, checked, indeterminate)
- Customizable appearance and behavior
- Easy integration with VBA UserForms
- Reusable class module for consistent checkbox behavior

## Default Icons and Colors

- Default icons: Uses Segoe MDL2 Assets font with character codes 59193 (unchecked), 59194 (checked), and 59195 (indeterminate)
- Default color: Black (vbBlack) for all states
- Hover effect: Lightens the icon color when hovering over the checkbox

## Files

- `clsMultiStateCheckBox.cls` - Main class implementation
- `frmTestClass.frm` - Test form demonstrating usage
- `modShowForms.bas` - Module with form display functions
- `multi_state_checkbox_v2.xlsm` - Excel workbook containing the implementation

## Usage

### Initializing in a UserForm

To use the multi-state checkbox in your UserForm:

1. Add a Label control to your form that will serve as the checkbox
2. Import the `clsMultiStateCheckBox.cls` class module into your VBA project
3. In your form's code, declare an instance of the class:

```vba
Dim mCheckBox As clsMultiStateCheckBox

Private Sub UserForm_Initialize()
    ' Initialize with default parameters
    Set mCheckBox = New clsMultiStateCheckBox
    mCheckBox.Initialize Me.Label1 ' Replace Label1 with your label control
End Sub

' Example with all optional parameters:
' Initialize(LabelBtn As MSForms.Label, _
'         Optional Item As Byte, _
'         Optional arrIcon As Variant, _
'         Optional ArrIconColor As Variant)
'
' Where:
' - LabelBtn: The label control to use as the checkbox
' - Item: Initial state (0 for unchecked, 1 for checked, 2 for indeterminate, etc.)
' - arrIcon: Array of icon character codes (default is Array(59193, 59194, 59195))
' - ArrIconColor: Array of colors for each state (default is Array(vbBlack))

Private Sub UserForm_Initialize()
    ' Example with all optional parameters specified
    Set mCheckBox = New clsMultiStateCheckBox
    
    ' Define custom icons (in this case, using Segoe MDL2 Assets)
    Dim customIcons As Variant
    customIcons = Array(59193, 59194, 59195) ' Unchecked, checked, indeterminate
    
    ' Define custom colors for each state
    Dim customColors As Variant
    customColors = Array(vbBlack, vbBlue, vbRed) ' Colors for each state
    
    ' Initialize with custom parameters
    mCheckBox.Initialize Me.Label1, 0, customIcons, customColors
End Sub

Private Sub UserForm_Terminate()
    Set mCheckBox = Nothing
End Sub
```

4. The checkbox will now function with three states that cycle when clicked

## License

This project is licensed under the terms specified in the [LICENSE](./LICENSE) file.
# clsMultiStateCheckBox Class Documentation

## Table of Contents
1. [Overview](#overview)
2. [System Requirements](#system-requirements)
3. [Installation and Setup](#installation-and-setup)
4. [API Reference](#api-reference)
5. [Implementation Guides](#implementation-guides)
6. [Performance Considerations](#performance-considerations)
7. [Error Handling Strategies](#error-handling-strategies)
8. [Integration Examples](#integration-examples)
9. [Configuration Options](#configuration-options)
10. [Version History](#version-history)
11. [State Transition Diagrams](#state-transition-diagrams)
12. [Sample Projects](#sample-projects)
13. [Testing Methodologies](#testing-methodologies)
14. [Office Version Compatibility](#office-version-compatibility)
15. [Security Considerations](#security-considerations)

## Overview

The `clsMultiStateCheckBox` class is an implementation of a multi-state checkbox for VBA UserForms with three or more states: unchecked, checked, and indeterminate, as well as other custom states. This class provides customizable appearance and behavior with easy integration into VBA UserForms.

### Key Features
- Three-state checkbox functionality (unchecked, checked, indeterminate)
- Customizable appearance and behavior
- Easy integration with VBA UserForms
- Reusable class module for consistent checkbox behavior
- Cyclic/non-cyclic state switching (property `Cyclic`)
- Ability to set/get state by text (property `StateText`)
- Method to get all available states (`GetAllStates`)
- Improved error handling and validation
- Ability to set font name and size factor (properties `FontName` and `FontSizeFactor`)
- Ability to get/set current icon (property `CurrentIcon`)
- Method to set color for specific state (`SetStateColor`)
- Method to reset to initial state (`ResetToInitialState`)

## System Requirements

- Microsoft Excel (2010 or higher)
- Enabled macros
- VBA permissions (medium or low security level)

## Installation and Setup

1. Import the `clsMultiStateCheckBox.cls` file into your VBA project
2. Add a Label control to your form that will serve as the checkbox
3. Declare a class variable in your form

```vba
Private mCheckBox As clsMultiStateCheckBox
```

4. Initialize the control in the `UserForm_Initialize` event

```vba
Private Sub UserForm_Initialize()
    Set mCheckBox = New clsMultiStateCheckBox
    mCheckBox.Initialize Me.Label1 ' Replace Label1 with your Label control
End Sub
```

## API Reference

### Events

#### Click
Event raised when the multi-state checkbox is clicked

**Syntax:**
```vba
Private Sub mCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
```

**Parameters:**
- `control` - The label control that was clicked
- `Item` - The current state index (0-based)
- `CodeIcon` - The character code of the current icon
- `StateText` - The text representation of the current state

### Properties

#### Version
Gets the version information of the clsMultiStateCheckBox class

**Type:** String
**Access:** Read-only

#### StateText
Gets or sets the text of the current state

**Type:** String
**Access:** Read/Write

**Usage example:**
```vba
' Get current state
Dim currentState As String
currentState = mCheckBox.StateText

' Set state
mCheckBox.StateText = ChrW$(59194) ' Set to second state
```

#### Item
Gets or sets the current state index

**Type:** Byte
**Access:** Read/Write

#### Count
Gets the number of available states

**Type:** Byte
**Access:** Read-only

#### ArrayIcon
Gets or sets the array of icons for each state

**Type:** Variant
**Access:** Read/Write

#### Cyclic
Gets or sets whether the checkbox will cycle through states cyclically

**Type:** Boolean
**Access:** Read/Write

**Usage example:**
```vba
' Disable cycling
mCheckBox.Cyclic = False
```

#### FontName
Gets or sets the font name for the icon

**Type:** String
**Access:** Read/Write

**Usage example:**
```vba
' Set font
mCheckBox.FontName = "Wingdings"
```

#### FontSizeFactor
Gets or sets the font size factor for scaling

**Type:** Double
**Access:** Read/Write

**Usage example:**
```vba
' Set font size to 80% of width
mCheckBox.FontSizeFactor = 0.8
```

#### CurrentIcon
Gets or sets the current icon by code

**Type:** Long
**Access:** Read/Write

**Usage example:**
```vba
' Get current icon
Dim currentIcon As Long
currentIcon = mCheckBox.CurrentIcon

' Set current icon
mCheckBox.CurrentIcon = 59194
```

### Methods

#### Initialize
Initializes the multi-state checkbox control

**Syntax:**
```vba
Public Sub Initialize(ByRef LabelBtn As MSForms.Label, _
        Optional Item As Byte, _
        Optional arrIcon As Variant, _
        Optional ArrIconColor As Variant, _
        Optional fontName As String = "Segoe MDL2 Assets", _
        Optional fontSizeFactor As Double = 0.9)
```

**Parameters:**
- `LabelBtn` - The label control to use as the checkbox
- `Item` - Initial state (0 for unchecked, 1 for checked, 2 for indeterminate, etc.)
- `arrIcon` - Array of icon character codes (default is Array(59193, 59194, 59195))
- `ArrIconColor` - Array of colors for each state (default is Array(vbBlack))
- `fontName` - Font name for icons (default "Segoe MDL2 Assets")
- `fontSizeFactor` - Font size factor (default 0.9)

**Usage example:**
```vba
Dim customIcons As Variant
customIcons = Array(59193, 59194, 59195) ' Unchecked, checked, indeterminate

Dim customColors As Variant
customColors = Array(vbBlack, vbBlue, vbRed) ' Colors for each state

mCheckBox.Initialize Me.Label1, 0, customIcons, customColors, "Segoe MDL2 Assets", 0.8
```

#### UpdateIconColors
Updates icon colors after initialization

**Syntax:**
```vba
Public Sub UpdateIconColors(ArrIconColor As Variant)
```

**Parameters:**
- `ArrIconColor` - Array of colors for each state

#### SetStateColor
Sets color for a specific state

**Syntax:**
```vba
Public Sub SetStateColor(StateIndex As Byte, Color As Long)
```

**Parameters:**
- `StateIndex` - State index (0-based)
- `Color` - Color to set

**Usage example:**
```vba
' Set red color for second state
mCheckBox.SetStateColor 1, RGB(255, 0, 0)
```

#### ResetToInitialState
Resets the checkbox to initial state

**Syntax:**
```vba
Public Sub ResetToInitialState(InitialItem As Byte)
```

**Parameters:**
- `InitialItem` - Initial state index

**Usage example:**
```vba
' Reset checkbox to first state
mCheckBox.ResetToInitialState 0
```

#### GetAllStates
Returns an array of all available state texts

**Syntax:**
```vba
Public Function GetAllStates() As Variant
```

**Return Value:**
- Array of strings representing all available states

**Usage example:**
```vba
Dim allStates As Variant
allStates = mCheckBox.GetAllStates
```

## Implementation Guides

### Basic Implementation

The simplest way to use the multi-state checkbox:

1. Add a Label control to your form
2. Declare a class variable:
```vba
Private mCheckBox As clsMultiStateCheckBox
```

3. Initialize the control:
```vba
Private Sub UserForm_Initialize()
    Set mCheckBox = New clsMultiStateCheckBox
    mCheckBox.Initialize Me.Label1
End Sub
```

4. Handle the Click event:
```vba
Private Sub mCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
    Debug.Print "State: " & Item & ", Text: " & StateText
End Sub
```

### Advanced Implementation with Custom Icons

For customizing the appearance of the checkbox:

1. Define an icon array:
```vba
Dim customIcons As Variant
customIcons = Array(59193, 59194, 59195) ' Unicode character codes
```

2. Define a color array:
```vba
Dim customColors As Variant
customColors = Array(vbBlack, vbBlue, vbRed)
```

3. Initialize the control with custom parameters:
```vba
Private Sub UserForm_Initialize()
    Set mCheckBox = New clsMultiStateCheckBox
    
    Dim customIcons As Variant
    customIcons = Array(59193, 59194, 59195)
    
    Dim customColors As Variant
    customColors = Array(vbBlack, vbBlue, vbRed)
    
    mCheckBox.Initialize Me.Label1, 0, customIcons, customColors
End Sub
```

### Implementation with Cycling Disabled

For creating a progressive indicator that doesn't return to the first state:

```vba
Private Sub UserForm_Initialize()
    Set mCheckBox = New clsMultiStateCheckBox
    mCheckBox.Initialize Me.Label1
    mCheckBox.Cyclic = False ' Disable cycling
End Sub
```

## Performance Considerations

### Initialization Optimization

- Use the minimum required number of states for your task
- Use standard fonts if special typography is not required
- Initialize the control once in the UserForm_Initialize event

### Update Optimization

- Avoid frequent calls to UpdateIconColors if not needed
- Use SetStateColor only when you need to change color of a single state
- Minimize UI updates in loops

## Error Handling Strategies

### Parameter Validation

The class includes built-in parameter validation:

- Check that LabelBtn is not Nothing
- Check that icon and color arrays are arrays
- Check array bounds
- Check state index validity

### Error Handling

The class raises the following errors:

- `vbObjectError + 100`: LabelBtn is Nothing
- `vbObjectError + 1001`: arrIcon is not an array
- `vbObjectError + 1002`: arrIcon is empty
- `vbObjectError + 103`: ArrIconColor is not an array
- `vbObjectError + 1004`: ArrIconColor is empty
- `vbObjectError + 105`: ArrIconColor is not an array (in UpdateIconColors)
- `vbObjectError + 1006`: StateText not found in icon array
- `vbObjectError + 1007`: CurrentIcon not found in icon array
- `vbObjectError + 107`: StateIndex exceeds the number of available states

### Recommended Error Handling

```vba
Private Sub InitializeCheckBox()
    On Error GoTo ErrorHandler
    
    Set mCheckBox = New clsMultiStateCheckBox
    mCheckBox.Initialize Me.Label1
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error initializing checkbox: " & Err.Description, vbCritical
End Sub
```

## Integration Examples

### UserForm Integration

```vba
' In the form module
Private mCheckBox As clsMultiStateCheckBox

Private Sub UserForm_Initialize()
    Set mCheckBox = New clsMultiStateCheckBox
    
    ' Initialize with custom icons and colors
    Dim icons As Variant
    icons = Array(59193, 59194, 59195) ' Unchecked, checked, indeterminate
    
    Dim colors As Variant
    colors = Array(vbRed, vbGreen, vbBlue) ' Red, green, blue
    
    mCheckBox.Initialize Me.Label1, 0, icons, colors
End Sub

Private Sub mCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
    Select Case Item
        Case 0
            MsgBox "Checkbox unchecked"
        Case 1
            MsgBox "Checkbox checked"
        Case 2
            MsgBox "Checkbox in indeterminate state"
    End Select
End Sub
```

### UserForm Integration with Multiple Checkboxes

```vba
' In the form module
Private mCheckBox1 As clsMultiStateCheckBox
Private mCheckBox2 As clsMultiStateCheckBox

Private Sub UserForm_Initialize()
    ' First checkbox
    Set mCheckBox1 = New clsMultiStateCheckBox
    mCheckBox1.Initialize Me.Label1
    
    ' Second checkbox with different parameters
    Set mCheckBox2 = New clsMultiStateCheckBox
    Dim icons As Variant
    icons = Array(59193, 59194) ' Only two states
    mCheckBox2.Initialize Me.Label2, 0, icons
End Sub
```

## Configuration Options

### Appearance Configuration

The class allows configuring:

- Icons for each state
- Colors for each state
- Icon fonts
- Font size
- State switching mode

### Configuration Examples

```vba
' Configure icons and colors
Dim icons As Variant
icons = Array(59193, 59194, 59195, 59634) ' Four states

Dim colors As Variant
colors = Array(vbRed, vbGreen, vbBlue, vbYellow) ' Colors for each state

mCheckBox.Initialize Me.Label1, 0, icons, colors, "Wingdings", 0.7

' Change font after initialization
mCheckBox.FontName = "Webdings"
mCheckBox.FontSizeFactor = 0.85

' Set color for specific state
mCheckBox.SetStateColor 2, RGB(128, 0, 128) ' Purple for third state
```

## Version History

### 1.0.4
- Added CurrentIcon property for getting/setting current icon
- Added SetStateColor method for setting color for specific state
- Added ResetToInitialState method for resetting to initial state
- Added FontName and FontSizeFactor properties for font configuration
- Improved error handling

### 1.0.3
- Fixed BlendColor method for correct BGR color handling
- Added input validation in Initialize method
- Improved error handling in all methods
- Added StateText property for setting state by text
- Added Cyclic property for controlling state cycling
- Added GetAllStates method for getting all available states
- Improved UpdateIconColors method with validation checks

### 1.0.2
- Added default icons for three states
- Implemented state switching on click
- Added Click event
- Added support for custom colors

### 1.0.1
- Basic class implementation
- Support for three states (unchecked, checked, indeterminate)
- Using Segoe MDL2 Assets font for icons

### 1.0.0
- Initial release
- Basic multi-state checkbox functionality

## State Transition Diagrams

### Cyclic Mode (Cyclic = True)
```
State 0 → State 1 → State 2 → ... → State N → State 0
    ↑                                           ↓
    └───────────────────────────┘
```

### Non-Cyclic Mode (Cyclic = False)
```
State 0 → State 1 → State 2 → ... → State N
    ↑         ↑         ↑              ↑
   Start   Continue   Continue     Last (stays)
```

## Sample Projects

### Project: Task Management

Using the multi-state checkbox for tracking task status:

```vba
' Task status states:
' 0 - Not Started (empty square)
' 1 - In Progress (clock)
' 2 - Completed (checkmark)
' 3 - Cancelled (cross)

Private Sub InitializeTaskCheckBox()
    Set mCheckBox = New clsMultiStateCheckBox
    
    ' Using icons from Segoe MDL2 Assets font
    Dim taskIcons As Variant
    taskIcons = Array(59193, 59244, 59194, 59226) ' Not Started, In Progress, Completed, Cancelled
    
    Dim taskColors As Variant
    taskColors = Array(vbGray, vbBlue, vbGreen, vbRed) ' Colors for each status
    
    mCheckBox.Initialize Me.Label1, 0, taskIcons, taskColors
End Sub

Private Sub mCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
    Dim statusText As String
    Select Case Item
        Case 0: statusText = "Not Started"
        Case 1: statusText = "In Progress"
        Case 2: statusText = "Completed"
        Case 3: statusText = "Cancelled"
    End Select
    
    MsgBox "Task status changed to: " & statusText
End Sub
```

## Testing Methodologies

### Unit Testing

For testing the class functionality, it's recommended to create a test form with multiple class instances:

```vba
' Testing initialization
Sub TestInitialization()
    Dim checkBox As clsMultiStateCheckBox
    Set checkBox = New clsMultiStateCheckBox
    Dim form As Object
    Set form = CreateObject("Forms.Label") ' Mock control
    
    On Error Resume Next
    checkBox.Initialize form
    If Err.Number <> 0 Then
        Debug.Print "Initialization error: " & Err.Description
    Else
        Debug.Print "Initialization successful"
    End If
    On Error GoTo 0
End Sub

' Testing state cycling
Sub TestStateCycling()
    Dim checkBox As clsMultiStateCheckBox
    Set checkBox = New clsMultiStateCheckBox
    Dim form As Object
    Set form = CreateObject("Forms.Label")
    
    checkBox.Initialize form
    checkBox.Cyclic = True
    
    ' Check cyclic switching
    Dim i As Integer
    For i = 0 To 5 ' Test 6 clicks
        checkBox.mLabelBtn_Click ' Mock click (if possible)
        Debug.Print "State after click " & i + 1 & ": " & checkBox.Item
    Next i
End Sub
```

### Validation Testing

```vba
' Testing parameter validation
Sub TestValidation()
    Dim checkBox As clsMultiStateCheckBox
    Set checkBox = New clsMultiStateCheckBox
    Dim form As Object
    Set form = CreateObject("Forms.Label")
    
    ' Test with invalid icon array
    On Error Resume Next
    checkBox.Initialize form, 0, "not an array"
    If Err.Number <> 0 Then
        Debug.Print "Correctly handled error with invalid icon array"
    End If
    On Error GoTo 0
End Sub
```

## Office Version Compatibility

### Supported Versions

- Microsoft Excel 2010 and above
- Microsoft Word 2010 and above (for forms with controls)
- Microsoft Access 2010 and above
- Microsoft PowerPoint 2010 and above

### Version-Specific Features

#### Excel 2010-2013
- Full support for all class functions
- Recommended to use system-available fonts

#### Excel 2016/2019
- Full support for all class functions
- Enhanced Segoe MDL2 Assets font support

#### Excel 365
- Full support for all class functions
- Best support for modern fonts and icons

### Compatibility Recommendations

- Use standard system fonts for better compatibility
- Check for Segoe MDL2 Assets font availability before use
- Alternative fonts: Wingdings, Webdings, Symbol

## Security Considerations

### Macros and Security

- Ensure macro security level allows custom classes to run
- Use digital signatures for VBA projects in enterprise environments
- Check file and object access permissions

### Protection from Improper Usage

The class includes built-in checks:

- Nothing check for passed objects
- Type checking for parameter arrays
- Array bounds checking
- Valid index value checking

### Security Recommendations

- Don't pass unvalidated data to class methods
- Handle errors during initialization and usage
- Use security checks in enterprise applications
- Update class when potential vulnerabilities are identified

## Conclusion

The clsMultiStateCheckBox class provides a powerful and flexible way to implement multi-state checkboxes in VBA applications. With it, you can create intuitive interfaces with extended functionality compared to traditional checkboxes.

The class supports extensive customization, including custom icons, colors, fonts, and switching modes. With built-in validation and error handling, it is a reliable solution for production applications.
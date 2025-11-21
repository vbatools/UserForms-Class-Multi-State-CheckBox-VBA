# User Guide for clsMultiStateCheckBox

## Table of Contents
1. [Introduction](#introduction)
2. [System Requirements](#system-requirements)
3. [Installation and Setup](#installation-and-setup)
4. [Quick Start](#quick-start)
5. [Main Features](#main-features)
6. [Working with Controls](#working-with-controls)
7. [Style Configuration](#style-configuration)
8. [Creating Toggle Switches](#creating-toggle-switches)
9. [Adding Icons](#adding-icons)
10. [Working with Style Collection](#working-with-style-collection)
11. [Troubleshooting](#troubleshooting)
12. [Frequently Asked Questions](#frequently-asked-questions)

## Introduction

The `clsMultiStateCheckBox` class implements a multi-state checkbox for VBA UserForms with three or more states: unchecked, checked, and indeterminate, as well as other custom states. This class provides customizable appearance and behavior with easy integration into VBA UserForms.

### What This Project Can Do:
- Apply modern design to checkbox controls
- Provide multi-state functionality beyond standard checkbox
- Configure colors and fonts for checkboxes
- Add icons and visual elements
- Manage visibility and state of checkbox controls

## System Requirements

- Microsoft Excel (2010 or newer recommended)
- VBA support enabled
- Microsoft Forms 2.0 Object Library
- Windows 7 or newer

## Installation and Setup

### Step 1: Import the Class
1. Open Excel and go to the VBA editor (press Alt+F11)
2. In the menu, select "File" > "Import File"
3. Select the `clsMultiStateCheckBox.cls` file from the `vba-files/Class/` directory
4. Click "Open" to import the class

### Step 2: Configure References
1. In the VBA editor, select "Tools" > "References"
2. Find and check the box next to "Microsoft Forms 2.0 Object Library"
3. Click "OK" to save changes

### Step 3: Create a User Form
1. In the VBA editor, create a new user form
2. Add a Label control that will serve as the checkbox
3. Add a class variable to the form:
```vba
Dim MultiStateCheckBox As clsMultiStateCheckBox
```

## Quick Start

### Simple Usage Example
1. Create a new user form in Excel
2. Add a Label control to serve as the checkbox
3. In the `UserForm_Initialize` event, add the following code:
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    Call MultiStateCheckBox.Initialize(Me.Label1)
End Sub
```
4. Run the form to see the multi-state checkbox control

### Example with Color Configuration
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    Dim icons As Variant
    icons = Array(59193, 59194, 59195) ' Unchecked, checked, indeterminate
    Dim colors As Variant
    colors = Array(vbRed, vbGreen, vbBlue) ' Colors for each state
    Call MultiStateCheckBox.Initialize(Me.Label1, 0, icons, colors)
End Sub
```

## Main Features

### Checkbox Initialization
The `Initialize` method is the main way to set up the multi-state checkbox:
- Sets the initial state (unchecked, checked, indeterminate, etc.)
- Configures appearance properties
- Sets up event handlers for interaction
- Creates visual elements for different states

### Supported States
- Unchecked state (typically represented by empty square)
- Checked state (typically represented by checkmark)
- Indeterminate state (typically represented by dash or filled square)
- Additional custom states can be defined

## Working with Controls

### Checkbox Configuration
For checkboxes, the class automatically:
- Applies the specified state icons
- Creates visual elements for different states
- Handles click events to cycle through states
- Provides visual feedback during interaction
- Handles mouse events for interaction

### State Management
The class allows configuring:
- Initial state
- State transition behavior (cyclic/non-cyclic)
- Icons for each state
- Colors for each state
- Text representation of states

### Event Handling
The class provides events for:
- State changes during interaction
- Click events on the checkbox
- Visual feedback during state changes

## Style Configuration

### Color Schemes
The class provides extensive color configuration options:
- State-specific colors
- Icon colors for different states
- Background colors (if applicable)
- Text colors for state labels

### Configuration via Initialize Method
Colors can be configured during initialization:
```vba
Dim colors As Variant
colors = Array(vbRed, vbGreen, vbBlue)
MultiStateCheckBox.Initialize Me.Label1, 0, Array(59193, 59194, 59195), colors
```

### Dynamic Configuration
After initialization, you can change properties:
```vba
With MultiStateCheckBox
    .SetStateColor 1, RGB(255, 0, 0)  ' Set red color for second state
    .CurrentIcon = 59194  ' Set current icon to checked state
End With
```

### Font Configuration
The class allows:
- Setting font name for icons (typically Segoe MDL2 Assets)
- Configuring font size factor via `FontSizeFactor` property

## Creating Toggle Switches

The `clsMultiStateCheckBox` class can be used to create toggle switches by:
- Setting appropriate icons for each state
- Configuring the `Cyclic` property to control state cycling
- Using the `SetStateColor` method to customize appearance for each state

### Simple Toggle Switch
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    ' Create a simple on/off toggle with custom icons
    Dim icons As Variant
    icons = Array(60434, 60433) ' Toggle off/on icons
    Call MultiStateCheckBox.Initialize(Me.Label1, 0, icons)
    MultiStateCheckBox.Cyclic = True  ' Enable cycling between states
End Sub
```

## Adding Icons

### Using Built-in Icons
The class uses the Segoe MDL2 Assets font for icons:
- Unchecked state: 59193 (empty square)
- Checked state: 59194 (square with checkmark)
- Indeterminate: 59195 (square with dash)
- Various other icons available in the enumIcons enumeration

### Setting Icons
To customize the checkbox icons:
1. Create an array of icon codes
2. Pass it to the Initialize method:
```vba
Dim customIcons As Variant
customIcons = Array(59193, 59194, 59195) ' Unchecked, checked, indeterminate
MultiStateCheckBox.Initialize Me.Label1, 0, customIcons
```

### Icon Configuration
- Icons are displayed using the Segoe MDL2 Assets font by default
- Icon color changes depending on the checkbox state
- Icons automatically scale to the control size

## Working with Style Collection

### Accessing Individual Elements
After initialization, the checkbox provides access to its properties:
```vba
' Getting checkbox properties
Dim currentState As Byte
currentState = MultiStateCheckBox.Item
Dim stateText As String
stateText = MultiStateCheckBox.StateText
```

### Getting the Number of Elements
The checkbox class manages a single control element, but provides access to its properties.

### Iterating Through All Elements
The checkbox class focuses on a single control, so iteration is not typically necessary.

### Changing Properties of Individual Elements
```vba
' Changing checkbox properties
With MultiStateCheckBox
    .Item = 1  ' Set to checked state
    .StateText = ChrW(59194)  ' Set state text to checked icon
    .Cyclic = False  ' Disable cycling
End With
```

## Troubleshooting

### Display Issues
- Ensure Microsoft Forms 2.0 Object Library is enabled in references
- Check that the Label control is added before calling the Initialize method
- Ensure the MultiUse property is set to True for the class

### Interaction Issues
- Check that control events are not overloaded with other handlers
- Ensure control properties are not changed manually while the class is running
- Verify that the class is not initialized multiple times

### Performance Issues
- Reduce the number of state changes during initialization
- Avoid frequent calls to property getters during interaction
- Use visibility and availability properties appropriately

### Common Errors
- "Object variable not set" - ensure the class variable is properly initialized
- "Method or data member not found" - check that the class is properly imported
- "Can't assign to property" - avoid direct assignment to nested objects without checking for Nothing

## Frequently Asked Questions

### Question: How to change colors after initialization?
**Answer:** Use the `SetStateColor` method to set colors for specific states after initialization.

### Question: Are all control types supported?
**Answer:** The class specifically supports multi-state checkbox functionality using Label controls.

### Question: Can multiple class instances be used?
**Answer:** Yes, you can create multiple class instances for different checkboxes, but each instance manages a single checkbox control.

### Question: How to add custom icons?
**Answer:** The class uses Unicode character values for icons. For custom icons, you can use different character codes or consider using image controls.

### Question: Is the class compatible with different Excel versions?
**Answer:** The class is tested with Excel 2010 and newer. Compatibility with earlier versions is not guaranteed.

### Question: Can animation be configured?
**Answer:** The current version provides visual feedback during state changes, but direct animation configuration is not available.

### Question: How to handle events of checkbox elements?
**Answer:** The class provides Click and state change events that can be handled in the form module.
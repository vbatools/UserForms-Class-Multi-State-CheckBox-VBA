# Technical Documentation for clsMultiStateCheckBox

## Table of Contents
1. [Class Overview](#class-overview)
2. [Class Architecture](#class-architecture)
3. [Properties](#properties)
4. [Methods](#methods)
5. [Events](#events)
6. [Constants and Enumerations](#constants-and-enumerations)
7. [Implementation Details](#implementation-details)
8. [Dependencies](#dependencies)

## Class Overview

The `clsMultiStateCheckBox` class implements a multi-state checkbox for VBA UserForms with three or more states: unchecked, checked, and indeterminate, as well as other custom states. This class provides customizable appearance and behavior with easy integration into VBA UserForms.

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

## Class Architecture

### Core Class Properties

| Property | Type | Description |
|----------|------|-------------|
| `Item` | Byte | Current state index |
| `Count` | Byte | Number of available states |
| `Cyclic` | Boolean | Whether the checkbox cycles through states or stops at the last state |
| `StateText` | String | Current state text |
| `CurrentIcon` | Long | Current icon code |
| `FontName` | String | Font name for the checkbox |
| `FontSizeFactor` | Double | Font size factor for scaling |
| `ArrayIcon` | Variant | Array of icons for each state |
| `Value` | Boolean | Current value of the checkbox |

## Properties

### Core Properties
- `Item` - Gets or sets the current state index
- `Count` - Gets the number of available states
- `Cyclic` - Gets or sets whether the checkbox cycles through states or stops at the last state
- `StateText` - Gets or sets the current state text
- `CurrentIcon` - Gets or sets the current icon code
- `FontName` - Gets or sets the font name for the checkbox
- `FontSizeFactor` - Gets or sets the font size factor for scaling
- `ArrayIcon` - Gets or sets the array of icons for each state
- `Value` - Gets or sets the current value of the checkbox

### Additional Properties
- `Version` - Gets version information about the class
- `Items` - Gets the collection of all checkbox items
- `Count` - Gets the number of items in the collection

## Methods

### Initialize
Initializes the multi-state checkbox control.

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

### UpdateIconColors
Updates icon colors after initialization.

**Syntax:**
```vba
Public Sub UpdateIconColors(ArrIconColor As Variant)
```

**Parameters:**
- `ArrIconColor` - Array of colors for each state

### SetStateColor
Sets color for a specific state.

**Syntax:**
```vba
Public Sub SetStateColor(StateIndex As Byte, Color As Long)
```

**Parameters:**
- `StateIndex` - State index (0-based)
- `Color` - Color to set

### ResetToInitialState
Resets the checkbox to initial state.

**Syntax:**
```vba
Public Sub ResetToInitialState(InitialItem As Byte)
```

**Parameters:**
- `InitialItem` - Initial state index

### GetAllStates
Returns an array of all available state texts.

**Syntax:**
```vba
Public Function GetAllStates() As Variant
```

**Return Value:**
- Array of strings representing all available states

## Events

### Click
Event raised when the multi-state checkbox is clicked.

**Syntax:**
```vba
Private Sub mLabelBtn_Click()
```

## Constants and Enumerations

The class does not define specific enumerations but uses standard VBA data types for state management.

## Implementation Details

### State Management
The class supports multiple states through an array of icons and colors:
- Each state has a corresponding icon and color
- States can be cycled through or stop at the last state
- State transitions are handled automatically

### Visual Elements
The checkbox consists of multiple visual elements:
- Label control for the checkbox button
- Optional label for the checkbox caption
- Visual elements for different states

### Event Handling
The class automatically handles checkbox events:
- `Click` - toggles the state and updates the visual appearance
- State changes are reflected in the UI immediately

## Dependencies

- MSForms.Label control
- VBA runtime environment
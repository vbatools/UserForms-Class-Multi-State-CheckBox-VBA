# VBA Multi-State Checkbox Class

![Project Demo](User_Forms.gif)

This repository contains a VBA implementation of a multi-state checkbox control that can be used in Excel UserForms. The checkbox class provides three or more states (unchecked, checked, indeterminate) with customizable appearance and behavior.

## Table of Contents
1. [Features](#features)
2. [Components](#components)
3. [Installation](#installation)
4. [Quick Start](#quick-start)
5. [Main Functions](#main-functions)
6. [Working with Controls](#working-with-controls)
7. [Style Configuration](#style-configuration)
8. [Troubleshooting](#troubleshooting)

## Features

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

## Components

- `clsMultiStateCheckBox.cls`: The main checkbox class implementation
- `frmTestClass.frm`: Test form demonstrating usage
- `modShowForms.bas`: Module containing form display functions
- Documentation in the `docs/` folder:
  - [`docs/technical_documentation_eng.md`](docs/technical_documentation_eng.md) - Technical documentation in English
 - [`docs/technical_documentation_rus.md`](docs/technical_documentation_rus.md) - Technical documentation in Russian
 - [`docs/user_guide_eng.md`](docs/user_guide_eng.md) - User guide in English
 - [`docs/user_guide_rus.md`](docs/user_guide_rus.md) - User guide in Russian
  - [`docs/implementation_examples_eng.md`](docs/implementation_examples_eng.md) - Implementation examples in English
  - [`docs/implementation_examples_rus.md`](docs/implementation_examples_rus.md) - Implementation examples in Russian

## Installation

1. Download the `clsMultiStateCheckBox.cls` file from the `vba-files/Class/` directory
2. Import the class into your VBA project
3. Ensure you have the Microsoft Forms 2.0 Object Library referenced in your project

## Quick Start

### Simple Usage Example
```vba
' Create an instance of clsMultiStateCheckBox class
Set MultiStateCheckBox = New clsMultiStateCheckBox

' Initialize the checkbox using a Label control
Call MultiStateCheckBox.Initialize(Me.Label1)

' The class automatically applies multi-state functionality to the control
```

## Main Functions

- **Checkbox Initialization**: The `Initialize` method allows setting the initial state and configuring the checkbox appearance
- **State Management**: Support for three or more states (unchecked, checked, indeterminate and others)
- **Color Configuration**: Ability to configure colors for each state
- **Icon Support**: Using icons for various checkbox states
- **Cyclic Switching**: Ability to enable/disable cyclic switching between states

## Working with Controls

The `clsMultiStateCheckBox` class converts a regular Label control into a multi-state checkbox with capabilities:
- Setting initial state
- Configuring appearance for each state
- Handling click events for state switching
- Accessing current checkbox state

## Style Configuration

The class allows customization of:
- Colors for each checkbox state
- Icons for each state
- Fonts and text sizes
- Interaction behavior (cyclic behavior, reset to initial state)

Example of color configuration:
```vba
' Configure colors for different states
Dim colors As Variant
colors = Array(vbRed, vbGreen, vbBlue) ' Colors for each state
MultiStateCheckBox.Initialize Me.Label1, 0, Array(59193, 59194, 59195), colors
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

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.

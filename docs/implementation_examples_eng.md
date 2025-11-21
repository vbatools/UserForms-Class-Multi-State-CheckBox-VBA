# Implementation Examples for clsMultiStateCheckBox

## Table of Contents
1. [Basic Implementation](#basic-implementation)
2. [Advanced Configuration](#advanced-configuration)
3. [Custom Styling](#custom-styling)
4. [Event Handling](#event-handling)
5. [Integration with Forms](#integration-with-forms)
6. [Dynamic Checkbox Creation](#dynamic-checkbox-creation)
7. [Working with Multiple Checkboxes](#working-with-multiple-checkboxes)
8. [Advanced Examples](#advanced-examples)

## Basic Implementation

### Simple Checkbox Usage
```vba
' In the form module
Dim MultiStateCheckBox As clsMultiStateCheckBox

Private Sub UserForm_Initialize()
    ' Create an instance of the checkbox class
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Initialize the checkbox with default parameters
    Call MultiStateCheckBox.Initialize(Me.Label1)
End Sub
```

### Checkbox with State Display
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Initialize with visible state label
    Call MultiStateCheckBox.Initialize(Me.Label1)
    
    ' Update a label with current state
    LabelState.Caption = "Current State: " & MultiStateCheckBox.Item
End Sub

Private Sub MultiStateCheckBox_Click()
    ' Update when checkbox state changes
    LabelState.Caption = "Current State: " & MultiStateCheckBox.Item & " - " & MultiStateCheckBox.StateText
End Sub
```

## Advanced Configuration

### Configuring Colors and Appearance
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Initialize with custom icons and colors
    Dim icons As Variant
    icons = Array(59193, 59194, 59195) ' Unchecked, checked, indeterminate
    Dim colors As Variant
    colors = Array(vbRed, vbGreen, vbBlue) ' Colors for each state
    
    Call MultiStateCheckBox.Initialize(Me.Label1, 0, icons, colors)
        
    ' Further customize appearance
    With MultiStateCheckBox
        .SetStateColor 1, RGB(0, 150, 0)  ' Green for checked state
        .SetStateColor 2, RGB(150, 150, 150)  ' Gray for indeterminate state
    End With
End Sub
```

### Cyclic vs Non-Cyclic Behavior
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Initialize with cyclic behavior (default) - cycles through states
    Call MultiStateCheckBox.Initialize(Me.Label1)
    MultiStateCheckBox.Cyclic = True
    
    ' Or non-cyclic behavior - stops at the last state
    ' MultiStateCheckBox.Cyclic = False
End Sub
```

## Custom Styling

### Custom State Icons
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Initialize with custom state icons
    Dim customIcons As Variant
    customIcons = Array(59193, 59194, 59195, 59634) ' Four different states
    Call MultiStateCheckBox.Initialize(Me.Label1, 0, customIcons)
    
    ' This will display four different states
End Sub
```

### Checkbox with Custom Font
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Initialize with custom font settings
    Call MultiStateCheckBox.Initialize(Me.Label1, 0, , , "Wingdings", 0.8)
End Sub
```

## Event Handling

### Handling State Changes
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    Call MultiStateCheckBox.Initialize(Me.Label1)
End Sub

Private Sub MultiStateCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
    ' Handle click event
    Debug.Print "Checkbox clicked, current state: " & Item & ", Text: " & StateText
    
    ' Perform actions based on state
    Select Case Item
        Case 0
            Debug.Print "Unchecked state"
        Case 1
            Debug.Print "Checked state"
        Case 2
            Debug.Print "Indeterminate state"
    End Select
End Sub
```

### Checkbox Locking
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    Call MultiStateCheckBox.Initialize(Me.Label1)
    
    ' Initially lock the checkbox
    MultiStateCheckBox.Locked = True
End Sub

Private Sub LockCheckboxButton_Click()
    ' Toggle checkbox lock state
    MultiStateCheckBox.Locked = Not MultiStateCheckBox.Locked
    If MultiStateCheckBox.Locked Then
        LockButton.Caption = "Unlock Checkbox"
    Else
        LockButton.Caption = "Lock Checkbox"
    End If
End Sub
```

## Integration with Forms

### Complete Form Example
```vba
Dim AgreementCheckbox As clsMultiStateCheckBox
Dim ConsentCheckbox As clsMultiStateCheckBox

Private Sub UserForm_Initialize()
    ' Initialize agreement checkbox
    Set AgreementCheckbox = New clsMultiStateCheckBox
    Dim agreementIcons As Variant
    agreementIcons = Array(59193, 59194, 59195) ' Unchecked, checked, indeterminate
    Call AgreementCheckbox.Initialize(Me.AgreementLabel, 0, agreementIcons)
    
    ' Initialize consent checkbox
    Set ConsentCheckbox = New clsMultiStateCheckBox
    Dim consentIcons As Variant
    consentIcons = Array(59193, 59194) ' Just unchecked and checked
    Call ConsentCheckbox.Initialize(Me.ConsentLabel, 0, consentIcons)
    
    ' Update form controls with initial states
    UpdateAgreementDisplay
    UpdateConsentDisplay
End Sub

Private Sub UpdateAgreementDisplay()
    AgreementStatusLabel.Caption = "Agreement: " & GetStateDescription(AgreementCheckbox.Item)
    ' Enable/disable submit button based on agreement state
    SubmitButton.Enabled = (AgreementCheckbox.Item = 1)
End Sub

Private Sub UpdateConsentDisplay()
    ConsentStatusLabel.Caption = "Consent: " & GetStateDescription(ConsentCheckbox.Item)
End Sub

Private Function GetStateDescription(stateIndex As Byte) As String
    Select Case stateIndex
        Case 0: GetStateDescription = "Unchecked"
        Case 1: GetStateDescription = "Checked"
        Case 2: GetStateDescription = "Indeterminate"
        Case Else: GetStateDescription = "Unknown"
    End Select
End Function

Private Sub AgreementLabel_Click()
    UpdateAgreementDisplay
End Sub

Private Sub ConsentLabel_Click()
    UpdateConsentDisplay
End Sub
```

## Dynamic Checkbox Creation

### Creating Checkboxes Programmatically
```vba
Private Sub UserForm_Initialize()
    ' Create multiple checkboxes dynamically
    Dim i As Integer
    Dim newLabel As MSForms.Label
    Dim newCheckbox As clsMultiStateCheckBox
    
    For i = 1 To 3
        ' Create a new label for the checkbox
        Set newLabel = Me.Controls.Add("Forms.Label.1", "DynamicCheckbox" & i, True)
        With newLabel
            .Left = 20
            .Top = 30 + (i - 1) * 40
            .Width = 20
            .Height = 20
            .BackColor = RGB(240, 240, 240)
            .Caption = ""
        End With
        
        ' Create and initialize checkbox
        Set newCheckbox = New clsMultiStateCheckBox
        Call newCheckbox.Initialize(newLabel, 0, Array(59193, 59194, 59195))
        
        ' Store reference in a collection for later access
        CheckboxCollection.Add newCheckbox
    Next i
End Sub
```

## Working with Multiple Checkboxes

### Managing Multiple Checkboxes
```vba
Dim CheckboxCollection As Collection

Private Sub UserForm_Initialize()
    Set CheckboxCollection = New Collection
    
    ' Initialize multiple checkboxes
    InitializeCheckboxes
End Sub

Private Sub InitializeCheckboxes()
    Dim Task1Checkbox As clsMultiStateCheckBox
    Dim Task2Checkbox As clsMultiStateCheckBox
    Dim Task3Checkbox As clsMultiStateCheckBox
    
    ' Create task checkboxes
    Set Task1Checkbox = New clsMultiStateCheckBox
    Set Task2Checkbox = New clsMultiStateCheckBox
    Set Task3Checkbox = New clsMultiStateCheckBox
    
    ' Initialize with different parameters
    Call Task1Checkbox.Initialize(Me.Task1Label, 0, Array(59193, 59194, 59195)) ' Unchecked, checked, indeterminate
    Call Task2Checkbox.Initialize(Me.Task2Label, 0, Array(59193, 59194))        ' Unchecked, checked
    Call Task3Checkbox.Initialize(Me.Task3Label, 0, Array(59193, 59194, 59195)) ' Unchecked, checked, indeterminate
    
    ' Add to collection
    CheckboxCollection.Add Task1Checkbox
    CheckboxCollection.Add Task2Checkbox
    CheckboxCollection.Add Task3Checkbox
    
    ' Update task summary
    UpdateTaskSummary
End Sub

Private Sub UpdateTaskSummary()
    ' Count completed tasks
    Dim completedTasks As Integer
    Dim i As Integer
    
    For i = 1 To CheckboxCollection.Count
        If CheckboxCollection(i).Item = 1 Then ' Checked state
            completedTasks = completedTasks + 1
        End If
    Next i
    
    ' Update summary label
    TaskSummaryLabel.Caption = "Completed Tasks: " & completedTasks & "/" & CheckboxCollection.Count
End Sub

Private Sub Task1Label_Click()
    UpdateTaskSummary
End Sub

Private Sub Task2Label_Click()
    UpdateTaskSummary
End Sub

Private Sub Task3Label_Click()
    UpdateTaskSummary
End Sub
```

## Advanced Examples

### Checkbox with Animation Effect
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    Call MultiStateCheckBox.Initialize(Me.Label1)
End Sub

Private Sub AnimateCheckboxEffect()
    Dim originalColor As Long
    originalColor = MultiStateCheckBox.CurrentIcon
    
    ' Create animation effect by changing color temporarily
    MultiStateCheckBox.SetStateColor 1, RGB(255, 255, 0) ' Yellow for checked state
    
    ' Wait briefly to show effect
    Application.Wait (Now + TimeValue("0:00:0.2"))
    
    ' Restore original color
    MultiStateCheckBox.SetStateColor 1, RGB(0, 0, 0) ' Black for checked state
End Sub

Private Sub MultiStateCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
    ' Trigger animation on click
    AnimateCheckboxEffect
End Sub
```

### Responsive Checkbox Panel
```vba
Private Sub UserForm_Resize()
    ' Adjust checkbox positions when form resizes
    AdjustCheckboxLayout
End Sub

Private Sub AdjustCheckboxLayout()
    If Not MultiStateCheckBox Is Nothing Then
        With MultiStateCheckBox
            ' Adjust checkbox to maintain proper positioning
            .control.Width = 20
            .control.Height = 20
            .control.Left = Me.Width - 50
            .control.Top = 30
        End With
    End If
End Sub
```

These examples demonstrate various ways to implement and use the `clsMultiStateCheckBox` class in VBA UserForms. Each example can be adapted to specific application requirements.
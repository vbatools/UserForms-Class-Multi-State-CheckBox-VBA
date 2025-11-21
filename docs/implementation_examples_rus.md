# Примеры реализации для clsMultiStateCheckBox

## Содержание
1. [Базовая реализация](#базовая-реализация)
2. [Расширенная настройка](#расширенная-настройка)
3. [Пользовательский стиль](#пользовательский-стиль)
4. [Обработка событий](#обработка-событий)
5. [Интеграция с формами](#интеграция-с-формами)
6. [Динамическое создание флажков](#динамическое-создание-флажков)
7. [Работа с несколькими флажками](#работа-с-несколькими-флажками)
8. [Продвинутые примеры](#продвинутые-примеры)

## Базовая реализация

### Простое использование флажка
```vba
' В модуле формы
Dim MultiStateCheckBox As clsMultiStateCheckBox

Private Sub UserForm_Initialize()
    ' Создание экземпляра класса флажка
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Инициализация флажка с параметрами по умолчанию
    Call MultiStateCheckBox.Initialize(Me.Label1)
End Sub
```

### Флажок с отображением состояния
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Инициализация с видимой меткой состояния
    Call MultiStateCheckBox.Initialize(Me.Label1)
    
    ' Обновление метки с текущим состоянием
    LabelState.Caption = "Текущее состояние: " & MultiStateCheckBox.Item
End Sub

Private Sub MultiStateCheckBox_Click()
    ' Обновление при изменении состояния флажка
    LabelState.Caption = "Текущее состояние: " & MultiStateCheckBox.Item & " - " & MultiStateCheckBox.StateText
End Sub
```

## Расширенная настройка

### Настройка цветов и внешнего вида
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Инициализация с пользовательскими значками и цветами
    Dim icons As Variant
    icons = Array(59193, 59194, 59195) ' Снят, установлен, неопределён
    Dim colors As Variant
    colors = Array(vbRed, vbGreen, vbBlue) ' Цвета для каждого состояния
    
    Call MultiStateCheckBox.Initialize(Me.Label1, 0, icons, colors)
        
    ' Дальнейшая настройка внешнего вида
    With MultiStateCheckBox
        .SetStateColor 1, RGB(0, 150, 0)  ' Зеленый для состояния установки
        .SetStateColor 2, RGB(150, 150, 150)  ' Серый для неопределенного состояния
    End With
End Sub
```

### Циклическое и нециклическое поведение
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Инициализация с циклическим поведением (по умолчанию) - переключается между состояниями
    Call MultiStateCheckBox.Initialize(Me.Label1)
    MultiStateCheckBox.Cyclic = True
    
    ' Или нециклическое поведение - останавливается на последнем состоянии
    ' MultiStateCheckBox.Cyclic = False
End Sub
```

## Пользовательский стиль

### Пользовательские значки состояний
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Инициализация с пользовательскими значками состояний
    Dim customIcons As Variant
    customIcons = Array(59193, 59194, 59195, 59634) ' Четыре разных состояния
    Call MultiStateCheckBox.Initialize(Me.Label1, 0, customIcons)
    
    ' Это будет отображать четыре разных состояния
End Sub
```

### Флажок с пользовательским шрифтом
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    
    ' Инициализация с пользовательскими настройками шрифта
    Call MultiStateCheckBox.Initialize(Me.Label1, 0, , , "Wingdings", 0.8)
End Sub
```

## Обработка событий

### Обработка изменений состояния
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    Call MultiStateCheckBox.Initialize(Me.Label1)
End Sub

Private Sub MultiStateCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
    ' Обработка события клика
    Debug.Print "Флажок нажат, текущее состояние: " & Item & ", Текст: " & StateText
    
    ' Выполнение действий в зависимости от состояния
    Select Case Item
        Case 0
            Debug.Print "Состояние снятия"
        Case 1
            Debug.Print "Состояние установки"
        Case 2
            Debug.Print "Неопределенное состояние"
    End Select
End Sub
```

### Блокировка флажка
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    Call MultiStateCheckBox.Initialize(Me.Label1)
    
    ' Изначально заблокировать флажок
    MultiStateCheckBox.Locked = True
End Sub

Private Sub LockCheckboxButton_Click()
    ' Переключение состояния блокировки флажка
    MultiStateCheckBox.Locked = Not MultiStateCheckBox.Locked
    If MultiStateCheckBox.Locked Then
        LockButton.Caption = "Разблокировать флажок"
    Else
        LockButton.Caption = "Заблокировать флажок"
    End If
End Sub
```

## Интеграция с формами

### Полный пример формы
```vba
Dim AgreementCheckbox As clsMultiStateCheckBox
Dim ConsentCheckbox As clsMultiStateCheckBox

Private Sub UserForm_Initialize()
    ' Инициализация флажка соглашения
    Set AgreementCheckbox = New clsMultiStateCheckBox
    Dim agreementIcons As Variant
    agreementIcons = Array(59193, 59194, 59195) ' Снят, установлен, неопределён
    Call AgreementCheckbox.Initialize(Me.AgreementLabel, 0, agreementIcons)
    
    ' Инициализация флажка согласия
    Set ConsentCheckbox = New clsMultiStateCheckBox
    Dim consentIcons As Variant
    consentIcons = Array(59193, 59194) ' Только снят и установлен
    Call ConsentCheckbox.Initialize(Me.ConsentLabel, 0, consentIcons)
    
    ' Обновление элементов управления формы начальными состояниями
    UpdateAgreementDisplay
    UpdateConsentDisplay
End Sub

Private Sub UpdateAgreementDisplay()
    AgreementStatusLabel.Caption = "Соглашение: " & GetStateDescription(AgreementCheckbox.Item)
    ' Включение/отключение кнопки отправки на основе состояния соглашения
    SubmitButton.Enabled = (AgreementCheckbox.Item = 1)
End Sub

Private Sub UpdateConsentDisplay()
    ConsentStatusLabel.Caption = "Согласие: " & GetStateDescription(ConsentCheckbox.Item)
End Sub

Private Function GetStateDescription(stateIndex As Byte) As String
    Select Case stateIndex
        Case 0: GetStateDescription = "Снят"
        Case 1: GetStateDescription = "Установлен"
        Case 2: GetStateDescription = "Неопределён"
        Case Else: GetStateDescription = "Неизвестно"
    End Select
End Function

Private Sub AgreementLabel_Click()
    UpdateAgreementDisplay
End Sub

Private Sub ConsentLabel_Click()
    UpdateConsentDisplay
End Sub
```

## Динамическое создание флажков

### Создание флажков программно
```vba
Private Sub UserForm_Initialize()
    ' Создание нескольких флажков динамически
    Dim i As Integer
    Dim newLabel As MSForms.Label
    Dim newCheckbox As clsMultiStateCheckBox
    
    For i = 1 To 3
        ' Создание новой метки для флажка
        Set newLabel = Me.Controls.Add("Forms.Label.1", "DynamicCheckbox" & i, True)
        With newLabel
            .Left = 20
            .Top = 30 + (i - 1) * 40
            .Width = 20
            .Height = 20
            .BackColor = RGB(240, 240, 240)
            .Caption = ""
        End With
        
        ' Создание и инициализация флажка
        Set newCheckbox = New clsMultiStateCheckBox
        Call newCheckbox.Initialize(newLabel, 0, Array(59193, 59194, 59195))
        
        ' Сохранение ссылки в коллекции для последующего доступа
        CheckboxCollection.Add newCheckbox
    Next i
End Sub
```

## Работа с несколькими флажками

### Управление несколькими флажками
```vba
Dim CheckboxCollection As Collection

Private Sub UserForm_Initialize()
    Set CheckboxCollection = New Collection
    
    ' Инициализация нескольких флажков
    InitializeCheckboxes
End Sub

Private Sub InitializeCheckboxes()
    Dim Task1Checkbox As clsMultiStateCheckBox
    Dim Task2Checkbox As clsMultiStateCheckBox
    Dim Task3Checkbox As clsMultiStateCheckBox
    
    ' Создание флажков задач
    Set Task1Checkbox = New clsMultiStateCheckBox
    Set Task2Checkbox = New clsMultiStateCheckBox
    Set Task3Checkbox = New clsMultiStateCheckBox
    
    ' Инициализация с разными параметрами
    Call Task1Checkbox.Initialize(Me.Task1Label, 0, Array(59193, 59194, 59195)) ' Снят, установлен, неопределён
    Call Task2Checkbox.Initialize(Me.Task2Label, 0, Array(59193, 59194))        ' Снят, установлен
    Call Task3Checkbox.Initialize(Me.Task3Label, 0, Array(59193, 59194, 59195)) ' Снят, установлен, неопределён
    
    ' Добавление в коллекцию
    CheckboxCollection.Add Task1Checkbox
    CheckboxCollection.Add Task2Checkbox
    CheckboxCollection.Add Task3Checkbox
    
    ' Обновление сводки задач
    UpdateTaskSummary
End Sub

Private Sub UpdateTaskSummary()
    ' Подсчет выполненных задач
    Dim completedTasks As Integer
    Dim i As Integer
    
    For i = 1 To CheckboxCollection.Count
        If CheckboxCollection(i).Item = 1 Then ' Состояние установки
            completedTasks = completedTasks + 1
        End If
    Next i
    
    ' Обновление метки сводки
    TaskSummaryLabel.Caption = "Выполненные задачи: " & completedTasks & "/" & CheckboxCollection.Count
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

## Продвинутые примеры

### Флажок с эффектом анимации
```vba
Private Sub UserForm_Initialize()
    Set MultiStateCheckBox = New clsMultiStateCheckBox
    Call MultiStateCheckBox.Initialize(Me.Label1)
End Sub

Private Sub AnimateCheckboxEffect()
    Dim originalColor As Long
    originalColor = MultiStateCheckBox.CurrentIcon
    
    ' Создание эффекта анимации путем временного изменения цвета
    MultiStateCheckBox.SetStateColor 1, RGB(255, 255, 0) ' Желтый для состояния установки
    
    ' Краткое ожидание для отображения эффекта
    Application.Wait (Now + TimeValue("0:00:0.2"))
    
    ' Восстановление исходного цвета
    MultiStateCheckBox.SetStateColor 1, RGB(0, 0, 0) ' Черный для состояния установки
End Sub

Private Sub MultiStateCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
    ' Запуск анимации при клике
    AnimateCheckboxEffect
End Sub
```

### Адаптивная панель флажков
```vba
Private Sub UserForm_Resize()
    ' Настройка позиций флажков при изменении размера формы
    AdjustCheckboxLayout
End Sub

Private Sub AdjustCheckboxLayout()
    If Not MultiStateCheckBox Is Nothing Then
        With MultiStateCheckBox
            ' Настройка флажка для сохранения правильного позиционирования
            .control.Width = 20
            .control.Height = 20
            .control.Left = Me.Width - 50
            .control.Top = 30
        End With
    End If
End Sub
```

Эти примеры демонстрируют различные способы реализации и использования класса `clsMultiStateCheckBox` в VBA UserForms. Каждый пример может быть адаптирован под конкретные требования приложения.
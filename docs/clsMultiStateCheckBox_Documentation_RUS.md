# Документация по классу clsMultiStateCheckBox

## Содержание
1. [Обзор](#обзор)
2. [Системные требования](#системные-требования)
3. [Установка и настройка](#установка-и-настройка)
4. [Справочник по API](#справочник-по-api)
5. [Руководства по реализации](#руководства-по-реализации)
6. [Рассмотрение производительности](#рассмотрение-производительности)
7. [Стратегии обработки ошибок](#стратегии-обработки-ошибок)
8. [Примеры интеграции](#примеры-интеграции)
9. [Варианты настройки](#варианты-настройки)
10. [История версий](#история-версий)
11. [Диаграммы переходов состояний](#диаграммы-переходов-состояний)
12. [Примеры проектов](#примеры-проектов)
13. [Методы тестирования](#методы-тестирования)
14. [Совместимость с версиями Office](#совместимость-с-версиями-office)
15. [Рассмотрение безопасности](#рассмотрение-безопасности)

## Обзор

Класс `clsMultiStateCheckBox` - это реализация многопозиционного флажка для VBA UserForms с тремя и более состояниями: снят, установлен и неопределён, а также другими пользовательскими состояниями. Этот класс предоставляет настраиваемый внешний вид и поведение с возможностью легкой интеграции в VBA UserForms.

### Основные возможности
- Трехпозиционный флажок (снят, установлен, неопределён)
- Настраиваемый внешний вид и поведение
- Простая интеграция с VBA UserForms
- Повторно используемый модуль класса для согласованного поведения флажка
- Циклическое/непрерывное переключение состояний (свойство `Cyclic`)
- Возможность установки/получения состояния по тексту (свойство `StateText`)
- Метод для получения всех доступных состояний (`GetAllStates`)
- Улучшенная обработка ошибок и проверки
- Возможность установки имени шрифта и коэффициента размера (свойства `FontName` и `FontSizeFactor`)
- Возможность получения/установки текущего значка (свойство `CurrentIcon`)
- Метод для установки цвета для конкретного состояния (`SetStateColor`)
- Метод для сброса в начальное состояние (`ResetToInitialState`)

## Системные требования

- Microsoft Excel (2010 или выше)
- Включенные макросы
- Разрешения VBA (уровень средней или низкой безопасности)

## Установка и настройка

1. Импортируйте файл `clsMultiStateCheckBox.cls` в ваш VBA проект
2. Добавьте элемент управления Label на вашу форму, который будет служить флажком
3. Объявите переменную класса в вашей форме

```vba
Private mCheckBox As clsMultiStateCheckBox
```

4. Инициализируйте элемент управления в событии `UserForm_Initialize`

```vba
Private Sub UserForm_Initialize()
    Set mCheckBox = New clsMultiStateCheckBox
    mCheckBox.Initialize Me.Label1 ' Замените Label1 на ваш элемент управления Label
End Sub
```

## Справочник по API

### События

#### Click
Событие, возникающее при клике по многопозиционному флажку

**Синтаксис:**
```vba
Private Sub mCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
```

**Параметры:**
- `control` - Элемент управления Label, который был нажат
- `Item` - Текущий индекс состояния (начиная с 0)
- `CodeIcon` - Код символа текущего значка
- `StateText` - Текстовое представление текущего состояния

### Свойства

#### Version
Получает информацию о версии класса clsMultiStateCheckBox

**Тип:** String
**Доступ:** Только чтение

#### StateText
Получает или устанавливает текст текущего состояния

**Тип:** String
**Доступ:** Чтение/запись

**Пример использования:**
```vba
' Получение текущего состояния
Dim currentState As String
currentState = mCheckBox.StateText

' Установка состояния
mCheckBox.StateText = ChrW$(59194) ' Установка во второе состояние
```

#### Item
Получает или устанавливает текущий индекс состояния

**Тип:** Byte
**Доступ:** Чтение/запись

#### Count
Получает количество доступных состояний

**Тип:** Byte
**Доступ:** Только чтение

#### ArrayIcon
Получает или устанавливает массив значков для каждого состояния

**Тип:** Variant
**Доступ:** Чтение/запись

#### Cyclic
Получает или устанавливает, будет ли флажок циклически переключаться между состояниями

**Тип:** Boolean
**Доступ:** Чтение/запись

**Пример использования:**
```vba
' Отключить цикличность
mCheckBox.Cyclic = False
```

#### FontName
Получает или устанавливает имя шрифта для значка

**Тип:** String
**Доступ:** Чтение/запись

**Пример использования:**
```vba
' Установить шрифт
mCheckBox.FontName = "Wingdings"
```

#### FontSizeFactor
Получает или устанавливает коэффициент размера шрифта для масштабирования

**Тип:** Double
**Доступ:** Чтение/запись

**Пример использования:**
```vba
' Установить размер шрифта 80% от ширины
mCheckBox.FontSizeFactor = 0.8
```

#### CurrentIcon
Получает или устанавливает текущий значок по коду

**Тип:** Long
**Доступ:** Чтение/запись

**Пример использования:**
```vba
' Получить текущий значок
Dim currentIcon As Long
currentIcon = mCheckBox.CurrentIcon

' Установить текущий значок
mCheckBox.CurrentIcon = 59194
```

### Методы

#### Initialize
Инициализирует элемент управления многопозиционного флажка

**Синтаксис:**
```vba
Public Sub Initialize(ByRef LabelBtn As MSForms.Label, _
        Optional Item As Byte, _
        Optional arrIcon As Variant, _
        Optional ArrIconColor As Variant, _
        Optional fontName As String = "Segoe MDL2 Assets", _
        Optional fontSizeFactor As Double = 0.9)
```

**Параметры:**
- `LabelBtn` - Элемент управления Label, который будет использоваться как флажок
- `Item` - Начальное состояние (0 для снятого, 1 для установленного, 2 для неопределенного и т.д.)
- `arrIcon` - Массив кодов символов значков (по умолчанию Array(59193, 59194, 59195))
- `ArrIconColor` - Массив цветов для каждого состояния (по умолчанию Array(vbBlack))
- `fontName` - Имя шрифта для значков (по умолчанию "Segoe MDL2 Assets")
- `fontSizeFactor` - Коэффициент размера шрифта (по умолчанию 0.9)

**Пример использования:**
```vba
Dim customIcons As Variant
customIcons = Array(59193, 59194, 59195) ' Снят, установлен, неопределен

Dim customColors As Variant
customColors = Array(vbBlack, vbBlue, vbRed) ' Цвета для каждого состояния

mCheckBox.Initialize Me.Label1, 0, customIcons, customColors, "Segoe MDL2 Assets", 0.8
```

#### UpdateIconColors
Обновляет цвета значков после инициализации

**Синтаксис:**
```vba
Public Sub UpdateIconColors(ArrIconColor As Variant)
```

**Параметры:**
- `ArrIconColor` - Массив цветов для каждого состояния

#### SetStateColor
Устанавливает цвет для конкретного состояния

**Синтаксис:**
```vba
Public Sub SetStateColor(StateIndex As Byte, Color As Long)
```

**Параметры:**
- `StateIndex` - Индекс состояния (начиная с 0)
- `Color` - Цвет для установки

**Пример использования:**
```vba
' Установить красный цвет для второго состояния
mCheckBox.SetStateColor 1, RGB(255, 0, 0)
```

#### ResetToInitialState
Сбрасывает флажок в начальное состояние

**Синтаксис:**
```vba
Public Sub ResetToInitialState(InitialItem As Byte)
```

**Параметры:**
- `InitialItem` - Индекс начального состояния

**Пример использования:**
```vba
' Сбросить флажок в первое состояние
mCheckBox.ResetToInitialState 0
```

#### GetAllStates
Возвращает массив всех доступных текстовых состояний

**Синтаксис:**
```vba
Public Function GetAllStates() As Variant
```

**Возвращаемое значение:**
- Массив строк, представляющих все доступные состояния

**Пример использования:**
```vba
Dim allStates As Variant
allStates = mCheckBox.GetAllStates
```

## Руководства по реализации

### Базовая реализация

Самый простой способ использовать многопозиционный флажок:

1. Добавьте элемент управления Label на вашу форму
2. Объявите переменную класса:
```vba
Private mCheckBox As clsMultiStateCheckBox
```

3. Инициализируйте элемент управления:
```vba
Private Sub UserForm_Initialize()
    Set mCheckBox = New clsMultiStateCheckBox
    mCheckBox.Initialize Me.Label1
End Sub
```

4. Обработайте событие Click:
```vba
Private Sub mCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
    Debug.Print "Состояние: " & Item & ", Текст: " & StateText
End Sub
```

### Расширенная реализация с пользовательскими значками

Для настройки внешнего вида флажка:

1. Определите массив значков:
```vba
Dim customIcons As Variant
customIcons = Array(59193, 59194, 59195) ' Коды символов Unicode
```

2. Определите массив цветов:
```vba
Dim customColors As Variant
customColors = Array(vbBlack, vbBlue, vbRed)
```

3. Инициализируйте элемент управления с пользовательскими параметрами:
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

### Реализация с отключенной цикличностью

Для создания прогрессивного индикатора, который не возвращается к первому состоянию:

```vba
Private Sub UserForm_Initialize()
    Set mCheckBox = New clsMultiStateCheckBox
    mCheckBox.Initialize Me.Label1
    mCheckBox.Cyclic = False ' Отключить цикличность
End Sub
```

## Рассмотрение производительности

### Оптимизация при инициализации

- Используйте минимально необходимое количество состояний для вашей задачи
- Используйте стандартные шрифты, если не требуется специальная типографика
- Инициализируйте элемент управления один раз в событии UserForm_Initialize

### Оптимизация при обновлении

- Избегайте частых вызовов UpdateIconColors, если не требуется
- Используйте SetStateColor только когда нужно изменить цвет одного состояния
- Минимизируйте количество обновлений внешнего вида в циклах

## Стратегии обработки ошибок

### Проверка параметров

Класс включает встроенную проверку параметров:

- Проверка, что LabelBtn не равен Nothing
- Проверка, что массивы значков и цветов являются массивами
- Проверка границ массивов
- Проверка корректности индексов состояний

### Обработка ошибок

Класс генерирует следующие ошибки:

- `vbObjectError + 100`: LabelBtn равен Nothing
- `vbObjectError + 1001`: arrIcon не является массивом
- `vbObjectError + 1002`: arrIcon пустой
- `vbObjectError + 1003`: ArrIconColor не является массивом
- `vbObjectError + 1004`: ArrIconColor пустой
- `vbObjectError + 105`: ArrIconColor не является массивом (в UpdateIconColors)
- `vbObjectError + 1006`: StateText не найден в массиве значков
- `vbObjectError + 1007`: CurrentIcon не найден в массиве значков
- `vbObjectError + 107`: StateIndex превышает количество доступных состояний

### Рекомендуемая обработка ошибок

```vba
Private Sub InitializeCheckBox()
    On Error GoTo ErrorHandler
    
    Set mCheckBox = New clsMultiStateCheckBox
    mCheckBox.Initialize Me.Label1
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при инициализации флажка: " & Err.Description, vbCritical
End Sub
```

## Примеры интеграции

### Интеграция с UserForm

```vba
' В модуле формы
Private mCheckBox As clsMultiStateCheckBox

Private Sub UserForm_Initialize()
    Set mCheckBox = New clsMultiStateCheckBox
    
    ' Инициализация с пользовательскими значками и цветами
    Dim icons As Variant
    icons = Array(59193, 59194, 59195) ' Снят, установлен, неопределен
    
    Dim colors As Variant
    colors = Array(vbRed, vbGreen, vbBlue) ' Красный, зеленый, синий
    
    mCheckBox.Initialize Me.Label1, 0, icons, colors
End Sub

Private Sub mCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
    Select Case Item
        Case 0
            MsgBox "Флажок снят"
        Case 1
            MsgBox "Флажок установлен"
        Case 2
            MsgBox "Флажок в неопределенном состоянии"
    End Select
End Sub
```

### Интеграция с UserForm с несколькими флажками

```vba
' В модуле формы
Private mCheckBox1 As clsMultiStateCheckBox
Private mCheckBox2 As clsMultiStateCheckBox

Private Sub UserForm_Initialize()
    ' Первый флажок
    Set mCheckBox1 = New clsMultiStateCheckBox
    mCheckBox1.Initialize Me.Label1
    
    ' Второй флажок с другими параметрами
    Set mCheckBox2 = New clsMultiStateCheckBox
    Dim icons As Variant
    icons = Array(59193, 59194) ' Только два состояния
    mCheckBox2.Initialize Me.Label2, 0, icons
End Sub
```

## Варианты настройки

### Настройка внешнего вида

Класс позволяет настраивать:

- Значки для каждого состояния
- Цвета для каждого состояния
- Шрифт значков
- Размер шрифта
- Цикличность переключения

### Примеры настройки

```vba
' Настройка значков и цветов
Dim icons As Variant
icons = Array(59193, 59194, 59195, 59634) ' Четыре состояния

Dim colors As Variant
colors = Array(vbRed, vbGreen, vbBlue, vbYellow) ' Цвета для каждого состояния

mCheckBox.Initialize Me.Label1, 0, icons, colors, "Wingdings", 0.7

' Изменение шрифта после инициализации
mCheckBox.FontName = "Webdings"
mCheckBox.FontSizeFactor = 0.85

' Установка цвета для конкретного состояния
mCheckBox.SetStateColor 2, RGB(128, 0, 128) ' Фиолетовый для третьего состояния
```

## История версий

### 1.0.4
- Добавлено свойство CurrentIcon для получения/установки текущего значка
- Добавлен метод SetStateColor для установки цвета для конкретного состояния
- Добавлен метод ResetToInitialState для сброса в начальное состояние
- Добавлены свойства FontName и FontSizeFactor для настройки шрифта
- Улучшена обработка ошибок

### 1.0.3
- Исправлен метод BlendColor для корректной обработки цветов в формате BGR
- Добавлена валидация входных данных в метод Initialize
- Улучшена обработка ошибок во всех методах
- Добавлено свойство StateText для установки состояния по тексту
- Добавлено свойство Cyclic для управления цикличностью состояний
- Добавлен метод GetAllStates для получения всех доступных состояний
- Улучшен метод UpdateIconColors с проверками корректности данных

### 1.0.2
- Добавлены значки по умолчанию для трех состояний
- Реализовано переключение между состояниями по клику
- Добавлено событие Click
- Добавлена поддержка пользовательских цветов

### 1.0.1
- Базовая реализация класса
- Поддержка трех состояний (снят, установлен, неопределен)
- Использование шрифта Segoe MDL2 Assets для значков

### 1.0.0
- Первоначальный выпуск
- Базовая функциональность многопозиционного флажка

## Диаграммы переходов состояний

### Циклический режим (Cyclic = True)
```
Состояние 0 → Состояние 1 → Состояние 2 → ... → Состояние N → Состояние 0
    ↑                                                                 ↓
    └─────────────────────────────────────────────────────────────────┘
```

### Нециклический режим (Cyclic = False)
```
Состояние 0 → Состояние 1 → Состояние 2 → ... → Состояние N
    ↑         ↑         ↑                    ↑
   Начало   Продолжение  Продолжение      Последнее (остается)
```

## Примеры проектов

### Проект: Управление задачами

Использование многопозиционного флажка для отслеживания статуса задач:

```vba
' Состояния задачи:
' 0 - Не начата (пустой квадрат)
' 1 - В процессе (часы)
' 2 - Завершена (галочка)
' 3 - Отменена (крестик)

Private Sub InitializeTaskCheckBox()
    Set mCheckBox = New clsMultiStateCheckBox
    
    ' Использование значков из шрифта Segoe MDL2 Assets
    Dim taskIcons As Variant
    taskIcons = Array(59193, 59244, 59194, 59226) ' Не начата, В процессе, Завершена, Отменена
    
    Dim taskColors As Variant
    taskColors = Array(vbGray, vbBlue, vbGreen, vbRed) ' Цвета для каждого статуса
    
    mCheckBox.Initialize Me.Label1, 0, taskIcons, taskColors
End Sub

Private Sub mCheckBox_Click(control As Object, Item As Byte, ByVal CodeIcon As Long, ByVal StateText As String)
    Dim statusText As String
    Select Case Item
        Case 0: statusText = "Не начата"
        Case 1: statusText = "В процессе"
        Case 2: statusText = "Завершена"
        Case 3: statusText = "Отменена"
    End Select
    
    MsgBox "Статус задачи изменен на: " & statusText
End Sub
```

## Методы тестирования

### Модульное тестирование

Для тестирования функциональности класса рекомендуется создать тестовую форму с несколькими экземплярами класса:

```vba
' Тестирование инициализации
Sub TestInitialization()
    Dim checkBox As clsMultiStateCheckBox
    Set checkBox = New clsMultiStateCheckBox
    Dim form As Object
    Set form = CreateObject("Forms.Label") ' Имитация элемента управления
    
    On Error Resume Next
    checkBox.Initialize form
    If Err.Number <> 0 Then
        Debug.Print "Ошибка инициализации: " & Err.Description
    Else
        Debug.Print "Инициализация прошла успешно"
    End If
    On Error GoTo 0
End Sub

' Тестирование переключения состояний
Sub TestStateCycling()
    Dim checkBox As clsMultiStateCheckBox
    Set checkBox = New clsMultiStateCheckBox
    Dim form As Object
    Set form = CreateObject("Forms.Label")
    
    checkBox.Initialize form
    checkBox.Cyclic = True
    
    ' Проверка циклического переключения
    Dim i As Integer
    For i = 0 To 5 ' Проверим 6 кликов
        checkBox.mLabelBtn_Click ' Имитация клика (если возможно)
        Debug.Print "Состояние после клика " & i + 1 & ": " & checkBox.Item
    Next i
End Sub
```

### Проверка валидации

```vba
' Тестирование валидации параметров
Sub TestValidation()
    Dim checkBox As clsMultiStateCheckBox
    Set checkBox = New clsMultiStateCheckBox
    Dim form As Object
    Set form = CreateObject("Forms.Label")
    
    ' Тест с недопустимым массивом значков
    On Error Resume Next
    checkBox.Initialize form, 0, "not an array"
    If Err.Number <> 0 Then
        Debug.Print "Корректно обработана ошибка с недопустимым массивом значков"
    End If
    On Error GoTo 0
End Sub
```

## Совместимость с версиями Office

### Поддерживаемые версии

- Microsoft Excel 2010 и выше
- Microsoft Word 2010 и выше (для форм с элементами управления)
- Microsoft Access 2010 и выше
- Microsoft PowerPoint 2010 и выше

### Особенности разных версий

#### Excel 2010-2013
- Полная поддержка всех функций класса
- Рекомендуется использовать шрифты, доступные в системе

#### Excel 2016/2019
- Полная поддержка всех функций класса
- Улучшенная поддержка шрифтов Segoe MDL2 Assets

#### Excel 365
- Полная поддержка всех функций класса
- Наилучшая поддержка современных шрифтов и значков

### Рекомендации по совместимости

- Используйте стандартные системные шрифты для лучшей совместимости
- Проверяйте наличие шрифта Segoe MDL2 Assets перед использованием
- Альтернативные шрифты: Wingdings, Webdings, Symbol

## Рассмотрение безопасности

### Макросы и безопасность

- Убедитесь, что уровень безопасности макросов позволяет запускать пользовательские классы
- Используйте цифровую подпись для VBA проектов в корпоративной среде
- Проверяйте права доступа к файлам и объектам системы

### Защита от неправильного использования

Класс включает встроенные проверки:

- Проверка на Nothing для передаваемых объектов
- Проверка типов для массивов параметров
- Проверка границ массивов
- Проверка допустимых значений индексов

### Рекомендации по безопасности

- Не передавайте непроверенные данные в методы класса
- Обрабатывайте ошибки при инициализации и использовании
- Используйте проверки безопасности в корпоративных приложениях
- Обновляйте класс при выявлении потенциальных уязвимостей

## Заключение

Класс clsMultiStateCheckBox предоставляет мощный и гибкий способ реализации многопозиционных флажков в VBA приложениях. С его помощью можно создавать интуитивно понятные интерфейсы с расширенной функциональностью по сравнению с традиционными флажками.

Класс поддерживает широкую настраиваемость, включая пользовательские значки, цвета, шрифты и режимы переключения. Благодаря встроенной валидации и обработке ошибок, класс является надежным решением для использования в производственных приложениях.
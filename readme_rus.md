# Multi-State CheckBox VBA Class

![Пример User Forms](User_Forms.gif)

Реализация VBA класса, который предоставляет элемент управления многопозиционным флажком с тремя и более состояниями: снят, установлен и неопределён и другие.

## Возможности

- Функциональность трехпозиционного флажка (снят, установлен, неопределён)
- Настраиваемый внешний вид и поведение
- Простая интеграция с VBA UserForms
- Повторно используемый модуль класса для согласованного поведения флажка

## Значки и цвета по умолчанию

- Значки по умолчанию: Использует шрифт Segoe MDL2 Assets с кодами символов 59193 (снят), 59194 (установлен) и 59195 (неопределён)
- Цвет по умолчанию: Черный (vbBlack) для всех состояний
- Эффект при наведении: Осветляет цвет значка при наведении курсора на флажок

## Файлы

- `clsMultiStateCheckBox.cls` - Основная реализация класса
- `frmTestClass.frm` - Тестовая форма, демонстрирующая использование
- `modShowForms.bas` - Модуль с функциями отображения форм
- `multi_state_checkbox_v2.xlsm` - Книга Excel, содержащая реализацию

## Использование

### Инициализация в пользовательской форме

Для использования многопозиционного флажка в вашей UserForm:

1. Добавьте элемент управления Label на вашу форму, который будет служить флажком
2. Импортируйте модуль класса `clsMultiStateCheckBox.cls` в ваш VBA проект
3. В коде вашей формы объявите экземпляр класса:

```vba
Dim mCheckBox As clsMultiStateCheckBox

Private Sub UserForm_Initialize()
    ' Инициализация с параметрами по умолчанию
    Set mCheckBox = New clsMultiStateCheckBox
    mCheckBox.Initialize Me.Label1 ' Замените Label1 на ваш элемент управления Label
End Sub

' Пример со всеми необязательными параметрами:
' Initialize(LabelBtn As MSForms.Label, _
'         Optional Item As Byte, _
'         Optional arrIcon As Variant, _
'         Optional ArrIconColor As Variant)
'
' Где:
' - LabelBtn: Элемент управления Label, который будет использоваться как флажок
' - Item: Начальное состояние (0 - снят, 1 - установлен, 2 - неопределён, и т.д.)
' - arrIcon: Массив кодов символов значков (по умолчанию Array(59193, 59194, 59195))
' - ArrIconColor: Массив цветов для каждого состояния (по умолчанию Array(vbBlack))

Private Sub UserForm_Initialize()
    ' Пример с указанием всех необязательных параметров
    Set mCheckBox = New clsMultiStateCheckBox
    
    ' Определение пользовательских значков (в данном случае, с использованием Segoe MDL2 Assets)
    Dim customIcons As Variant
    customIcons = Array(59193, 59194, 59195) ' Снят, установлен, неопределён
    
    ' Определение пользовательских цветов для каждого состояния
    Dim customColors As Variant
    customColors = Array(vbBlack, vbBlue, vbRed) ' Цвета для каждого состояния
    
    ' Инициализация с пользовательскими параметрами
    mCheckBox.Initialize Me.Label1, 0, customIcons, customColors
End Sub

Private Sub UserForm_Terminate()
    Set mCheckBox = Nothing
End Sub
```

4. Флажок теперь будет функционировать с тремя состояниями, которые переключаются при клике

## Лицензия

Этот проект лицензирован в соответствии с условиями, указанными в файле [LICENSE](./LICENSE).
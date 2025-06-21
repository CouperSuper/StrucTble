Attribute VB_Name = "ModulesAddins"
Sub DelEmptyColMod()
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

    LastColumn = ActiveSheet.UsedRange.Column - 1 + ActiveSheet.UsedRange.Columns.count    'определяем размеры таблицы
    For r = LastColumn To 1 Step -1           'проходим от последнего столбца до первого
        If Application.CountA(Columns(r)) = 0 Then Columns(r).Delete   'если в столбце пусто - удаляем его
    Next r
    LastColumn = ActiveSheet.UsedRange.Column - 1 + ActiveSheet.UsedRange.Columns.count
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub
Sub DelEmptyRowMod()
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

    LastRow = ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.count    'определяем размеры таблицы
    For r = LastRow To 1 Step -1           'проходим от последней строчки до первой
        If Application.CountA(Rows(r)) = 0 Then Rows(r).Delete   'если в строчке пусто - удаляем ее
    Next r
    LastRow = ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.count
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub
Sub LvlUnloadMod()

Dim LvlCol As String:    Dim TrigCol As Integer
Dim LvlNumb As Integer:  Dim FoundTrigCol As Integer
Dim Col As Integer:      Dim A As Object
Application.ScreenUpdating = False

Range("A1").Select: ActiveCell.Columns("A:A").EntireColumn.Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.ColumnWidth = 8
Range("A1").Select
Selection = "Уровень"

LvlCol = 1
ActiveCell.Offset(0, LvlCol).Select

'Счет столбцов
Do While Selection <> 0
Col = Col + 1
ActiveCell.Offset(0, 1).Select
Loop
ActiveCell.Offset(0, -Col).Select

'Поиск триггеров нулевого уровня
Do While Selection <> 0
    FoundTrigCol = FoundTrigCol + 1:   ActiveCell.Offset(0, 1).Select
    If ActiveCell.Text = "Статус операции" Then Exit Do
    If ActiveCell.Text = "Идентификатор операции" Then Exit Do
Loop
ActiveCell.Offset(0, -FoundTrigCol).Select

'Назначение нулевого уровня
If Col > FoundTrigCol Then TrigCol = FoundTrigCol + 1
If Col = FoundTrigCol Then TrigCol = InputBox("Нулевой уровень", "По какому столбцу отслеживать нулевой уровень?", Default)

'Цикл назначения уровня
Do While Selection <> 0
    LvlNumb = 0
    ActiveCell.Offset(1, 0).Select:    If Selection = 0 Then Exit Do
    
'Формула рассчета повторяющихся пробелов

For LvlNumb = Len(Selection) To 1 Step -1
 If InStr(Selection, Space$(LvlNumb)) Then Exit For
Next
        
    If LvlNumb = 1 Then LvlNumb = 0
    LvlNumb = (LvlNumb + 2) / 2
    ActiveCell.Offset(0, -LvlCol).Select
    Selection = LvlNumb
    ActiveCell.Offset(0, TrigCol).Select
    If Selection <> 0 Then LvlNumb = 10
    ActiveCell.Offset(0, -TrigCol).Select
    Selection = LvlNumb
ActiveCell.Offset(0, LvlCol).Select
Loop

'Пост обработка

Selection.AutoFilter
Application.ScreenUpdating = True

End Sub
Sub LvlPivotMod()
Dim lvl As Integer
Dim StepY As Integer
Application.ScreenUpdating = False

Range("A1").Select
ActiveCell.Columns("A:A").EntireColumn.Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("A1").Select
Selection = "Уровень"

Do Until ActiveCell.Offset(StepY, 1).Text = ""
StepY = StepY + 1
lvl = ActiveCell.Offset(StepY, 1).IndentLevel
ActiveCell.Offset(StepY, 0).Value = lvl + 1
Loop

Application.ScreenUpdating = True
End Sub
Sub SummaryMod()
    Dim Col As Integer
    Dim ML As Integer
    Dim CountCL As Integer
    Dim currentRow As Range
    Dim level As Integer
    Dim funcType As String
    Dim funcChoice As Integer

    Col = InputBox("В какую колонку добавить суммирования структуры", "Колонка суммирования структуры", 1)
    
    ' Запрос выбора функции по номеру
    funcChoice = InputBox("Выберите функцию для расчета: 1 - SUMIF, 2 - AVERAGEIF, 3 - SUMIF + COUNTIF, 4 - MAX, 5 - MIN, 6 - COUNTIF", "Выбор функции", 1)

    ' Определение функции по номеру
    Select Case funcChoice
        Case 1
            funcType = "SUMIF"
        Case 2
            funcType = "AVERAGEIF"
        Case 3
            funcType = "SUMIF_COUNTIF"
        Case 4
            funcType = "MAX"
        Case 5
            funcType = "MIN"
        Case 6
            funcType = "COUNTIF"
        Case Else
            MsgBox "Недопустимый выбор функции. Пожалуйста, выберите от 1 до 6."
            Exit Sub
    End Select

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Определяем максимальный уровень в структуре
    ML = Application.WorksheetFunction.Max(Range("A:A"))

    For level = 1 To ML - 1
        Set currentRow = Selection ' Начинаем с первой строки данных
        
        Do While Not IsEmpty(currentRow)
            If currentRow.Value = level Then
                ' Начинаем подсчет строк на следующем уровне
                CountCL = 0
                Do While Not IsEmpty(currentRow.Offset(CountCL + 1, 0)) And currentRow.Offset(CountCL + 1, 0).Value > level
                    CountCL = CountCL + 1
                Loop

                ' Проверка: на последнем уровне используем COUNTIF, на уровнях выше - SUMIF
                If CountCL > 0 Then
                    If funcType = "SUMIF_COUNTIF" Then
                        If level = ML - 1 Then
                            ' COUNTIF на последнем уровне
                            currentRow.Offset(0, Col).Formula = "=COUNTIF(R[1]C[-" & Col & "]:R[" & CountCL & "]C[-" & Col & "], " & level + 1 & ")"
                        Else
                            ' SUMIF на уровнях выше
                            currentRow.Offset(0, Col).Formula = "=SUMIF(R[1]C[-" & Col & "]:R[" & CountCL & "]C[-" & Col & "], " & level + 1 & ", R[1]C:R[" & CountCL & "]C)"
                        End If
                    ElseIf funcType = "MAX" Or funcType = "MIN" Then
                        ' Массивная формула для MIN и MAX
                        currentRow.Offset(0, Col).FormulaArray = "=" & funcType & "(IF(R[1]C[-" & Col & "]:R[" & CountCL & "]C[-" & Col & "]=" & level + 1 & ", R[1]C:R[" & CountCL & "]C))"
                    ElseIf funcType = "COUNTIF" Then
                        ' Формула COUNTIF требует только диапазон и условие
                        currentRow.Offset(0, Col).Formula = "=" & funcType & "(R[1]C[-" & Col & "]:R[" & CountCL & "]C[-" & Col & "], " & level + 1 & ")"
                    Else
                        ' Формулы с IF для SUMIF и AVERAGEIF
                        currentRow.Offset(0, Col).Formula = "=" & funcType & "(R[1]C[-" & Col & "]:R[" & CountCL & "]C[-" & Col & "]," & level + 1 & ",R[1]C:R[" & CountCL & "]C)"
                    End If
                End If
            End If
            
            ' Переход к следующей строке
            Set currentRow = currentRow.Offset(1, 0)
        Loop
    Next level

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
Sub GroupMod()

    Dim groupLevel As Integer
    Dim currentRow As Range
    Dim level As Integer
    Dim startRow As Range
    Dim groupSize As Integer

    groupLevel = InputBox("Введите уровень группировки", "Уровень группировки", 1)
    Application.ScreenUpdating = False

    For level = 1 To groupLevel
        Set currentRow = Selection ' Начинаем с первой строки данных

        Do While Not IsEmpty(currentRow)
            ' Проверяем, соответствует ли уровень ячейки текущему уровню группировки
            If currentRow.Value = level Then
                Set startRow = currentRow.Offset(1, 0) ' Начинаем группировку со следующей строки
                groupSize = 0

                ' Находим строки для группировки, пока уровень больше текущего
                Do While Not IsEmpty(startRow.Offset(groupSize, 0)) And startRow.Offset(groupSize, 0).Value > level
                    groupSize = groupSize + 1
                Loop

                ' Выполняем группировку, если найдено несколько строк для объединения
                If groupSize > 0 Then
                    startRow.Resize(groupSize).Rows.Group
                End If

                ' Переход к следующей строке после группировки
                Set currentRow = startRow.Offset(groupSize, 0)
            Else
                ' Если уровень не совпадает, просто переходим к следующей строке
                Set currentRow = currentRow.Offset(1, 0)
            End If
        Loop
    Next level

    ' Настройка стилей контура
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlAbove
        .SummaryColumn = xlLeft
    End With
    
    Application.ScreenUpdating = True
    
End Sub
Sub UnGroupMod()
    Selection.ClearOutline
End Sub
Sub RightOpenMod()
    With ActiveSheet.Outline
        .SummaryColumn = xlLeft
    End With
End Sub
Sub DownOpenMod()
    With ActiveSheet.Outline
        .SummaryRow = xlAbove
    End With
End Sub
Sub Форматирование_дат()
'
Dim P As Range
Dim i As Date
Dim CountEmpty As Integer
Dim StepY As Integer

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Do Until CountEmpty = 10000
    
    StepY = StepY + 1
    If ActiveCell.Offset(StepY, 0).Value = "" Then: CountEmpty = CountEmpty + 1: GoTo x
    For Each P In ActiveCell.Offset(StepY, 0)
    P = Replace(Application.Trim(Replace(P.Value, "A", " ")), " ", "A")
    P = Replace(Application.Trim(Replace(P.Value, "*", " ")), " ", "*")
    i = P.Value
    P.Value = i
    Next
x:
    
Loop
Range("A1").Select
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


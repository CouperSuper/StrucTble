Attribute VB_Name = "Smartart"
Option Explicit

Public Sub BuildSmartArtFromPivot()
    Dim oSmartArt As Office.Smartart
    Dim ws As Object
    Dim rngData As Range
    Dim rCell As Range
    Dim dictLastNode As Object   ' Словарь для хранения последнего узла для каждого indent-уровня
    Dim currentLevel As Long
    Dim parentNode As Office.SmartArtNode
    Dim newNode As Office.SmartArtNode

    Application.ScreenUpdating = False

    Set ws = ActiveSheet
    
    ' Получаем SmartArt-объект на активном листе (предполагается, что он уже создан)
    Set oSmartArt = GetSmartArtObject()
    If oSmartArt Is Nothing Then
        MsgBox "На активном листе не найден SmartArt?объект.", vbExclamation
        Exit Sub
    End If
    
    ' Очистка SmartArt: оставляем только заголовок (первый узел с индексом 1)
    Do While oSmartArt.AllNodes.count > 1
        oSmartArt.AllNodes(oSmartArt.AllNodes.count).Delete
    Loop
    
    ' Предполагается, что сводная таблица с данными находится в столбце A, начиная со второй строки.
    ' Первая строка – заголовок (SmartArt уже содержит заголовок с индексом 1)
    Set rngData = ws.Range("A2", ws.Cells(ws.Rows.count, "A").End(xlUp))
    
    ' Создаем словарь для запоминания последнего узла для каждого indent-уровня.
    Set dictLastNode = CreateObject("Scripting.Dictionary")
    ' Для indent = 0 (заголовок) сохраняем объект узла с индексом 1.
    Set dictLastNode(0) = oSmartArt.AllNodes(1)
    
    ' Проходим по каждой строке сводной таблицы.
    For Each rCell In rngData
        If Trim(rCell.Value) <> "" Then
            ' Получаем уровень отступа ячейки
            currentLevel = rCell.IndentLevel + 1
            
            ' Для indent=1 родитель – это всегда заголовок (уровень 0)
            If currentLevel = 1 Then
                Set parentNode = dictLastNode(0)
            Else
                ' Для indent > 1 ищем родительский узел с уровнем (currentLevel - 1)
                If dictLastNode.exists(currentLevel - 1) Then
                    Set parentNode = dictLastNode(currentLevel - 1)
                Else
                    Set parentNode = dictLastNode(0)
                End If
            End If
            
            ' Добавляем новый узел под найденным родительским узлом
            Set newNode = parentNode.AddNode(Position:=msoSmartArtNodeBelow)
            newNode.TextFrame2.TextRange.Text = rCell.Value
            
            ' Сохраняем объект нового узла для текущего indent-уровня
            Set dictLastNode(currentLevel) = newNode
        End If
    Next rCell
    
    Set dictLastNode = Nothing
    
    Application.ScreenUpdating = True
End Sub
Private Function GetSmartArtObject() As Office.Smartart
    Dim oShp As Shape
    For Each oShp In ActiveSheet.Shapes
        If oShp.Type = msoSmartArt Then
            Set GetSmartArtObject = oShp.Smartart
            Exit Function
        End If
    Next oShp
    Set GetSmartArtObject = Nothing
End Function
Sub SetIndentByValue()
    Dim cell As Range
    
    ' Проверяем, есть ли выделенные ячейки
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False ' Отключаем обновление экрана для ускорения

    ' Проходим по каждой выделенной ячейке
    For Each cell In Selection
        If IsNumeric(cell.Value) Then ' Проверяем, является ли значение числом
            If cell.Value >= 1 And cell.Value <= 10 Then
                cell.IndentLevel = cell.Value - 1 ' Устанавливаем отступ на (значение - 1)
            ElseIf cell.Value > 10 Then
                cell.IndentLevel = 10 ' Максимальный отступ (Excel не поддерживает больше 10)
            Else
                cell.IndentLevel = 0 ' Отрицательные числа и нули сбрасывают отступ
            End If
        End If
    Next cell

    Application.ScreenUpdating = True ' Включаем обновление экрана
End Sub


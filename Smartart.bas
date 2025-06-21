Attribute VB_Name = "Smartart"
Option Explicit

Public Sub BuildSmartArtFromPivot()
    Dim oSmartArt As Office.Smartart
    Dim ws As Object
    Dim rngData As Range
    Dim rCell As Range
    Dim dictLastNode As Object   ' ������� ��� �������� ���������� ���� ��� ������� indent-������
    Dim currentLevel As Long
    Dim parentNode As Office.SmartArtNode
    Dim newNode As Office.SmartArtNode

    Application.ScreenUpdating = False

    Set ws = ActiveSheet
    
    ' �������� SmartArt-������ �� �������� ����� (��������������, ��� �� ��� ������)
    Set oSmartArt = GetSmartArtObject()
    If oSmartArt Is Nothing Then
        MsgBox "�� �������� ����� �� ������ SmartArt?������.", vbExclamation
        Exit Sub
    End If
    
    ' ������� SmartArt: ��������� ������ ��������� (������ ���� � �������� 1)
    Do While oSmartArt.AllNodes.count > 1
        oSmartArt.AllNodes(oSmartArt.AllNodes.count).Delete
    Loop
    
    ' ��������������, ��� ������� ������� � ������� ��������� � ������� A, ������� �� ������ ������.
    ' ������ ������ � ��������� (SmartArt ��� �������� ��������� � �������� 1)
    Set rngData = ws.Range("A2", ws.Cells(ws.Rows.count, "A").End(xlUp))
    
    ' ������� ������� ��� ����������� ���������� ���� ��� ������� indent-������.
    Set dictLastNode = CreateObject("Scripting.Dictionary")
    ' ��� indent = 0 (���������) ��������� ������ ���� � �������� 1.
    Set dictLastNode(0) = oSmartArt.AllNodes(1)
    
    ' �������� �� ������ ������ ������� �������.
    For Each rCell In rngData
        If Trim(rCell.Value) <> "" Then
            ' �������� ������� ������� ������
            currentLevel = rCell.IndentLevel + 1
            
            ' ��� indent=1 �������� � ��� ������ ��������� (������� 0)
            If currentLevel = 1 Then
                Set parentNode = dictLastNode(0)
            Else
                ' ��� indent > 1 ���� ������������ ���� � ������� (currentLevel - 1)
                If dictLastNode.exists(currentLevel - 1) Then
                    Set parentNode = dictLastNode(currentLevel - 1)
                Else
                    Set parentNode = dictLastNode(0)
                End If
            End If
            
            ' ��������� ����� ���� ��� ��������� ������������ �����
            Set newNode = parentNode.AddNode(Position:=msoSmartArtNodeBelow)
            newNode.TextFrame2.TextRange.Text = rCell.Value
            
            ' ��������� ������ ������ ���� ��� �������� indent-������
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
    
    ' ���������, ���� �� ���������� ������
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False ' ��������� ���������� ������ ��� ���������

    ' �������� �� ������ ���������� ������
    For Each cell In Selection
        If IsNumeric(cell.Value) Then ' ���������, �������� �� �������� ������
            If cell.Value >= 1 And cell.Value <= 10 Then
                cell.IndentLevel = cell.Value - 1 ' ������������� ������ �� (�������� - 1)
            ElseIf cell.Value > 10 Then
                cell.IndentLevel = 10 ' ������������ ������ (Excel �� ������������ ������ 10)
            Else
                cell.IndentLevel = 0 ' ������������� ����� � ���� ���������� ������
            End If
        End If
    Next cell

    Application.ScreenUpdating = True ' �������� ���������� ������
End Sub


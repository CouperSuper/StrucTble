Attribute VB_Name = "LvlFormat"
Public RGB1 As String:  Public CFont1 As String:    Public WFont1 As String
Public RGB2 As String:  Public CFont2 As String:    Public WFont2 As String
Public RGB3 As String:  Public CFont3 As String:    Public WFont3 As String
Public RGB4 As String:  Public CFont4 As String:    Public WFont4 As String
Public RGB5 As String:  Public CFont5 As String:    Public WFont5 As String
Public RGB6 As String:  Public CFont6 As String:    Public WFont6 As String
Public RGB7 As String:  Public CFont7 As String:    Public WFont7 As String
Public RGB8 As String:  Public CFont8 As String:    Public WFont8 As String

Public Pallete As String
Public Sub ColorFormat(control As IRibbonControl, selectedId As String)
Pallete = selectedId
    Select Case selectedId
    
        Case "������������� ��������"
            RGB1 = RGB(221, 11, 34):            CFont1 = RGB(255, 255, 255):        WFont1 = xlBold   ' ������������ ������� (�������� ���������)
            RGB2 = RGB(255, 255, 255):          CFont2 = RGB(0, 0, 0):              WFont2 = xlNormal ' ����� (�������)
            RGB3 = RGB(157, 157, 157):          CFont3 = RGB(0, 0, 0):              WFont3 = xlBold   ' ������������� ����� (������ ��������)
            RGB4 = RGB(60, 60, 60):             CFont4 = RGB(255, 255, 255):        WFont4 = xlBold   ' �����-����� (���������)
            RGB5 = RGB(87, 87, 87):             CFont5 = RGB(255, 255, 255):        WFont5 = xlThick  ' �������� ����� (������������)
            RGB6 = RGB(111, 111, 111):          CFont6 = RGB(255, 255, 255):        WFont6 = xlThick  ' ������� ����� (��� ���������)
            RGB7 = RGB(198, 198, 198):          CFont7 = RGB(0, 0, 0):              WFont7 = xlNormal ' ������-����� (������� ��������)
            RGB8 = RGB(237, 237, 237):          CFont8 = RGB(0, 0, 0):              WFont8 = xlNormal ' ����� ������� ����� (���������)
            
        Case "�������� ��������"
            RGB1 = RGB(237, 237, 237):          CFont1 = RGB(0, 0, 0):              WFont1 = xlNormal ' ����� ������� ����� (�������, �����������)
            RGB2 = RGB(218, 218, 218):          CFont2 = RGB(0, 0, 0):              WFont2 = xlNormal ' ������-����� (������, ����������������)
            RGB3 = RGB(198, 198, 198):          CFont3 = RGB(0, 0, 0):              WFont3 = xlNormal ' ������� ����� (���� ������, �������� ��� ����������)
            RGB4 = RGB(178, 178, 178):          CFont4 = RGB(255, 255, 255):        WFont4 = xlNormal ' �������� ����� (��� ���������)
            RGB5 = RGB(157, 157, 157):          CFont5 = RGB(255, 255, 255):        WFont5 = xlBold   ' ������������� ����� (�������� ���������)
            RGB6 = RGB(87, 87, 87):             CFont6 = RGB(255, 255, 255):        WFont6 = xlBold   ' �������� ���������� (��� ������� ����������)
            RGB7 = RGB(60, 60, 60):             CFont7 = RGB(255, 255, 255):        WFont7 = xlBold   ' ׸���� ������ (��� ������ ����������)
            RGB8 = RGB(221, 11, 34):            CFont8 = RGB(255, 255, 255):        WFont8 = xlBold   ' ������������ ������� (��������� ������)

        Case "������-�����"
            RGB1 = RGB(218, 226, 248):           CFont1 = RGB(0, 0, 0):            WFont1 = xlNormal ' �����-���������� �����
            RGB2 = RGB(190, 209, 240):           CFont2 = RGB(0, 0, 0):            WFont2 = xlNormal ' ���������� ������
            RGB3 = RGB(163, 192, 233):           CFont3 = RGB(0, 0, 0):            WFont3 = xlNormal ' ������� � ������� ��������
            RGB4 = RGB(135, 170, 222):           CFont4 = RGB(255, 255, 255):      WFont4 = xlNormal ' �������� �����
            RGB5 = RGB(109, 147, 210):           CFont5 = RGB(255, 255, 255):      WFont5 = xlBold   ' �������� ����-�����
            RGB6 = RGB(92, 126, 196):            CFont6 = RGB(255, 255, 255):      WFont6 = xlBold   ' ������������ ����������� �����
            RGB7 = RGB(72, 87, 160):             CFont7 = RGB(255, 255, 255):      WFont7 = xlBold   ' ��������-�����
            RGB8 = RGB(156, 81, 182):            CFont8 = RGB(255, 255, 255):      WFont8 = xlBold   ' ��������� ������
        
        Case "������ ������"
            RGB1 = RGB(38, 70, 83):              CFont1 = RGB(255, 255, 255):      WFont1 = xlBold   ' �������� ����-����� (������������, ��������)
            RGB2 = RGB(42, 157, 143):            CFont2 = RGB(255, 255, 255):      WFont2 = xlBold   ' ������� ����� (�������� ����� ������)
            RGB3 = RGB(233, 196, 106):           CFont3 = RGB(0, 0, 0):            WFont3 = xlNormal ' ������� �������� (������ ������)
            RGB4 = RGB(244, 162, 97):            CFont4 = RGB(0, 0, 0):            WFont4 = xlNormal ' ��������-������������ (����� � �����)
            RGB5 = RGB(231, 111, 81):            CFont5 = RGB(255, 255, 255):      WFont5 = xlBold   ' ������ ���������� (����������� ������)
            RGB6 = RGB(255, 245, 233):           CFont6 = RGB(0, 0, 0):            WFont6 = xlNormal ' �������� ����� (������ ���)
            RGB7 = RGB(252, 227, 138):           CFont7 = RGB(0, 0, 0):            WFont7 = xlNormal ' ������� �� (�������������� �������)
            RGB8 = RGB(129, 178, 154):           CFont8 = RGB(0, 0, 0):            WFont8 = xlNormal ' ���������� ����������-���������

        Case "�������� �������������"
            RGB1 = RGB(197, 225, 251):            CFont1 = RGB(0, 0, 0):            WFont1 = xlNormal ' ������ �������� �������
            RGB2 = RGB(158, 158, 158):            CFont2 = RGB(0, 0, 0):            WFont2 = xlNormal ' ����� ���
            RGB3 = RGB(207, 216, 220):            CFont3 = RGB(0, 0, 0):            WFont3 = xlNormal ' ����������� ����
            RGB4 = RGB(255, 255, 255):            CFont4 = RGB(0, 0, 0):            WFont4 = xlBold   ' ������ ����� ���� (������ ��������)
            RGB5 = RGB(179, 229, 252):            CFont5 = RGB(0, 0, 0):            WFont5 = xlNormal ' ������� ������
            RGB6 = RGB(100, 181, 246):            CFont6 = RGB(255, 255, 255):      WFont6 = xlBold   ' ������ �������� ����� (�����������)
            RGB7 = RGB(174, 213, 229):            CFont7 = RGB(0, 0, 0):            WFont7 = xlNormal ' ������� �����
            RGB8 = RGB(38, 50, 56):               CFont8 = RGB(255, 255, 255):      WFont8 = xlBold   ' �������� ����������� ����-����� (������ ��� ����������)
        
        Case "������� �������"
            RGB1 = RGB(204, 213, 174):           CFont1 = RGB(0, 0, 0):            WFont1 = xlNormal ' Ҹ���� ��������� (������ ������)
            RGB2 = RGB(233, 237, 201):           CFont2 = RGB(0, 0, 0):            WFont2 = xlNormal ' ������� ����������-������ (��������)
            RGB3 = RGB(254, 250, 224):           CFont3 = RGB(0, 0, 0):            WFont3 = xlNormal ' �������� ������� (���, ���)
            RGB4 = RGB(250, 237, 205):           CFont4 = RGB(0, 0, 0):            WFont4 = xlNormal ' Ҹ���� ������-������ (�������������)
            RGB5 = RGB(212, 163, 115):           CFont5 = RGB(0, 0, 0):            WFont5 = xlBold   ' ������-����������� (�����)
            RGB6 = RGB(180, 136, 91):            CFont6 = RGB(255, 255, 255):      WFont6 = xlBold   ' ������� �������� (������)
            RGB7 = RGB(145, 103, 63):            CFont7 = RGB(255, 255, 255):      WFont7 = xlBold   ' �������-���������� (�������)
            RGB8 = RGB(114, 77, 38):             CFont8 = RGB(255, 255, 255):      WFont8 = xlBold   ' Ҹ���� ��������� (��������)

        Case "�������� ��������"
            RGB1 = RGB(95, 15, 64):              CFont1 = RGB(255, 255, 255):      WFont1 = xlBold   ' �������� �������-��������� (�����, �� �� ������� ����)
            RGB2 = RGB(154, 3, 30):              CFont2 = RGB(255, 255, 255):      WFont2 = xlBold   ' ��������� �������� (���������� ������� ���)
            RGB3 = RGB(251, 139, 36):            CFont3 = RGB(0, 0, 0):            WFont3 = xlNormal ' �����-��������� (������ � ������)
            RGB4 = RGB(227, 100, 20):            CFont4 = RGB(255, 255, 255):      WFont4 = xlBold   ' �������� �������� (����������� � ������)
            RGB5 = RGB(15, 76, 92):              CFont5 = RGB(255, 255, 255):      WFont5 = xlBold   ' ������� ���� ����� ������ (�������� ������)
            RGB6 = RGB(255, 183, 77):            CFont6 = RGB(0, 0, 0):            WFont6 = xlNormal ' ������-��������� (������� ������ ������)
            RGB7 = RGB(255, 204, 128):           CFont7 = RGB(0, 0, 0):            WFont7 = xlNormal ' ���������� ����������� (������, ������)
            RGB8 = RGB(239, 108, 0):             CFont8 = RGB(255, 255, 255):      WFont8 = xlBold   ' �������-��������� (����������� ������)

        Case Else
            MsgBox "������� ����������� �������"
    End Select
End Sub
Sub FormatMod()

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Range("A1").Select
Set Rng = Range("A1", Selection.End(xlDown))
Fcol = Selection(1).CurrentRegion.Columns.count
Rng.Select

If Pallete = Empty Then GoTo x Else GoTo Y

x:
    
    RGB1 = RGB(221, 11, 34):            CFont1 = RGB(255, 255, 255):        WFont1 = xlBold
    RGB2 = RGB(255, 255, 255):          CFont2 = RGB(0, 0, 0):              WFont2 = xlThick
    RGB3 = RGB(157, 157, 157):          CFont3 = RGB(0, 0, 0):              WFont4 = xlNormal
    RGB4 = RGB(60, 60, 60):             CFont4 = RGB(255, 255, 255):        WFont5 = xlNormal
    RGB5 = RGB(87, 87, 87):             CFont5 = RGB(255, 255, 255):        WFont6 = xlNormal
    RGB6 = RGB(111, 111, 111):          CFont6 = RGB(255, 255, 255):        WFont7 = xlNormal
    RGB7 = RGB(198, 198, 198):          CFont7 = RGB(0, 0, 0):              WFont8 = xlNormal
    RGB8 = RGB(237, 237, 237):          CFont8 = RGB(0, 0, 0):              WFont9 = xlNormal
            
Y:
For Each A In Rng
      A.Select
       Select Case A
        Case 0:  Range(ActiveCell, ActiveCell.Offset(0, Fcol - 1)).Select
                                    With Selection.Interior:  .Color = RGB1:   End With
                                    With Selection.Font:      .Color = CFont1: End With
                                    With Selection.Font:      .Bold = (WFont1 = xlBold): End With
        Case 1:  Range(ActiveCell, ActiveCell.Offset(0, Fcol - 1)).Select
                                    With Selection.Interior:  .Color = RGB1:   End With
                                    With Selection.Font:      .Color = CFont1: End With
                                    With Selection.Font:      .Bold = (WFont1 = xlBold): End With
        Case 2:  Range(ActiveCell, ActiveCell.Offset(0, Fcol - 1)).Select
                                    With Selection.Interior:  .Color = RGB2:   End With
                                    With Selection.Font:      .Color = CFont2: End With
                                    With Selection.Font:      .Bold = (WFont2 = xlBold): End With
        Case 3:  Range(ActiveCell, ActiveCell.Offset(0, Fcol - 1)).Select
                                    With Selection.Interior:  .Color = RGB3:   End With
                                    With Selection.Font:      .Color = CFont3: End With
                                    With Selection.Font:      .Bold = (WFont3 = xlBold): End With
        Case 4:  Range(ActiveCell, ActiveCell.Offset(0, Fcol - 1)).Select
                                    With Selection.Interior:  .Color = RGB4:   End With
                                    With Selection.Font:      .Color = CFont4: End With
                                    With Selection.Font:      .Bold = (WFont4 = xlBold): End With
        Case 5:  Range(ActiveCell, ActiveCell.Offset(0, Fcol - 1)).Select
                                    With Selection.Interior:  .Color = RGB5:   End With
                                    With Selection.Font:      .Color = CFont5: End With
                                    With Selection.Font:      .Bold = (WFont5 = xlBold): End With
        Case 6:  Range(ActiveCell, ActiveCell.Offset(0, Fcol - 1)).Select
                                    With Selection.Interior:  .Color = RGB6:   End With
                                    With Selection.Font:      .Color = CFont6: End With
                                    With Selection.Font:      .Bold = (WFont6 = xlBold): End With
        Case 7:  Range(ActiveCell, ActiveCell.Offset(0, Fcol - 1)).Select
                                    With Selection.Interior:  .Color = RGB7:   End With
                                    With Selection.Font:      .Color = CFont7: End With
                                    With Selection.Font:      .Bold = (WFont7 = xlBold): End With
        Case 8:  Range(ActiveCell, ActiveCell.Offset(0, Fcol - 1)).Select
                                    With Selection.Interior:  .Color = RGB8:   End With
                                    With Selection.Font:      .Color = CFont8: End With
                                    With Selection.Font:      .Bold = (WFont8 = xlBold): End With
    End Select
Next
Range("A1").Select: Selection.CurrentRegion.Select
               With Selection.Borders(xlEdgeLeft):          .LineStyle = xlContinuous:   .Weight = xlThin:     End With
               With Selection.Borders(xlEdgeTop):           .LineStyle = xlContinuous:   .Weight = xlThin:     End With
               With Selection.Borders(xlEdgeBottom):        .LineStyle = xlContinuous:   .Weight = xlThin:     End With
               With Selection.Borders(xlEdgeRight):         .LineStyle = xlContinuous:   .Weight = xlThin:     End With
               With Selection.Borders(xlInsideVertical):    .LineStyle = xlContinuous:   .Weight = xlThin:     End With
               With Selection.Borders(xlInsideHorizontal):  .LineStyle = xlContinuous:   .Weight = xlThin:     End With

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

RGB1 = Empty:  RGB2 = Empty:  RGB3 = Empty
RGB4 = Empty:  RGB5 = Empty:  RGB6 = Empty
RGB7 = Empty:  RGB8 = Empty:  Pallete = Empty

CFont1 = Empty: CFont2 = Empty: CFont3 = Empty
CFont4 = Empty: CFont5 = Empty: CFont6 = Empty
CFont7 = Empty: CFont8 = Empty

WFont1 = Empty: WFont2 = Empty: WFont3 = Empty
WFont4 = Empty: WFont5 = Empty: WFont6 = Empty
WFont7 = Empty: WFont8 = Empty

End Sub

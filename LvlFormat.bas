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
    
        Case "Корпоративный брэндбук"
            RGB1 = RGB(221, 11, 34):            CFont1 = RGB(255, 255, 255):        WFont1 = xlBold   ' Ализариновый красный (основной акцентный)
            RGB2 = RGB(255, 255, 255):          CFont2 = RGB(0, 0, 0):              WFont2 = xlNormal ' Белый (фоновый)
            RGB3 = RGB(157, 157, 157):          CFont3 = RGB(0, 0, 0):              WFont3 = xlBold   ' Перламутровый серый (мягкий контраст)
            RGB4 = RGB(60, 60, 60):             CFont4 = RGB(255, 255, 255):        WFont4 = xlBold   ' Темно-серый (заголовки)
            RGB5 = RGB(87, 87, 87):             CFont5 = RGB(255, 255, 255):        WFont5 = xlThick  ' Глубокий серый (подзаголовки)
            RGB6 = RGB(111, 111, 111):          CFont6 = RGB(255, 255, 255):        WFont6 = xlThick  ' Средний серый (для выделения)
            RGB7 = RGB(198, 198, 198):          CFont7 = RGB(0, 0, 0):              WFont7 = xlNormal ' Светло-серый (фоновые элементы)
            RGB8 = RGB(237, 237, 237):          CFont8 = RGB(0, 0, 0):              WFont8 = xlNormal ' Очень светлый серый (градиенты)
            
        Case "Брэндбук монохром"
            RGB1 = RGB(237, 237, 237):          CFont1 = RGB(0, 0, 0):              WFont1 = xlNormal ' Очень светлый серый (Фоновый, нейтральный)
            RGB2 = RGB(218, 218, 218):          CFont2 = RGB(0, 0, 0):              WFont2 = xlNormal ' Светло-серый (Мягкий, сбалансированный)
            RGB3 = RGB(198, 198, 198):          CFont3 = RGB(0, 0, 0):              WFont3 = xlNormal ' Средний серый (Чуть глубже, подходит для разделений)
            RGB4 = RGB(178, 178, 178):          CFont4 = RGB(255, 255, 255):        WFont4 = xlNormal ' Глубокий серый (Для выделений)
            RGB5 = RGB(157, 157, 157):          CFont5 = RGB(255, 255, 255):        WFont5 = xlBold   ' Перламутровый серый (Основной акцентный)
            RGB6 = RGB(87, 87, 87):             CFont6 = RGB(255, 255, 255):        WFont6 = xlBold   ' Глубокий графитовый (Для сильных контрастов)
            RGB7 = RGB(60, 60, 60):             CFont7 = RGB(255, 255, 255):        WFont7 = xlBold   ' Чёрный графит (Для важных заголовков)
            RGB8 = RGB(221, 11, 34):            CFont8 = RGB(255, 255, 255):        WFont8 = xlBold   ' Ализариновый красный (Фирменный акцент)

        Case "Бизнес-синий"
            RGB1 = RGB(218, 226, 248):           CFont1 = RGB(0, 0, 0):            WFont1 = xlNormal ' Нежно-лавандовый синий
            RGB2 = RGB(190, 209, 240):           CFont2 = RGB(0, 0, 0):            WFont2 = xlNormal ' Пастельный индиго
            RGB3 = RGB(163, 192, 233):           CFont3 = RGB(0, 0, 0):            WFont3 = xlNormal ' Голубой с лиловым подтоном
            RGB4 = RGB(135, 170, 222):           CFont4 = RGB(255, 255, 255):      WFont4 = xlNormal ' Пудровый синий
            RGB5 = RGB(109, 147, 210):           CFont5 = RGB(255, 255, 255):      WFont5 = xlBold   ' Глубокий серо-синий
            RGB6 = RGB(92, 126, 196):            CFont6 = RGB(255, 255, 255):      WFont6 = xlBold   ' Приглушённый королевский синий
            RGB7 = RGB(72, 87, 160):             CFont7 = RGB(255, 255, 255):      WFont7 = xlBold   ' Фиалково-синий
            RGB8 = RGB(156, 81, 182):            CFont8 = RGB(255, 255, 255):      WFont8 = xlBold   ' Пурпурный сапфир
        
        Case "Теплый акцент"
            RGB1 = RGB(38, 70, 83):              CFont1 = RGB(255, 255, 255):      WFont1 = xlBold   ' Глубокий серо-синий (приглушённый, солидный)
            RGB2 = RGB(42, 157, 143):            CFont2 = RGB(255, 255, 255):      WFont2 = xlBold   ' Морская волна (основной тёплый акцент)
            RGB3 = RGB(233, 196, 106):           CFont3 = RGB(0, 0, 0):            WFont3 = xlNormal ' Светлый янтарный (летнее солнце)
            RGB4 = RGB(244, 162, 97):            CFont4 = RGB(0, 0, 0):            WFont4 = xlNormal ' Лососево-апельсиновый (живой и тёплый)
            RGB5 = RGB(231, 111, 81):            CFont5 = RGB(255, 255, 255):      WFont5 = xlBold   ' Теплый коралловый (вибрирующий акцент)
            RGB6 = RGB(255, 245, 233):           CFont6 = RGB(0, 0, 0):            WFont6 = xlNormal ' Кремовый белый (мягкий фон)
            RGB7 = RGB(252, 227, 138):           CFont7 = RGB(0, 0, 0):            WFont7 = xlNormal ' Светлый мёд (поддерживающий элемент)
            RGB8 = RGB(129, 178, 154):           CFont8 = RGB(0, 0, 0):            WFont8 = xlNormal ' Пастельный зеленовато-бирюзовый

        Case "Холодный аналитический"
            RGB1 = RGB(197, 225, 251):            CFont1 = RGB(0, 0, 0):            WFont1 = xlNormal ' Нежный морозный голубой
            RGB2 = RGB(158, 158, 158):            CFont2 = RGB(0, 0, 0):            WFont2 = xlNormal ' Серый лед
            RGB3 = RGB(207, 216, 220):            CFont3 = RGB(0, 0, 0):            WFont3 = xlNormal ' Серебристый иней
            RGB4 = RGB(255, 255, 255):            CFont4 = RGB(0, 0, 0):            WFont4 = xlBold   ' Чистый белый снег (важные элементы)
            RGB5 = RGB(179, 229, 252):            CFont5 = RGB(0, 0, 0):            WFont5 = xlNormal ' Ледяная бирюза
            RGB6 = RGB(100, 181, 246):            CFont6 = RGB(255, 255, 255):      WFont6 = xlBold   ' Чистый морозный синий (контрастный)
            RGB7 = RGB(174, 213, 229):            CFont7 = RGB(0, 0, 0):            WFont7 = xlNormal ' Голубая дымка
            RGB8 = RGB(38, 50, 56):               CFont8 = RGB(255, 255, 255):      WFont8 = xlBold   ' Глубокий арктический серо-синий (жирный для заголовков)
        
        Case "Осенняя палитра"
            RGB1 = RGB(204, 213, 174):           CFont1 = RGB(0, 0, 0):            WFont1 = xlNormal ' Тёплый оливковый (мягкая основа)
            RGB2 = RGB(233, 237, 201):           CFont2 = RGB(0, 0, 0):            WFont2 = xlNormal ' Светлый травянисто-зелёный (легкость)
            RGB3 = RGB(254, 250, 224):           CFont3 = RGB(0, 0, 0):            WFont3 = xlNormal ' Кремовый светлый (фон, уют)
            RGB4 = RGB(250, 237, 205):           CFont4 = RGB(0, 0, 0):            WFont4 = xlNormal ' Тёплый бежево-желтый (натуральность)
            RGB5 = RGB(212, 163, 115):           CFont5 = RGB(0, 0, 0):            WFont5 = xlBold   ' Светло-карамельный (тепло)
            RGB6 = RGB(180, 136, 91):            CFont6 = RGB(255, 255, 255):      WFont6 = xlBold   ' Осенний ореховый (акцент)
            RGB7 = RGB(145, 103, 63):            CFont7 = RGB(255, 255, 255):      WFont7 = xlBold   ' Кофейно-коричневый (глубина)
            RGB8 = RGB(114, 77, 38):             CFont8 = RGB(255, 255, 255):      WFont8 = xlBold   ' Тёмный древесный (контраст)

        Case "Песчаный градиент"
            RGB1 = RGB(95, 15, 64):              CFont1 = RGB(255, 255, 255):      WFont1 = xlBold   ' Глубокий бордово-пурпурный (тепло, но не слишком ярко)
            RGB2 = RGB(154, 3, 30):              CFont2 = RGB(255, 255, 255):      WFont2 = xlBold   ' Бархатный бордовый (насыщенный осенний тон)
            RGB3 = RGB(251, 139, 36):            CFont3 = RGB(0, 0, 0):            WFont3 = xlNormal ' Медно-оранжевый (листья в солнце)
            RGB4 = RGB(227, 100, 20):            CFont4 = RGB(255, 255, 255):      WFont4 = xlBold   ' Глубокий янтарный (контрастный и теплый)
            RGB5 = RGB(15, 76, 92):              CFont5 = RGB(255, 255, 255):      WFont5 = xlBold   ' Осеннее небо перед дождем (холодный баланс)
            RGB6 = RGB(255, 183, 77):            CFont6 = RGB(0, 0, 0):            WFont6 = xlNormal ' Светло-оранжевый (осенний теплый акцент)
            RGB7 = RGB(255, 204, 128):           CFont7 = RGB(0, 0, 0):            WFont7 = xlNormal ' Золотистый карамельный (мягкий, нежный)
            RGB8 = RGB(239, 108, 0):             CFont8 = RGB(255, 255, 255):      WFont8 = xlBold   ' Огненно-оранжевый (контрастный акцент)

        Case Else
            MsgBox "Выбрана неизвестная палитра"
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

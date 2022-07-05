Option Explicit
Dim lastrow         As Long, LastRow2 As Long
Public tip_spiska_IKK As Integer
Dim lastCol         As Integer
Sub Obrabotka_IKK()
    Dim Get_List    As Worksheet, Input_List As Worksheet, response As Integer
    'Dim t
    't = timer
    response = MsgBox("Начать обработку?", vbOKCancel, "Обработка списка")
    If response = vbCancel Then Exit Sub
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = FALSE
    Application.EnableEvents = FALSE
    ActiveSheet.DisplayPageBreaks = FALSE
    Application.DisplayStatusBar = FALSE
    Set Get_List = ActiveSheet
    Get_List.Activate
    If Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues) Is Nothing Then
        MsgBox "Лист: " & Get_List.name & " не содержит информации.": Exit Sub
    End If
    If Cells.Find(What:="снилс", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues) Is Nothing Then
        MsgBox "Колонка СНИЛС на вкладке " & ws & " не найдена!" & "" & Chr(10) & "" & "Данные с этой вкладки не будут обработанны." _
             & "" & Chr(10) & "" & "Озаглавьте столбец СНИЛС и попытайтесь еще раз.": Exit Sub
    End If
    If Get_List.FilterMode = TRUE Then Get_List.ShowAllData
    Cells.UnMerge
    ActiveSheet.UsedRange.EntireRow.Hidden = FALSE
    Call New_List2
    Set Input_List = Sheets(1)
    Get_List.Activate
    Call Main_IKK
    Input_List.Activate
    ActiveSheet.DisplayPageBreaks = FALSE
    Call Nomer_Akta: Call Chahge_SNILS: Call Naim_po_kontrakt
    Select Case tip_spiska_IKK
        Case 1, 2, 3, 4, 5, 8
            Call Parametr_FIZLICO
    End Select
    Call Formatirovan1
    Call Date_of_birth1
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = TRUE
    Application.EnableEvents = TRUE
    Application.DisplayStatusBar = TRUE
    Call data_naprav1
    'MsgBox Format(timer - t, "00.00000000") & " секунд"
End Sub
Private Sub Chahge_SNILS()
    Dim i           As Long, Col As Integer, rep_area As Range
    lastrow = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = Range("SNILS").Column
    'Col = Cells.Find(What:="снилс", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
    Set rep_area = Range(Cells(2, Col), Cells(lastrow, Col))
    rep_area.Replace What:="-", Replacement:=""
    rep_area.Replace What:=" ", Replacement:=""
    rep_area.Replace What:=",", Replacement:=""
    rep_area.Replace What:="C", Replacement:=""
    rep_area.Replace What:="С", Replacement:=""
    On Error Resume Next
    For i = 2 To lastrow
        If Len(Cells(i, Col)) = 11 Then
            Cells(i, Col).Value = Mid(Cells(i, Col), 1, 3) & "-" & Mid(Cells(i, Col), 4, 3) _
                   & "-" & Mid(Cells(i, Col), 7, 3) & " " & Mid(Cells(i, Col), 10, 2)
        ElseIf Len(Cells(i, Col)) = 10 Then
            Cells(i, Col).Value = "0" & Mid(Cells(i, Col), 1, 2) & "-" & Mid(Cells(i, Col), 3, 3) _
                   & "-" & Mid(Cells(i, Col), 6, 3) & " " & Mid(Cells(i, Col), 9, 2)
        ElseIf Len(Cells(i, Col)) = 9 Then
            Cells(i, Col).Value = "00" & Mid(Cells(i, Col), 1, 1) & "-" & Mid(Cells(i, Col), 2, 3) _
                   & "-" & Mid(Cells(i, Col), 5, 3) & Space(1) & Mid(Cells(i, Col), 8, 2)
        End If
    Next i
End Sub
Private Sub Nomer_Akta()
    Dim filial      As Integer, nedelya As Integer, i As Long, mathes As Integer, a As Integer
    lastrow = Range("FIO").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    nedelya = WorksheetFunction.WeekNum(Now)
    
    If lastrow = 2 Then filial = Range("fil")(2)
    For i = 2 To lastrow
        If IsNumeric(Val(Range("fil")(i))) = TRUE And Val(Range("fil")(i)) = Val(Range("fil")(i + 1)) And Val(Range("fil")(i).Value) > 0 Then
            filial = Range("fil")(i)
            Exit For
        End If
    Next i
    For i = 2 To lastrow
        Range("Acts")(i).Value = nedelya & "/" & filial & "/"
    Next i
End Sub
Private Sub New_List2()
    Dim a           As String, b As String, c As String
    Worksheets.Add Before:=Sheets(1)
    Select Case tip_spiska_IKK
        Case 1: a = "КоляскиОтОбодаКолесаКоличество": b = "НаименованиеПоКонтракту": c = "ШиринаКоляски"
        Case 2: a = "КоляскиДляДетейСДЦПКоличество": b = "НаименованиеПоКонтракту": c = "ШиринаКоляски"
        Case 3: a = "КоляскиСОткидн.СпинкойКоличество": b = "НаименованиеПоКонтракту": c = "ШиринаКоляски"
        Case 4: a = "КоляскиСРычаж.ПриводомКоличество": b = "НаименованиеПоКонтракту": c = "ШиринаКоляски"
        Case 5: a = "КоляскиПовыш.Груз.Количество": b = "НаименованиеПоКонтракту": c = "ШиринаКоляски"
        Case 6: a = "КоляскиСЭлектроприводомКоличество": b = "НаименованиеПоКонтракту": c = "ШиринаКоляски"
        Case 7: a = "КоляскиДляВыс.Ампут.Количество": b = "НаименованиеПоКонтракту": c = "ШиринаКоляски"
        Case 8: a = "КоляскиОттоБокКоличество": b = "НаименованиеПоКонтракту": c = "ШиринаКоляски"
        Case 9: a = "Кресла-СтульясСан.Осн.Количество": b = "ТипКреслаСтула": c = "НаименованиеПоКонтракту"
        Case Else: a = "КоляскиОтОбодаКолесаКоличество"
    End Select
    With Sheets(1)
        .Range("A:A").name = "Acts": .Range("A:A").ColumnWidth = 9.8: .Range("A:A").NumberFormat = "@"
        .Range("B:B").name = "FIO": .Range("B:B").ColumnWidth = 24
        .Range("C:C").name = "DateOfBirth": .Range("C:C").ColumnWidth = 10
        .Range("D:D").name = "SNILS": .Range("D:D").ColumnWidth = 14
        .Range("E:E").name = "Propiska": .Range("E:E").ColumnWidth = 20
        .Range("F:F").name = "Fakt": .Range("F:F").ColumnWidth = 20
        .Range("G:G").name = "TEL": .Range("G:G").ColumnWidth = 15
        .Range("H:H").name = "DopChar": .Range("H:H").ColumnWidth = 23
        .Range("I:I").name = "Kolvo": .Range("I:I").ColumnWidth = 8
        .Range("J:J").name = "Ves": .Range("J:J").ColumnWidth = 6.5
        .Range("K:K").name = "OT": .Range("K:K").ColumnWidth = 6.5
        .Range("L:L").name = "Fil": .Range("L:L").ColumnWidth = 6.5
        .Range("M:M").name = "Pasport": .Range("M:M").ColumnWidth = 21
        .Range("N:N").name = "DateNap": .Range("N:N").ColumnWidth = 10
        .Range("O:O").name = "NomerNap": .Range("O:O").ColumnWidth = 7
        .Range("P:P").name = "NaimKontr": .Range("P:P").ColumnWidth = 13
        .Cells(1, 1).Value = "НомерАкта": .Cells(1, 2).Value = "ФИО": .Cells(1, 3).Value = "ДатаРождения"
        .Cells(1, 4).Value = "СНИЛС": .Cells(1, 5).Value = "АдресПоПрописке": .Cells(1, 6).Value = "АдресФактический"
        .Cells(1, 7).Value = "Телефон": .Cells(1, 8).Value = "ДопХарактеристики": .Cells(1, 9).Value = a
        .Cells(1, 10).Value = "ВесФизлицо": .Cells(1, 11).Value = "ОбъемТалииФизлицо": .Cells(1, 12).Value = "Филиал"
        .Cells(1, 13).Value = "Паспорт": .Cells(1, 14).Value = "ДатаНаправления": .Cells(1, 15).Value = "НомерНаправления"
        .Cells(1, 16).Value = b: .Cells(1, 17).Value = "ПаспортПредставителя"
        .Cells(1, 18).Value = "Наспункт": .Cells(1, 19).Value = "ВесКоляски": .Cells(1, 20).Value = c
        .Cells(1, 21).Value = "Номенклатура": .Cells(1, 22).Value = "Примечание"
        .Range("A:B").NumberFormat = "@": .Range("D:H").NumberFormat = "@"
        .Range("I:K").NumberFormat = "0": .Range("L:M").NumberFormat = "@": .Range("N:N").NumberFormat = "m/d/yyyy":
        .Range("O:O").NumberFormat = "@": .Range("C:C").NumberFormat = "m/d/yyyy"
    End With
    
    With Sheets(1).Range(Cells(1, 1), Cells(1, 22))
        .WrapText = TRUE
        .AutoFilter
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 42
        .Font.name = "Cambria"
        .Font.Size = 10
    End With
End Sub
Private Sub Formatirovan1()
    Dim format_area As Range
    lastrow = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    lastCol = Cells.Find(What:="*", SearchOrder:=xlColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
    Set format_area = Range(Cells(2, 1), Cells(lastrow, lastCol))
    format_area.Replace What:="" & Chr(10) & "", Replacement:=" ", SearchOrder:=xlByColumns
    With format_area
        .WrapText = TRUE
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .RowHeight = 40
        .Font.name = "Cambria"
        .Font.Size = 9
        .Value = Application.Trim(.Value)
    End With
    With Range(Cells(1, 1), Cells(lastrow, lastCol))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = TRUE
End Sub
Private Sub Naim_po_kontrakt()
    Dim Col         As Integer, rep_area As Range
    lastrow = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = Range("NaimKontr").Column
    Set rep_area = Range(Cells(2, Col), Cells(lastrow, Col))
    Select Case tip_spiska_IKK
        Case 1
            With rep_area
                .Replace What:="*руч*баз*комнат*", Replacement:="Кресло-коляска с ручным приводом базовая комнатная"
                .Replace What:="*руч*баз*прогул*", Replacement:="Кресло-коляска с ручным приводом базовая прогулочная"
                .Replace What:="*комнат*обод*колес*", Replacement:="Кресло-коляска с ручным приводом базовая комнатная"
                .Replace What:="*прогул*обод*колес*", Replacement:="Кресло-коляска с ручным приводом базовая прогулочная"
            End With
        Case 2
            With rep_area
                .Replace What:="*ДЦП*комнат*", Replacement:="Кресло-коляска с ручным приводом для больных ДЦП комнатная"
                .Replace What:="*ДЦП*прогул*", Replacement:="Кресло-коляска с ручным приводом для больных ДЦП прогулочная"
            End With
        Case 3
            With rep_area
                .Replace What:="*откидн*спин*комнат*", Replacement:="Кресло-коляска с ручным приводом с откидной спинкой комнатная"
                .Replace What:="*откидн*спин*прогул*", Replacement:="Кресло-коляска с ручным приводом с откидной спинкой прогулочная"
            End With
        Case 4
            With rep_area
                .Replace What:="*рычаж*", Replacement:="Кресло-коляска с рычажным приводом"
                .Replace What:="*", Replacement:="Кресло-коляска с рычажным приводом"
            End With
        Case 5
            With rep_area
                .Replace What:="*больш*вес*комнат*", Replacement:="Кресло-коляска с ручным приводом для лиц с большим весом комнатная"
                .Replace What:="*больш*вес*прогул*", Replacement:="Кресло-коляска с ручным приводом для лиц с большим весом прогулочная"
            End With
        Case 6
            With rep_area
                .Replace What:="*электропр*комнат*", Replacement:="Кресло-коляска с электроприводом комнатная"
                .Replace What:="*электропр*прогул*", Replacement:="Кресло-коляска с электроприводом прогулочная"
            End With
        Case 7
            With rep_area
                .Replace What:="*выс*ампут*", Replacement:=""
                .Replace What:="*малогабарит*", Replacement:=""
            End With
        Case 8
            With rep_area
                .Replace What:="*комнат*", Replacement:="Кресло-коляска для инв. СТАРТ с ручным приводом комнатная"
                .Replace What:="*прогул*", Replacement:="Кресло-коляска для инв. СТАРТ с ручным приводом прогулочная"
            End With
        Case 9
            With rep_area
                .Replace What:="*груз*", Replacement:="ПГ"
                .Replace What:="*повыш*", Replacement:="ПГ"
                .Replace What:="*актив*", Replacement:="А"
                .Replace What:="*без*кол*", Replacement:="БК"
                .Replace What:="*с*колес*", Replacement:="СК"
            End With
        Case Else
            rep_area.Replace What:="*", Replacement:=""
    End Select
End Sub
Private Sub Parametr_FIZLICO()
    Dim i           As Long, s As String, m As String, t As Integer, talia As String, ves As String, param As String
    Dim all         As Variant, a As Integer, b As Integer, c As Integer, f As Integer
    Dim re          As RegExp, Pattern As String, element As match, Col_match As MatchCollection
    Set re = New RegExp
    re.Global = TRUE
    re.IgnoreCase = TRUE
    lastrow = Range("FIO").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    'On Error Resume Next
    For i = 2 To lastrow
        re.Pattern = "(\b\d{1,3})\s?(-|\/)\s?\d{2,3}\s?(-|\/)\s?(\d{1,3}\b)"
        talia = "": ves = "": a = 0: b = 0: c = 0: f = 0
        If re.test(Range("V:V")(i).Value) = TRUE Then
            param = Replace(Replace(Range("V:V")(i).Value, "/", "-"), ",", "-")
            Set Col_match = re.Execute(param)
            For Each element In Col_match
                param = element.Value
            Next element
            all = Split(param, "-")
            a = all(0)
            b = all(1)
            c = all(2)
            If Val(a) < Val(b) Then
                f = a
            Else
                f = b
            End If
            If Val(c) < Val(f) Then
                ves = c
            Else
                ves = f
            End If
            If ves = c Then
                If Val(a) < Val(b) Then
                    talia = a
                Else
                    talia = b
                End If
            Else
                If f = b Then
                    If Val(a) < Val(c) Then
                        talia = a
                    Else
                        talia = c
                    End If
                Else
                    If Val(b) < Val(c) Then
                        talia = b
                    Else
                        talia = c
                    End If
                End If
            End If
            Range("OT")(i).Value = talia
            Range("Ves")(i).Value = ves
        Else
            re.Pattern = "В([.,/=-]|ЕС)?\s?[.,/=-]?\s?(ДО)?\s?\d+\b|\b\d+\s?КГ"
            s = ""
            If re.test(Range("V:V")(i).Value) = TRUE Then
                param = Range("V:V")(i).Value
                Set Col_match = re.Execute(param)
                For Each element In Col_match
                    param = element.Value
                Next element
                For t = 1 To Len(param) + 1
                    m = Mid(param, t, 1)
                    If m Like "[0-9]" Then
                        s = s & m
                    Else
                        ves = ves & s
                        s = ""
                    End If
                Next t
                Range("Ves")(i).Value = ves
            End If
            re.Pattern = "(О([-.,/=])?(Б(ЕДЕР|\.Б)?|[-.,/= ]Т(АЛИИ|\.)?|БЪЕМ)|V)([-.,/=])?\s?(Б(ЕДЕР|\.)|Т(АЛИИ)?)?\s?[-.,/=]?\s?\d+\b"
            s = ""
            If re.test(Range("V:V")(i).Value) = TRUE Then
                param = Range("V:V")(i).Value
                Set Col_match = re.Execute(param)
                For Each element In Col_match
                    param = element.Value
                Next element
                For t = 1 To Len(param) + 1
                    m = Mid(param, t, 1)
                    If m Like "[0-9]" Then
                        s = s & m
                    Else
                        talia = talia & s
                        s = ""
                    End If
                Next t
                Range("OT")(i).Value = talia
            End If
        End If
    Next i
End Sub
Private Sub data_naprav1()
    Dim message     As String, title As String, Col As Integer, t_f As Boolean, data As Variant
    Dim d           As Date
    lastrow = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = Range("DateNap").Column
    d = DateValue(Now)
    message = "Напишите дату направления " & "" & Chr(10) & "" & "(например: " & d & ")"
    title = "Дата направления"
    Wrong_Input:
    data = InputBox(message, title)
    t_f = IsDate(data)
    If data = "" Then
        'MsgBox "Дата направления не проставлена"
    ElseIf t_f = FALSE Then
        MsgBox "Некорректная дата"
        GoTo Wrong_Input
    Else
        Range(Cells(2, Col), Cells(lastrow, Col)).Value = DateValue(data)
    End If
End Sub
Private Sub Date_of_birth1()
    Dim i           As Long
    lastrow = Range("FIO").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    On Error Resume Next
    For i = 2 To lastrow
        If IsDate(Range("DateOfBirth")(i)) Then Range("DateOfBirth")(i) = DateValue(Range("DateOfBirth")(i).Value)
    Next i
End Sub
Private Sub Main_IKK()
    Dim get_range   As Range
    Dim Temp_Arr()  As Variant, Temp_arr1() As Variant, N As Integer, f As Integer, z As Long, p As Integer
    Dim h           As Integer, v As Long, mathes As Long, k As Long, first_row As Integer, LastRow2 As Long, Col As Integer
    Dim up_c        As Integer, low_c As Integer, up_r As Long, low_r As Integer        'верхние и нижние границы массива
    'Dim t ' таймер
    Dim snils       As Variant, tel As String
    Dim seriya_pas  As Integer, nomer_pas As Integer, date_pas As Integer, kemvydan_pas As Integer, all_pasport As Integer, nomerNap As Integer
    Dim re          As RegExp, Pattern As String
    Set re = New RegExp
    re.Global = TRUE
    re.IgnoreCase = TRUE
    't = timer
    lastrow = Range(Cells(1, 1), Cells(2500, 21)).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    lastCol = Range(Cells(1, 1), Cells(2000, 21)).Find(What:="*", SearchOrder:=xlColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
    first_row = Range(Cells(1, 1), Cells(2500, 26)).Find(What:="снилс", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = Range(Cells(1, 1), Cells(2500, 26)).Find(What:="снилс", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
    For v = first_row + 1 To lastrow
        If IsEmpty(Cells(v, Col)) Then
            first_row = first_row + 1
        Else: Exit For
        End If
    Next v
    re.Pattern = "(№|НОМ(ЕР)?)\s+(ЗАЯВ(КИ)?|НАПРАВ(ЛЕНИЯ)?)"
    nomerNap = 0
    For h = 1 To lastCol
        tel = Cells(first_row, h).Value
        If re.test(tel) = TRUE Then
            nomerNap = h
            Exit For
        End If
    Next h
    Set get_range = Range(Cells(first_row + 1, 1), Cells(lastrow, lastCol))
    Temp_arr1 = get_range.Resize(lastrow - first_row, lastCol)
    If nomerNap <> 0 Then Temp_arr1(1, nomerNap) = Temp_arr1(1, nomerNap) & "номер направления"
    LastRow2 = Sheets(1).Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    '---------Передача переменным границ массива1--------------
    up_c = UBound(Temp_arr1, 2)
    low_c = LBound(Temp_arr1, 2)
    up_r = UBound(Temp_arr1, 1)
    low_r = LBound(Temp_arr1, 1)
    '--------------------------------------------------------------------
    ' --------------Заполнение второго массива непустыми строками и столбцами------------
    For v = low_r To up_r
        mathes = 0
        For h = low_c To up_c
            If IsEmpty(Temp_arr1(v, h)) = TRUE Then
                mathes = mathes + 1
            ElseIf IsNumeric(Temp_arr1(v, h)) = TRUE And Val(Temp_arr1(v, h)) >= 1 And Val(Temp_arr1(v, h)) <= 20 Then mathes = mathes + 1
        End If
    Next h
    If mathes > 10 Then N = N + 1
Next v
For h = low_c To up_c
    mathes = 0
    For v = low_r To up_r
        If IsEmpty(Temp_arr1(v, h)) = TRUE Then mathes = mathes + 1
    Next v
    If mathes >= up_r - (up_r * 0.05) Then f = f + 1
Next h
ReDim Temp_Arr(1 To up_r - N, 1 To up_c - f)
N = 0: f = 0

For v = low_r To up_r
    f = 0
    mathes = 0
    For p = low_c To up_c
        If IsEmpty(Temp_arr1(v, p)) = TRUE Then
            mathes = mathes + 1
        ElseIf IsNumeric(Temp_arr1(v, p)) = TRUE And Val(Temp_arr1(v, p)) >= 1 And Val(Temp_arr1(v, p)) <= 20 Then mathes = mathes + 1
    End If
Next p
If mathes <= 10 Then
    N = N + 1
    For h = low_c To up_c
        mathes = 0
        For z = low_r To up_r
            If IsEmpty(Temp_arr1(z, h)) = TRUE Then mathes = mathes + 1
        Next z
        If mathes < up_r - (up_r * 0.05) Then
            f = f + 1
            Temp_Arr(N, f) = Temp_arr1(v, h)
        End If
    Next h
End If
Next v
Erase Temp_arr1
'--------------------------------------------------------------------
'---------Передача переменным границ массива2--------------
up_c = UBound(Temp_Arr, 2)
low_c = LBound(Temp_Arr, 2)
up_r = UBound(Temp_Arr, 1)
low_r = LBound(Temp_Arr, 1)
'--------------------------------------------------------------------
'    --------------Поиск ФИО------------
re.Pattern = "([А-Яа-я]+\s+[А-Яа-я]+\s+[А-Яа-я]+([Вв][Ии][Чч]|[Вв][Нн][Аа]|[Чч][Нн][Аа]|[Ьь][Ии][Чч]))"
For h = low_c To up_c
    mathes = 0
    For v = low_r To up_r
        If re.test(Temp_Arr(v, h)) = TRUE Then mathes = mathes + 1
        '            If UCase(RTrim(Right(Temp_arr(v, h), 15))) Like UCase("*вич") Then
        '            mathes = mathes + 1
        '            ElseIf UCase(RTrim(Right(Temp_arr(v, h), 15))) Like UCase("*вна") Then
        '            mathes = mathes + 1
        '            End If
    Next v
    If mathes >= up_r - Round(up_r * 0.15) Then        'округление применяется для очень коротких списков (меньше 5 элементов) для избежания ошибки
    For k = low_r To up_r
        Sheets(1).Range("FIO")(LastRow2 + k) = Temp_Arr(k, h)
    Next k
    Exit For
End If
Next h
'--------------------------------------------------------------------
'--------------Поиск Даты Рождения------------
For h = low_c To up_c
    mathes = 0
    For v = low_r To up_r
        If IsDate(Temp_Arr(v, h)) = TRUE Then
            If DateValue(Temp_Arr(v, h)) < 41640 And DateValue(Temp_Arr(v, h)) > 7306 Then mathes = mathes + 1        ' проверка на дату, а также чтобы дата была раньше 2013-го и позже 1920-го
        End If
    Next v
    If mathes >= up_r - Round(up_r * 0.15) Then
        For k = low_r To up_r
            Sheets(1).Range("DateOfBirth")(LastRow2 + k) = Temp_Arr(k, h)
        Next k
        Exit For
    End If
Next h
'--------------------------------------------------------------------
'--------------Поиск СНИЛС------------
For h = low_c To up_c
    mathes = 0
    For v = low_r To up_r        ' паттерн для поиска (^\d{11}$)|(^(\d{3})(\-)(\d{3})(\-)(\d{3})(\s)(\d{2})$)|(^\d{9}$)
        snils = Replace(Replace(Replace(Temp_Arr(v, h), "-", "", 1, 3), " ", "", 1, 3), "C", "")
        If IsNumeric(snils) = TRUE And Val(snils) > 100000000 Then mathes = mathes + 1
    Next v
    If mathes >= up_r - Round(up_r * 0.15) Then
        For k = low_r To up_r
            Sheets(1).Range("SNILS")(LastRow2 + k) = Temp_Arr(k, h)
        Next k
        snils = h
        Exit For
    End If
Next h
'--------------------------------------------------------------------
'--------------Поиск Адреса------------
'    re.Pattern = "КРАЙ|\sГ((\.|\s)|ОР(ОД\s|\.\s)?)|СТ(-|АНИ)ЦА|Х(\.|УТОР)|П(ОС(\.|ЕЛОК))|ПГТ|УЛ(ИЦА|(\.|\s))|ПР(ОСПЕКТ|ОЕЗД|\.)|БУЛЬВАР"
re.Pattern = "Моск(овская|ва)(\s|,)|Краснодар(ский)?(\s|,)|Твер(ь|ская)(\s|,)|Крым(,|\s)|Севастополь(,|\s)|Тул(а|ьская)(,|\s)|Санкт-Петербург(,|\s)"
For h = low_c To up_c
    mathes = 0
    For v = low_r To up_r
        If re.test(Temp_Arr(v, h)) = TRUE Then mathes = mathes + 1
    Next v
    If mathes >= up_r - Round(up_r * 0.2) Then
        For k = low_r To up_r
            Sheets(1).Range("Propiska")(LastRow2 + k) = Temp_Arr(k, h)
            Sheets(1).Range("Fakt")(LastRow2 + k) = Temp_Arr(k, h)
        Next k
        Exit For
    End If
Next h
'--------------------------------------------------------------------
'--------------Поиск Телефона------------
For h = low_c To up_c
    mathes = 0
    For v = low_r To up_r
        tel = Temp_Arr(v, h)
        re.Pattern = "(-|\s|\+|\(|\)|\/|[а-яА-Я]|\*)"
        tel = re.Replace(tel, "")
        '            re.Pattern = "((\b8)|([9][0-9][0-9]))+(\d{3})(\d{2})(\d{2}\b)|((\b\d{5})(\d{2}\b|\b))"
        re.Pattern = "[1-9]{1}[\s\-\(]?\(?\d{3}\)?[\s\.\-]?(\d{3}|\d{2})[\s\.\-]?(\d{4}|(\d{3}|\d{2}[\s\.\-]?\d{2}))"
        If re.test(tel) = TRUE Then mathes = mathes + 1
    Next v
    If mathes >= up_r - Round(up_r * 0.35) And h <> snils Then
        For k = low_r To up_r
            Sheets(1).Range("TEL")(LastRow2 + k) = Temp_Arr(k, h)
        Next k
        Exit For
    End If
Next h
'--------------------------------------------------------------------
'--------------Поиск ДопХарактеристики------------
For h = low_c To up_c
    mathes = 0
    For v = low_r To up_r
        If UCase(Temp_Arr(v, h)) Like UCase("*кресл*коляс*") Or Temp_Arr(v, h) Like "*ДЦП*" Or UCase(Temp_Arr(v, h)) Like UCase("*комнатн*") _
           Or UCase(Temp_Arr(v, h)) Like UCase("*прогулочн*") Or UCase(Temp_Arr(v, h)) Like UCase("*электроприв*") _
           Or UCase(Temp_Arr(v, h)) Like UCase("*рычаж*") Or UCase(Temp_Arr(v, h)) Like UCase("*крес*стул*") Then
        mathes = mathes + 1
    End If
Next v
If mathes >= up_r - Round(up_r * 0.25) Then
    For k = low_r To up_r
        Sheets(1).Range("DopChar")(LastRow2 + k) = Temp_Arr(k, h)
        Sheets(1).Range("NaimKontr")(LastRow2 + k) = Temp_Arr(k, h)
    Next k
    Exit For
End If
Next h
'--------------------------------------------------------------------
'--------------Количество------------
For k = low_r To up_r
    Sheets(1).Range("Kolvo")(LastRow2 + k) = 1
Next k
'--------------------------------------------------------------------
'--------------Поиск Филиала------------
re.Pattern = "^([1-9]|1\d|20)$"
For h = low_c To up_c
    mathes = 0
    If up_r = 1 Then
        If re.test(Temp_Arr(low_r, h)) = TRUE Then
            Sheets(1).Range("Fil")(LastRow2 + 1) = Temp_Arr(low_r, h): Exit For: End If
        Else
            For v = up_r To low_r Step -1
                If re.test(Temp_Arr(v, h)) = TRUE Then
                    Select Case v
                        Case Is > 1: If Val(Temp_Arr(v, h)) = Val(Temp_Arr(v - 1, h)) Then mathes = mathes + 1
                        Case Else: If Val(Temp_Arr(v, h)) = Val(Temp_Arr(v + 1, h)) Then mathes = mathes + 1
                    End Select
                End If
            Next v
            If mathes >= up_r - 1 - Round(up_r * 0.25) Then
                For k = low_r To up_r
                    Sheets(1).Range("Fil")(LastRow2 + k) = Temp_Arr(k, h)
                Next k
                Exit For
            End If
        End If
    Next h
    '--------------------------------------------------------------------
    '--------------Поиск №Направления------------------------------------
    For h = low_c To up_c
        If Temp_Arr(1, h) Like "*номер направления" Then
            For k = low_r To up_r
                If Val(Temp_Arr(k, h)) <> 0 Then Sheets(1).Range("NomerNap")(LastRow2 + k) = Val(Temp_Arr(k, h))
            Next k
        End If
    Next h
    '--------------------------------------------------------------------
    '--------------Поиск Примечание------------
    re.Pattern = "(\b\d{2,3})\s?(-|\/)\s?\d{2,3}\s?(-|\/)\s?(\d{2,3}\b)|Р([.,/=-]?|ОС[.,/=-]?Т?)\s?\d+|В([.,/=-]|ЕС)[.,/=-]?\s?(ДО)?\s?\d+|О([.,/=-])?(Т(АЛИИ)?|Б(ЕДЕР)?|БЪЕМ)([.,/=-])?\s?\d+"
    For h = up_c To low_c Step -1
        mathes = 0
        For v = low_r To up_r
            If re.test(Temp_Arr(v, h)) = TRUE Then mathes = mathes + 1
        Next v
        If mathes >= up_r - Round(up_r * 0.45) Then
            For k = low_r To up_r
                Sheets(1).Range("V:V")(LastRow2 + k) = Temp_Arr(k, h)
            Next k
            Exit For
        End If
    Next h
    '--------------------------------------------------------------------
    '--------------Поиск Паспорта в одной ячейке------------
    re.Pattern = "((\b\d{2}\s?(0|1|9)\d\b)|((^[1-4]+|[IVX]+|П)\s?-\s?[А-Я]{2}))\s?([А-Я ,./;:№]+)?\s?(0?0?\d{6})\s?([А-Я ,./;:]+)?\s?(\d{2}.(0|1)\d.((1|2)(9|0)\d{2}|\d{2})\b)?"
    For h = low_c To up_c
        mathes = 0
        For v = low_r To up_r
            If re.test(Temp_Arr(v, h)) = TRUE Then
                If Not UCase(Temp_Arr(v, h)) Like UCase("*вна *") And Not UCase(Temp_Arr(v, h)) Like UCase("*вич *") Then mathes = mathes + 1
            End If
        Next v
        If mathes >= up_r - Round(up_r * 0.2) Then
            For k = low_r To up_r
                Sheets(1).Range("Pasport")(LastRow2 + k) = Temp_Arr(k, h)
            Next k
            all_pasport = 1
            Exit For
        End If
    Next h
    '--------------------------------------------------------------------
    If all_pasport = 0 Then        ' поиск паспорта, если данные не в одной ячейке
    '---------серия паспорта------------
    re.Pattern = "((^\d{2}\s?(0|1|9)\d\b)|(^([1-4]+|[IVX]+|П)\s?-\s?[А-Я]{2}))"
    For h = up_c To low_c Step -1
        mathes = 0
        For v = low_r To up_r
            If re.test(Temp_Arr(v, h)) = TRUE Then mathes = mathes + 1
            '            If Trim(Temp_arr(v, h)) Like "##[ ,]##" Or Trim(Temp_arr(v, h)) Like "*-*АГ" Then
            '            mathes = mathes + 1
            '            ElseIf Trim(Temp_arr(v, h)) Like "####" Then
            '            mathes = mathes + 1
            '            End If
        Next v
        If mathes >= up_r - Round(up_r * 0.25) Then
            seriya_pas = h
            Exit For
        End If
    Next h
    '--------------------------------------------------------------------
    '---------номер паспорта------------
    For h = up_c To low_c Step -1
        mathes = 0
        For v = low_r To up_r
            If Len(Trim(Temp_Arr(v, h))) > 8 Then
            ElseIf Trim(Temp_Arr(v, h)) Like "######" Then
                mathes = mathes + 1
            ElseIf Len(Trim(Temp_Arr(v, h))) = 8 And Trim(Temp_Arr(v, h)) Like "00######" Then
                mathes = mathes + 1
            End If
        Next v
        If mathes >= up_r - Round(up_r * 0.25) Then
            nomer_pas = h
            Exit For
        End If
    Next h
    '--------------------------------------------------------------------
    '---------дата выдачи паспорта------------
    For h = up_c To seriya_pas Step -1
        mathes = 0
        For v = low_r To up_r
            If IsDate(Trim(Temp_Arr(v, h))) = TRUE And Temp_Arr(v, h) > 34700 Then mathes = mathes + 1
        Next v
        If mathes >= up_r - Round(up_r * 0.25) Then
            date_pas = h
            Exit For
        End If
    Next h
    '--------------------------------------------------------------------
    '---------кем выдан паспорт------------
    For h = up_c To low_c Step -1
        mathes = 0
        For v = low_r To up_r        'ОТД(\.|ЕЛ)?(ЕНИЕ(М)?|ОМ)?\s?МИЛИЦИИ|ГЕН(\.)КОНСУЛЬСТВО?|(ОТД(\.|ЕЛ)?(ОМ)?\s?)?((Р?Г?[УО](ПРАВЛЕНИЕ(М)?\s)В?(НУТРЕННИХ\s)?Д(ЕЛ)?)+|О?ЗАГС(А)?|О?[ОУ]?Ф(ЕДЕРАЛЬНОЙ\s)?М(ИГРАЦИОННОЙ\s)?С(ЛУЖБ(А|ОЙ|Ы))?|О?ПВС)
            If Temp_Arr(v, h) Like "*ОВД*" Or Temp_Arr(v, h) Like "*УВД*" Or Temp_Arr(v, h) Like "*УФМС*" Or Temp_Arr(v, h) Like "*ЗАГС*" Or Temp_Arr(v, h) Like UCase("*запис*актов*") _
               Or Temp_Arr(v, h) Like UCase("*внутрен*дел*") Or Temp_Arr(v, h) Like UCase("*миграц*служ*") Or Temp_Arr(v, h) Like "*ПОМ*" Or Temp_Arr(v, h) Like UCase("*отд*милиц*") Or Temp_Arr(v, h) Like UCase("*консульств*") Then
            If Not UCase(Temp_Arr(v, h)) Like UCase("*вна *") And Not UCase(Temp_Arr(v, h)) Like UCase("*вич *") Then mathes = mathes + 1
        End If
    Next v
    If mathes >= up_r - Round(up_r * 0.25) Then
        kemvydan_pas = h
        Exit For
    End If
Next h
'--------------------------------------------------------------------
If seriya_pas <> 0 And nomer_pas <> 0 And date_pas <> 0 And kemvydan_pas <> 0 Then
    On Error Resume Next
    For k = low_r To up_r
        Sheets(1).Range("Pasport")(LastRow2 + k) = Temp_Arr(k, seriya_pas) & " " & Temp_Arr(k, nomer_pas) & " " & Temp_Arr(k, date_pas) & " " & Temp_Arr(k, kemvydan_pas)
    Next k
Else
    On Error Resume Next
    For k = low_r To up_r
        Sheets(1).Range("Pasport")(LastRow2 + k) = Temp_Arr(k, seriya_pas) & " " & Temp_Arr(k, nomer_pas Or seriya_pas + 1) & " " & Temp_Arr(k, date_pas Or seriya_pas + 2) & " " & Temp_Arr(k, kemvydan_pas Or seriya_pas + 3)
    Next k
End If
End If
'MsgBox Format(timer - t, "00.000000000") & " секунд"
End Sub
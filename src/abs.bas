Option Explicit
Public times_do     As Integer
Public last_time_do As Integer
Public ws           As String
Public tip_spiska   As Integer, tip_obrabotki As Integer, pnumber As Long
Dim lastrow         As Long, LastRow2 As Long
Dim lastCol         As Integer

Sub Obrabotka_main()
    Dim Get_List    As Worksheet, Input_List As Worksheet
    'Dim t
    't = timer
    Set Get_List = Sheets(ws)
    Get_List.Activate
    If Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues) Is Nothing Then
        MsgBox "Лист: " & ws & " не содержит информации.": Exit Sub
    End If
    If Cells.Find(What:="снилс", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues) Is Nothing Then
        MsgBox "Колонка СНИЛС на вкладке " & ws & " не найдена!" & "" & Chr(10) & "" & "Данные с этой вкладки не будут обработанны." _
             & "" & Chr(10) & "" & "Озаглавьте столбец СНИЛС и попытайтесь еще раз.": times_do = 0: Exit Sub
    End If
    If Get_List.FilterMode = TRUE Then Get_List.ShowAllData
    Cells.UnMerge
    ActiveSheet.UsedRange.EntireRow.Hidden = FALSE
    Select Case tip_obrabotki
        Case 1
            Call New_List1
        Case 2
            If times_do = 1 Then Call New_List1
    End Select
    Set Input_List = Sheets(1)
    Get_List.Activate
    Call Main_ABS
    Select Case tip_obrabotki
        Case 1: last_time_do = 0
        Case 2: End Select
            If last_time_do = 0 Then
                Input_List.Activate
                ActiveSheet.DisplayPageBreaks = FALSE
                Call Nomer_Akta: Call Chahge_SNILS: Call Razmer
                Select Case tip_spiska
                    Case 1, 2, 4
                        Call Vpit_FIZLICO
                    Case 3, 5
                End Select
                Call Formatirovan1
                Call Date_of_birth1
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = TRUE
                Application.EnableEvents = TRUE
                Application.DisplayStatusBar = TRUE
                Call data_naprav1
                Select Case tip_spiska
                    Case 1
                        Call SummaPoVpit
                    Case Else
                End Select
                Select Case tip_obrabotki
                    Case 1
                        Application.Calculation = xlCalculationManual
                        Application.ScreenUpdating = FALSE
                        Application.EnableEvents = FALSE
                        Application.DisplayStatusBar = FALSE
                    Case 2: End Select
                    End If
End Sub

Private Sub Chahge_SNILS()
    Dim i As Long, Col As Integer, rep_area As Range
    lastrow = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = Range("SNILS").Column
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
    Dim filial As Integer, nedelya As Integer, i As Long
    lastrow = Range("FIO").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    'nedelya = DateDiff("w", "01.01." & Year(Now), Now) + 2
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

Private Sub SummaPoVpit()
    Dim a As Integer, b As Integer, i As Long, first_col As Integer, get_range As Range, Temp_Arr() As Variant, Last_col As Integer
    Dim Collect As New Collection, c As Long, Collect2 As New Collection, t As Long
    lastrow = Range("FIO").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    first_col = Range("Kolvo").Column
    Last_col = Range("Size").Column
    Set get_range = Range(Cells(2, first_col), Cells(lastrow, Last_col))
    Temp_Arr = get_range.Resize(lastrow - 1, Last_col - first_col + 1)
    On Error Resume Next
    For i = 1 To UBound(Temp_Arr, 1)
        Collect.Add Temp_Arr(i, 3), CStr(Temp_Arr(i, 2) & " " & Temp_Arr(i, 3))
        Collect2.Add Temp_Arr(i, 2), CStr(Temp_Arr(i, 2) & " " & Temp_Arr(i, 3))
    Next i
    b = Collect.count
    For a = 1 To b
        Range("Size")(lastrow + 1 + a).Value = Collect(a)
        Range("Vpit")(lastrow + 1 + a).Value = Collect2(a)
        c = 0
        For t = 1 To UBound(Temp_Arr, 1)
            If Temp_Arr(t, 3) = Collect(a) And Temp_Arr(t, 2) = Collect2(a) Then
                c = c + Val(Temp_Arr(t, 1))
            End If
        Next t
        Range("Kolvo")(lastrow + 1 + a).Value = c
    Next a
End Sub

Private Sub New_List1()
    Dim b As String, d As String, a As String, c As String, f As String, i As Integer
    Worksheets.Add Before:=Sheets(1)
    Select Case tip_spiska
        Case 1, 3: b = "ВпитФизлицо": d = "Впитываемость": a = "Номенклатура": i = 20: c = "Подгузники(впит.)Количество": f = "Размер"
        Case 2: b = "Впитываемость": d = "Номенклатура": a = "": i = 19: c = "Пеленки(впит.)Количество": f = "Размер"
        Case 4: b = "Впитываемость": d = "Номенклатура": a = "": i = 19: c = "ПрокладкиКоличество": f = "Пол"
        Case 5: b = "Впитываемость": d = "Номенклатура": a = "": i = 19: c = "ТрусыКоличество": f = "Размер"
        Case Else: b = "ВпитФизлицо": d = "Впитываемость": a = "Номенклатура": i = 20:: c = "Подгузники(впит.)Количество": f = "Размер"
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
        .Range("J:J").name = "Vpit": .Range("J:J").ColumnWidth = 8
        .Range("K:K").name = "Size": .Range("K:K").ColumnWidth = 7
        .Range("L:L").name = "Fil": .Range("L:L").ColumnWidth = 6.5
        .Range("M:M").name = "Pasport": .Range("M:M").ColumnWidth = 21
        .Range("N:N").name = "DateNap": .Range("N:N").ColumnWidth = 10
        .Range("O:O").name = "NomerNap": .Range("O:O").ColumnWidth = 7
        .Cells(1, 1).Value = "НомерАкта": .Cells(1, 2).Value = "ФИО": .Cells(1, 3).Value = "ДатаРождения"
        .Cells(1, 4).Value = "СНИЛС": .Cells(1, 5).Value = "АдресПоПрописке": .Cells(1, 6).Value = "АдресФактический"
        .Cells(1, 7).Value = "Телефон": .Cells(1, 8).Value = "ДопХарактеристики": .Cells(1, 9).Value = c
        .Cells(1, 10).Value = b: .Cells(1, 11).Value = f: .Cells(1, 12).Value = "Филиал"
        .Cells(1, 13).Value = "Паспорт": .Cells(1, 14).Value = "ДатаНаправления": .Cells(1, 15).Value = "НомерНаправления"
        .Cells(1, 16).Value = "НаименованиеПоКонтракту": .Cells(1, 17).Value = "ПаспортПредставителя"
        .Cells(1, 18).Value = "НасПункт": .Cells(1, 19).Value = d: .Cells(1, 20).Value = a
        .Range("A:B").NumberFormat = "@": .Range("D:H").NumberFormat = "@"
        .Range("I:I").NumberFormat = "0": .Range("K:M").NumberFormat = "@": .Range("N:N").NumberFormat = "m/d/yyyy":
        .Range("O:O").NumberFormat = "@": .Range("C:C").NumberFormat = "m/d/yyyy"
    End With
    
    With Sheets(1).Range(Cells(1, 1), Cells(1, i))
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

Private Sub Razmer()
    Dim Col As Integer, rep_area As Range
    lastrow = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = Range("Size").Column
    Set rep_area = Range(Cells(2, Col), Cells(lastrow, Col))
    Select Case tip_spiska
        Case 1, 5
            With rep_area
                .Replace What:="Seni", Replacement:=""
                .Replace What:="Super", Replacement:=""
                .Replace What:="*Extra Large*", Replacement:="XL"
                .Replace What:="*Extra Small*", Replacement:="XS"
                .Replace What:="*XS*", Replacement:="qqq"
                .Replace What:="*ХS*", Replacement:="qqq"
                .Replace What:="*X S*", Replacement:="qqq"
                .Replace What:="*Х S*", Replacement:="qqq"
                .Replace What:="*XL*", Replacement:="www"
                .Replace What:="*X L*", Replacement:="www"
                .Replace What:="*Х L*", Replacement:="www"
                .Replace What:="*ХL*", Replacement:="www"
                .Replace What:="*S*", Replacement:="S"
                .Replace What:="*L*", Replacement:="L"
                .Replace What:="*qqq*", Replacement:="XS"
                .Replace What:="*www*", Replacement:="XL"
                .Replace What:="*M*", Replacement:="M"
                .Replace What:="*М*", Replacement:="M"
            End With
        Case 2
            With rep_area
                .Replace What:=" ", Replacement:=""
                .Replace What:="*60*x60*", Replacement:="60x60"
                .Replace What:="*60*х60*", Replacement:="60x60"
                .Replace What:="*60*~*60*", Replacement:="60x60"
                .Replace What:="*60*x40*", Replacement:="60x40"
                .Replace What:="*60*х40*", Replacement:="60x40"
                .Replace What:="*60*~*40*", Replacement:="60x40"
                .Replace What:="*40*x60*", Replacement:="60x40"
                .Replace What:="*40*х60*", Replacement:="60x40"
                .Replace What:="*40*~*60*", Replacement:="60x40"
                .Replace What:="*90*x60*", Replacement:="60x90"
                .Replace What:="*90*х60*", Replacement:="60x90"
                .Replace What:="*90*~*60*", Replacement:="60x90"
                .Replace What:="*90*х60*", Replacement:="60x90"
                .Replace What:="*60*x90*", Replacement:="60x90"
                .Replace What:="*60*х90*", Replacement:="60x90"
                .Replace What:="*60*~*90*", Replacement:="60x90"
                .Replace What:="*пелен*", Replacement:=""
            End With
        Case 3
            With rep_area
                .Replace What:="*7*18*",        '7-18"
                .Replace What:="*11*25*",        '11-25"
                .Replace What:="*15*30*",        '15-30"
                .Replace What:="*4*9*",        '4-9"
            End With
        Case 4
            With rep_area
                .Replace What:="*жен*", Replacement:="Ж"
                .Replace What:="*муж*", Replacement:="М"
                .Replace What:="*вкладыш*", Replacement:=""
            End With
        Case Else
            rep_area.Replace What:="*", Replacement:=""
    End Select
End Sub

Private Sub Vpit_FIZLICO()
    Dim i As Long, nachpoz As Integer, s As String, m As String, t As Integer, vpit As Integer
    lastrow = Range("FIO").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    On Error Resume Next
    For i = 2 To lastrow
        vpit = 0
        If InStr(1, Range("DopChar")(i), "впит", vbTextCompare) <> 0 Then
            s = ""
            nachpoz = InStr(1, Range("DopChar")(i), "впит", vbTextCompare) + 4
            For t = nachpoz To Len(Range("DopChar")(i)) + 1
                m = Mid(Range("DopChar")(i), t, 1)
                If m Like "[0-9]" Then
                    s = s & m
                Else
                    If Val(s) > 200 And Val(s) < 10000 Then
                        vpit = s
                        Exit For
                    Else: s = ""
                    End If
                End If
            Next t
            Range("Vpit")(i).Value = vpit
        ElseIf InStr(1, Range("DopChar")(i), "мл", vbTextCompare) <> 0 Then
            s = ""
            nachpoz = InStr(1, Range("DopChar")(i), "мл", vbTextCompare)
            For t = nachpoz To 1 Step -1
                m = Mid(Range("DopChar")(i), t, 1)
                If m Like "[0-9]" Then
                    s = m & s
                Else
                    If Val(s) > 200 And Val(s) < 10000 Then
                        vpit = s
                        s = ""
                        Exit For
                    Else: s = ""
                    End If
                End If
            Next t
            Range("Vpit")(i).Value = vpit
        ElseIf InStr(1, Range("DopChar")(i), "влагопоглощение", vbTextCompare) <> 0 Then
            s = ""
            nachpoz = InStr(1, Range("DopChar")(i), "влагопоглощение", vbTextCompare) + 13
            For t = nachpoz To Len(Range("DopChar")(i).Value)
                m = Mid(Range("DopChar")(i), t, 1)
                If m Like "[0-9]" Then
                    s = s & m
                Else
                    If Val(s) > 200 And Val(s) < 10000 Then
                        vpit = s
                        s = ""
                        Exit For
                    Else: s = ""
                    End If
                End If
            Next t
            Range("Vpit")(i).Value = vpit
        End If
        If vpit = 0 Then
            s = ""
            For t = 1 To Len(Range("DopChar")(i)) + 1
                m = Mid(Range("DopChar")(i), t, 1)
                If m Like "[0-9]" Then
                    s = s & m
                Else
                    If Val(s) > 200 And Val(s) < 10000 Then
                        vpit = s
                        s = ""
                        Exit For
                    Else: s = ""
                    End If
                End If
            Next t
            Range("Vpit")(i).Value = vpit
        End If
    Next i
    Select Case tip_spiska
        Case 4
            For i = 2 To lastrow
                Select Case Range("Size")(i).Value
                    Case "М"
                        Select Case Range("Vpit")(i).Value
                            Case 1 To 150
                                Range("Vpit")(i).Value = 150
                            Case 151 To 300
                                Range("Vpit")(i).Value = 300
                            Case 301 To 600
                                Range("Vpit")(i).Value = 600
                        End Select
                    Case "Ж"
                        Select Case Range("Vpit")(i).Value
                            Case 1 To 300
                                Range("Vpit")(i).Value = 300
                            Case 301 To 500
                                Range("Vpit")(i).Value = 500
                            Case 501 To 800
                                Range("Vpit")(i).Value = 800
                        End Select
                    Case Else
                        Select Case Range("Vpit")(i).Value
                            Case 1 To 1200
                                Range("Vpit")(i).Value = 1200
                            Case 1201 To 1600
                                Range("Vpit")(i).Value = 1600
                            Case 1601 To 2000
                                Range("Vpit")(i).Value = 2000
                        End Select
                End Select
            Next i
    End Select
End Sub

Private Sub data_naprav1()
    Dim message As String, title As String, Col As Integer, t_f As Boolean, data As String
    Dim d As Date
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
        On Error Resume Next
        Range(Cells(2, Col), Cells(lastrow, Col)).Value = DateValue(data)
    End If
End Sub
Private Sub Date_of_birth1()
    Dim i As Long
    lastrow = Range("FIO").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    On Error Resume Next
    For i = 2 To lastrow
        If IsDate(Range("DateOfBirth")(i)) Then Range("DateOfBirth")(i) = DateValue(Range("DateOfBirth")(i).Value)
    Next i
End Sub

Private Sub Main_ABS()
    Dim get_range As Range
    Dim Temp_Arr() As Variant, Temp_arr1() As Variant, N As Integer, f As Integer, z As Long, p As Integer, h As Integer, v As Long, mathes As Long, k As Long
    Dim first_row As Integer, LastRow2 As Long, IPR_Col As Integer, Col As Integer, filial As Integer
    Dim up_c As Integer, low_c As Integer, up_r As Long, low_r As Integer        'верхние и нижние границы массива
    'Dim t ' таймер
    Dim snils As Variant, tel As String
    Dim seriya_pas As Integer, nomer_pas As Integer, date_pas As Integer, kemvydan_pas As Integer, all_pasport As Integer, nomerNap As Integer
    Dim re As RegExp, Pattern As String
    Set re = New RegExp
    re.Global = TRUE
    re.IgnoreCase = TRUE
    't = timer
    lastrow = Range(Cells(1, 1), Cells(10000, 21)).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    lastCol = Range(Cells(1, 1), Cells(2000, 21)).Find(What:="*", SearchOrder:=xlColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
    first_row = Range(Cells(1, 1), Cells(10000, 26)).Find(What:="снилс", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = Range(Cells(1, 1), Cells(10000, 26)).Find(What:="снилс", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
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
    If Range(Cells(1, 1), Cells(first_row + 3, 26)).Find(What:="ипр", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues) Is Nothing Then
        IPR_Col = 0
    Else
        For v = 1 To lastrow
            For h = 1 To lastCol
                If UCase(Cells(v, h)) Like UCase("*ипр*") Then
                    For k = v To lastrow
                        If Cells(k, h) Like "##" Or Cells(k, h) Like "#" Or Cells(k, h) Like "###" Then
                            IPR_Col = h
                            Exit For
                        End If
                    Next k
                End If
                If IPR_Col = h Then Exit For
            Next h
            If IPR_Col = h Then Exit For
        Next v
    End If
    Set get_range = Range(Cells(first_row + 1, 1), Cells(lastrow, lastCol))
    Temp_arr1 = get_range.Resize(lastrow - first_row, lastCol)
    If IPR_Col <> 0 Then Temp_arr1(1, IPR_Col) = "ИПР"
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
     re.Pattern = "([А-Яа-я]+\s+[А-Яа-я]+\s+[А-Яа-я]+([Вв][Ии][Чч]|[Вв][Нн][Аа]|[Чч][Нн][Аа]))"
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
         Select Case tip_spiska
             Case 1, 2, 4, 5
                 If IsDate(Temp_Arr(v, h)) = TRUE Then
                     If DateValue(Temp_Arr(v, h)) < 40544 And DateValue(Temp_Arr(v, h)) > 7306 Then mathes = mathes + 1        ' проверка на дату, а также чтобы дата была раньше 2005-го и позже 1920-го
                 End If
             Case Else
                 If IsDate(Temp_Arr(v, h)) = TRUE Then
                     If DateValue(Temp_Arr(v, h)) < 41640 And DateValue(Temp_Arr(v, h)) > 7306 Then mathes = mathes + 1        ' проверка на дату, а также чтобы дата была раньше 2013-го и позже 1920-го
                 End If
         End Select
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
     If mathes >= up_r - Round(up_r * 0.2) Then
         For k = low_r To up_r
             Sheets(1).Range("SNILS")(LastRow2 + k) = Temp_Arr(k, h)
         Next k
         snils = h
         Exit For
     End If
     snils = 0
 Next h
 '--------------------------------------------------------------------
 '--------------Поиск Адреса------------
 re.Pattern = "Моск(овская|ва)(\s|,)|Краснодар(ский)?(\s|,)|Твер(ь|ская)(\s|,)|Крым(,|\s)|Севастополь(,|\s)|Тул(а|ьская)(,|\s)|Санкт-Петербург(,|\s)|Калу(га|жская)(,|\s)"
 For h = low_c To up_c
     mathes = 0
     For v = low_r To up_r
         If re.test(Temp_Arr(v, h)) = TRUE Then mathes = mathes + 1
     Next v
     '            MsgBox mathes
     If mathes >= up_r - Round(up_r * 0.3) Then
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
 re.Pattern = "(ПОДГ([./,]|УЗНИК)?|ДЕТ(ЕЙ|СКИЕ|[.,])?|ПЕЛ(Е|Ё)НКИ|ПРОСТЫНИ|ВПИТ(\s?|\.|ЫВАЕМОСТЬ)?|УРОЛОГИЧ|ПРОКЛАД(КИ)?|ВКЛАДЫШ|ТРУС)"
 For h = low_c To up_c
     mathes = 0
     For v = low_r To up_r
         If re.test(Temp_Arr(v, h)) = TRUE Then mathes = mathes + 1
     Next v
     If mathes >= up_r - Round(up_r * 0.25) Then
         For k = low_r To up_r
             Sheets(1).Range("DopChar")(LastRow2 + k) = Temp_Arr(k, h)
             Sheets(1).Range("Size")(LastRow2 + k) = Temp_Arr(k, h)
         Next k
         Exit For
     End If
 Next h
 '--------------------------------------------------------------------
 '--------------Поиск Филиала-----------------------------------------
 re.Pattern = "^([1-9]|1\d|2\d|3\d|4\d)$"
 For h = up_c To low_c Step -1
     mathes = 0
     If up_r = 1 Then
         If re.test(Temp_Arr(low_r, h)) = TRUE Then
             Sheets(1).Range("Fil")(LastRow2 + 1) = Temp_Arr(low_r, h): filial = h: Exit For: End If
         Else
             For v = up_r To low_r Step -1
                 If re.test(Temp_Arr(v, h)) = TRUE Then
                     Select Case v
                         Case Is > 1: If Val(Temp_Arr(v, h)) = Val(Temp_Arr(v - 1, h)) Then mathes = mathes + 1
                         Case Else: If Val(Temp_Arr(v, h)) = Val(Temp_Arr(v + 1, h)) Then mathes = mathes + 1
                     End Select
                 End If
             Next v
             If mathes >= up_r - Round(up_r * 0.4) And Temp_Arr(1, h) <> "ИПР" Then
                 For k = low_r To up_r
                     Sheets(1).Range("Fil")(LastRow2 + k) = Temp_Arr(k, h)
                 Next k
                 filial = h
                 Exit For
             End If
         End If
     Next h
     '--------------------------------------------------------------------
     '--------------Поиск Количества--------------------------------------
     If snils <> 0 Then
         p = snils
     Else
         p = low_c
     End If
     re.Pattern = "^0$|^\d{2,3}$"
     For h = p To up_c
         mathes = 0
         If up_r = 1 Then
             If re.test(Temp_Arr(low_r, h)) = TRUE And Temp_Arr(1, h) <> "ИПР" And h <> filial Then
                 Sheets(1).Range("Kolvo")(LastRow2 + 1) = Temp_Arr(low_r, h): Exit For: End If
             Else
                 For v = up_r To low_r Step -1
                     If re.test(Temp_Arr(v, h)) = TRUE Then
                         Select Case v
                             Case Is > 1: If Val(Temp_Arr(v, h)) <> Val(Temp_Arr(v, h)) - 1 Or Val(Temp_Arr(v, h)) = 0 Then mathes = mathes + 1
                             Case Else: If Val(Temp_Arr(v, h)) <> Val(Temp_Arr(v, h)) + 1 Or Val(Temp_Arr(v, h)) = 0 Then mathes = mathes + 1
                         End Select
                     End If
                 Next v
                 If mathes >= up_r - Round(up_r * 0.25) And Temp_Arr(1, h) <> "ИПР" And h <> filial Then
                     For k = low_r To up_r
                         Sheets(1).Range("Kolvo")(LastRow2 + k) = Temp_Arr(k, h)
                     Next k
                     z = h
                     Exit For
                 End If
             End If
         Next h
         If z = 0 Then
             For h = up_c To low_c Step -1
                 mathes = 0
                 If up_r = 1 Then
                     If re.test(Temp_Arr(low_r, h)) = TRUE And Temp_Arr(1, h) <> "ИПР" And h <> filial Then
                         Sheets(1).Range("Kolvo")(LastRow2 + 1) = Temp_Arr(low_r, h): Exit For: End If
                     Else
                         For v = up_r To low_r Step -1
                             If re.test(Temp_Arr(v, h)) = TRUE Then
                                 Select Case v
                                     Case Is > 1: If Val(Temp_Arr(v, h)) <> Val(Temp_Arr(v, h)) - 1 Or Val(Temp_Arr(v, h)) = 0 Then mathes = mathes + 1
                                     Case Else: If Val(Temp_Arr(v, h)) <> Val(Temp_Arr(v, h)) + 1 Or Val(Temp_Arr(v, h)) = 0 Then mathes = mathes + 1
                                 End Select
                             End If
                         Next v
                         If mathes >= up_r - Round(up_r * 0.25) And Temp_Arr(1, h) <> "ИПР" And h <> filial Then
                             For k = low_r To up_r
                                 Sheets(1).Range("Kolvo")(LastRow2 + k) = Temp_Arr(k, h)
                             Next k
                             Exit For
                         End If
                     End If
                 Next h
             End If
             '--------------------------------------------------------------------
             '--------------Поиск №Направления------------------------------------
             For h = low_c To up_c
                 If Temp_Arr(1, h) Like "*номер направления" Or Temp_Arr(1, h) = "Направление" Then
                     For k = low_r To up_r
                         If Val(Temp_Arr(k, h)) <> 0 Then Sheets(1).Range("NomerNap")(LastRow2 + k) = Val(Temp_Arr(k, h))
                     Next k
                 End If
             Next h
             '--------------------------------------------------------------------
             '--------------Поиск Паспорта в одной ячейке------------
             re.Pattern = "(^|\s)(\d{4}|([1-4]+|[IVX]+|П)\s?-\s?[А-ЯA-Z]{2})\s(номер\s)?\d{6}"
             For h = low_c To up_c
                 mathes = 0
                 For v = low_r To up_r
                     If re.test(Temp_Arr(v, h)) = TRUE Then
                         If Temp_Arr(v, h) Like "*ОВД*" Or Temp_Arr(v, h) Like "*УВД*" Or Temp_Arr(v, h) Like "*УФМС*" Or Temp_Arr(v, h) Like "*ЗАГС*" Or Temp_Arr(v, h) Like UCase("*запис*актов*") _
                            Or Temp_Arr(v, h) Like UCase("*внутрен*дел*") Or Temp_Arr(v, h) Like UCase("*миграц*служ*") Or Temp_Arr(v, h) Like "*ПОМ*" Or Temp_Arr(v, h) Like UCase("*отд*милиц*") Or Temp_Arr(v, h) Like UCase("*консульств*") Then
                         If Not UCase(Temp_Arr(v, h)) Like UCase("*вна *") And Not UCase(Temp_Arr(v, h)) Like UCase("*вич *") Then mathes = mathes + 1
                     End If
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
         '---------серия и номер вместе------------
         re.Pattern = "(^|\s)(\d{4}|([1-4]+|[IVX]+|П)\s?-\s?[А-ЯA-Z]{2})\s(номер\s)?\d{6}"
         For h = up_c To low_c Step -1
             mathes = 0
             For v = low_r To up_r
                 If re.test(Temp_Arr(v, h)) = TRUE Then mathes = mathes + 1
             Next v
             If mathes >= up_r - Round(up_r * 0.25) Then
                 seriya_pas = h
                 all_pasport = 2
                 Exit For
             End If
         Next h
         '--------------------------------------------------------------------
         If all_pasport = 0 Then
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
                 On Error Resume Next
                 For v = low_r To up_r
                     If IsDate(Trim(Temp_Arr(v, h))) = TRUE And Temp_Arr(v, h) > 34700 Then mathes = mathes + 1
                 Next v
                 If mathes >= up_r - Round(up_r * 0.25) Then
                     date_pas = h
                     Exit For
                 End If
             Next h
             '--------------------------------------------------------------------
         End If
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
     If all_pasport = 2 Then
         On Error Resume Next
         For k = low_r To up_r
             Sheets(1).Range("Pasport")(LastRow2 + k) = Temp_Arr(k, seriya_pas) & " " & Temp_Arr(k, kemvydan_pas)
         Next k
     Else
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
 End If
 'MsgBox Format(timer - t, "00.000000000") & " секунд"
End Sub

Function Is_Like(txt As String, Pattern As String) As Boolean
    Is_Like = txt Like Pattern
End Function
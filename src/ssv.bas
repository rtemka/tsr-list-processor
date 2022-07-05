Option Explicit
Dim lastrow         As Long, LastRow2 As Long
Public tip_spiska_SSV As Integer
Dim lastCol         As Integer
Sub Obrabotka_SSV()
    Dim Get_List    As Worksheet, Input_List As Worksheet, response As Integer
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
    Call New_List3
    Set Input_List = Sheets(1)
    Get_List.Activate
    Call Main_SSV
    Input_List.Activate
    ActiveSheet.DisplayPageBreaks = FALSE
    Call Nomer_Akta: Call Chahge_SNILS: Call Naim_po_kontrakt
    If tip_spiska_SSV = 14 Then
        Call SredstvaKolvo
    End If
    Select Case tip_spiska_SSV
        Case 1, 2, 3, 4, 5, 6, 7, 15, 16, 17, 21, 22, 23
            Call Diametr_FIZLICO
        Case 8
            Call Diametr_FIZLICO: Call PolSamokat
    End Select
    Call Formatirovan1
    Call Date_of_birth1
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = TRUE
    Application.EnableEvents = TRUE
    Application.DisplayStatusBar = TRUE
    Call data_naprav1
End Sub
Private Sub Chahge_SNILS()
    Dim i           As Long, Col As Integer, rep_area As Range
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
    Dim filial      As Integer, nedelya As Integer, i As Long
    lastrow = Range("FIO").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    nedelya = WorksheetFunction.WeekNum(Now)
    If lastrow > 2 Then
        For i = 2 To lastrow
            If IsNumeric(Val(Range("fil")(i))) = TRUE And Val(Range("fil")(i)) = Val(Range("fil")(i + 1)) And Val(Range("fil")(i).Value) > 0 Then
                filial = Range("fil")(i)
                Exit For
            End If
        Next i
    Else
        filial = Range("fil")(2)
    End If
    For i = 2 To lastrow
        Range("Acts")(i).Value = nedelya & "/" & filial & "/"
    Next i
End Sub
Private Sub SredstvaKolvo()
    Dim i           As Long, message As String, isSet As Boolean
    lastrow = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    For i = 2 To lastrow
        isSet = FALSE
        If UCase(Range("DopChar")(i)) Like UCase("*защит*пленк*флак*") Then
            Range("Kolvo5")(i).Value = Range("Kolvo")(i).Value
            Range("Kolvo")(i).Clear
            isSet = TRUE
        End If
        If UCase(Range("DopChar")(i)) Like UCase("*защит*пленк*салф*") Then
            Range("Kolvo6")(i).Value = Range("Kolvo")(i).Value
            Range("Kolvo")(i).Clear
            isSet = TRUE
        End If
        If UCase(Range("DopChar")(i)) Like UCase("*паста*тубе*") Then
            Range("Kolvo4")(i).Value = Range("Kolvo")(i).Value
            Range("Kolvo")(i).Clear
            isSet = TRUE
        End If
        If UCase(Range("DopChar")(i)) Like UCase("*паста*полоск*") Then
            Range("Kolvo7")(i).Value = Range("Kolvo")(i).Value
            Range("Kolvo")(i).Clear
            isSet = TRUE
        End If
        If UCase(Range("DopChar")(i)) Like UCase("*очиститель*флакон*") Then
            Range("Kolvo2")(i).Value = Range("Kolvo")(i).Value
            Range("Kolvo")(i).Clear
            isSet = TRUE
        End If
        If UCase(Range("DopChar")(i)) Like UCase("*очиститель*салфетк*") Then
            Range("Kolvo")(i).Value = Range("Kolvo")(i).Value
            isSet = TRUE
        End If
        If UCase(Range("DopChar")(i)) Like UCase("*крем*защит*") Then
            Range("Kolvo7")(i).Value = Range("Kolvo")(i).Value
            Range("Kolvo")(i).Clear
            isSet = TRUE
        End If
        If UCase(Range("DopChar")(i)) Like UCase("*пудра*порошок*") Then
            Range("Kolvo8")(i).Value = Range("Kolvo")(i).Value
            Range("Kolvo")(i).Clear
            isSet = TRUE
        End If
        If UCase(Range("DopChar")(i)) Like UCase("*нейтрализатор*") Then
            Range("Kolvo9")(i).Value = Range("Kolvo")(i).Value
            Range("Kolvo")(i).Clear
            isSet = TRUE
        End If
        If Not isSet Then
            Range("Kolvo")(i).Value = "Не определено"
            message = "В некоторых позициях не удалось определить количество. Проверьте позиции с записью ""не определено"" в столбце количество"
        End If
    Next i
    
    If message <> "" Then
        MsgBox (message)
    End If
End Sub

Private Sub SummaPoDiametr()
    Dim a           As Integer, b As Integer, i As Long, first_col As Integer, get_range As Range, Temp_Arr() As Variant, Last_col As Integer
    Dim Collect     As New Collection, c As Long, t As Long, k As Long
    lastrow = Range("FIO").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    first_col = Range("Kolvo").Column
    Last_col = Range("Diametr").Column
    Set get_range = Range(Cells(2, first_col), Cells(lastrow, Last_col))
    Temp_Arr = get_range.Resize(lastrow - 1, Last_col - first_col + 1)
    On Error Resume Next
    For i = 1 To UBound(Temp_Arr, 1)
        If Temp_Arr(i, 3) = "" Then Temp_Arr(i, 3) = "нет Д"
    Next i
    For i = 1 To UBound(Temp_Arr, 1)
        Collect.Add Temp_Arr(i, 3), CStr(Temp_Arr(i, 3))
    Next i
    b = Collect.count
    Select Case tip_spiska_SSV
        Case 3, 4
            For a = 1 To b
                Range("Diametr")(lastrow + 1 + a).Value = Collect(a)
                c = 0
                k = 0
                For t = 1 To UBound(Temp_Arr, 1)
                    If Temp_Arr(t, 3) = Collect(a) Then
                        c = c + Val(Temp_Arr(t, 2))
                        k = k + Val(Temp_Arr(t, 1))
                    End If
                Next t
                Range("Kolvo2")(lastrow + 1 + a).Value = c
                Range("Kolvo")(lastrow + 1 + a).Value = k
            Next a
        Case Else
            For a = 1 To b
                Range("Diametr")(lastrow + 1 + a).Value = Collect(a)
                k = 0
                For t = 1 To UBound(Temp_Arr, 1)
                    If Temp_Arr(t, 3) = Collect(a) Then
                        k = k + Val(Temp_Arr(t, 1))
                    End If
                Next t
                Range("Kolvo")(lastrow + 1 + a).Value = k
            Next a
    End Select
End Sub
Private Sub New_List3()
    Dim a           As String, b As String, c As String, d As String, e As String, f As String, i As Integer
    Worksheets.Add Before:=Sheets(1)
    
    Select Case tip_spiska_SSV
        Case 13, 14
            i = 23
        Case 21
            i = 21
        Case Else
            i = 19
    End Select
    
    Select Case tip_spiska_SSV
        Case 1: a = "ОдноКомпКалоКоличество": b = "Диаметр": c = "ДиаметрСтомыФЛ"
        Case 2: a = "1КомпУроКоличество": b = "Диаметр": c = "ДиаметрСтомыФЛ"
        Case 3, 4: a = "ПластиныКоличество": b = "МешкиКоличество": c = "Диаметр"
        Case 5: a = "КатетердляуростомФолеяКоличество": c = "Диаметр"
        Case 6: a = "КатетерыПеццераКоличество": c = "Диаметр"
        Case 7: a = "КатетерНефростомаКоличество": c = "Диаметр"
        Case 8: a = "КатетерСамокатеризацияКоличество": b = "Пол": c = "Диаметр"
        Case 9: a = "Набордляс/кКоличество"
        Case 10: a = "НочнойМешокКоличество"
        Case 11: a = "ДневнойМешокКоличество"
        Case 12: a = "ДневнойМешокКоличество": b = "НочнойМешокКоличество": c = "РемешкиКоличество"
        Case 15: a = "ТампоныКоличество": c = "Диаметр"
        Case 16: a = "ТампоныДляСтомыКоличество": c = "Диаметр"
        Case 17: a = "УропрезервативыКоличество": c = "ДиаметрУропрезерватива"
        Case 18: a = "ПоясКоличество"
        Case 19: a = "РемешкиКоличество"
        Case 20: a = "НаборКоличество"
        Case 21: a = "УропрезервативыКоличество": b = "ДневнойМешокКоличество": c = "ДиаметрУропрезерватива": d = "НочнойМешокКоличество": e = "РемешкиКоличество"
        Case 22: a = "УретерокутанеостомаКоличество": c = "Диаметр"
        Case 23: a = "КомплектыКоличество": b = "": c = "Диаметр"
        Case Else: a = "ОдноКомпКалоКоличество": b = "Диаметр": c = "ДиаметрСтомыФЛ"
    End Select
    
    With Sheets(1)
        .Range("A:A").name = "Acts": .Range("A:A").ColumnWidth = 9.8: .Range("A:A").NumberFormat = "@"
        .Range("B:B").name = "FIO": .Range("B:B").ColumnWidth = 24
        .Range("C:C").name = "DateOfBirth": .Range("C:C").ColumnWidth = 10: .Range("C:C").NumberFormat = "m/d/yyyy"
        .Range("D:D").name = "SNILS": .Range("D:D").ColumnWidth = 14
        .Range("E:E").name = "Propiska": .Range("E:E").ColumnWidth = 20
        .Range("F:F").name = "Fakt": .Range("F:F").ColumnWidth = 20
        .Range("G:G").name = "TEL": .Range("G:G").ColumnWidth = 15
        .Range("H:H").name = "DopChar": .Range("H:H").ColumnWidth = 23
        .Cells(1, 1).Value = "НомерАкта": .Cells(1, 2).Value = "ФИО": .Cells(1, 3).Value = "ДатаРождения"
        .Cells(1, 4).Value = "СНИЛС": .Cells(1, 5).Value = "АдресПоПрописке": .Cells(1, 6).Value = "АдресФактический"
        .Cells(1, 7).Value = "Телефон": .Cells(1, 8).Value = "ДопХарактеристики"
    End With
    
    Select Case tip_spiska_SSV
        Case 13, 14
            With Sheets(1)
                .Range("I:I").name = "Kolvo": .Range("I:I").ColumnWidth = 6.5
                .Range("J:J").name = "Kolvo2": .Range("J:J").ColumnWidth = 6.5
                .Range("K:K").name = "Kolvo3": .Range("K:K").ColumnWidth = 6.5
                .Range("L:L").name = "Kolvo4": .Range("L:L").ColumnWidth = 6.5
                .Range("M:M").name = "Kolvo5": .Range("M:M").ColumnWidth = 6.5
                .Range("N:N").name = "Kolvo6": .Range("N:N").ColumnWidth = 6.5
                .Range("O:O").name = "Kolvo7": .Range("O:O").ColumnWidth = 6.5
                .Range("P:P").name = "Kolvo8": .Range("P:P").ColumnWidth = 6.5
                .Range("Q:Q").name = "Kolvo9": .Range("Q:Q").ColumnWidth = 6.5
                .Range("R:R").name = "Fil": .Range("R:R").ColumnWidth = 6.5
                .Range("S:S").name = "Pasport": .Range("S:S").ColumnWidth = 21
                .Range("T:T").name = "DateNap": .Range("T:T").ColumnWidth = 10
                .Range("U:U").name = "NomerNap": .Range("U:U").ColumnWidth = 7
                .Range("V:V").name = "NaimKontr": .Range("V:V").ColumnWidth = 13
                .Cells(1, 18).Value = "Филиал": .Cells(1, 19).Value = "Паспорт": .Cells(1, 20).Value = "ДатаНаправления"
                .Cells(1, 22).Value = "НаименованиеПоКонтракту": .Cells(1, 23).Value = "ПаспортПредставителя"
                .Cells(1, 24).Value = "Номенклатура": .Cells(1, 25).Value = "Наспункт": .Cells(1, 21).Value = "НомерНаправления"
                .Range("A:B").NumberFormat = "@": .Range("D:H").NumberFormat = "@"
                .Range("I:Q").NumberFormat = "0": .Range("R:S").NumberFormat = "@": .Range("T:T").NumberFormat = "m/d/yyyy":
                .Range("U:U").NumberFormat = "@"
                .Cells(1, 9).Value = "ОчистительКоличество"
                .Cells(1, 10).Value = "Очиститель(Флакон)Количество"
                .Cells(1, 11).Value = "ПастаКоличество"
                .Cells(1, 12).Value = "Паста(Втубе)Количество"
                .Cells(1, 13).Value = "Пленка(Спрей)Количество"
                .Cells(1, 14).Value = "ПленкаКоличество"
                .Cells(1, 15).Value = "КремКоличество"
                .Cells(1, 16).Value = "ПорошокКоличество"
                .Cells(1, 17).Value = "НейтрализаторКоличество"
            End With
        Case Else
            With Sheets(1)
                .Range("I:I").name = "Kolvo": .Range("I:I").ColumnWidth = 8
                .Range("J:J").name = "Kolvo2": .Range("J:J").ColumnWidth = 6.5
                .Range("K:K").name = "Diametr": .Range("K:K").ColumnWidth = 7
                .Range("L:L").name = "Fil": .Range("L:L").ColumnWidth = 6.5
                .Range("M:M").name = "Pasport": .Range("M:M").ColumnWidth = 21
                .Range("N:N").name = "DateNap": .Range("N:N").ColumnWidth = 10
                .Range("O:O").name = "NomerNap": .Range("O:O").ColumnWidth = 7
                .Range("P:P").name = "NaimKontr": .Range("P:P").ColumnWidth = 13
                .Cells(1, 12).Value = "Филиал": .Cells(1, 13).Value = "Паспорт": .Cells(1, 14).Value = "ДатаНаправления"
                .Cells(1, 16).Value = "НаименованиеПоКонтракту": .Cells(1, 17).Value = "ПаспортПредставителя"
                .Cells(1, 18).Value = "Номенклатура": .Cells(1, 19).Value = "Наспункт": .Cells(1, 15).Value = "НомерНаправления"
                .Cells(1, 9).Value = a: .Cells(1, 10).Value = b: .Cells(1, 11).Value = c
                .Range("A:B").NumberFormat = "@": .Range("D:H").NumberFormat = "@"
                .Range("I:K").NumberFormat = "0": .Range("L:M").NumberFormat = "@": .Range("N:N").NumberFormat = "m/d/yyyy":
                .Range("O:O").NumberFormat = "@"
                .Cells(1, 20).Value = d: .Cells(1, 21).Value = e: .Cells(1, 22).Value = f
            End With
    End Select
    
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
    On Error Resume Next
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
    Select Case tip_spiska_SSV
        Case 3, 4, 23
            With rep_area
                .Replace What:="*дву*ком*кало*", Replacement:="Двухкомпонентный дренируемый калоприемник"
                .Replace What:="*2*ком*кало*", Replacement:="Двухкомпонентный дренируемый калоприемник"
                .Replace What:="*кало*2*ком*", Replacement:="Двухкомпонентный дренируемый калоприемник"
                .Replace What:="*кало*дву*", Replacement:="Двухкомпонентный дренируемый калоприемник"
                .Replace What:="*дву*ком*уро*", Replacement:="Двухкомпонентный дренируемый уроприемник"
                .Replace What:="*2*ком*уро*", Replacement:="Двухкомпонентный дренируемый уроприемник"
                .Replace What:="*уро*2*ком*", Replacement:="Двухкомпонентный дренируемый уроприемник"
                .Replace What:="*дву*ком*моче*", Replacement:="Двухкомпонентный дренируемый уроприемник"
                .Replace What:="*моче*2*ком*", Replacement:="Двухкомпонентный дренируемый уроприемник"
                .Replace What:="*2*ком*моче*", Replacement:="Двухкомпонентный дренируемый уроприемник"
                .Replace What:="*уро*дву*", Replacement:="Двухкомпонентный дренируемый уроприемник"
                .Replace What:="*моче*дву*", Replacement:="Двухкомпонентный дренируемый уроприемник"
            End With
        Case Else
            rep_area.Replace What:="*", Replacement:=""
    End Select
End Sub
Private Sub Diametr_FIZLICO()
    Dim i           As Long, s As String, m As String, t As Integer, diametr As String, param As String
    Dim re          As RegExp, Pattern As String, element As match, Col_match As MatchCollection
    Set re = New RegExp
    re.Global = TRUE
    re.IgnoreCase = TRUE
    lastrow = Range("FIO").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    'On Error Resume Next
    For i = 2 To lastrow
        Select Case tip_spiska_SSV
            Case 5, 6, 7
                re.Pattern = "№(\s+)?\d{1,3}|(Д([-.,/=]?|М[-.,/=]?|ИАМ(ЕТР)?(\sСТОМЫ)?)?|D)[-.,/= ]?\d{1,3}|РАЗМЕР((\s+)|\-)?\d{1,3}"
            Case Else
                re.Pattern = "((Д\s?[-.,/=]?(СТ|М[-.,/=]?)|Д(ИАМ)?(\.)?(ЕТР)?(\s?СТО(\.)?(МЫ)?)?)|D)[-.,/= ]?\d{1,3}|№(\s+)?\d{1,3}|\b\d{1,3}(\s+)?ММ"
        End Select
        diametr = "": s = ""
        If re.test(Range("DopChar")(i).Value) = TRUE Then
            param = Range("DopChar")(i).Value
            Set Col_match = re.Execute(param)
            For Each element In Col_match
                param = element.Value
            Next element
            For t = 1 To Len(param) + 1
                m = Mid(param, t, 1)
                If m Like "[0-9]" Then
                    s = s & m
                Else
                    diametr = diametr & s
                    s = ""
                End If
            Next t
            If Val(diametr) > 0 Then
                Select Case tip_spiska_SSV
                    Case 1, 2, 3, 4, 23
                        Select Case Val(diametr)
                            Case 1 To 40: diametr = 40
                            Case 41 To 50: diametr = 50
                            Case 51 To 60: diametr = 60
                            Case 61 To 70: diametr = 70
                            Case 71 To 80: diametr = 80
                            Case 81 To 90: diametr = 90
                            Case 91 To 100: diametr = 100
                        End Select
                    Case 15
                        Select Case Val(diametr)
                            Case 1 To 37: diametr = 37
                            Case Is >= 38: diametr = 45
                        End Select
                    Case 16
                        Select Case Val(diametr)
                            Case 1 To 35: diametr = 35
                            Case Is >= 36: diametr = 45
                        End Select
                    Case 17, 21
                        Select Case Val(diametr)
                            Case 1 To 20: diametr = 20
                            Case 21 To 25: diametr = 25
                            Case 26 To 30: diametr = 30
                            Case 31 To 36: diametr = 35
                            Case 37 To 45: diametr = 40
                        End Select
                    Case 5, 6, 7, 8, 22
                        Select Case Val(diametr)
                            Case 1 To 9: diametr = 0 & diametr
                        End Select
                End Select
                Select Case tip_spiska_SSV
                    Case 1, 2
                        Range("Diametr")(i).Value = diametr
                    Case Else
                        Range("Diametr")(i).Value = "D" & diametr
                End Select
            Else
                Range("Diametr")(i).Value = "D0"
            End If
        End If
    Next i
End Sub
Private Sub PolSamokat()
    Dim i           As Long
    lastrow = Range("FIO").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    For i = 2 To lastrow
        If UCase(Range("FIO")(i).Value) Like UCase("*вич") Then
            Range("Kolvo2")(i).Value = "М"
        ElseIf UCase(Range("FIO")(i).Value) Like UCase("*вна") Or UCase(Range("FIO")(i).Value) Like UCase("*чна") Then
            Range("Kolvo2")(i).Value = "Ж"
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
        Range("DateOfBirth")(i) = DateValue(Range("DateOfBirth")(i).Value)
    Next i
End Sub
Private Sub Main_SSV()
    Dim get_range   As Range
    Dim Temp_Arr()  As Variant, Temp_arr1() As Variant, N As Integer, f As Integer, z As Long, p As Integer, dchar As Integer
    Dim h           As Integer, v As Long, mathes As Long, k As Long, first_row As Integer, LastRow2 As Long, Col As Integer, filial As Integer, IPR_Col As Integer
    Dim up_c        As Integer, low_c As Integer, up_r As Long, low_r As Integer        'верхние и нижние границы массива
    'Dim t ' таймер
    Dim snils       As Variant, tel As String
    Dim seriya_pas  As Integer, nomer_pas As Integer, date_pas As Integer, kemvydan_pas As Integer, all_pasport As Integer, nomerNap As Integer
    Dim re          As RegExp, Pattern As String
    Set re = New RegExp
    re.Global = TRUE
    re.IgnoreCase = TRUE
    't = timer
    lastrow = Range(Cells(1, 1), Cells(2500, 23)).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    lastCol = Range(Cells(1, 1), Cells(2000, 23)).Find(What:="*", SearchOrder:=xlColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
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
    ' --------------Особые случаи по стоме------------
    Select Case tip_spiska_SSV
        Case 12
            If Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*ноч*меш*", SearchDirection:=xlPrevious) Is Nothing Then
                p = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*меш*ноч*", SearchDirection:=xlPrevious).Column
            Else
                p = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*ноч*меш*", SearchDirection:=xlPrevious).Column
            End If
            If Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*дн*меш*", SearchDirection:=xlPrevious) Is Nothing Then
                z = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*меш*дн*", SearchDirection:=xlPrevious).Column
            Else
                z = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*дн*меш*", SearchDirection:=xlPrevious).Column
            End If
            k = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*ремеш*", SearchDirection:=xlPrevious).Column
            Temp_arr1(1, p) = Temp_arr1(1, p) & "ночной мешок"
            Temp_arr1(1, z) = Temp_arr1(1, z) & "дневной мешок"
            Temp_arr1(1, k) = Temp_arr1(1, k) & "ремешок"
        Case 21
            On Error Resume Next
            If Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*ноч*меш*", SearchDirection:=xlPrevious) Is Nothing Then
                p = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*меш*ноч*", SearchDirection:=xlPrevious).Column
            Else
                p = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*ноч*меш*").Column
            End If
            If Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*дн*меш*", SearchDirection:=xlPrevious) Is Nothing Then
                z = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*меш*дн*", SearchDirection:=xlPrevious).Column
            Else
                z = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*дн*меш*", SearchDirection:=xlPrevious).Column
            End If
            k = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*уропрез*", SearchDirection:=xlPrevious).Column
            v = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*ремеш*", SearchDirection:=xlPrevious).Column
            Temp_arr1(1, p) = Temp_arr1(1, p) & "ночной мешок"
            Temp_arr1(1, z) = Temp_arr1(1, z) & "дневной мешок"
            Temp_arr1(1, k) = Temp_arr1(1, k) & "уропрезерватив"
            Temp_arr1(1, v) = Temp_arr1(1, v) & "ремешок"
        Case 13
            On Error Resume Next
            p = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*очист*", SearchDirection:=xlPrevious).Column
            z = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*пленк*", SearchDirection:=xlPrevious).Column
            k = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*крем*", SearchDirection:=xlPrevious).Column
            v = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*паста*", SearchDirection:=xlPrevious).Column
            h = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*порошок*", SearchDirection:=xlPrevious).Column
            mathes = Range(Cells(1, 1), Cells(first_row, 21)).Find(What:="*нейтрал*", SearchDirection:=xlPrevious).Column
            Temp_arr1(1, p) = Temp_arr1(1, p) & "очиститель"
            Temp_arr1(1, z) = Temp_arr1(1, z) & "пленка"
            Temp_arr1(1, k) = Temp_arr1(1, k) & "крем защитный"
            Temp_arr1(1, v) = Temp_arr1(1, v) & "паста"
            Temp_arr1(1, h) = Temp_arr1(1, h) & "порошок"
            Temp_arr1(1, mathes) = Temp_arr1(1, mathes) & "нейтрализатор"
    End Select
    '--------------------------------------------------------------------
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
re.Pattern = "([А-Яа-яё]+\s+[А-Яа-яё]+\s+[А-Яа-яё]+([Вв][Ии][Чч]|[Вв][Нн][Аа]|[Чч][Нн][Аа]|[Ьь][Ии][Чч]))"
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
            If DateValue(Temp_Arr(v, h)) < 42005 And DateValue(Temp_Arr(v, h)) > 7306 Then mathes = mathes + 1        ' проверка на дату, а также чтобы дата была раньше 2013-го и позже 1920-го
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
re.Pattern = "^\d{2}-\d{2}|ДИАМЕТР|(ДВУХ|ОДНО)КОМПОНЕНТ|(КАЛО|УРО|МОЧЕ)(ПРИ(Е|Ё)М|ПРЕЗЕРВ)|УХОД|ИРРИГАЦ|КАТЕТЕР|ПОЯС|РЕМЕШ|ТАМПОН|ОЧИСТИТ|НЕЙТРАЛИЗ|ПАСТА"
For h = low_c To up_c
    mathes = 0
    For v = low_r To up_r
        If re.test(Temp_Arr(v, h)) = TRUE Then mathes = mathes + 1
    Next v
    If mathes >= up_r - Round(up_r * 0.25) Then
        For k = low_r To up_r
            Sheets(1).Range("DopChar")(LastRow2 + k) = Temp_Arr(k, h)
            Sheets(1).Range("NaimKontr")(LastRow2 + k) = Temp_Arr(k, h)
        Next k
        dchar = h
        Exit For
    End If
Next h
'--------------------------------------------------------------------
'--------------Поиск Филиала------------
re.Pattern = "^([1-9]|1\d|20)$"
For h = up_c To low_c Step -1
    mathes = 0
    If up_r = 1 Then
        If re.test(Temp_Arr(low_r, h)) = TRUE Then
            Sheets(1).Range("Fil")(LastRow2 + 1) = Temp_Arr(low_r, h)
            filial = h
            Exit For
        End If
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
'--------------Поиск №Направления------------------------------------
For h = low_c To up_c
    If Temp_Arr(1, h) Like "*номер направления" Then
        For k = low_r To up_r
            If Val(Temp_Arr(k, h)) <> 0 Then Sheets(1).Range("NomerNap")(LastRow2 + k) = Val(Temp_Arr(k, h))
        Next k
        nomerNap = h
    End If
Next h
'--------------------------------------------------------------------
'--------------Поиск Количества------------
Select Case tip_spiska_SSV
    Case 1, 2, 5, 6, 7, 8, 9, 10, 11, 14, 15, 16, 17, 18, 19, 20, 22, 23
        re.Pattern = "^0$|^\d{1,3}$"
        For h = up_c To low_c Step -1
            mathes = 0
            If up_r = 1 Then
                If re.test(Temp_Arr(low_r, h)) = TRUE And Temp_Arr(1, h) <> "ИПР" And h <> filial And h <> nomerNap Then
                    Sheets(1).Range("Kolvo")(LastRow2 + 1) = Temp_Arr(low_r, h): Exit For: End If
                Else
                    For v = up_r To low_r Step -1
                        If re.test(Temp_Arr(v, h)) = TRUE Then
                            If tip_spiska_SSV <> 14 Then
                                Select Case v
                                    Case Is > 1: If Val(Temp_Arr(v, h)) - 1 <> Val(Temp_Arr(v - 1, h)) Or Val(Temp_Arr(v, h)) = 0 Then mathes = mathes + 1
                                    Case Else: If Val(Temp_Arr(v, h)) + 1 <> Val(Temp_Arr(v + 1, h)) Or Val(Temp_Arr(v, h)) = 0 Then mathes = mathes + 1
                                End Select
                            Else
                                mathes = mathes + 1
                            End If
                        End If
                    Next v
                    If tip_spiska_SSV <> 14 Then
                        If mathes >= up_r - Round(up_r * 0.35) And Temp_Arr(1, h) <> "ИПР" And h <> filial And h <> nomerNap Then
                            For k = low_r To up_r
                                Sheets(1).Range("Kolvo")(LastRow2 + k) = Temp_Arr(k, h)
                            Next k
                            Exit For
                        End If
                    Else
                        If mathes >= up_r - Round(up_r * 0.35) And Temp_Arr(1, h) <> "ИПР" And h <> nomerNap Then
                            For k = low_r To up_r
                                Sheets(1).Range("Kolvo")(LastRow2 + k) = Temp_Arr(k, h)
                            Next k
                            Exit For
                        End If
                    End If
                End If
            Next h
        Case 3, 4
            re.Pattern = "^0$|^\d{2,3}$"
            N = 0: f = 0
            For z = 1 To 2
                For h = up_c To low_c Step -1
                    mathes = 0
                    If up_r = 1 Then
                        If re.test(Temp_Arr(low_r, h)) = TRUE And Temp_Arr(1, h) <> "ИПР" And h <> filial And h <> nomerNap And N = 0 Then
                            N = h
                            Exit For
                        ElseIf re.test(Temp_Arr(low_r, h)) = TRUE And Temp_Arr(1, h) <> "ИПР" And h <> filial And h <> nomerNap And N <> 0 And N <> h Then
                            f = h
                            Exit For
                        End If
                    Else
                        For v = up_r To low_r Step -1
                            If re.test(Temp_Arr(v, h)) = TRUE Then
                                Select Case v
                                    Case Is > 1: If Val(Temp_Arr(v, h)) - 1 <> Val(Temp_Arr(v - 1, h)) Or Val(Temp_Arr(v, h)) = 0 Then mathes = mathes + 1
                                    Case Else: If Val(Temp_Arr(v, h)) + 1 <> Val(Temp_Arr(v + 1, h)) Or Val(Temp_Arr(v, h)) = 0 Then mathes = mathes + 1
                                End Select
                            End If
                        Next v
                        If mathes >= up_r - Round(up_r * 0.25) And Temp_Arr(1, h) <> "ИПР" And h <> filial And h <> nomerNap And N = 0 Then
                            N = h
                            Exit For
                        ElseIf mathes >= up_r - Round(up_r * 0.25) And Temp_Arr(1, h) <> "ИПР" And h <> filial And h <> nomerNap And N <> 0 And N <> h Then
                            f = h
                            Exit For
                        End If
                    End If
                Next h
            Next z
            mathes = 0
            For k = low_r To up_r
                If Temp_Arr(k, N) > Temp_Arr(k, f) Then mathes = mathes + 1
            Next k
            If mathes >= up_r - Round(up_r * 0.5) Then
                For k = low_r To up_r
                    Sheets(1).Range("Kolvo")(LastRow2 + k) = Temp_Arr(k, f)
                    Sheets(1).Range("Kolvo2")(LastRow2 + k) = Temp_Arr(k, N)
                Next k
            Else
                For k = low_r To up_r
                    Sheets(1).Range("Kolvo")(LastRow2 + k) = Temp_Arr(k, N)
                    Sheets(1).Range("Kolvo2")(LastRow2 + k) = Temp_Arr(k, f)
                Next k
            End If
        Case 12
            For h = up_c To low_c Step -1
                If Temp_Arr(1, h) Like "*ночной мешок" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("Kolvo2")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                ElseIf Temp_Arr(1, h) Like "*дневной мешок" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("Kolvo")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                ElseIf Temp_Arr(1, h) Like "*ремешок" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("Diametr")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                End If
            Next h
        Case 21
            For h = up_c To low_c Step -1
                If Temp_Arr(1, h) Like "*ночной мешок" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("T:T")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                ElseIf Temp_Arr(1, h) Like "*дневной мешок" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("Kolvo2")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                ElseIf Temp_Arr(1, h) Like "*уропрезерватив" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("Kolvo")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                ElseIf Temp_Arr(1, h) Like "*ремешок" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("U:U")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                End If
            Next h
        Case 13
            On Error Resume Next
            For h = up_c To low_c Step -1
                If Temp_Arr(1, h) Like "*очиститель" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("Kolvo")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                ElseIf Temp_Arr(1, h) Like "*крем защитный" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("T:T")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                ElseIf Temp_Arr(1, h) Like "*паста" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("Kolvo2")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                ElseIf Temp_Arr(1, h) Like "*порошок" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("U:U")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                ElseIf Temp_Arr(1, h) Like "*пленка" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("Diametr")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                ElseIf Temp_Arr(1, h) Like "*нейтрализатор" And h <> dchar Then
                    For k = low_r To up_r
                        Sheets(1).Range("V:V")(LastRow2 + k) = Val(Temp_Arr(k, h))
                    Next k
                End If
            Next h
    End Select
    '--------------------------------------------------------------------
    '--------------Поиск Примечание------------
    re.Pattern = "БЕССРОЧНО|Б\/(С|П)|((Д\s?[-.,/=]?(СТ|М[-.,/=]?)|ДИАМ(\.)?(ЕТР)?(\s?СТО(\.)?(МЫ)?)?)|D)[-.,/= ]?\d{1,3}|№(\s+)?\d{1,3}(\s|$)|\b\d{1,3}(\s+)?ММ"
    For h = up_c To low_c Step -1
        mathes = 0
        For v = low_r To up_r
            If re.test(Temp_Arr(v, h)) = TRUE Then mathes = mathes + 1
        Next v
        If mathes >= up_r - Round(up_r * 0.45) And dchar <> h Then
            For k = low_r To up_r
                Sheets(1).Range("DopChar")(LastRow2 + k) = Sheets(1).Range("DopChar")(LastRow2 + k).Value & " " & Temp_Arr(k, h)
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
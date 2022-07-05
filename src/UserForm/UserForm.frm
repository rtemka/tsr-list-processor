Option Explicit
Dim sheet           As Worksheet
Private Sub CommandButton1_Click()
    Dim i           As Integer, p As Integer, k As Integer, response As Integer
    Dim ss()        As Variant
    For i = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) = TRUE Then
            p = p + 1
        End If
    Next i
    If p = 0 Then
        MsgBox "Не выбран ни один лист из списка"
        Exit Sub: End If
        response = MsgBox("Начать обработку?", vbOKCancel, "Обработка списка")
        If response = vbCancel Then Exit Sub
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = FALSE
        Application.EnableEvents = FALSE
        ActiveSheet.DisplayPageBreaks = FALSE
        Application.DisplayStatusBar = FALSE
        Me.Hide
        ReDim ss(1 To p)
        For i = 0 To Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = TRUE Then
                k = k + 1
                ss(k) = ListBox1.List(i)
            End If
        Next i
        last_time_do = 0
        times_do = 0
        pnumber = 0
        For i = 1 To p
            last_time_do = p - i
            times_do = times_do + 1
            ws = ss(i)
            Call Obrabotka_main
        Next i
        Unload Me
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = TRUE
        Application.EnableEvents = TRUE
        Application.DisplayStatusBar = TRUE
        'MsgBox "Обработка выполнена"
    End Sub
    Private Sub CloseButton_Click()
        Unload UserForm1
    End Sub
    Private Sub KinderButton_Click()
        Me.ListBox1.Clear
        For Each sheet In ActiveWorkbook.Sheets
            If Not UCase(sheet.name) Like "*[SLM]*" And Not sheet.name Like "*М*" And Not sheet.name Like "*пелен*" And Not sheet.name Like "*60*#0*" _
               And Not sheet.name Like "*40*60*" Then
            Me.ListBox1.AddItem (sheet.name)
        End If
    Next sheet
    tip_spiska = 3
End Sub
Private Sub PantsButton_Click()
    Me.ListBox1.Clear
    For Each sheet In ActiveWorkbook.Sheets
        Me.ListBox1.AddItem (sheet.name)
    Next sheet
    tip_spiska = 5
End Sub

Private Sub PelenkiButton_Click()
    Me.ListBox1.Clear
    For Each sheet In ActiveWorkbook.Sheets
        If Not sheet.name Like "*дет*" And Not UCase(sheet.name) Like "*[SLM]*" And Not sheet.name Like "*7*18*" And Not sheet.name Like "*11*25*" _
           And Not sheet.name Like "*15*30*" And Not sheet.name Like "*4*9*" And Not sheet.name Like "*М*" Then
        Me.ListBox1.AddItem (sheet.name)
    End If
Next sheet
tip_spiska = 2
End Sub
Private Sub ProkladkiButton_Click()
    Me.ListBox1.Clear
    For Each sheet In ActiveWorkbook.Sheets
        Me.ListBox1.AddItem (sheet.name)
    Next sheet
    tip_spiska = 4
End Sub

Private Sub VzrosButton_Click()
    Me.ListBox1.Clear
    For Each sheet In ActiveWorkbook.Sheets
        If Not sheet.name Like "*дет*" And Not UCase(sheet.name) Like "*пелен*" And Not sheet.name Like "*7*18*" And Not sheet.name Like "*11*25*" _
           And Not sheet.name Like "*15*30*" And Not sheet.name Like "*4*9*" And Not sheet.name Like "*60*#0*" And Not sheet.name Like "*40*60*" Then
        Me.ListBox1.AddItem (sheet.name)
    End If
Next sheet
tip_spiska = 1
End Sub
Private Sub ListBox1_Click()
    
End Sub
Private Sub OneListButton_Click()
    tip_obrabotki = 1
End Sub
Private Sub SeveralListsButton_Click()
    tip_obrabotki = 2
End Sub
Private Sub UserForm_Initialize()
    For Each sheet In ActiveWorkbook.Sheets
        Me.ListBox1.AddItem (sheet.name)
    Next sheet
    tip_obrabotki = 2
    tip_spiska = 1
End Sub
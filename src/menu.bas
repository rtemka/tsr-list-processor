Option Explicit
Sub Add_Sub_Menu()
    Dim Bar         As CommandBar
    Dim New_Menu    As CommandBarPopup
    Dim New_Submenu As CommandBarButton, IKK_Popup As CommandBarPopup, SSV_Popup As CommandBarPopup
    Dim IKK_Submenu As CommandBarButton, SSV_Submenu As CommandBarButton, Kateter_Popup As CommandBarPopup, Kateter_Submenu As CommandBarButton
    Dim Odnokomp_Popup As CommandBarPopup, Odnokomp_Submenu As CommandBarButton, Twicekomp_Popup As CommandBarPopup, Twicekomp_Submenu As CommandBarButton
    Dim Tampon_Popup As CommandBarPopup, Tampon_Submenu As CommandBarButton, Ukhod_Popup As CommandBarPopup, Ukhod_Submenu As CommandBarButton
    Dim Meshki_Popup As CommandBarPopup, Meshki_Submenu As CommandBarButton
    Delete_Submenu
    Set Bar = CommandBars("Cell")
    Set New_Menu = Bar.Controls.Add(Type:=msoControlPopup, temporary:=True)
    New_Menu.Caption = "Об&работать список ТСР"
    New_Menu.BeginGroup = TRUE
    Set New_Submenu = New_Menu.Controls.Add(Type:=msoControlButton)
    With New_Submenu
        .Caption = "&АБС"
        .OnAction = "Show_Form_ABS"
    End With
    Set IKK_Popup = New_Menu.Controls.Add(Type:=msoControlPopup, temporary:=True)
    IKK_Popup.Caption = "&ИКК"
    IKK_Popup.BeginGroup = TRUE
    Set SSV_Popup = New_Menu.Controls.Add(Type:=msoControlPopup, temporary:=True)
    SSV_Popup.Caption = "&ССВ"
    SSV_Popup.BeginGroup = TRUE
    Set IKK_Submenu = IKK_Popup.Controls.Add(Type:=msoControlButton)
    With IKK_Submenu
        .Caption = "Коляски &Базовые"
        .OnAction = "Baza"
    End With
    Set IKK_Submenu = IKK_Popup.Controls.Add(Type:=msoControlButton)
    With IKK_Submenu
        .Caption = "Коляски &ДЦП"
        .OnAction = "DCP"
    End With
    Set IKK_Submenu = IKK_Popup.Controls.Add(Type:=msoControlButton)
    With IKK_Submenu
        .Caption = "Коляски с &Откидной Спинкой"
        .OnAction = "Otkid_Spin"
    End With
    Set IKK_Submenu = IKK_Popup.Controls.Add(Type:=msoControlButton)
    With IKK_Submenu
        .Caption = "Коляски &Повышенной Груз."
        .OnAction = "Povis_Gruz"
    End With
    Set IKK_Submenu = IKK_Popup.Controls.Add(Type:=msoControlButton)
    With IKK_Submenu
        .Caption = "Коляски с &Рычажным Приводом"
        .OnAction = "Ruchazka"
    End With
    Set IKK_Submenu = IKK_Popup.Controls.Add(Type:=msoControlButton)
    With IKK_Submenu
        .Caption = "Коляски с &Электроприводом"
        .OnAction = "Elektro"
    End With
    Set IKK_Submenu = IKK_Popup.Controls.Add(Type:=msoControlButton)
    With IKK_Submenu
        .Caption = "Коляски &Малогабаритные"
        .OnAction = "Telezka"
    End With
    Set IKK_Submenu = IKK_Popup.Controls.Add(Type:=msoControlButton)
    With IKK_Submenu
        .Caption = "Коляски От&тоБок"
        .OnAction = "OttoBok"
    End With
    Set IKK_Submenu = IKK_Popup.Controls.Add(Type:=msoControlButton)
    With IKK_Submenu
        .Caption = "Санитарные Кресла-&Стулья"
        .OnAction = "Stuliya"
    End With
    Set Odnokomp_Popup = SSV_Popup.Controls.Add(Type:=msoControlPopup, temporary:=True)
    Odnokomp_Popup.Caption = "&Однокомпонентные"
    Odnokomp_Popup.BeginGroup = TRUE
    Set Odnokomp_Submenu = Odnokomp_Popup.Controls.Add(Type:=msoControlButton)
    With Odnokomp_Submenu
        .Caption = "Кало"
        .OnAction = "Kalo1"
    End With
    Set Odnokomp_Submenu = Odnokomp_Popup.Controls.Add(Type:=msoControlButton)
    With Odnokomp_Submenu
        .Caption = "Уро"
        .OnAction = "Uro1"
    End With
    Set Twicekomp_Popup = SSV_Popup.Controls.Add(Type:=msoControlPopup, temporary:=True)
    Twicekomp_Popup.Caption = "&Двухкомпонентные"
    Twicekomp_Popup.BeginGroup = TRUE
    Set Twicekomp_Submenu = Twicekomp_Popup.Controls.Add(Type:=msoControlButton)
    With Twicekomp_Submenu
        .Caption = "Кало"
        .OnAction = "Kalo2"
    End With
    Set Twicekomp_Submenu = Twicekomp_Popup.Controls.Add(Type:=msoControlButton)
    With Twicekomp_Submenu
        .Caption = "Уро"
        .OnAction = "Uro2"
    End With
    Set Twicekomp_Submenu = Twicekomp_Popup.Controls.Add(Type:=msoControlButton)
    With Twicekomp_Submenu
        .Caption = "Комплекты"
        .OnAction = "Copmlekt2"
    End With
    Set Kateter_Popup = SSV_Popup.Controls.Add(Type:=msoControlPopup, temporary:=True)
    Kateter_Popup.Caption = "&Катетеры"
    Kateter_Popup.BeginGroup = TRUE
    Set Kateter_Submenu = Kateter_Popup.Controls.Add(Type:=msoControlButton)
    With Kateter_Submenu
        .Caption = "Фолея"
        .OnAction = "Folea"
    End With
    Set Kateter_Submenu = Kateter_Popup.Controls.Add(Type:=msoControlButton)
    With Kateter_Submenu
        .Caption = "Пеццера"
        .OnAction = "Peccera"
    End With
    Set Kateter_Submenu = Kateter_Popup.Controls.Add(Type:=msoControlButton)
    With Kateter_Submenu
        .Caption = "Нефростома"
        .OnAction = "Nefrostoma"
    End With
    Set Kateter_Submenu = Kateter_Popup.Controls.Add(Type:=msoControlButton)
    With Kateter_Submenu
        .Caption = "Самокатетеризация"
        .OnAction = "Samokat"
    End With
    Set Kateter_Submenu = Kateter_Popup.Controls.Add(Type:=msoControlButton)
    With Kateter_Submenu
        .Caption = "Уретерокутанеостома"
        .OnAction = "Ureterocutaneo"
    End With
    Set Kateter_Submenu = Kateter_Popup.Controls.Add(Type:=msoControlButton)
    With Kateter_Submenu
        .Caption = "Наборы с/к"
        .OnAction = "Samokat_nabor"
    End With
    Set Meshki_Popup = SSV_Popup.Controls.Add(Type:=msoControlPopup, temporary:=True)
    Meshki_Popup.Caption = "&Мешки"
    Meshki_Popup.BeginGroup = TRUE
    Set Meshki_Submenu = Meshki_Popup.Controls.Add(Type:=msoControlButton)
    With Meshki_Submenu
        .Caption = "Дневные"
        .OnAction = "Meshki_D"
    End With
    Set Meshki_Submenu = Meshki_Popup.Controls.Add(Type:=msoControlButton)
    With Meshki_Submenu
        .Caption = "Ночные"
        .OnAction = "Meshki_N"
    End With
    Set Meshki_Submenu = Meshki_Popup.Controls.Add(Type:=msoControlButton)
    With Meshki_Submenu
        .Caption = "Наборы мочеприемные"
        .OnAction = "NaborMoche"
    End With
    Set Meshki_Submenu = Meshki_Popup.Controls.Add(Type:=msoControlButton)
    With Meshki_Submenu
        .Caption = "Урокомплекты"
        .OnAction = "UroComplect"
    End With
    Set Ukhod_Popup = SSV_Popup.Controls.Add(Type:=msoControlPopup, temporary:=True)
    Ukhod_Popup.Caption = "&Средства ухода"
    Ukhod_Popup.BeginGroup = TRUE
    Set Ukhod_Submenu = Ukhod_Popup.Controls.Add(Type:=msoControlButton)
    With Ukhod_Submenu
        .Caption = "В одном файле"
        .OnAction = "Ukhod_Together"
    End With
    Set Ukhod_Submenu = Ukhod_Popup.Controls.Add(Type:=msoControlButton)
    With Ukhod_Submenu
        .Caption = "Раздельно"
        .OnAction = "Ukhod_Separate"
    End With
    Set Tampon_Popup = SSV_Popup.Controls.Add(Type:=msoControlPopup, temporary:=True)
    Tampon_Popup.Caption = "&Тампоны"
    Tampon_Popup.BeginGroup = TRUE
    Set Tampon_Submenu = Tampon_Popup.Controls.Add(Type:=msoControlButton)
    With Tampon_Submenu
        .Caption = "Анальные"
        .OnAction = "Tampon_Anal"
    End With
    Set Tampon_Submenu = Tampon_Popup.Controls.Add(Type:=msoControlButton)
    With Tampon_Submenu
        .Caption = "Для Стомы"
        .OnAction = "Tampon_Stoma"
    End With
    Set SSV_Submenu = SSV_Popup.Controls.Add(Type:=msoControlButton)
    With SSV_Submenu
        .Caption = "Уропрезервативы"
        .OnAction = "Prezervativ"
    End With
    Set SSV_Submenu = SSV_Popup.Controls.Add(Type:=msoControlButton)
    With SSV_Submenu
        .Caption = "Пояс"
        .OnAction = "Belt"
    End With
    Set SSV_Submenu = SSV_Popup.Controls.Add(Type:=msoControlButton)
    With SSV_Submenu
        .Caption = "Ремешки"
        .OnAction = "Remeshki"
    End With
    Set SSV_Submenu = SSV_Popup.Controls.Add(Type:=msoControlButton)
    With SSV_Submenu
        .Caption = "Ирригация"
        .OnAction = "Irrigacia"
    End With
End Sub
Private Sub Show_Form_ABS()
    UserForm1.Show
End Sub
Private Sub Baza()
    tip_spiska_IKK = 1
    Call Obrabotka_IKK
End Sub
Private Sub DCP()
    tip_spiska_IKK = 2
    Call Obrabotka_IKK
End Sub
Private Sub Otkid_Spin()
    tip_spiska_IKK = 3
    Call Obrabotka_IKK
End Sub
Private Sub Povis_Gruz()
    tip_spiska_IKK = 5
    Call Obrabotka_IKK
End Sub
Private Sub Ruchazka()
    tip_spiska_IKK = 4
    Call Obrabotka_IKK
End Sub
Private Sub Elektro()
    tip_spiska_IKK = 6
    Call Obrabotka_IKK
End Sub
Private Sub Telezka()
    tip_spiska_IKK = 7
    Call Obrabotka_IKK
End Sub
Private Sub OttoBok()
    tip_spiska_IKK = 8
    Call Obrabotka_IKK
End Sub
Private Sub Stuliya()
    tip_spiska_IKK = 9
    Call Obrabotka_IKK
End Sub

Private Sub Kalo1()
    tip_spiska_SSV = 1
    Call Obrabotka_SSV
End Sub
Private Sub Uro1()
    tip_spiska_SSV = 2
    Call Obrabotka_SSV
End Sub
Private Sub Kalo2()
    tip_spiska_SSV = 3
    Call Obrabotka_SSV
End Sub
Private Sub Uro2()
    tip_spiska_SSV = 4
    Call Obrabotka_SSV
End Sub
Private Sub Copmlekt2()
    tip_spiska_SSV = 23
    Call Obrabotka_SSV
End Sub
Private Sub Folea()
    tip_spiska_SSV = 5
    Call Obrabotka_SSV
End Sub
Private Sub Peccera()
    tip_spiska_SSV = 6
    Call Obrabotka_SSV
End Sub
Private Sub Nefrostoma()
    tip_spiska_SSV = 7
    Call Obrabotka_SSV
End Sub
Private Sub Ureterocutaneo()
    tip_spiska_SSV = 22
    Call Obrabotka_SSV
End Sub
Private Sub Samokat()
    tip_spiska_SSV = 8
    Call Obrabotka_SSV
End Sub
Private Sub Samokat_nabor()
    tip_spiska_SSV = 9
    Call Obrabotka_SSV
End Sub
Private Sub Meshki_N()
    tip_spiska_SSV = 10
    Call Obrabotka_SSV
End Sub
Private Sub Meshki_D()
    tip_spiska_SSV = 11
    Call Obrabotka_SSV
End Sub
Private Sub NaborMoche()
    tip_spiska_SSV = 12
    Call Obrabotka_SSV
End Sub
Private Sub Ukhod_Together()
    tip_spiska_SSV = 13
    Call Obrabotka_SSV
End Sub
Private Sub Ukhod_Separate()
    tip_spiska_SSV = 14
    Call Obrabotka_SSV
End Sub
Private Sub Tampon_Anal()
    tip_spiska_SSV = 15
    Call Obrabotka_SSV
End Sub
Private Sub Tampon_Stoma()
    tip_spiska_SSV = 16
    Call Obrabotka_SSV
End Sub
Private Sub Prezervativ()
    tip_spiska_SSV = 17
    Call Obrabotka_SSV
End Sub
Private Sub Belt()
    tip_spiska_SSV = 18
    Call Obrabotka_SSV
End Sub
Private Sub Remeshki()
    tip_spiska_SSV = 19
    Call Obrabotka_SSV
End Sub
Private Sub Irrigacia()
    tip_spiska_SSV = 20
    Call Obrabotka_SSV
End Sub
Private Sub UroComplect()
    tip_spiska_SSV = 21
    Call Obrabotka_SSV
End Sub

Private Sub Delete_Submenu()
    On Error Resume Next
    CommandBars("Cell").Controls("Об&работать список ТСР").Delete
End Sub
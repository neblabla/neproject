Attribute VB_Name = "demo_sort_site"
Option Explicit
    Dim Ch As Selenium.ChromeDriver
    Dim FindBy As New Selenium.By
    Dim Chapter As Selenium.WebElement
    Dim URL As String
    Dim ImportURL As String
    Dim SortURL As String
    Dim IndexURL As String
    Dim Login As String
    Dim Password As String
        Dim Uploads As String
        Dim UploadFile As String
        Dim Sorting_done As String
        Dim SortByPrice As String
        Dim ShowAll As String
        Dim HideAll As String


    Dim i As Double
    Dim n As Double         'Счетчик обновленных позиций
    Dim Allt As Double      'Таймер для общей процедуры

Sub Sort_site()
Allt = Timer


        'Вводим входные данные
    Set Ch = New Selenium.ChromeDriver
        URL = "https://ural-soft.info/editor/"                                                                                  'Сайт
        Login = "login"                                                                                                         'Логин
        Password = "parol"                                                                                                      'Пароль
        ImportURL = "https://ural-soft.info/editor/structure/editsection/importexport/"                                         'Импорт цен
            Uploads = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\Прайсы\Выгрузки\"                         'Путь к файлам цен
        SortURL = "https://ural-soft.info/editor/structure/editsection/"                                                        'Сортировка каталога
        IndexURL = "https://ural-soft.info/editor/structure/indexthissite/"                                                     'Индексация поиска по сайту
            Sorting_done = "uss_editor_message_ok"                                                                              'Уведомление о успешной сортировке
            SortByPrice = "sortorder-4-styler"                                                                                  'Сортировка по цене по возрастанию
            ShowAll = "sortorder-11-styler"                                                                                     'Видимость. Показать все
            HideAll = "sortorder-16-styler"                                                                                     'Видимость. Скрыть все с нулевой


        'запускаем браузер и логинимся
    Ch.AddArgument "start-maximized"
    Ch.start
    Ch.Get URL
    Ch.FindElement(FindBy.Name("login")).SendKeys Login
    Ch.FindElement(FindBy.Name("password")).SendKeys Password
    Ch.FindElement(FindBy.Name("send_login_data")).Click


        'создаем массив разделов для сортировки
    Dim Category(31) As String

        Category(0) = "1 » LED экраны"
        Category(1) = "1 » Звуковое оборудование » Акустические системы » Adamson"
        Category(2) = "1 » Звуковое оборудование » Акустические системы » D&B Audiotechnik"
        Category(3) = "1 » Звуковое оборудование » Акустические системы » L-Acoustics"
        Category(4) = "1 » Звуковое оборудование » Акустические системы » Martin Audio"
        Category(5) = "1 » Звуковое оборудование » Акустические системы » ProTone"
        Category(6) = "1 » Звуковое оборудование » Микшерные пульты » DiGiCo"
        Category(7) = "1 » Звуковое оборудование » Процессоры управления акустическими системами » L-Acoustics"
        Category(8) = "1 » Звуковое оборудование » Процессоры управления акустическими системами » Martin Audio"
        Category(9) = "1 » Звуковое оборудование » Усилители » Adamson"
        Category(10) = "1 » Звуковое оборудование » Усилители » D&B Audiotechnik"
        Category(11) = "1 » Звуковое оборудование » Усилители » L-Acoustics"
        Category(12) = "1 » Звуковое оборудование » Усилители » Martin Audio"
        Category(13) = "1 » Звуковое оборудование » Усилители » ProTone"
        Category(14) = "1 » Коммутация и рэки » Кабель в бухтах » Акустический кабель » Draka"
        Category(15) = "1 » Коммутация и рэки » Кабель в бухтах » Акустический кабель » Wiring Parts"
        Category(16) = "1 » Коммутация и рэки » Кабель в бухтах » Акустический кабель » XSE"
        Category(17) = "1 » Коммутация и рэки » Лючки сценические » EDS"
        Category(18) = "1 » Коммутация и рэки » Силовое распределение » Дистрибьюторы мобильные » EDS"
        Category(19) = "1 » Коммутация и рэки » Силовое распределение » Дистрибьюторы рэковые » EDS"
        Category(20) = "1 » Коммутация и рэки » Устройства передачи сигнала » EDS"
        Category(21) = "1 » Механика сцены и фермы"
        Category(22) = "1 » Механика сцены и фермы » Лебедки" 'Скрыть все с нулевой ценой
        Category(23) = "1 » Механика сцены и фермы » Лебедки » Imlight"
        Category(24) = "1 » Механика сцены и фермы » Сценические фермы » Imlight" 'Скрыть все с нулевой ценой
        Category(25) = "1 » Микрофоны и радиосистемы » Цифровые радиосистемы » Shure » Серия Shure Axient"
        Category(26) = "1 » Микрофоны и радиосистемы » Цифровые радиосистемы » Shure » Серия Shure ULXD"
        Category(27) = "1 » Световое оборудование » Вращающиеся головы » DTS"
        Category(28) = "1 » Световое оборудование » Приборы заливного света » ETC"
        Category(29) = "1 » Световое оборудование » Приборы управления светом » Световые консоли » ETC"
        Category(30) = "1 » Световое оборудование » Театральные прожекторы » DTS"
        Category(31) = "1 » Световое оборудование » Театральные прожекторы » ETC"


        'Открываем все позиции
    Ch.Get SortURL
        Ch.FindElementById(SortByPrice).Click
        Ch.FindElementById(ShowAll).Click
        Ch.FindElementByName("sortpos").Click
            If Not Ch.IsElementPresent(FindBy.Class(Sorting_done)) Then
                MsgBox "Плохое интернет соединение или ошибка в коде", vbExclamation
                Exit Sub
            End If
        Debug.Print "1.1 Выполнено: Показать все позиции."


        'Скрываем с нулевыми ценами
    Ch.Get SortURL
        Ch.FindElementById(HideAll).Click
        Ch.FindElementByName("sortpos").Click
            If Not Ch.IsElementPresent(FindBy.Class(Sorting_done)) Then
                MsgBox "Плохое интернет соединение или ошибка в коде", vbExclamation
                Exit Sub
            End If
        Debug.Print "1.2 Выполнено: Скрыть все позиции с нулевой ценой."


        'Открываем и закрываем нужные разделы каталога
    For i = 0 To UBound(Category)
    
    
        Ch.Get SortURL
            Ch.FindElement(FindBy.Class("jq-selectbox__trigger-arrow")).Click
                If Not Ch.IsElementPresent(FindBy.XPath("/html/body/div[3]/ul/li[text()='" & Category(i) & "']")) Then
                    MsgBox Category(i) & " - не находится в списке", vbExclamation
                    Exit Sub
                End If
            Set Chapter = Ch.FindElementByXPath("/html/body/div[3]/ul/li[text()='" & Category(i) & "']")
            Chapter.Click
            Ch.FindElementById(ShowAll).Click
                If i = 22 Or i = 24 Then
                    Ch.FindElementById(HideAll).Click
                End If
            Ch.FindElementByName("sortpos").Click
            
                If Not Ch.IsElementPresent(FindBy.Class(Sorting_done)) Then
                    MsgBox "Плохое интернет соединение или ошибка в коде. не отсортирован раздел " & Category(i), vbExclamation
                    Exit Sub
                End If
            Debug.Print "2." & i + 1 & " Отсортирован раздел: " & Category(i)
    
    Next i


        'Индексируем поиск по сайту
    Ch.Get IndexURL
    Ch.Quit
    
Allt = Timer - Allt
    Debug.Print "Загрузка и сортировка каталога прошла успешно! (" & Format(Allt / 60, "#") & "мин " & Format(Allt Mod 60, "#") & "с)"
    MsgBox "Загрузка и сортировка каталога прошла успешно! (" & Format(Allt / 60, "#") & "мин " & Format(Allt Mod 60, "#") & "с)"
    
End Sub


Attribute VB_Name = "demo_upload_items"
Option Explicit
    Dim Ch As Selenium.ChromeDriver
    Dim FindBy As New Selenium.By
    Dim URL As String, URL_site As String
    Dim Login As String
    Dim Password As String
    Dim DownloadFolder As String

    Dim i As Integer, j As Integer, j1 As Integer, k As Integer, k1 As Integer      'Счетчики для циклов
    Dim loaded_items As String, error_items As String, unfinded_items As String     'Итоговые позиции
    Dim n As Integer, n1 As Integer, n2 As Integer                                  'Счетчики позиций
    Dim t As Double                                                                 'Таймер для внутренних процедур
    Dim Allt As Double                                                              'Таймер для общей процедуры

Sub Upload_items()
Allt = Timer
loaded_items = "": error_items = "": unfinded_items = ""
n = 0: n1 = 0: n2 = 0

'======================================================================================
'        Вводим входные данные
'======================================================================================
    Set Ch = New Selenium.ChromeDriver
        URL = "https://ural-soft.info/editor/"                                                                  'Сайт
        URL_site = "https://imlight.ru/"                                                                        'Сайт производителя
        Login = "login"                                                                                         'Логин
        Password = "parol"                                                                                      'Пароль
        DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\Загрузка позиций\Изображения\"


'======================================================================================
'        Запускаем браузер, вкладки для работы и логинимся
'======================================================================================
Debug.Print "Запуск браузера"
    With Ch
        .AddArgument "start-maximized"
        .SetPreference "download.default_directory", (DownloadFolder)
        .start
        .Get URL
        .ExecuteScript "window.open(arguments[0])", URL_site
        .ExecuteScript "window.open(arguments[0])", "https://www.lunapic.com/editor/?action=smart-crop"
        .Windows(2).Activate
        .SwitchToPreviousWindow
        .FindElement(FindBy.Name("login"), 5000).SendKeys Login
        .FindElement(FindBy.Name("password"), 5000).SendKeys Password
        .FindElement(FindBy.Name("send_login_data"), 5000).Click
    End With


'======================================================================================
'        Отбираем позиции для сортировки
'======================================================================================
    Dim url_catalog As String
        url_catalog = "https://ural-soft.info/editor/structure/editsection/?reference=10137023"

    Dim key_search As String
        key_search = "#uss_filter > tbody > tr:nth-child(1) > td.uss_filter_input > div:nth-child(1) > input[type=text]"

    Dim items_name As String                'Поиск позиций
        items_name = InputBox("Что загружаем?") ', Xpos:=32000, Ypos:=5000)
Debug.Print "Отбор позиций: " & items_name

    With Ch
        .Get url_catalog
        .FindElement(FindBy.Css(key_search), 5000).SendKeys items_name
        .FindElement(FindBy.Class("jq-selectbox__trigger"), 5000).Click
        .FindElement(FindBy.Css("body > div.jq-selectbox__dropdown > ul > li:nth-child(2)"), 5000).Click
        .FindElement(FindBy.Css("#uss_filter > tbody > tr:nth-child(2) > td.uss_filter_line > div:nth-child(2) > label:nth-child(2) > div"), 5000).Click
        .FindElement(FindBy.Name("set_filters"), 5000).Click
    End With


'======================================================================================
'        Сбор названий и ссылок
'======================================================================================
    Dim items As WebElements
    Dim Items_library() As String
    
    Set items = Ch.FindElements(FindBy.Css("div [class=uss_editor_pos_title] > a"))
    
    If items.Count = 0 Then
        Debug.Print "Позиции не найдены, или ошибка поиска"
        Exit Sub
    End If
    
    ReDim Items_library(1, items.Count - 1)
    For i = 0 To items.Count - 1
        Items_library(0, i) = items(i + 1).Text
        Items_library(1, i) = items(i + 1).Attribute("href")
        Debug.Print vbTab; Items_library(0, i), Items_library(1, i)
    Next i

        Debug.Print "Отсортировано: " & items.Count & " позиций"




'============================================================================================================================================================================
'============================================================================================================================================================================
'Блок сбора информации
'============================================================================================================================================================================
'============================================================================================================================================================================
Debug.Print vbNewLine & vbNewLine & "Обработка позиций ..."

    Dim item_art As String
    
    Dim content_el As WebElements
    Dim el As WebElement
        Dim image_name() As String
        Dim content_image() As String
        Dim content_desc As String
        Dim content_char As String
        Dim content_video As String
    
   For i = 0 To UBound(Items_library, 2)

t = Timer
Debug.Print vbNewLine & i + 1 & ". " & Items_library(0, i)


'=================================================
'=================================================
'=================================================
'=================================================
    
        'Выбор артикула (для Imlight)
'=================================================
    Ch.Get Items_library(1, i)
        item_art = Ch.FindElement(FindBy.Css("#explanationid"), 5000).Text     '(IM-
        If InStr(1, item_art, "(IM-", 1) = 0 Then
            Ch.Windows(2).Activate
            Ch.SwitchToNextWindow
            GoTo Unfinding_item
        End If
        item_art = Replace(Right(item_art, Len(item_art) - (InStr(1, item_art, "(IM-", 1)) - 3), ")", "")


        'Переход на сайт
    Ch.Windows(2).Activate
    Ch.SwitchToNextWindow
    On Error GoTo Unfinding_item


Debug.Print vbTab & "Поиск... ";
        'Поиск необходимой позиции
    Ch.FindElement(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > button"), 5000).Click
    Ch.FindElement(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > form > div > input"), 5000).SendKeys item_art


        'обработка отсутствующих позиций
    If Ch.IsElementPresent(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > form > ul > span"), 2500) Or Len(item_art) < 4 Then
            Debug.Print "Позиция не найдена"
            Ch.FindElement(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > form > div > input"), 5000).Clear
            Ch.FindElement(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > form > div > span.search-close > i"), 5000).Click
            GoTo Unfinding_item
    Else:   Debug.Print "Позиция найдена"
    End If
    
    Ch.Get (Ch.FindElement(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > form > ul > li:nth-child(1) > a"), 5000).Attribute("href"))





content_image:
'======================================================================================
'       Изображения
'======================================================================================
Debug.Print vbTab & "Сбор изображений";




        'Дополнительные изображения
Set content_el = Ch.FindElements(FindBy.XPath("//div[@class='product-specifications__images-item-wrap slick-slide slick-current slick-active']/div/a | //div[@class='product-specifications__images-item-wrap slick-slide slick-active']/div/a | //div[@class='product-specifications__images-item-wrap slick-slide']/div/a"))
    If content_el.Count > 15 Then
        k = 15
    Else: k = content_el.Count
    End If
ReDim content_image(k)
ReDim image_name(k)


        'Основное изображение
    If (Not Ch.IsElementPresent(FindBy.XPath("//div[@class='product-specifications__image-big d-print-inline-block hidden-print']/a[1]"), 500)) Then
        GoTo Unfinding_item
    End If


        'Проверка на отсутствие картинок
        'И создание массива с ссылками
    content_image(0) = Ch.FindElement(FindBy.XPath("//div[@class='product-specifications__image-big d-print-inline-block hidden-print']/a[1]"), 5000).Attribute("href")

    For j = 1 To k
        content_image(j) = content_el(j).Attribute("href")
    Next j

    If content_image(0) = "" Then
        content_image(0) = "https://www.imlight.ru/images/cms/data/logo/Roxtone/roxtone_400x400_new.png"
    End If
    

        'Создание объектов для скачивания
    Dim httpObject As Object
    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    Dim binaryStream As Object
    Set binaryStream = CreateObject("ADODB.Stream")


        'Переключение на окно обрезки изображений
'    Ch.SwitchToPreviousWindow
    On Error GoTo 0
Debug.Print " & обрезка и скачивание изображений"


        'Очистка папки загрузки
    If n <> 0 Then
        On Error Resume Next
        Kill DownloadFolder & "*"
        On Error GoTo 0
    End If

        'Цикл скачивания
    For j = 0 To UBound(content_image)
            
        If Right(content_image(j), 4) = "webp" Then
            image_name(j) = Items_library(0, i) & " " & j & ".png"
        ElseIf Right(content_image(j), 4) = "jpeg" Then
            image_name(j) = Items_library(0, i) & " " & j & ".jpeg"
        Else
            image_name(j) = Items_library(0, i) & " " & j & Right(content_image(j), 4)
        End If
                      
                    
        image_name(j) = Replace(image_name(j), "/", "")
        image_name(j) = Replace(image_name(j), "\", "")
        image_name(j) = Replace(image_name(j), "?", "")
        image_name(j) = Replace(image_name(j), ":", "")
        image_name(j) = Replace(image_name(j), "*", "")


            'Скачивание изображения
        httpObject.Open "GET", content_image(j), False
        httpObject.send
            'Write the image to a file
        binaryStream.Open
        binaryStream.Type = 1
        binaryStream.Write httpObject.responseBody
        binaryStream.SaveToFile DownloadFolder & image_name(j), 2
            'Clean up
        binaryStream.Close

    Next j

Set binaryStream = Nothing
Set httpObject = Nothing


'    Ch.SwitchToPreviousWindow
    On Error GoTo Unfinding_item




Content_description:
'======================================================================================
'       Описание
'======================================================================================




Debug.Print vbTab & "Поиск описания...";
    content_desc = ""
        Set content_el = Ch.FindElements(FindBy.XPath("//*[@id='description']/div/div[2]//*"))

            If content_el.Count = 0 Then
                GoTo Content_features
            End If
    content_el.Last.ScrollIntoView
                Debug.Print " копирование"

        For Each el In content_el
            If el.Text = "" Or el.Text = " " Or el.Text = "  " Or el.tagName = "ul" Or el.IsElementPresent(FindBy.Css("strong")) Then
            Else
                content_desc = content_desc & "<" & el.tagName & " style='text-align: justify;'>" & el.Text & "</" & el.tagName & ">"
            End If
        Next el


            'Копирование особенностей при наличии
Content_features:
        Set content_el = Ch.FindElements(FindBy.XPath("//*[@id='features']/div/div[2]/ul/li"))

            If content_el.Count = 0 Then
                GoTo Content_characteristics
            End If
    content_el.Last.ScrollIntoView
    
    
    content_desc = content_desc & "<h5 style='text-align: justify;'>Особенности:</h5><ul>"
        For Each el In content_el
            content_desc = content_desc & "<li style='text-align: justify;'>" & el.Text & "</li>"
        Next el
    content_desc = content_desc & "</ul>"




Content_characteristics:
'======================================================================================
'       Характеристики (таблицы)
'======================================================================================
Debug.Print vbTab & "Поиск характеристик...";
    content_char = "<table class='uss_table_darkgrey10' style='width: 100%;' dir='ltr' border='0'><tbody>"

                
                'Проверка наличия таблицы
            If Not Ch.IsElementPresent(FindBy.XPath("//*[@id='ttx']/div/div[2]/div[1]/table")) Then
                Debug.Print " отсутствуют"
                content_char = ""
                GoTo content_video
            End If
                Debug.Print " копирование"
                
        Set content_el = Ch.FindElements(FindBy.XPath("//*[@id='ttx']/div/div[2]/div[1]//tr"))
    content_el(1).ScrollIntoView

        k = content_el.Count


    If k = 2 Then
    
            'Обработка 4 столбцовых перевернутых таблиц
        For j = 1 To content_el(1).FindElements(FindBy.Css("td")).Count
            content_char = content_char & "<tr>"
            content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(1).FindElements(FindBy.Css("td"))(j).Text & "</td>"
            content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(2).FindElements(FindBy.Css("td"))(j).Text & "</td>"
            content_char = content_char & "</tr>"
        Next j
    
    
            'Обработка 4 столбцовых таблиц
    ElseIf content_el(1).FindElements(FindBy.Css("td")).Count = 4 _
    Or content_el(2).FindElements(FindBy.Css("td")).Count = 4 _
    Or content_el(3).FindElements(FindBy.Css("td")).Count = 4 Then

        For j = 1 To k
            content_char = content_char & "<tr>"
            
                    'Пропуск заголовка из 3 ячеек
                If content_el(j).FindElements(FindBy.Css("td")).Count = 3 And j = 1 Then
                
                     'Добавление от первого столбца в название (Imlight)
                ElseIf content_el(j).FindElements(FindBy.Css("td")).Count = 4 Then
                    content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' colspan='2' rowspan='1' data-sheets-value='{'><strong>" & content_el(j).FindElements(FindBy.Css("td"))(1).Text & "</strong></td>"
                    content_char = content_char & "</tr><tr>"
                    content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(j).FindElements(FindBy.Css("td"))(2).Text & "</td>"
                    content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(j).FindElements(FindBy.Css("td"))(3).Text & content_el(j).FindElements(FindBy.Css("td"))(4).Text & "</td>"
                
                ElseIf content_el(j).FindElements(FindBy.Css("td")).Count = 3 Then
                    content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(j).FindElements(FindBy.Css("td"))(1).Text & "</td>"
                    content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(j).FindElements(FindBy.Css("td"))(2).Text & content_el(j).FindElements(FindBy.Css("td"))(3).Text & "</td>"
                End If
            
            content_char = content_char & "</tr>"
        Next j
        

    Else        'Обработка 2-3 столбцовых таблиц
        For j = 1 To k
                
            If content_el(j).FindElements(FindBy.Css("td")).Count = 3 And ( _
            LCase(content_el(j).FindElements(FindBy.Css("td"))(1).Text) = "упаковка" Or _
            InStr(1, LCase(content_el(j).FindElements(FindBy.Css("td"))(1).Text), "комплект") > 0 Or _
            InStr(1, LCase(content_el(j).FindElements(FindBy.Css("td"))(1).Text), "срок") > 0 Or _
            InStr(1, LCase(content_el(j).FindElements(FindBy.Css("td"))(1).Text), "сертификат") > 0 _
            ) Then
                Exit For
            End If
            content_char = content_char & "<tr>"
            
                     'Добавление от первого столбца в название (Imlight)
                If content_el(j).FindElements(FindBy.Css("td")).Count = 3 Then
                
                    content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' colspan='2' rowspan='1' data-sheets-value='{'><strong>" & content_el(j).FindElements(FindBy.Css("td"))(1).Text & "</strong></td>"
                    content_char = content_char & "</tr><tr>"
                    
                    content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(j).FindElements(FindBy.Css("td"))(2).Text & "</td>"
                    content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(j).FindElements(FindBy.Css("td"))(3).Text & "</td>"
                
                ElseIf content_el(j).FindElements(FindBy.Css("td")).Count = 2 Then
                    content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(j).FindElements(FindBy.Css("td"))(1).Text & "</td>"
                    content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(j).FindElements(FindBy.Css("td"))(2).Text & "</td>"
                
                ElseIf content_el(j).FindElements(FindBy.Css("td")).Count = 1 Then
                    content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' colspan='2' rowspan='1' data-sheets-value='{'><strong>" & content_el(j).FindElements(FindBy.Css("td"))(1).Text & "</strong></td>"
                End If
                
            content_char = content_char & "</tr>"
        Next j
    End If
        content_char = content_char & "</tbody></table>"




content_video:
'======================================================================================
'       Видео
'======================================================================================
Debug.Print vbTab & "Поиск видео...";
    content_video = ""
    Set content_el = Ch.FindElements(FindBy.Css("#video > div.container > div.slick-video.mt-20.mt-xs-30.slick-initialized.slick-slider > div > div > div > iframe"))


            If content_el.Count = 0 Then
                Debug.Print " отсутствует"
                content_video = ""
                GoTo Upload_content
            End If
    content_el.Last.ScrollIntoView
                Debug.Print " копирование"

    content_video = content_el(1).Attribute("src")




Upload_content:
'============================================================================================================================================================================
'============================================================================================================================================================================
'Блок загрузки информации
'============================================================================================================================================================================
'============================================================================================================================================================================
Debug.Print "Загрузка на сайт: ";

        'Переход на страницу позиции
    Ch.SwitchToNextWindow
    Ch.Get Items_library(1, i)
    
On Error GoTo loading_error

j1 = 0


Input_image:
'======================================================================================
'       Вставка изображений
'======================================================================================

Debug.Print "основное изображение";
    Ch.FindElement(FindBy.Css("#imageid"), 5000).SendKeys DownloadFolder & image_name(0)


    If UBound(image_name) < 1 Then GoTo Input_description
    
Debug.Print ", дополнительные изображения";
    Ch.FindElement(FindBy.Css("#multi_img_wrapper > div > div > div > input"), 5000).SendKeys DownloadFolder & image_name(1)

    If UBound(image_name) < 2 Then GoTo Input_description

    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(22)"), 5000).ScrollIntoView
        For j = 2 To UBound(image_name)
            Ch.FindElement(FindBy.Css("div.add_multi_i"), 5000).Click
            Ch.FindElement(FindBy.Css("#multi_img_wrapper > div:nth-child(" & j & ") > div > input"), 5000).SendKeys DownloadFolder & image_name(j)
        Next j



Input_description:
'======================================================================================
'       Вставка описания в модуль
'======================================================================================

    If content_desc = "" Then GoTo Input_characteristics
Debug.Print ", описание";

    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(27)"), 5000).LocationInView
    Ch.FindElement(FindBy.Css("#mceu_32-button"), 5000).Click
    Ch.Wait 250
    
    Call string_splitter(content_desc, "#mceu_58")
    
    Ch.FindElement(FindBy.Css("#mceu_60-button"), 5000).Click
j1 = j1 + 7

Input_characteristics:
'======================================================================================
'Вставка характеристик в модуль
'======================================================================================

    If content_char = "" Then GoTo Input_video
Debug.Print ", характеристики";

    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(27)"), 5000).LocationInView
    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(25) > div.uss_editor_addnewtab > a"), 5000).Click
    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(27)"), 5000).LocationInView
    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(25) > div.tabsWrap > div > div.tabItem > input")).SendKeys "Характеристики"
    Ch.FindElement(FindBy.Css("#mceu_" & (87 + j1) & "-button"), 5000).Click
    Ch.Wait 250
    
    Call string_splitter(content_char, "#mceu_" & (113 + j1))
    
    Ch.FindElement(FindBy.Css("#mceu_" & (115 + j1) & "-button"), 5000).Click
j1 = j1 + 62


Input_video:
'======================================================================================
'       Вставка видео в модуль
'======================================================================================

    If content_video = "" Then GoTo Save_item
Debug.Print ", видео";

        'Настраивание зависимости от наличия характеристик
    If content_char <> "" Then
        j = 3
    Else
        j = 1
    End If

    With Ch
        .FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(27)"), 5000).LocationInView
        .FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(25) > div.uss_editor_addnewtab > a"), 5000).Click
        .FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(27)"), 5000).LocationInView
        .FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(25) > div.tabsWrap > div > div:nth-child(" & j & ") > input"), 5000).SendKeys "Видео"
        .FindElement(FindBy.Css("#mceu_" & (77 + j1) & "-button"), 5000).Click
        .Wait 250
        .FindElement(FindBy.Css("#mceu_" & (114 + j1) & "-inp"), 5000).SendKeys content_video
        .FindElement(FindBy.Css("#mceu_" & (116 + j1)), 5000).Click
        .FindElement(FindBy.Css("#mceu_" & (116 + j1)), 5000).Clear
        .FindElement(FindBy.Css("#mceu_" & (116 + j1)), 5000).SendKeys 860
        .FindElement(FindBy.Css("#mceu_" & (118 + j1)), 5000).Clear
        .FindElement(FindBy.Css("#mceu_" & (118 + j1)), 5000).SendKeys 485
        .FindElement(FindBy.Css("#mceu_" & (126 + j1) & "-button"), 5000).Click
    End With




Save_item:
Debug.Print
Debug.Print "Cохранение... ";
    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div.buttonsWrap > div > input.submit.save"), 5000).Click
    
On Error GoTo 0
n = n + 1

    If loaded_items = "" Then
        loaded_items = Items_library(0, i)
        Else: loaded_items = loaded_items & ", " & Items_library(0, i)
    End If
    
    t = Format(Timer - t, "#.##")
    Debug.Print "Позиция """ & Items_library(0, i) & """ загружена успешно(" & t & "с)"
    GoTo Next_item


Unfinding_item:             'Обработка ненайденных позиций
    Ch.SwitchToNextWindow
    n1 = n1 + 1

    Debug.Print ""; "Страница отсутствует на сайте"
    Debug.Print

    If unfinded_items = "" Then
        unfinded_items = Items_library(0, i)
        Else: unfinded_items = unfinded_items & ", " & Items_library(0, i)
    End If

    GoTo Next_item


loading_error:             'Обработка ошибок загрузки позиций
    Debug.Print "Ошибка - произошло гавно!"
    n2 = n2 + 1
    If error_items = "" Then
        error_items = Items_library(0, i)
        Else: error_items = error_items & ", " & Items_library(0, i)
    End If
    
GoTo Next_item



Next_item:
Next i




Allt = Timer - Allt
Debug.Print vbNewLine & vbNewLine & "Загрузка окончена(" & Format(Allt / 60, "#") & "мин " & Format(Allt Mod 60, "#") & "с)"


uploaded_print:
        If n = 0 Then GoTo unfinded_print
        
        If n Mod 10 = 1 Then
                Debug.Print "Загружена " & n & " позиция:"
            ElseIf n Mod 10 = 2 Or n Mod 10 = 3 Or n Mod 10 = 4 Then
                Debug.Print "Загружено " & n & " позиции:"
            Else: Debug.Print "Загружено " & n & " позиций:"
        End If
Debug.Print loaded_items & vbNewLine


unfinded_print:
        If n1 = 0 Then GoTo errors_print
        
        If n1 Mod 10 = 1 Then
                Debug.Print "Отсутствуют на сайте " & n1 & " позиция:"
            ElseIf n1 Mod 10 = 2 Or n1 Mod 10 = 3 Or n1 Mod 10 = 4 Then
                Debug.Print "Отсутствуют на сайте " & n1 & " позиции:"
            Else: Debug.Print "Отсутствуют на сайте " & n1 & " позиций:"
        End If
Debug.Print unfinded_items


errors_print:
        If n2 = 0 Then GoTo finish_sub
        
        If n2 Mod 10 = 1 Then
                Debug.Print "Ошибка загрузки " & n2 & " позиции:"
            ElseIf n2 Mod 10 = 2 Or n2 Mod 10 = 3 Or n2 Mod 10 = 4 Then
                Debug.Print "Ошибка загрузки " & n2 & " позиций:"
            Else: Debug.Print "Ошибка загрузки " & n2 & " позиций:"
        End If
Debug.Print error_items



finish_sub:
Debug.Print ("Среднее время загрузки позиции: " & Format(Allt / n, "#") & "с")
MsgBox ("Загрузка окончена(" & Format(Allt / 60, "#") & "мин " & Format(Allt Mod 60, "#") & "с)" & vbCrLf & _
        "Среднее время загрузки позиции: " & Format(Allt / n, "#") & "с")
End Sub




'======================================================================================
'======================================================================================
'       Вспомогательные функции
'======================================================================================
'======================================================================================




'======================================================================================
'       string_splitter
'======================================================================================
Function string_splitter(str As String, selector As String)

    Dim str_i As Integer, srt_n As Integer, str_len As Integer
    str_len = 500
    srt_n = Len(str) / str_len
    
    If srt_n < Len(str) / str_len Then
        srt_n = srt_n + 1
    End If

    If Len(str) < 32767 Then
        For str_i = 1 To srt_n
            Ch.FindElement(FindBy.Css(selector), 5000).SendKeys Mid(str, 1 + (str_i - 1) * str_len, str_len)
        Next str_i
    Else
        Ch.FindElement(FindBy.Css(selector), 5000).SendKeys str
    End If
        

End Function




Attribute VB_Name = "demo_upload_items"
Option Explicit
    Dim Ch As Selenium.ChromeDriver
    Dim FindBy As New Selenium.By
    Dim URL As String, URL_site As String
    Dim Login As String
    Dim Password As String
    Dim DownloadFolder As String

    Dim i As Integer, j As Integer, j1 As Integer, k As Integer, k1 As Integer      '�������� ��� ������
    Dim loaded_items As String, error_items As String, unfinded_items As String     '�������� �������
    Dim n As Integer, n1 As Integer, n2 As Integer                                  '�������� �������
    Dim t As Double                                                                 '������ ��� ���������� ��������
    Dim Allt As Double                                                              '������ ��� ����� ���������

Sub Upload_items()
Allt = Timer
loaded_items = "": error_items = "": unfinded_items = ""
n = 0: n1 = 0: n2 = 0

'======================================================================================
'        ������ ������� ������
'======================================================================================
    Set Ch = New Selenium.ChromeDriver
        URL = "https://ural-soft.info/editor/"                                                                  '����
        URL_site = "https://imlight.ru/"                                                                        '���� �������������
        Login = "login"                                                                                         '�����
        Password = "parol"                                                                                      '������
        DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\�������� �������\�����������\"


'======================================================================================
'        ��������� �������, ������� ��� ������ � ���������
'======================================================================================
Debug.Print "������ ��������"
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
'        �������� ������� ��� ����������
'======================================================================================
    Dim url_catalog As String
        url_catalog = "https://ural-soft.info/editor/structure/editsection/?reference=10137023"

    Dim key_search As String
        key_search = "#uss_filter > tbody > tr:nth-child(1) > td.uss_filter_input > div:nth-child(1) > input[type=text]"

    Dim items_name As String                '����� �������
        items_name = InputBox("��� ���������?") ', Xpos:=32000, Ypos:=5000)
Debug.Print "����� �������: " & items_name

    With Ch
        .Get url_catalog
        .FindElement(FindBy.Css(key_search), 5000).SendKeys items_name
        .FindElement(FindBy.Class("jq-selectbox__trigger"), 5000).Click
        .FindElement(FindBy.Css("body > div.jq-selectbox__dropdown > ul > li:nth-child(2)"), 5000).Click
        .FindElement(FindBy.Css("#uss_filter > tbody > tr:nth-child(2) > td.uss_filter_line > div:nth-child(2) > label:nth-child(2) > div"), 5000).Click
        .FindElement(FindBy.Name("set_filters"), 5000).Click
    End With


'======================================================================================
'        ���� �������� � ������
'======================================================================================
    Dim items As WebElements
    Dim Items_library() As String
    
    Set items = Ch.FindElements(FindBy.Css("div [class=uss_editor_pos_title] > a"))
    
    If items.Count = 0 Then
        Debug.Print "������� �� �������, ��� ������ ������"
        Exit Sub
    End If
    
    ReDim Items_library(1, items.Count - 1)
    For i = 0 To items.Count - 1
        Items_library(0, i) = items(i + 1).Text
        Items_library(1, i) = items(i + 1).Attribute("href")
        Debug.Print vbTab; Items_library(0, i), Items_library(1, i)
    Next i

        Debug.Print "�������������: " & items.Count & " �������"




'============================================================================================================================================================================
'============================================================================================================================================================================
'���� ����� ����������
'============================================================================================================================================================================
'============================================================================================================================================================================
Debug.Print vbNewLine & vbNewLine & "��������� ������� ..."

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
    
        '����� �������� (��� Imlight)
'=================================================
    Ch.Get Items_library(1, i)
        item_art = Ch.FindElement(FindBy.Css("#explanationid"), 5000).Text     '(IM-
        If InStr(1, item_art, "(IM-", 1) = 0 Then
            Ch.Windows(2).Activate
            Ch.SwitchToNextWindow
            GoTo Unfinding_item
        End If
        item_art = Replace(Right(item_art, Len(item_art) - (InStr(1, item_art, "(IM-", 1)) - 3), ")", "")


        '������� �� ����
    Ch.Windows(2).Activate
    Ch.SwitchToNextWindow
    On Error GoTo Unfinding_item


Debug.Print vbTab & "�����... ";
        '����� ����������� �������
    Ch.FindElement(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > button"), 5000).Click
    Ch.FindElement(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > form > div > input"), 5000).SendKeys item_art


        '��������� ������������� �������
    If Ch.IsElementPresent(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > form > ul > span"), 2500) Or Len(item_art) < 4 Then
            Debug.Print "������� �� �������"
            Ch.FindElement(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > form > div > input"), 5000).Clear
            Ch.FindElement(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > form > div > span.search-close > i"), 5000).Click
            GoTo Unfinding_item
    Else:   Debug.Print "������� �������"
    End If
    
    Ch.Get (Ch.FindElement(FindBy.Css("body > header > div.header-top-wrap > div.header-center > div > div.header-center-inner > div > div.col-md-1 > div > form > ul > li:nth-child(1) > a"), 5000).Attribute("href"))





content_image:
'======================================================================================
'       �����������
'======================================================================================
Debug.Print vbTab & "���� �����������";




        '�������������� �����������
Set content_el = Ch.FindElements(FindBy.XPath("//div[@class='product-specifications__images-item-wrap slick-slide slick-current slick-active']/div/a | //div[@class='product-specifications__images-item-wrap slick-slide slick-active']/div/a | //div[@class='product-specifications__images-item-wrap slick-slide']/div/a"))
    If content_el.Count > 15 Then
        k = 15
    Else: k = content_el.Count
    End If
ReDim content_image(k)
ReDim image_name(k)


        '�������� �����������
    If (Not Ch.IsElementPresent(FindBy.XPath("//div[@class='product-specifications__image-big d-print-inline-block hidden-print']/a[1]"), 500)) Then
        GoTo Unfinding_item
    End If


        '�������� �� ���������� ��������
        '� �������� ������� � ��������
    content_image(0) = Ch.FindElement(FindBy.XPath("//div[@class='product-specifications__image-big d-print-inline-block hidden-print']/a[1]"), 5000).Attribute("href")

    For j = 1 To k
        content_image(j) = content_el(j).Attribute("href")
    Next j

    If content_image(0) = "" Then
        content_image(0) = "https://www.imlight.ru/images/cms/data/logo/Roxtone/roxtone_400x400_new.png"
    End If
    

        '�������� �������� ��� ����������
    Dim httpObject As Object
    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    Dim binaryStream As Object
    Set binaryStream = CreateObject("ADODB.Stream")


        '������������ �� ���� ������� �����������
'    Ch.SwitchToPreviousWindow
    On Error GoTo 0
Debug.Print " & ������� � ���������� �����������"


        '������� ����� ��������
    If n <> 0 Then
        On Error Resume Next
        Kill DownloadFolder & "*"
        On Error GoTo 0
    End If

        '���� ����������
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


            '���������� �����������
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
'       ��������
'======================================================================================




Debug.Print vbTab & "����� ��������...";
    content_desc = ""
        Set content_el = Ch.FindElements(FindBy.XPath("//*[@id='description']/div/div[2]//*"))

            If content_el.Count = 0 Then
                GoTo Content_features
            End If
    content_el.Last.ScrollIntoView
                Debug.Print " �����������"

        For Each el In content_el
            If el.Text = "" Or el.Text = " " Or el.Text = "  " Or el.tagName = "ul" Or el.IsElementPresent(FindBy.Css("strong")) Then
            Else
                content_desc = content_desc & "<" & el.tagName & " style='text-align: justify;'>" & el.Text & "</" & el.tagName & ">"
            End If
        Next el


            '����������� ������������ ��� �������
Content_features:
        Set content_el = Ch.FindElements(FindBy.XPath("//*[@id='features']/div/div[2]/ul/li"))

            If content_el.Count = 0 Then
                GoTo Content_characteristics
            End If
    content_el.Last.ScrollIntoView
    
    
    content_desc = content_desc & "<h5 style='text-align: justify;'>�����������:</h5><ul>"
        For Each el In content_el
            content_desc = content_desc & "<li style='text-align: justify;'>" & el.Text & "</li>"
        Next el
    content_desc = content_desc & "</ul>"




Content_characteristics:
'======================================================================================
'       �������������� (�������)
'======================================================================================
Debug.Print vbTab & "����� �������������...";
    content_char = "<table class='uss_table_darkgrey10' style='width: 100%;' dir='ltr' border='0'><tbody>"

                
                '�������� ������� �������
            If Not Ch.IsElementPresent(FindBy.XPath("//*[@id='ttx']/div/div[2]/div[1]/table")) Then
                Debug.Print " �����������"
                content_char = ""
                GoTo content_video
            End If
                Debug.Print " �����������"
                
        Set content_el = Ch.FindElements(FindBy.XPath("//*[@id='ttx']/div/div[2]/div[1]//tr"))
    content_el(1).ScrollIntoView

        k = content_el.Count


    If k = 2 Then
    
            '��������� 4 ���������� ������������ ������
        For j = 1 To content_el(1).FindElements(FindBy.Css("td")).Count
            content_char = content_char & "<tr>"
            content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(1).FindElements(FindBy.Css("td"))(j).Text & "</td>"
            content_char = content_char & "<td style='width: 50%; text-align: left; vertical-align: top;' data-sheets-value='{'>" & content_el(2).FindElements(FindBy.Css("td"))(j).Text & "</td>"
            content_char = content_char & "</tr>"
        Next j
    
    
            '��������� 4 ���������� ������
    ElseIf content_el(1).FindElements(FindBy.Css("td")).Count = 4 _
    Or content_el(2).FindElements(FindBy.Css("td")).Count = 4 _
    Or content_el(3).FindElements(FindBy.Css("td")).Count = 4 Then

        For j = 1 To k
            content_char = content_char & "<tr>"
            
                    '������� ��������� �� 3 �����
                If content_el(j).FindElements(FindBy.Css("td")).Count = 3 And j = 1 Then
                
                     '���������� �� ������� ������� � �������� (Imlight)
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
        

    Else        '��������� 2-3 ���������� ������
        For j = 1 To k
                
            If content_el(j).FindElements(FindBy.Css("td")).Count = 3 And ( _
            LCase(content_el(j).FindElements(FindBy.Css("td"))(1).Text) = "��������" Or _
            InStr(1, LCase(content_el(j).FindElements(FindBy.Css("td"))(1).Text), "��������") > 0 Or _
            InStr(1, LCase(content_el(j).FindElements(FindBy.Css("td"))(1).Text), "����") > 0 Or _
            InStr(1, LCase(content_el(j).FindElements(FindBy.Css("td"))(1).Text), "����������") > 0 _
            ) Then
                Exit For
            End If
            content_char = content_char & "<tr>"
            
                     '���������� �� ������� ������� � �������� (Imlight)
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
'       �����
'======================================================================================
Debug.Print vbTab & "����� �����...";
    content_video = ""
    Set content_el = Ch.FindElements(FindBy.Css("#video > div.container > div.slick-video.mt-20.mt-xs-30.slick-initialized.slick-slider > div > div > div > iframe"))


            If content_el.Count = 0 Then
                Debug.Print " �����������"
                content_video = ""
                GoTo Upload_content
            End If
    content_el.Last.ScrollIntoView
                Debug.Print " �����������"

    content_video = content_el(1).Attribute("src")




Upload_content:
'============================================================================================================================================================================
'============================================================================================================================================================================
'���� �������� ����������
'============================================================================================================================================================================
'============================================================================================================================================================================
Debug.Print "�������� �� ����: ";

        '������� �� �������� �������
    Ch.SwitchToNextWindow
    Ch.Get Items_library(1, i)
    
On Error GoTo loading_error

j1 = 0


Input_image:
'======================================================================================
'       ������� �����������
'======================================================================================

Debug.Print "�������� �����������";
    Ch.FindElement(FindBy.Css("#imageid"), 5000).SendKeys DownloadFolder & image_name(0)


    If UBound(image_name) < 1 Then GoTo Input_description
    
Debug.Print ", �������������� �����������";
    Ch.FindElement(FindBy.Css("#multi_img_wrapper > div > div > div > input"), 5000).SendKeys DownloadFolder & image_name(1)

    If UBound(image_name) < 2 Then GoTo Input_description

    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(22)"), 5000).ScrollIntoView
        For j = 2 To UBound(image_name)
            Ch.FindElement(FindBy.Css("div.add_multi_i"), 5000).Click
            Ch.FindElement(FindBy.Css("#multi_img_wrapper > div:nth-child(" & j & ") > div > input"), 5000).SendKeys DownloadFolder & image_name(j)
        Next j



Input_description:
'======================================================================================
'       ������� �������� � ������
'======================================================================================

    If content_desc = "" Then GoTo Input_characteristics
Debug.Print ", ��������";

    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(27)"), 5000).LocationInView
    Ch.FindElement(FindBy.Css("#mceu_32-button"), 5000).Click
    Ch.Wait 250
    
    Call string_splitter(content_desc, "#mceu_58")
    
    Ch.FindElement(FindBy.Css("#mceu_60-button"), 5000).Click
j1 = j1 + 7

Input_characteristics:
'======================================================================================
'������� ������������� � ������
'======================================================================================

    If content_char = "" Then GoTo Input_video
Debug.Print ", ��������������";

    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(27)"), 5000).LocationInView
    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(25) > div.uss_editor_addnewtab > a"), 5000).Click
    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(27)"), 5000).LocationInView
    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(25) > div.tabsWrap > div > div.tabItem > input")).SendKeys "��������������"
    Ch.FindElement(FindBy.Css("#mceu_" & (87 + j1) & "-button"), 5000).Click
    Ch.Wait 250
    
    Call string_splitter(content_char, "#mceu_" & (113 + j1))
    
    Ch.FindElement(FindBy.Css("#mceu_" & (115 + j1) & "-button"), 5000).Click
j1 = j1 + 62


Input_video:
'======================================================================================
'       ������� ����� � ������
'======================================================================================

    If content_video = "" Then GoTo Save_item
Debug.Print ", �����";

        '������������ ����������� �� ������� �������������
    If content_char <> "" Then
        j = 3
    Else
        j = 1
    End If

    With Ch
        .FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(27)"), 5000).LocationInView
        .FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(25) > div.uss_editor_addnewtab > a"), 5000).Click
        .FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(27)"), 5000).LocationInView
        .FindElement(FindBy.Css("#eshopposeditform_id > div:nth-child(25) > div.tabsWrap > div > div:nth-child(" & j & ") > input"), 5000).SendKeys "�����"
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
Debug.Print "C���������... ";
    Ch.FindElement(FindBy.Css("#eshopposeditform_id > div.buttonsWrap > div > input.submit.save"), 5000).Click
    
On Error GoTo 0
n = n + 1

    If loaded_items = "" Then
        loaded_items = Items_library(0, i)
        Else: loaded_items = loaded_items & ", " & Items_library(0, i)
    End If
    
    t = Format(Timer - t, "#.##")
    Debug.Print "������� """ & Items_library(0, i) & """ ��������� �������(" & t & "�)"
    GoTo Next_item


Unfinding_item:             '��������� ����������� �������
    Ch.SwitchToNextWindow
    n1 = n1 + 1

    Debug.Print ""; "�������� ����������� �� �����"
    Debug.Print

    If unfinded_items = "" Then
        unfinded_items = Items_library(0, i)
        Else: unfinded_items = unfinded_items & ", " & Items_library(0, i)
    End If

    GoTo Next_item


loading_error:             '��������� ������ �������� �������
    Debug.Print "������ - ��������� �����!"
    n2 = n2 + 1
    If error_items = "" Then
        error_items = Items_library(0, i)
        Else: error_items = error_items & ", " & Items_library(0, i)
    End If
    
GoTo Next_item



Next_item:
Next i




Allt = Timer - Allt
Debug.Print vbNewLine & vbNewLine & "�������� ��������(" & Format(Allt / 60, "#") & "��� " & Format(Allt Mod 60, "#") & "�)"


uploaded_print:
        If n = 0 Then GoTo unfinded_print
        
        If n Mod 10 = 1 Then
                Debug.Print "��������� " & n & " �������:"
            ElseIf n Mod 10 = 2 Or n Mod 10 = 3 Or n Mod 10 = 4 Then
                Debug.Print "��������� " & n & " �������:"
            Else: Debug.Print "��������� " & n & " �������:"
        End If
Debug.Print loaded_items & vbNewLine


unfinded_print:
        If n1 = 0 Then GoTo errors_print
        
        If n1 Mod 10 = 1 Then
                Debug.Print "����������� �� ����� " & n1 & " �������:"
            ElseIf n1 Mod 10 = 2 Or n1 Mod 10 = 3 Or n1 Mod 10 = 4 Then
                Debug.Print "����������� �� ����� " & n1 & " �������:"
            Else: Debug.Print "����������� �� ����� " & n1 & " �������:"
        End If
Debug.Print unfinded_items


errors_print:
        If n2 = 0 Then GoTo finish_sub
        
        If n2 Mod 10 = 1 Then
                Debug.Print "������ �������� " & n2 & " �������:"
            ElseIf n2 Mod 10 = 2 Or n2 Mod 10 = 3 Or n2 Mod 10 = 4 Then
                Debug.Print "������ �������� " & n2 & " �������:"
            Else: Debug.Print "������ �������� " & n2 & " �������:"
        End If
Debug.Print error_items



finish_sub:
Debug.Print ("������� ����� �������� �������: " & Format(Allt / n, "#") & "�")
MsgBox ("�������� ��������(" & Format(Allt / 60, "#") & "��� " & Format(Allt Mod 60, "#") & "�)" & vbCrLf & _
        "������� ����� �������� �������: " & Format(Allt / n, "#") & "�")
End Sub




'======================================================================================
'======================================================================================
'       ��������������� �������
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




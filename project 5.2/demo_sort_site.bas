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
    Dim n As Double         '������� ����������� �������
    Dim Allt As Double      '������ ��� ����� ���������

Sub Sort_site()
Allt = Timer


        '������ ������� ������
    Set Ch = New Selenium.ChromeDriver
        URL = "https://ural-soft.info/editor/"                                                                                  '����
        Login = "login"                                                                                                         '�����
        Password = "parol"                                                                                                      '������
        ImportURL = "https://ural-soft.info/editor/structure/editsection/importexport/"                                         '������ ���
            Uploads = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\��������\"                         '���� � ������ ���
        SortURL = "https://ural-soft.info/editor/structure/editsection/"                                                        '���������� ��������
        IndexURL = "https://ural-soft.info/editor/structure/indexthissite/"                                                     '���������� ������ �� �����
            Sorting_done = "uss_editor_message_ok"                                                                              '����������� � �������� ����������
            SortByPrice = "sortorder-4-styler"                                                                                  '���������� �� ���� �� �����������
            ShowAll = "sortorder-11-styler"                                                                                     '���������. �������� ���
            HideAll = "sortorder-16-styler"                                                                                     '���������. ������ ��� � �������


        '��������� ������� � ���������
    Ch.AddArgument "start-maximized"
    Ch.start
    Ch.Get URL
    Ch.FindElement(FindBy.Name("login")).SendKeys Login
    Ch.FindElement(FindBy.Name("password")).SendKeys Password
    Ch.FindElement(FindBy.Name("send_login_data")).Click


        '������� ������ �������� ��� ����������
    Dim Category(31) As String

        Category(0) = "1 � LED ������"
        Category(1) = "1 � �������� ������������ � ������������ ������� � Adamson"
        Category(2) = "1 � �������� ������������ � ������������ ������� � D&B Audiotechnik"
        Category(3) = "1 � �������� ������������ � ������������ ������� � L-Acoustics"
        Category(4) = "1 � �������� ������������ � ������������ ������� � Martin Audio"
        Category(5) = "1 � �������� ������������ � ������������ ������� � ProTone"
        Category(6) = "1 � �������� ������������ � ��������� ������ � DiGiCo"
        Category(7) = "1 � �������� ������������ � ���������� ���������� ������������� ��������� � L-Acoustics"
        Category(8) = "1 � �������� ������������ � ���������� ���������� ������������� ��������� � Martin Audio"
        Category(9) = "1 � �������� ������������ � ��������� � Adamson"
        Category(10) = "1 � �������� ������������ � ��������� � D&B Audiotechnik"
        Category(11) = "1 � �������� ������������ � ��������� � L-Acoustics"
        Category(12) = "1 � �������� ������������ � ��������� � Martin Audio"
        Category(13) = "1 � �������� ������������ � ��������� � ProTone"
        Category(14) = "1 � ���������� � ���� � ������ � ������ � ������������ ������ � Draka"
        Category(15) = "1 � ���������� � ���� � ������ � ������ � ������������ ������ � Wiring Parts"
        Category(16) = "1 � ���������� � ���� � ������ � ������ � ������������ ������ � XSE"
        Category(17) = "1 � ���������� � ���� � ����� ����������� � EDS"
        Category(18) = "1 � ���������� � ���� � ������� ������������� � ������������� ��������� � EDS"
        Category(19) = "1 � ���������� � ���� � ������� ������������� � ������������� ������� � EDS"
        Category(20) = "1 � ���������� � ���� � ���������� �������� ������� � EDS"
        Category(21) = "1 � �������� ����� � �����"
        Category(22) = "1 � �������� ����� � ����� � �������" '������ ��� � ������� �����
        Category(23) = "1 � �������� ����� � ����� � ������� � Imlight"
        Category(24) = "1 � �������� ����� � ����� � ����������� ����� � Imlight" '������ ��� � ������� �����
        Category(25) = "1 � ��������� � ������������ � �������� ������������ � Shure � ����� Shure Axient"
        Category(26) = "1 � ��������� � ������������ � �������� ������������ � Shure � ����� Shure ULXD"
        Category(27) = "1 � �������� ������������ � ����������� ������ � DTS"
        Category(28) = "1 � �������� ������������ � ������� ��������� ����� � ETC"
        Category(29) = "1 � �������� ������������ � ������� ���������� ������ � �������� ������� � ETC"
        Category(30) = "1 � �������� ������������ � ����������� ���������� � DTS"
        Category(31) = "1 � �������� ������������ � ����������� ���������� � ETC"


        '��������� ��� �������
    Ch.Get SortURL
        Ch.FindElementById(SortByPrice).Click
        Ch.FindElementById(ShowAll).Click
        Ch.FindElementByName("sortpos").Click
            If Not Ch.IsElementPresent(FindBy.Class(Sorting_done)) Then
                MsgBox "������ �������� ���������� ��� ������ � ����", vbExclamation
                Exit Sub
            End If
        Debug.Print "1.1 ���������: �������� ��� �������."


        '�������� � �������� ������
    Ch.Get SortURL
        Ch.FindElementById(HideAll).Click
        Ch.FindElementByName("sortpos").Click
            If Not Ch.IsElementPresent(FindBy.Class(Sorting_done)) Then
                MsgBox "������ �������� ���������� ��� ������ � ����", vbExclamation
                Exit Sub
            End If
        Debug.Print "1.2 ���������: ������ ��� ������� � ������� �����."


        '��������� � ��������� ������ ������� ��������
    For i = 0 To UBound(Category)
    
    
        Ch.Get SortURL
            Ch.FindElement(FindBy.Class("jq-selectbox__trigger-arrow")).Click
                If Not Ch.IsElementPresent(FindBy.XPath("/html/body/div[3]/ul/li[text()='" & Category(i) & "']")) Then
                    MsgBox Category(i) & " - �� ��������� � ������", vbExclamation
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
                    MsgBox "������ �������� ���������� ��� ������ � ����. �� ������������ ������ " & Category(i), vbExclamation
                    Exit Sub
                End If
            Debug.Print "2." & i + 1 & " ������������ ������: " & Category(i)
    
    Next i


        '����������� ����� �� �����
    Ch.Get IndexURL
    Ch.Quit
    
Allt = Timer - Allt
    Debug.Print "�������� � ���������� �������� ������ �������! (" & Format(Allt / 60, "#") & "��� " & Format(Allt Mod 60, "#") & "�)"
    MsgBox "�������� � ���������� �������� ������ �������! (" & Format(Allt / 60, "#") & "��� " & Format(Allt Mod 60, "#") & "�)"
    
End Sub


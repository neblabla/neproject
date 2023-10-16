Attribute VB_Name = "demo_download_prices"
Option Explicit
    Dim Ch As Selenium.ChromeDriver                     '���������� ��� ������ � ���������
    Dim FindBy As New Selenium.By                       '���������� ��������� ������� By
            '���������� ��� ��������� ����������
    Dim article As String, URL As String, URL2 As String, Login As String, Password As String, LoginCSS As String, PasswordCSS As String, ButtonCSS As String, FileName As String, filePath As String
    Dim DownloadFolder As String, Prices As String    '����� ��� �������� � ������
    
    Dim Request(2, 4) As String     '��� ���������� ������� � �����   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Openbook As Boolean         '��� ��������� �������� ���� � ��������
    Dim ErrorI As Double           '��� �������� ������
    Dim ErrorMessage As String      '��������� �� �������
    Dim successmail As Boolean        '���� ������ � �����
    Dim t As Double                 '������ ��� ���������� ��������
    Dim Allt As Double              '������ ��� ����� ���������

'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : DownloadFromMail
'
' Purpose   : Support Function
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Function DownloadFromMail()

On Error GoTo ErrorHandl
Debug.Print "���������� ������� � ����� ..."

    URL = "https://mail.yandex.ru/?uid=1130000029312481#inbox"
    Login = "login"
    Password = "parol"

    Set Ch = New Selenium.ChromeDriver
        DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\��������� ��� ����������\"
        Ch.SetPreference "download.default_directory", (DownloadFolder)
        Ch.AddArgument "start-maximized"
        Ch.start
        Ch.Get URL
        
        If Ch.IsElementPresent(FindBy.XPath("//*[@id='js-button']"), 3000) Then
            Stop
        End If


' ������� �� �������� �����������
'======================================================================================
ButtonCSS = "#root > div.Background_1BqcnsQB2gBAMrnbsU_Okz > div > header > div.SmallHeader_10gL9KctcHHvOwoPL2Nnmy > div.ActionButtons_1KQUh4y2uqGFcS5C_M9sDV > a.Button2.Button2_type_link.Button2_view_default.Button2_size_m"
    If Not Ch.IsElementPresent(FindBy.Css(ButtonCSS)) Then
        Debug.Print "�������� ����� ������ �����": GoTo ErrorHandl
    End If: Ch.FindElement(FindBy.Css(ButtonCSS)).Click

ButtonCSS = "div.AuthLoginInputToggle-wrapper > div:nth-child(1) > button"
    Ch.FindElement(FindBy.Css(ButtonCSS)).Click


' ���� ������
'======================================================================================
LoginCSS = "#passp-field-login"
    If Not Ch.IsElementPresent(FindBy.Css(LoginCSS), 3000) Then
        Debug.Print "�������� ����� ���� ����� ������" ': GoTo ErrorHandl
    End If
    Ch.FindElement(FindBy.Css(LoginCSS)).SendKeys Login

ButtonCSS = "div.passp-button.passp-sign-in-button"
    If Not Ch.IsElementPresent(FindBy.Css(ButtonCSS), 3000) Then
        Debug.Print "�������� ����� ������ ������" ': GoTo ErrorHandl
    End If: Ch.FindElement(FindBy.Css(ButtonCSS)).Click


' ���� ������
'======================================================================================
PasswordCSS = "#passp-field-passwd"
    If Not Ch.IsElementPresent(FindBy.Css(PasswordCSS), 3000) Then
        Debug.Print "�������� ����� ���� ����� ������" ': GoTo ErrorHandl
    End If: Ch.FindElement(FindBy.Css(PasswordCSS)).SendKeys Password

ButtonCSS = "#passp\:sign-in"
    If Not Ch.IsElementPresent(FindBy.Css(ButtonCSS), 3000) Then
        Debug.Print "�������� ����� ������ ������" ': GoTo ErrorHandl
    End If: Ch.FindElement(FindBy.Css(ButtonCSS)).Click


    Debug.Print "����������� ������ �������"

' ����� ��������� � ���������� ������
'======================================================================================
    Request(0, 0) = "LO":       Request(1, 0) = "������� �� ��� ""�������� �����"""
    Request(0, 1) = "AR":       Request(1, 1) = "����� ����������"
    Request(0, 2) = "NS":       Request(1, 2) = "�����-���� ��� ������� ����-�����"
    Request(0, 3) = "MI":       Request(1, 3) = "MixArt (����"
    Request(0, 4) = "CTC":      Request(1, 4) = "������� ���"
    


    Dim n As Integer
    For n = 0 To UBound(Request, 2)
    
                '����� ������ � �������
                '======================================================================================
                    Dim SearchClass As String
                        SearchClass = "textinput__control"
                    If Not Ch.IsElementPresent(FindBy.Class(SearchClass), 10000) Then
                        Debug.Print Request(0, n) & " - �������� ����� ���� ������" ': GoTo ErrorHandl
                    End If: Ch.FindElement(FindBy.Class(SearchClass)).SendKeys Request(1, n)
                
                    ButtonCSS = "div.search-input.search-input_search-icon_right > form > button"
                    If Not Ch.IsElementPresent(FindBy.Css(ButtonCSS), 10000) Then
                        Debug.Print Request(0, n) & " - �������� ����� ������ ������" ': GoTo ErrorHandl
                    End If: Ch.FindElement(FindBy.Css(ButtonCSS)).Click
        
                Debug.Print "����� ������ " & Request(0, n) & " �� ������: " & Request(1, n)


                '������� � ������ � ���������� ������
                '======================================================================================
                                
                Dim first_mail_xpath As String: first_mail_xpath = "//*[@id='js-apps-container']/div[2]/div[7]/div/div[2]/div[1]/div[2]/div/main/div[7]/div[1]/div/div/div[3]/div/div[2]/div/div/div/a/div/span[2]/div/span/span[1]"
                Dim i_do As Integer
                i_do = 0
                Do Until InStr(1, Ch.FindElement(FindBy.XPath(first_mail_xpath)).Text, Request(1, n)) >= 1 _
                        Or InStr(1, Ch.FindElement(FindBy.XPath(first_mail_xpath)).Text, "�������") >= 1 _
                        Or i_do = 10
                i_do = i_do + 1
                Ch.Wait (300)
                Loop
                
                    Dim mailXpatch As String: mailXpatch = "//div[@class='ns-view-container-desc mail-MessagesList js-messages-list']/div[2]/div/div/div/a/div/span[1]"
                    If Not Ch.IsElementPresent(FindBy.XPath(mailXpatch), 10000) Then
                        Debug.Print Request(0, n) & " - �������� ����� ������" ': GoTo ErrorHandl
                    End If: Ch.FindElement(FindBy.XPath(mailXpatch)).Click
                Ch.Wait 300
                                                                       '
                    Dim downloadB_Xpatch As String: downloadB_Xpatch = "//*[@id='js-apps-container']/div[2]/div[7]/div/div[2]/div[1]/div[2]/div/main/div[7]/div[2]/div/div/div/div/div[2]/div/div[3]/div/div/div[2]/a"
                    If Not Ch.IsElementPresent(FindBy.XPath(downloadB_Xpatch), 3000) Then
                        Debug.Print Request(0, n) & " - �������� ����� ������ ����������" ': GoTo ErrorHandl
                    End If: Ch.FindElement(FindBy.XPath(downloadB_Xpatch)).Click
                    
                Debug.Print "���������� ������ - " & Request(0, n) & "..."
                    
                    Dim Filename_Xpatch As String: Filename_Xpatch = "//*[@id='js-apps-container']/div[2]/div[7]/div/div[2]/div[1]/div[2]/div/main/div[7]/div[2]/div/div/div/div/div[2]/div/div[3]/div/div/div[1]/a/div/div[2]/span"
                    If Not Ch.IsElementPresent(FindBy.XPath(Filename_Xpatch), 3000) Then
                        Debug.Print Request(0, n) & " - �������� ����� ����� �����" ': GoTo ErrorHandl
                    End If: Request(2, n) = Ch.FindElement(FindBy.XPath(Filename_Xpatch)).Attribute("title")
                Ch.Wait 300
    Next n
    
    Dim i As Integer
    Do Until Dir(DownloadFolder & Request(2, n - 1), vbDirectory) <> vbNullString
        Ch.Wait 100
        i = i + 1
        If i > 100 Then
            GoTo ErrorHandl
        End If
    Loop
    Ch.Quit
    Debug.Print "������ " & Request(0, 0) & ", " & Request(0, 1) & ", " & Request(0, 2) & ", " & Request(0, 3) & " ������� � ����� �������!"
    Debug.Print ""
    successmail = True
Exit Function

ErrorHandl:
        Debug.Print "������ �������� � �����!"
        successmail = False
End Function


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : DownloadFiles
'
' Purpose   : Support Function
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Function DownloadFiles(D_URL As String, D_FileName As String, Optional D_Login As String, Optional D_Password As String, _
    Optional D_LoginCSS As String, Optional D_PasswordCSS As String, Optional D_ButtonCSS As String, Optional D_URL2 As String)

On Error GoTo ErrorHandl
    
    If Dir(DownloadFolder & D_FileName, vbDirectory) <> vbNullString Then
        Kill DownloadFolder & D_FileName
    End If

    Set Ch = New Selenium.ChromeDriver
        Ch.SetPreference "download.default_directory", (DownloadFolder)
        Ch.AddArgument "start-maximized"
        Ch.start
        Ch.Get D_URL


' �����������
'======================================================================================
    If D_LoginCSS <> "" And D_PasswordCSS <> "" And D_ButtonCSS <> "" Then

            If Not Ch.IsElementPresent(FindBy.Css(D_LoginCSS)) Then
                Debug.Print article & " ��������� ����� ���� ����� ������": Exit Function
            End If: Ch.FindElement(FindBy.Css(D_LoginCSS)).SendKeys D_Login

            If Not Ch.IsElementPresent(FindBy.Css(D_PasswordCSS)) Then
                Debug.Print article & " ��������� ����� ���� ����� ������": Exit Function
            End If: Ch.FindElement(FindBy.Css(D_PasswordCSS)).SendKeys D_Password

            If Not Ch.IsElementPresent(FindBy.Css(D_ButtonCSS)) Then
                Debug.Print article & " ��������� ����� ������ �����������": Exit Function
            End If: Ch.FindElement(FindBy.Css(D_ButtonCSS)).Click

    End If

    If Not D_URL2 = "" Then
        Ch.Get D_URL2
    End If

    Dim i As Integer
    Do Until Dir(DownloadFolder & D_FileName, vbDirectory) <> vbNullString
        Ch.Wait (100)
        i = i + 1
        If i > 100 Then
            GoTo ErrorHandl
        End If
    Loop: Ch.Quit
Exit Function

ErrorHandl:
        Ch.Quit
        Debug.Print "������ ��������! " & "������� " & article & " �� ��������. ���������� ������ ��� ��� �����"
End Function


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : UnRAR
'
' Purpose   : Support Function
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Function UnRAR(sArchiveName As String)
    Const sWinRarAppPath As String = "C:\Program Files\WinRAR\WinRAR.exe"
    '��������� ������ �� ������ � ������� ����(vbHide)
    '� ����������� ������������ ������ (-o+)
    Dim sWinRarApp As String
        sWinRarApp = sWinRarAppPath & " E -o+ "
    '��������� ������� �������, ��� �������� ��� �������� � ������ ����� � ����, ������� �������� �������.
    '��� ������� ������� �����������
    UnRAR = shell(sWinRarApp & " """ & DownloadFolder & "\" & sArchiveName & """ """ & DownloadFolder & """ ") ', vbHide)
End Function


Sub All_Downloader()
Allt = Timer
ErrorI = 0
ErrorMessage = ""


    DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\��������� ��� ����������\"
            Prices = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\"

    On Error Resume Next
        Kill DownloadFolder & "*"                                                 '������� ����� ��������
        Kill Prices & "*"
    On Error GoTo 0


    Call DownloadFromMail

    If successmail Then
        Call Download_AR    '���������� ����� ����� --------------------- + ������ PQ
        Call Download_LO    '���������� ����� ����� --------------------- + ������ PQ + �������� ���������
        Call Download_NS    '���������� ����� ����� --------------------- + ������ PQ + �������� ���������
        Call Download_MI    '���������� ����� ����� --------------------- + ������ PQ
        Call Download_CTC   '���������� ����� ����� ----------------------+ ������ PQ + �������� ���� ���������
    End If

    Call Download_AN        '���������� � ����� ------------------------- + ������ PQ + �������� ���� ���������

    Call Download_SL        '���������� DownloadFiles + ���������� ������ + ������ PQ + �������� ���������
    Call Download_FP        '���������� DownloadFiles + ���������� ������ + ������ PQ
    Call Download_OA        '���������� DownloadFiles + ���������� ������ + ������ PQ
    Call Download_AT        '���������� DownloadFiles + ���������� ������ + ������ PQ

    Call Download_SHA       '���������� DownloadFiles ------------------- + ������ PQ + �������� ���������
    Call Download_PA        '���������� DownloadFiles ------------------- + ������ PQ + �������� ���������
    Call Download_D         '���������� DownloadFiles ------------------- + ������ PQ + �������� ���� ���������
    Call Download_AU        '���������� DownloadFiles ------------------- + ������ PQ + �������� ���� ���������

    Call Download_CVG       '---------------------------------------------- ������ PQ + �������� ���������
    Call Download_LU        '---------------------------------------------- ������ PQ + �������� ���������
    Call Download_ARP       '---------------------------------------------- ������ PQ + �������� ���� ���������
    Call Download_LTM       '---------------------------------------------- ������ PQ
    Call Download_GM        '---------------------------------------------- ������ PQ
    Call Download_IM        '---------------------------------------------- ������ PQ
    Call Download_IV        '---------------------------------------------- ������ PQ
    
    Allt = Timer - Allt
    Debug.Print "�������� ������� ��������� (" & Format(Allt / 60, "#") & "��� " & Format(Allt Mod 60, "#") & "�)"
    If ErrorI <> 0 Then
        If ErrorI = 1 Or ErrorI = 21 Then
            Debug.Print "�� ��������:"
            Debug.Print "    " & ErrorI & " ����� - " & Trim(ErrorMessage) & "."
        End If
        If (1 < ErrorI And ErrorI < 5) Or (21 < ErrorI And ErrorI < 25) Then
            Debug.Print "�� ���������:"
            Debug.Print "    " & ErrorI & " ������ - " & Trim(ErrorMessage) & "."
        End If
        If 4 < ErrorI And ErrorI < 21 Then
            Debug.Print "�� ���������:"
            Debug.Print "    " & ErrorI & " ������� - " & Trim(ErrorMessage) & "."
        End If
    Else: Debug.Print "��� ������ ���������� �������"
    End If
    
'    MsgBox "�������� ������� ���������"

End Sub




'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "Artimusic Dictribution"
'
' Article   : AR
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_AR()
On Error GoTo ErrorHandl
t = Timer
        article = "AR "
        FileName = Request(2, 1)
'    DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\��������� ��� ����������\"
'            Prices = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(TDSheet1,8)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������������� �������"" = Table.DuplicateColumn(#""���������� ���������"", ""������������"", ""����� ������������"")," & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"" = Table.RenameColumns(#""������������� �������"",{{""����� ������������"", ""������ ������������""}, {""������������"", ""������������""}, {""���������"", ""�������""}, {""���"", ""�������""}})," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""��������������� �������"",{""�������"", ""������������"", ""������ ������������"", ""�������""})," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ ��������� �������"",{{""�������"", Currency.Type}})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""���������� ���"", each ([�������] <> null))" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [TDSheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "Logos"
'
' Article   : LO
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_LO()

'On Error GoTo ErrorHandl
t = Timer
        article = "LO "
        FileName = Request(2, 0)
'    DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\��������� ��� ����������\"
'            Prices = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & "    #""���������� ���������"" = Table.PromoteHeaders(TDSheet1, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""���������� ���������"",{{""��� ������������"", type text}, {""�������������"", type text}, {""������������"", type text}, {""�������������� ��������"", type text}, {""����������"", Int64.Type}, {""���� �������, ���"", type text}, {""���� �����, ���"", type text}})," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���"",{""��� ������������"", ""�������������"", ""������������"", ""�������������� ��������"", ""���� �������, ���""})," & Chr(13) & "" & Chr(10) & _
                "    #""������������ �������"" = Table.CombineColumns(#""������ ��������� �������"",{""�������������"", ""������������""},Combiner.CombineTextByDelimiter("" "", QuoteStyle.None),""������������"")," & Chr(13) & "" & Chr(10) & _
                "    #""������������� �������"" = Table.DuplicateColumn(#""������������ �������"", ""������������"", ""����� ������������"")," & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"" = Table.ReorderColumns(#""������������� �������"",{""��� ������������"", ""������������"", ""����� ������������"", ""�������������� ��������"", ""���� �������, ���""})," & Chr(13) & "" & Chr(10) & _
                "    #""������������ �������1"" = Table.CombineColumns(#""����������������� �������"",{""����� ������������"", ""�������������� ��������""},Combiner.CombineTextByDelimiter("" - "", QuoteStyle.None),""������ ������������"")" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""������������ �������1"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [TDSheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "TDSheet"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
Exit Sub

ErrorHandl2:
        Debug.Print "������� " & article & " �� ��������. ������� �� ����������� �� ��������� ������ ��� ���������� �������� �����"
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "���� �����"
'
' Article   : NS
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_NS()

On Error GoTo ErrorHandl
t = Timer
        article = "NS "
        FileName = Request(2, 2)


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(TDSheet1,6)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""�������"", ""������������"", ""�������������"", ""����""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each [������������] <> null and [������������] <> """")," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ � ����������� ��������"",{{""����"", Currency.Type}})," & Chr(13) & "" & Chr(10) & _
                "    #""������������� �������"" = Table.DuplicateColumn(#""���������� ���"", ""�������"", ""����� �������"")," & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"" = Table.ReorderColumns(#""������������� �������"",{""�������"", ""�������������"", ""����� �������"", ""������������"", ""����""})," & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"" = Table.RenameColumns(#""����������������� �������"",{{""������������"", ""������ ������������""}})," & Chr(13) & "" & Chr(10) & _
                "    #""������������ �������"" = Table.CombineColumns(#""��������������� �������"",{""�������������"", ""����� �������""},Combiner.CombineTextByDelimiter("" "", QuoteStyle.None),""������������"")" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""������������ �������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [TDSheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "TDSheet"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "MixArt"
'
' Article   : MI
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_MI()

On Error GoTo ErrorHandl
t = Timer
        article = "MI "
        FileName = Request(2, 3)
'          FileName = ""
    DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\��������� ��� ����������\"
'�===================================================

    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    Source = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet1 = Source{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""Renamed Columns"" = Table.RenameColumns(TDSheet1,{{""Column1"", ""�������""}, {""Column2"", ""�������������""}, {""Column3"", ""������������""}, {""Column6"", ""�������""}})," & Chr(13) & "" & Chr(10) & _
                "    #""Removed Top Rows"" = Table.Skip(#""Renamed Columns"",1)," & Chr(13) & "" & Chr(10) & _
                "    #""Filtered Rows"" = Table.SelectRows(#""Removed Top Rows"", each ([�������] <> null))," & Chr(13) & "" & Chr(10) & _
                "    #""Removed Other Columns"" = Table.SelectColumns(#""Filtered Rows"",{""�������"", ""�������������"", ""������������"", ""�������""})" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""Removed Other Columns"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [TDSheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "TDSheet"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "Anzhee"
'
' Article   : AN
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_AN()

'On Error GoTo ErrorHandl
t = Timer
        article = "AN "
        URL = "https://drive.google.com/drive/folders/1yl7BKEpFoU1zqfQM9mVcE-BurxcktCJ9"
        FileName = ""


    DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\��������� ��� ����������\"
        Prices = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\"

    Set Ch = New Selenium.ChromeDriver
        Ch.SetPreference "download.default_directory", (DownloadFolder)
        Ch.AddArgument "start-maximized"
        Ch.start
        Ch.Get URL

' ������������ ���� !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'======================================================================================
    ButtonCSS = "//*[@id='drive_main_page']/div/div[3]/div[1]/div/div/div/div[2]/div/div[1]/div/div/div[2]/div/div[3]"
    
        If Ch.IsElementPresent(FindBy.XPath(ButtonCSS), 3000) Then
            If Ch.FindElement(FindBy.XPath(ButtonCSS)).Attribute("aria-label") = "�����" Then
                ButtonCSS = "#drive_main_page > div > div.g3Fmkb > div.S630me > div > div > div > div:nth-child(2) > div > div:nth-child(1) > div > div > div.a-s-tb-sc-Ja-Q.a-s-tb-sc-Ja-Q-Nm.a-Ba-Ed.a-s-Ba-dj.a-s-Ba-Ed-Be-nAm6yf > div > div:nth-child(3) > div > svg"
                Ch.FindElement(FindBy.Css(ButtonCSS), 3000).ClickAndHold
                Ch.Wait 100
                Ch.FindElement(FindBy.Css(ButtonCSS), 3000).Click
            End If
        End If


' ��������� ����� �����
'======================================================================================                                        '1
    ButtonCSS = "//*[@id=':1']/div/c-wiz/div[2]/c-wiz/div[1]/c-wiz/div/c-wiz/div[1]/c-wiz[2]/c-wiz/div/c-wiz[2]/div/div/div/div[1]/div[4]"
        If Not Ch.IsElementPresent(FindBy.XPath(ButtonCSS), 3000) Then
            Debug.Print article & " ��������� ����� ������ 1": GoTo ErrorHandl
        End If
        FileName = Ch.FindElement(FindBy.XPath(ButtonCSS)).Text


' ���������� �����
'======================================================================================
    ButtonCSS = "//*[@id=':1']/div/c-wiz/div[2]/c-wiz/div[1]/c-wiz/div/c-wiz/div[1]/c-wiz[2]/c-wiz/div/c-wiz[2]/div/div/div/div[2]/div[1]/div/div[2]/div[1]"
        If Not Ch.IsElementPresent(FindBy.XPath(ButtonCSS), 3000) Then
            Debug.Print article & " ���������� ���� �����": GoTo ErrorHandl
        End If
        Ch.FindElement(FindBy.XPath(ButtonCSS)).ReleaseMouse

    ButtonCSS = "//*[@id=':1']/div/c-wiz/div[2]/c-wiz/div[1]/c-wiz/div/c-wiz/div[1]/c-wiz[2]/c-wiz/div/c-wiz[2]/div/div/div/div[2]/div[2]"
        If Not Ch.IsElementPresent(FindBy.XPath(ButtonCSS), 3000) Then
            Debug.Print article & " ��������� ����� ������ 2": GoTo ErrorHandl
        End If
        Ch.FindElement(FindBy.XPath(ButtonCSS)).Click


    Dim i As Double
    Do Until Dir(DownloadFolder & FileName, vbDirectory) <> vbNullString
        Ch.Wait (100)
        i = i + 0.1
        If i > 30 Then
            GoTo ErrorHandl
        End If
    Loop
    Ch.Quit


    Workbooks.Add: Openbook = True
    
    '��� ��������� " ���." � ������� ��������
'            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
'                "let" & Chr(13) & "" & Chr(10) & _
'                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
'                "    TDSheet_Sheet = ��������{[Item=""TDSheet"",Kind=""Sheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
'                "    #""��������� ������� ������"" = Table.Skip(TDSheet_Sheet,7)," & Chr(13) & "" & Chr(10) & _
'                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
'                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""���"", ""������������, �����������"", ""������������ ��� ������"", ""����, ������� � ��� RUB""})," & Chr(13) & "" & Chr(10) & _
'                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each [#""������������, �����������""] <> null and [#""������������, �����������""] <> """")," & Chr(13) & "" & Chr(10) & _
'                "    #""���������� ��������"" = Table.ReplaceValue(#""������ � ����������� ��������"","" ���."","""",Replacer.ReplaceText,{""����, ������� � ��� RUB""})," & Chr(13) & "" & Chr(10) & _
'                "    #""���������� ��� � ������"" = Table.TransformColumnTypes(#""���������� ��������"", {{""����, ������� � ��� RUB"", Currency.Type}}, ""en-US"")" & Chr(13) & "" & Chr(10) & _
'                "in" & Chr(13) & "" & Chr(10) & _
'                "    #""���������� ��� � ������"""
                
    '��� ���������
            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet_Sheet = ��������{[Item=""TDSheet"",Kind=""Sheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(TDSheet_Sheet,7)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""���"", ""������������, �����������"", ""������������ ��� ������"", ""����, ������� � ��� RUB""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each [#""������������, �����������""] <> null and [#""������������, �����������""] <> """")," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ��� � ������"" = Table.TransformColumnTypes(#""������ � ����������� ��������"", {{""����, ������� � ��� RUB"", Currency.Type}}, ""en-US"")" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""���������� ��� � ������"""
                
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [TDSheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "TDSheet"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"            '���� �������� � ������� ������, �� ���� ��� ���������� ( ��� ��� �������)

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "CVGAudio"
'
' Article   : CVG
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_CVG()

On Error GoTo ErrorHandl
t = Timer
        article = "CVG"
        URL = "https://www.cvg.ru/docs/CVGAUDIO_DEALER_pricelist.xlsx"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="�����-���� CVGAUDIO", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(Web.Contents(""" & URL & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    #""�����-���� CVGAUDIO_Sheet"" = ��������{[Item=""�����-���� CVGAUDIO"",Kind=""Sheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(#""�����-���� CVGAUDIO_Sheet"",5)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������1"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������1"",{""�������"", ""������"", ""������������"", ""������� (���.)""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each ([�������] <> null))," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ � ����������� ��������"",{{""�������"", type text}, {""������"", type text}, {""������������"", type text}, {""������� (���.)"", Currency.Type}})," & Chr(13) & "" & Chr(10) & _
                "    #""�������� ���������������� ������"" = Table.AddColumn(#""���������� ���"", ""������ ������������"", each [������] & "" - "" & [������������])," & Chr(13) & "" & Chr(10) & _
                "    #""��������� �������"" = Table.RemoveColumns(#""�������� ���������������� ������"",{""������������""})," & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"" = Table.ReorderColumns(#""��������� �������"",{""�������"", ""������"", ""������ ������������"", ""������� (���.)""})," & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"" = Table.RenameColumns(#""����������������� �������"",{{""�������"", ""�������""}, {""������"", ""������������""}})" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""�����-���� CVGAUDIO"";Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [�����-���� CVGAUDIO]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "�����_����_CVGAUDIO"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "Slami"
'
' Article   : SL
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_SL()

On Error GoTo ErrorHandl
t = Timer
        article = "SL "
        URL = "http://www.slami.ru/info/dilprice_slami.zip"
        FileName = "dilprice_slami.zip"


'    DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\��������� ��� ����������\"
'            Prices = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\"

    Call DownloadFiles(URL, FileName, Login, Password)
    Call UnRAR(FileName)
    Application.Wait Now + TimeSerial(0, 0, 1)

    Dim i As Integer
    Do Until Dir(DownloadFolder & "SLAMI PL " & Format(Date - i, "dd.mm.yyyy") & ".xls", vbDirectory) <> vbNullString
        i = i + 1
            If i = 7 Then
                GoTo ErrorHandl2
            End If
    Loop
    FileName = "SLAMI PL " & Format(Date - i, "dd.mm.yyyy") & ".xls"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(TDSheet1,4)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""���������� ���������"", each ([����������] <> ""���� � ������������""))," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""������ � ����������� ��������"",{""���"", ""������������"", ""���""})," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ ��������� �������"",{{""���"", Currency.Type}})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������1"" = Table.SelectRows(#""���������� ���"", each ([���] <> null))," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ��������"" = Table.ReplaceValue(#""������ � ����������� ��������1"",""  "","" - "",Replacer.ReplaceText,{""������������""})," & Chr(13) & "" & Chr(10) & _
                "    #""����������� ����� ����� ������������"" = Table.AddColumn(#""���������� ��������"", ""����� ����� ������������"", each Text.BeforeDelimiter([������������], "" -""), type text)," & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"" = Table.RenameColumns(#""����������� ����� ����� ������������"",{{""����� ����� ������������"", ""������� ������������""}})," & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"" = Table.ReorderColumns(#""��������������� �������"",{""���"", ""������� ������������"", ""������������"", ""���""})" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [TDSheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "TDSheet"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"
    
URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
Exit Sub

ErrorHandl2:
        Debug.Print "������� " & article & " �� ��������. ������� �� ����������� �� ��������� ������ ��� ���������� �������� �����"
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "���� �����"
'
' Article   : OA
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_OA()

'On Error GoTo ErrorHandl
t = Timer
        article = "OA "
        URL = "https://price.okno-audio.ru/Price.zip"
        FileName = "Price.zip"


    Call DownloadFiles(URL, FileName, Login, Password)
    Call UnRAR(FileName)
    FileName = "Price.xls"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="�������� �����-����", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    #""�������� �����-����1"" = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(#""�������� �����-����1"",7)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""�������"", ""������������"", ""��������� ����, ���.""})," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ ��������� �������"",{{""��������� ����, ���."", Currency.Type}})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""���������� ���"", each [#""�������""] <> null and [#""�������""] <> """")" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""�������� �����-����"";Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [�������� �����-����]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "��������_�����_����"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "����� ���"
'
' Article   : FP
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_FP()

On Error GoTo ErrorHandl
t = Timer
        article = "FP "
        URL = "http://www.pop-music.ru/files/focuspro.zip"
        FileName = "focuspro.zip"
    DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\��������� ��� ����������\"
            Prices = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\"


    Call DownloadFiles(URL, FileName, Login, Password)
    Call UnRAR(FileName)
    Application.Wait Now + TimeSerial(0, 0, 1)
        Dim i As Integer


    FileName = "������� �������� "
        Do Until Dir(DownloadFolder & FileName & Format(Date - i, "dd.mm") & ".xls", vbDirectory) <> vbNullString
            i = i + 1
                If i = 7 Then
                    GoTo VariantName2
                End If
        Loop
GoTo nextstep

VariantName2:
    i = 0
    FileName = "��������"
        Do Until Dir(DownloadFolder & FileName & Format(Date - i, "dd.mm") & ".xls", vbDirectory) <> vbNullString
            i = i + 1
                If i = 7 Then
                    GoTo VariantName3
                End If
        Loop
GoTo nextstep

VariantName3:
    i = 0
    FileName = "�������� "
        Do Until Dir(DownloadFolder & FileName & Format(Date - i, "dd.mm") & ".xls", vbDirectory) <> vbNullString
            i = i + 1
                If i = 7 Then
                    GoTo VariantName4
                End If
        Loop
GoTo nextstep

VariantName4:
    i = 0
    FileName = "������� "
        Do Until Dir(DownloadFolder & FileName & Format(Date - i, "dd.mm") & ".xls", vbDirectory) <> vbNullString
            i = i + 1
                If i = 7 Then
                    GoTo ErrorHandl2
                End If
        Loop


nextstep:
    FileName = FileName & Format(Date - i, "dd.mm") & ".xls"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(TDSheet1,7)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""������������.������� "", ""������������"", ""���� ���������""})," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ ��������� �������"",{{""���� ���������"", Currency.Type}})," & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"" = Table.RenameColumns(#""���������� ���"",{{""���� ���������"", ""�������""}, {""������������.������� "", ""�������""}})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""��������������� �������"", each [�������] <> null and [�������] <> """")" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [TDSheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "TDSheet"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
Exit Sub

ErrorHandl2:
        Debug.Print "������� " & article & " �� ��������. ������� �� ����������� �� ��������� ������ ��� ���������� �������� �����"
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "�&T Trade"
'
' Article   : AT
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_AT()

On Error GoTo ErrorHandl
t = Timer
        article = "AT "
        URL = "http://www.attrade.ru/pr30fe16ff-24c4-4bd1-be40-69cbd4593924/TempFolder/Price_Stock.zip"
        FileName = "Price_Stock.zip"
        Login = "login"
        Password = "parol"


    Call DownloadFiles(URL, FileName, Login, Password)
    Call UnRAR(FileName)
    FileName = "Price_Stock.XLS"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="�����-����, ������� �� ������", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    #""�����-����, ������� �� ������1"" = ��������{[Name=""�����-����, ������� �� ������""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(#""�����-����, ������� �� ������1"",17)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""���"", ""������������"", ""���������"", ""�����"", ""���������� �������""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each ([������������] <> null and [������������] <> """") and ([���������] <> ""����� � ������������""))," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ � ����������� ��������"",{{""���������� �������"", Currency.Type}})," & Chr(13) & "" & Chr(10) & _
                "    #""��������� �������"" = Table.RemoveColumns(#""���������� ���"",{""���������""})" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""��������� �������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""�����-����, ������� �� ������"";Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [�����-����, ������� �� ������]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "�����_����__�������_��_������"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "Dynatone"
'
' Article   : D
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_D()

'On Error GoTo ErrorHandl
t = Timer
        article = "D  "
        URL = "https://apidnt.ru/v2/download/DynaTone_dealer_price.xls?key=1b82fEC8-277-5123-554b-737B-af820D65B263"
        FileName = "DynaTone_dealer_price.xls"
        Login = "login"
        Password = "parol"

'    DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\��������� ��� ����������\"
'            Prices = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\"


    Call DownloadFiles(URL, FileName, Login, Password)

    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(TDSheet1,6)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"" = Table.RenameColumns(#""���������� ���������"",{{""Column1"", ""�������""}})," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""��������������� �������"",{""�������"", ""������������ �������"", ""��������� ���."", ""��������""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each [�������] <> null and [�������] <> """")," & Chr(13) & "" & Chr(10) & _
                "    #""��������� �������"" = Table.RemoveColumns(#""������ � ����������� ��������"",{""��������""})" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""��������� �������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [TDSheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "TDSheet"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "Show Atelier"
'
' Article   : SHA
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_SHA()

On Error GoTo ErrorHandl
t = Timer
        article = "SHA"
        URL = "https://showatelier.ru/login.html"
        URL2 = "https://showatelier.ru/download.html"
        FileName = "SAStock.xls"
        Login = "login"
        Password = "parol"
        LoginCSS = "label.loginUsernameLabel > input"
        PasswordCSS = "label.loginPasswordLabel > input"
        ButtonCSS = "span > input[type=submit]"

    Call DownloadFiles(URL, FileName, Login, Password, LoginCSS, PasswordCSS, ButtonCSS, URL2)


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(TDSheet1,8)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""�������"", ""�����"", ""������������"", ""���������""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each [�������] <> null and [�������] <> """")," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ � ����������� ��������"",{{""�������"", type text}, {""�����"", type text}, {""������������"", type text}, {""���������"", Currency.Type}})," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� �� �����������"" = Table.SplitColumn(#""���������� ���"", ""������������"", Splitter.SplitTextByDelimiter(""#(lf)"", QuoteStyle.Csv), {""������������.1"", ""������������.2"", ""������������.3"", ""������������.4""})," & Chr(13) & "" & Chr(10) & _
                "    #""������������ �������"" = Table.CombineColumns(#""��������� ������� �� �����������"",{""�����"", ""������������.1""},Combiner.CombineTextByDelimiter("" "", QuoteStyle.None),""������������"")," & Chr(13) & "" & Chr(10) & _
                "    #""������������ �������1"" = Table.CombineColumns(#""������������ �������"",{""������������.2"", ""������������.3"", ""������������.4""},Combiner.CombineTextByDelimiter("" "", QuoteStyle.None),""��������"")," & Chr(13) & "" & Chr(10) & _
                "    #""�������� ���������������� ������"" = Table.AddColumn(#""������������ �������1"", ""������ ������������"", each [������������] & "" - "" & [��������])," & Chr(13) & "" & Chr(10) & _
                "    #""��������� �������"" = Table.RemoveColumns(#""�������� ���������������� ������"",{""��������""})," & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"" = Table.ReorderColumns(#""��������� �������"",{""�������"", ""������������"", ""������ ������������"", ""���������""})" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [TDSheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "TDSheet"
                .Refresh BackgroundQuery:=False
            End With

    Range("E2").Select
        ActiveCell.FormulaR1C1 = "=TRIM([@������������])"
    Range("F2").Select
        ActiveCell.FormulaR1C1 = "=TRIM([@[������ ������������]])"

    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

    Range("B1:C1").Cut Range("E1")
    Columns("E:F").Cut Destination:=Columns("B:C")

        Columns("A:A").ColumnWidth = 13.71
        Columns("A:A").NumberFormat = "@"
        Columns("B:B").ColumnWidth = 35.57
        Columns("C:C").ColumnWidth = 59.29
        Columns("D:D").EntireColumn.AutoFit
        Range("A1:D1").Font.Bold = True
        Range("A1").Select

    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "PASystem"
'
' Article   : PA
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_PA()

On Error GoTo ErrorHandl
t = Timer
        article = "PA "
        URL = "https://dealers:pasrussia@dealer.pasystem.ru/Ostatki_price.xlsx"
        FileName = "Ostatki_price.xlsx"
        Login = "login"
        Password = "parol"
        
    Call DownloadFiles(URL, FileName, Login, Password)


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet_Sheet = ��������{[Item=""TDSheet"",Kind=""Sheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(TDSheet_Sheet,3)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""������� / ������"", ""������������"", ""��������� ����,        ���.""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each [������������] <> null and [������������] <> """")," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ � ����������� ��������"",{{""������� / ������"", type text}, {""������������"", type text}, {""��������� ����,        ���."", Currency.Type}})," & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"" = Table.RenameColumns(#""���������� ���"",{{""������� / ������"", ""�������""}, {""��������� ����,        ���."", ""��������� ����""}})" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [TDSheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "TDSheet"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "AUVIX"
'
' Article   : AU
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_AU()

On Error GoTo ErrorHandl
t = Timer
        article = "AU "
        URL = "https://b2b.auvix.ru/pricelist/dealer/Price_AUVIX_dealer_xlsx-xlsx/"
        FileName = "Price_AUVIX_dealer_xlsx.xlsx"
        Login = "login"
        Password = "parol"
        LoginCSS = "div:nth-child(1) > input"
        PasswordCSS = "div:nth-child(2) > input"
        ButtonCSS = "div:nth-child(2) > button"


    Call DownloadFiles(URL, FileName, Login, Password, LoginCSS, PasswordCSS, ButtonCSS)


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="Worksheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    Worksheet_Sheet = ��������{[Item=""Worksheet"",Kind=""Sheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(Worksheet_Sheet,9)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������1"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������1"",{""�������"", ""������������"", ""��������� ����, ���.""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each [�������] <> null and [�������] <> """")" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Worksheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [Worksheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "Worksheet"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "Aris"
'
' Article   : ARP
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_ARP()

On Error GoTo ErrorHandl
t = Timer
        article = "ARP"
        URL = "https://arispro.ru/price/Aris_ostatki.xls"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="ARP", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(Web.Contents(""" & URL & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    ARP1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(ARP1,1)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""���"",""�������"", ""������������"", ""��������� (���.)""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each ([#""��������� (���.)""] <> null))," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ � ����������� ��������"",{{""��������� (���.)"", Currency.Type}})," & Chr(13) & "" & Chr(10) & _
                "    #""����������� �������"" = Table.UnpivotOtherColumns(#""���������� ���"", {""������������"", ""��������� (���.)""}, ""�������"", ""��������"")," & Chr(13) & "" & Chr(10) & _
                "    #""��������� �������"" = Table.RemoveColumns(#""����������� �������"",{""�������""})," & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"" = Table.RenameColumns(#""��������� �������"",{{""��������"", ""�������""}})," & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"" = Table.ReorderColumns(#""��������������� �������"",{""�������"", ""������������"", ""��������� (���.)""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������1"" = Table.SelectRows(#""����������������� �������"", each ([�������] <> ""          ""))" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������1"""
'            ActiveWorkbook.Queries.Add Name:="ARP", Formula:= _
'                "let" & Chr(13) & "" & Chr(10) & _
'                "    �������� = Excel.Workbook(Web.Contents(""" & URL & """), null, true)," & Chr(13) & "" & Chr(10) & _
'                "    ARP1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
'                "    #""��������� ������� ������"" = Table.Skip(ARP1,1)," & Chr(13) & "" & Chr(10) & _
'                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
'                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""�������"", ""������������"", ""��������� (���.)""})," & Chr(13) & "" & Chr(10) & _
'                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ ��������� �������"",{{""�������"", type text}, {""������������"", type text}, {""��������� (���.)"", Int64.Type}})," & Chr(13) & "" & Chr(10) & _
'                "    #""������ � ����������� ��������"" = Table.SelectRows(#""���������� ���"", each [�������] <> null and [�������] <> """")" & Chr(13) & "" & Chr(10) & _
'                "in" & Chr(13) & "" & Chr(10) & _
'                "    #""������ � ����������� ��������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=ARP;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [ARP]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "ARP"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "Grand Mystery"
'
' Article   : GM
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_GM()

On Error GoTo ErrorHandl
t = Timer
        article = "GM "
        URL = "https://grandm.ru/personal/upload/price.xls"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="GM", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(Web.Contents(""" & URL & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(TDSheet1,8)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""������������.������� "", ""������������/ �������������� ������������"", ""�������""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each [#""������������.������� ""] <> null and [#""������������.������� ""] <> """")," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ��������"" = Table.ReplaceValue(#""������ � ����������� ��������"","" RUB"","""",Replacer.ReplaceText,{""�������""})," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ��������1"" = Table.ReplaceValue(#""���������� ��������"",""      "","""",Replacer.ReplaceText,{""������������/ �������������� ������������""})," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""���������� ��������1"",{{""������������.������� "", type text}, {""������������/ �������������� ������������"", type text}, {""�������"", Int64.Type}})" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=GM;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [GM]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "GM"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "Imlight"
'
' Article   : IM
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_IM()

On Error GoTo ErrorHandl
t = Timer
        article = "IM "
        URL = "https://price.imlight.ru/IMLIGHT_price.xls"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="IM", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(Web.Contents(""" & URL & """), null, true),    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(TDSheet1,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}, {""Column5"", type text}, {""Column6"", type text}, {""Column7"", type text}, {""Column8"", type text}, {""Column9"", type text}, {""Column10"", type text}, {""Column11"", type text}, {""Column12"", type text}, {""Column13"", type text}, {""Column14"", type text}, {""Column15"", type text}, {""Column16"", type text}})," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(#""���������� ���"",3)," & Chr(13) & "" & Chr(10) & "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""���"", ""������������"", ""������������ ��� ������"", ""�������������"", ""�����-���� (RUB)""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each [������������ ��� ������] <> null and [������������ ��� ������] <> """")," & Chr(13) & "" & Chr(10) & _
                "    #""�������� ���������������� ������"" = Table.AddColumn(#""������ � ����������� ��������"", ""������������"", each [�������������]&"" ""&[������������])," & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"" = Table.ReorderColumns(#""�������� ���������������� ������"",{""���"", ""������������"", ""������������"", ""������������ ��� ������"", ""�������������"", ""�����-���� (RUB)""})" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=IM;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [IM]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "IM"
                .Refresh BackgroundQuery:=False
            End With
        Columns("B:C").Select
        Selection.ColumnWidth = 30
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub

ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "Lutner"
'
' Article   : LU
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_LU()

On Error GoTo ErrorHandl
t = Timer
        article = "LU "
        URL = "https://lutner.ru/bitrix/catalog_export/upload/lutner_new.csv"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="lutner_new", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Csv.Document(Web.Contents(""https://lutner.ru/bitrix/catalog_export/upload/lutner_new.csv""),[Delimiter="";"", Columns=39, Encoding=1251, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(��������, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""IE_PREVIEW_TEXT"", ""IP_PROP140"", ""IP_PROP114"", ""IP_PROP96"", ""CV_PRICE_18""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each ([IP_PROP140] = """"))," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ��� � ������"" = Table.TransformColumnTypes(#""������ � ����������� ��������"", {{""CV_PRICE_18"", Currency.Type}}, ""en-US"")," & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"" = Table.RenameColumns(#""���������� ��� � ������"",{{""IP_PROP114"", ""�����""}, {""IP_PROP96"", ""�������""}, {""CV_PRICE_18"", ""�������""}})," & Chr(13) & "" & Chr(10) & _
                "    #""�������� ���������������� ������"" = Table.AddColumn(#""��������������� �������"", ""������������"", each [�����] & "" "" & [�������])," & Chr(13) & "" & Chr(10) & _
                "    #""�������� ���������������� ������1"" = Table.AddColumn(#""�������� ���������������� ������"", ""������ ������������"", each [�����] & "" "" & [IE_PREVIEW_TEXT])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������1"" = Table.SelectColumns(#""�������� ���������������� ������1"",{""�������"", ""�������"", ""������������"", ""������ ������������""})," & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"" = Table.ReorderColumns(#""������ ��������� �������1"",{""�������"", ""������������"", ""������ ������������"", ""�������""})" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""����������������� �������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=lutner_new;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [lutner_new]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "lutner_new"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub
    
ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "CTC"
'
' Article   : CTC
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_CTC()

On Error GoTo ErrorHandl
t = Timer
        article = "CTC"
        FileName = Request(2, 4)
'    DownloadFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\��������� ��� ����������\"
'            Prices = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\������\"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="�������", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(File.Contents(""" & DownloadFolder & FileName & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "   TDSheet1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "   #""��������� ������� ������"" = Table.Skip(TDSheet1,4)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������������� �������"" = Table.DuplicateColumn(#""���������� ���������"", ""������������"", ""����� ������������"")," & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"" = Table.RenameColumns(#""������������� �������"",{{""����� ������������"", ""������ ������������""}, {""����, �����"", ""�������""}})," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""��������������� �������"",{""�������"", ""������������"", ""������ ������������"", ""�������""})," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ ��������� �������"",{{""�������"", Currency.Type}})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""���������� ���"", each ([������������] <> null))" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=�������;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [�������]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "�������"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"
    
URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub
    
ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article & ""
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "Invask"
'
' Article   : IV
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_IV()

On Error GoTo ErrorHandl
t = Timer
        article = "IV "
        URL = "https://invask.ru/downloads/Ostatki_tovara.xls?v=1"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="Sheet1", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(Web.Contents(""" & URL & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    Sheet1 = ��������{[Name=""Sheet1""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(Sheet1,11)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""���"", ""�����"", ""������������"", ""�������""})," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""������ ��������� �������"", each [���] <> null and [���] <> """")" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Sheet1;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [Sheet1]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub
    
ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub


'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
'
' Procedure : Download Price "LTM-Music"
'
' Article   : LTM
'
'=================================================================================================================================================================================================
'=================================================================================================================================================================================================
Sub Download_LTM()

On Error GoTo ErrorHandl
t = Timer
        article = "LTM"
        URL = "https://ltm-music.ru/upload/excel/ostatki_retail.xls?1663149639"


    Workbooks.Add: Openbook = True
            ActiveWorkbook.Queries.Add Name:="TDSheet", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & _
                "    �������� = Excel.Workbook(Web.Contents(""" & URL & """), null, true)," & Chr(13) & "" & Chr(10) & _
                "    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������"" = Table.Skip(TDSheet1,5)," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���������"" = Table.PromoteHeaders(#""��������� ������� ������"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
                "    #""������ ��������� �������"" = Table.SelectColumns(#""���������� ���������"",{""��� ������"", ""�������� �����"", ""Column7"", ""���������""})," & Chr(13) & "" & Chr(10) & _
                "    #""��������������� �������"" = Table.RenameColumns(#""������ ��������� �������"",{{""Column7"", ""���""}})," & Chr(13) & "" & Chr(10) & _
                "    #""��������� ������� ������1"" = Table.Skip(#""��������������� �������"",4)," & Chr(13) & "" & Chr(10) & _
                "    #""������ � ����������� ��������"" = Table.SelectRows(#""��������� ������� ������1"", each [���] <> null and [���] <> """")," & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"" = Table.TransformColumnTypes(#""������ � ����������� ��������"",{{""���������"", Currency.Type}})" & Chr(13) & "" & Chr(10) & _
                "in" & Chr(13) & "" & Chr(10) & _
                "    #""���������� ���"""
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet;Extended Properties=""""" _
                , Destination:=Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [TDSheet]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "TDSheet"
                .Refresh BackgroundQuery:=False
            End With
    ActiveWorkbook.Close savechanges:=True, FileName:=Prices & Trim(article) & " " & Format(Date, "dd.mm.yyyy") & ".xlsx"

URL = "": URL2 = "": Login = "": Password = "": FileName = "": LoginCSS = "": PasswordCSS = "": ButtonCSS = "": Openbook = False: t = Format(Timer - t, "#.##")
Debug.Print "������� " & article & " �������� ������� (" & t & "�)"
Exit Sub
    
ErrorHandl:
        Debug.Print "������� " & article & " �� ��������. ���������� ������ ��� �������"
        If Openbook = True Then
            ActiveWorkbook.Close savechanges:=False
        End If
        If ErrorI = 0 Then
            ErrorMessage = article
            Else: ErrorMessage = ErrorMessage & ", " & article
        End If
        ErrorI = ErrorI + 1
End Sub



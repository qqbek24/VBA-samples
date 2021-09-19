Attribute VB_Name = "DodajWiersze"
Public NrZP, KomOrg As String
Public xPUSTY As String
Public NrSprawyX As Long
Public xNr, xNr2, xNr3, xNr4, xNr5, xNr6, xNr7, xNr8, xNr9, xNrMozna, xNrMoznaS, xNrMoznaN, NastepnyX, AktRowNext As Long
Public WybralNazwaZwyczaj As String
Public IdxZnalezioneVer, xNrSprawRealizacja As Single
Public CzyWstawicNr As Byte
Public TekstSearchGoogle As String
Public xCol As Integer
Public StaraNowa_Numeracja As Integer

Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long
'********************************************************************************************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''autor: Jakub Koziorowski''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Wstaw_Podswietlanie_AktywnegoWiersza()
    Dim Wbk1 As Workbook
    Dim Wks1 As Worksheet
    Dim a1 As String, a2 As String, Formul As String
    Dim b1 As Long
    On Error GoTo error
        Set Wbk1 = Application.ActiveWorkbook
        Set Wks1 = Wbk1.ActiveSheet
        Call DodajKod_DoArkusza(Wbk1, Wks1) 'dodawanie kodu do modulu arkusza
            a1 = Wbk1.Name: a2 = "False"
'dodawanie formatowania warunkowego
            Formul = "=WIERSZ(A" & ActiveCell.Row & ")=AktywnyWiersz"
            If Selection.FormatConditions.Count > 0 Then
                For i = 1 To Selection.FormatConditions.Count
                    If Selection.FormatConditions.Item(i).Formula1 = Formul Then
                        If ForXC = 1 Then
                            Selection.FormatConditions.Item(i).Delete
                        Else
                            GoTo Names
                        End If
                    End If
                Next
            End If
            Selection.FormatConditions.Add Type:=xlExpression, Formula1:=Formul
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                With Selection.FormatConditions(1)
                    .Interior.ColorIndex = 15
                    .StopIfTrue = True
                End With
Names:
'dodawanie nazwy do menedzera nazw
                    b1 = Wbk1.Names.Count
                    If b1 = 0 Then
                        Wbk1.Names.Add Name:="AktywnyWiersz", RefersToR1C1:="=0"
                        Wbk1.Names.Item Index:="AktywnyWiersz"
                    End If
                    For Each Item In Wbk1.Names
                        If Item.Name = "AktywnyWiersz" Then
                            a2 = "True"
                            Exit For
                        End If
                    Next Item
                    If a2 = "False" Then
                        Wbk1.Names.Add Name:="AktywnyWiersz", RefersTo:="" ',RefersToR1C1:="=0"
                        Wbk1.Names.Item Index:="AktywnyWiersz"
                    End If
                    Wbk1.VBProject.MakeCompiledFile
                    Application.Calculate
error:
End Sub
Sub Columns_View_Hide()
    Dim lastCol As Integer
        lastCol = ActiveCell.Column + 1
        Range(Cells(1, lastCol), Cells(Rows.Count, Columns.Count)).EntireColumn.Hidden = True
End Sub
Sub Rows_View_Hide()
    Dim LastRow As Integer
        LastRow = ActiveCell.Row + 1
        Range(Cells(LastRow, 1), Cells(Rows.Count, Columns.Count)).EntireRow.Hidden = True
End Sub
Sub Reveal_ColumnsandRows()
        Cells.EntireColumn.Hidden = False
        Cells.EntireRow.Hidden = False
End Sub
Sub DodajKod_DoArkusza(Wbk1 As Workbook, Wks1 As Worksheet)
    Dim ArkName As String
    On Error GoTo error
    ArkName = Wks1.CodeName
    Set CodePan = Wbk1.VBProject.VBComponents(ArkName).CodeModule
        s = _
            "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbNewLine & _
            "    ActiveWorkbook.Names(""AktywnyWiersz"").RefersTo = ""="" & ActiveCell.Row" & vbNewLine & _
            "End Sub"
            If CodePan.Find("Worksheet_SelectionChange", 1, 1, CodePan.CountOfLines, 50) = False Then
                With CodePan
                    .InsertLines .CountOfLines + 1, s
                End With
                Wbk1.VBProject.MakeCompiledFile
            End If
error:
End Sub
Sub Dodaj_Wiersze()
    Dim Wks As Worksheet
    Dim xRow, a As Long
    Dim xOdp As Variant
    Application.Calculation = xlCalculationManual
    Set Wks = ThisWorkbook.ActiveSheet
        xOdp = MsgBox("1. Czy zaznaczy≥aú/eú ca≥y wiersz ?" & Chr(13) & _
        "2. Wiersze zostanπ dodane powyøej zaznaczonego wiersza." & Chr(13) & _
        "3. Formatowanie dodanych wierszy, bÍdzie takie jak format wiersza" & Chr(13) & _
        "   powyøej tego, ktÛry zaznaczy≥aú/eú", vbOKCancel)
    Select Case xOdp
        Case vbOK
            GoTo Dodaj
        Case vbCancel
            Application.Calculation = xlCalculationAutomatic: Exit Sub
    End Select
Dodaj:
    a = 0
    xRow = Application.InputBox("Ile wierszy mam dodaÊ.", _
         "Dodawanie wierszy", , 250, 75, "", , 1)
    If xRow = False Then
        Application.Calculation = xlCalculationAutomatic: Exit Sub
    ElseIf xRow = 0 Then
        Application.Calculation = xlCalculationAutomatic: Exit Sub
    Else
        For a = 1 To xRow
            ActiveCell.EntireRow.Select
            Selection.Insert Shift:=xlDown, copyorigin:=xlFormatFromLeftOrAbove
        Next
    End If
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub add_Multiple_ROWS()
    Dim Select1 As Range
    Dim Komorka As Range
    Dim x1, x2, x3 As Long
    Dim xRow, a As Long
    Dim yRow, b As Long
    Dim xOdp As Variant
    Dim Wbk1 As Workbook
    Dim Wks1 As Worksheet
    Application.Calculation = xlCalculationManual
        Set Wbk1 = Application.ActiveWorkbook
        Set Wks1 = Wbk1.ActiveSheet
        Set Select1 = Selection
            xOdp = MsgBox("Czy napewno chcesz dodaÊ wiÍkszπ iloúÊ wierszy ???", vbOKCancel)
        Select Case xOdp
            Case vbOK
                GoTo Dodaj
            Case vbCancel
                Application.Calculation = xlCalculationAutomatic: Exit Sub
        End Select
Dodaj:
    a = 0
    xRow = Application.InputBox("Ile wierszy mam dodaÊ.", _
         "Dodawanie wierszy", , 250, 75, "", , 1)
    If xRow = False Then
        Application.Calculation = xlCalculationAutomatic: Exit Sub
    ElseIf xRow = 0 Then
        Application.Calculation = xlCalculationAutomatic: Exit Sub
    Else
        yRow = Application.InputBox("Co ktÛry wiersz mam dodaÊ.", _
         "Dodawanie wierszy", , 250, 75, "", , 1)
        x3 = 0
        x1 = Select1.Rows.Count
        If yRow <= 0 Or yRow = "" Then
            yRow = 1
        End If
        For x2 = x1 + Select1.Row To Select1.Row Step ("-" & yRow) '-1
            Set Komorka = Wks1.Cells(x2, 1)
            If x3 = x1 Then
                Application.Calculation = xlCalculationAutomatic: Exit Sub
            Else
                For a = 1 To xRow
                    Komorka.Rows.EntireRow.Insert Shift:=xlDown, copyorigin:=xlFormatFromLeftOrAbove
                Next
            End If
            x3 = x3 + 1
        Next
    End If
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub Dodaj_Kolumny()
    Dim Wks As Worksheet
    Dim aCol, xCol, a As Long
    Dim xOdp As Variant
    Application.Calculation = xlCalculationManual
    Set Wks = ThisWorkbook.ActiveSheet
        xOdp = MsgBox("1. Czy zaznaczy≥aú/eú ca≥π kolumnÍ ?" & Chr(13) & _
        "2. Kolumny zostanπ dodane przed zaznaczonπ kolumnπ." & Chr(13) & _
        "3. Formatowanie dodanych kolumn, bÍdzie takie jak format kolumny" & Chr(13) & _
        "   przed tπ, ktÛrπ zaznaczy≥aú/eú", vbOKCancel)
    Select Case xOdp
        Case vbOK
            GoTo Dodaj
        Case vbCancel
            Application.Calculation = xlCalculationAutomatic: Exit Sub
    End Select
Dodaj:
    a = 0
    xCol = Application.InputBox("Ile kolumn mam dodaÊ.", _
         "Dodawanie wierszy", , 250, 75, "", , 1)
    If xCol = False Then
        Application.Calculation = xlCalculationAutomatic: Exit Sub
    ElseIf xCol = 0 Then
        Application.Calculation = xlCalculationAutomatic: Exit Sub
    Else
        For a = 1 To xCol
            ActiveCell.EntireColumn.Select
            Selection.Insert Shift:=xlLeft, copyorigin:=xlFormatFromLeftOrAbove
        Next
    End If
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub CzyscArkusz()
    Dim Wbk As Workbook
    Dim Wks As Worksheet
        Set Wbk = Application.ActiveWorkbook
        Set Wks = Wbk.ActiveSheet
            Wks.Columns.ColumnWidth = 8.43
            Wks.UsedRange.Clear
            Wks.Cells.Font.Size = 11
            Wks.Columns.AutoFit
End Sub
Sub Idz_do_wiersza()
    Dim xRowNr As Long
    Dim WksLoop As Worksheet
    Dim WbkLoop As Workbook
        Set WbkLoop = Application.ActiveWorkbook
        Set WksLoop = WbkLoop.ActiveSheet
            xRowNr = Application.InputBox("Przejdü do wiersza nr: ", _
                 "Wiersz Loop", , 250, 75, "", , 1)
            If xRowNr = 0 Then
                Exit Sub
            ElseIf xRowNr = False Then
                Exit Sub
            Else
                WksLoop.Cells(xRowNr, 1).Select
            End If
End Sub
Sub IleArkuszy_w_pliku()
    Dim xArk As Long
    Dim WksLoop As Worksheet
    Dim WbkLoop As Workbook
        Set WbkLoop = Application.ActiveWorkbook
        Set WksLoop = WbkLoop.ActiveSheet
            xArk = WbkLoop.Worksheets.Count
            MsgBox "Ten plik zawiera " & xArk & " arkuszy"
End Sub
Function czyWkolekcji(igla As Variant, stog As Collection) As Boolean
Dim element As Variant
czyWkolekcji = False
For Each element In stog
    If igla = element Then czyWkolekcji = True
Next element
End Function
Sub szukajUForm_otw()
    SzukajRejestrZP.Show vbModeless
End Sub
Sub Szukaj_w_rejestrze_ZP()
    Dim WksLoop As Worksheet
    Dim WbkLoop As Workbook
    Dim xRow, yRow, LastRow, n, m, o As Long
    Dim ZakresSzuk, ZakresSzuk2 As Range
    Dim xObj As Object
        Set WbkLoop = Application.ActiveWorkbook
        Set WksLoop = WbkLoop.ActiveSheet
            LastRow = WksLoop.Cells(Rows.Count, 3).End(xlUp).Row
            Set ZakresSzuk = WksLoop.Range(WksLoop.Cells(47, 3), WksLoop.Cells(LastRow, 3))
            Set ZakresSzuk2 = WksLoop.Range(WksLoop.Cells(47, 5), WksLoop.Cells(LastRow, 5))
        NrZP = SzukajRejestrZP.TextNRzp.Value
        KomOrg = UCase(SzukajRejestrZP.TextKOMorg.Value)
        m = 0: n = 47: o = 0
        Set xObj = ZakresSzuk.Find(NrZP, After:=WksLoop.Cells(n, 3), lookat:=xlWhole, MatchCase:=False)
        If xObj Is Nothing Then
            MsgBox "nie znalaz≥em"
            Exit Sub
        Else
            m = ZakresSzuk.Find(NrZP, After:=WksLoop.Cells(n, 3), lookat:=xlWhole, MatchCase:=False).Row
        End If
        If NastepnyX = 1 Then n = AktRowNext
        Do
            If n = LastRow Then Exit Sub
            If o > 1 Then
                If n = m Then
                    MsgBox "nie znalaz≥em"
                    Exit Sub
                End If
            End If
            If IsError(ZakresSzuk.Find(NrZP, After:=WksLoop.Cells(n, 3), lookat:=xlWhole, MatchCase:=False).Row) = True Then
                MsgBox "nie znalaz≥em"
                Exit Sub
            Else
                o = o + 1
                xRow = ZakresSzuk.Find(NrZP, After:=WksLoop.Cells(n, 3), lookat:=xlWhole, MatchCase:=False).Row
            End If
            If KomOrg <> "" Then
                If WksLoop.Cells(xRow, 5) = KomOrg Then
                    WksLoop.Cells(xRow, 3).Activate
                    Exit Sub
                Else
                    n = xRow
                    GoTo NastRow
                End If
            ElseIf KomOrg = "" Then
                WksLoop.Cells(xRow, 3).Activate
                Exit Sub
            End If
NastRow:
        Loop Until n = LastRow
Application.ScreenUpdating = True
End Sub
Sub Szukaj_W_RejZP_Next()
    NastepnyX = 1
    If SzukajRejestrZP.TextNrSprawy.Value <> "" Then
        If ChBoxNr = 2 Or ChBoxNr = 0 Then
            AktRowNext = ActiveCell.Row
            If SzukajRejestrZP.CheckBox2.Value = "False" Then       'Stare nr spraw
                Call Szukaj_NrSprawy_w_RejestrzeZP
            ElseIf SzukajRejestrZP.CheckBox2.Value = "True" Then    'Nowe nr spraw
                Call Szukaj_NrSprawy_w_RejestrzeZP_Nowe
            End If
        ElseIf ChBoxNr = 3 Then
            AktRowNext = ActiveCell.Row
            Call Szukaj_NrZZ_w_RejestrzeZP
        End If
    Else
        AktRowNext = ActiveCell.Row + 1
        Call Szukaj_w_rejestrze_ZP
    End If
End Sub
Sub Szukaj_ZP_POdacie()
    Dim WksLoop As Worksheet
    Dim WbkLoop As Workbook
    Dim xRow, LastRow, n As Long
    Dim ZakresSzuk As Range
        Set WbkLoop = Application.ActiveWorkbook
        Set WksLoop = WbkLoop.ActiveSheet
            LastRow = WksLoop.Cells(Rows.Count, 2).End(xlUp).Row
            Set ZakresSzuk = WksLoop.Range(WksLoop.Cells(47, 2), WksLoop.Cells(LastRow, 2))
        n = 47
        SzukanaDATA = CDate(SzukajRejestrZP.HelpLabel.Caption)
        Do
            n = n + 1
            If n = LastRow Then Exit Sub
            If WksLoop.Cells(n, 2) = SzukanaDATA Then WksLoop.Cells(n, 1).Activate
        Loop Until n = LastRow
End Sub
Sub Szukaj_NrSprawy_w_RejestrzeZP()
    Dim WksLoop As Worksheet
    Dim WbkLoop As Workbook
    Dim xRow As Long, yRow As Long, LastRow As Long, n As Long, m As Long, o As Long ', xCol As Long
    Dim ZakresSzuk As Range
    Dim xObj As Object
    Dim NrSprawyXX As String
        Set WbkLoop = Application.ActiveWorkbook
        Set WksLoop = WbkLoop.ActiveSheet
            LastRow = WksLoop.Cells(Rows.Count, 3).End(xlUp).Row
            Set ZakresSzuk = WksLoop.Range(WksLoop.Cells(47, 13), WksLoop.Cells(LastRow, 13))
        NrSprawyXX = SzukajRejestrZP.TextNrSprawy.Value
        m = 0: n = 53: o = 0
        Set xObj = ZakresSzuk.Find(NrSprawyXX, After:=WksLoop.Cells(n, 13), lookat:=xlWhole, MatchCase:=False)
        If xObj Is Nothing Then
            MsgBox "nie znalaz≥em"
            Exit Sub
        Else
            m = ZakresSzuk.Find(NrSprawyXX, After:=WksLoop.Cells(n, 13), lookat:=xlWhole, MatchCase:=False).Row
        End If
        If NastepnyX = 1 Then n = AktRowNext
        Do
            If n = LastRow Then Exit Sub
            If o > 1 Then
                If n = m Then
                    MsgBox "nie znalaz≥em"
                    Exit Sub
                End If
            End If
            If IsError(ZakresSzuk.Find(NrSprawyXX, After:=WksLoop.Cells(n, 13), lookat:=xlWhole, MatchCase:=False).Row) = True Then
                MsgBox "nie znalaz≥em"
                Exit Sub
            Else
                o = o + 1
                xRow = ZakresSzuk.Find(NrSprawyXX, After:=WksLoop.Cells(n, 13), lookat:=xlWhole, MatchCase:=False).Row
                    WksLoop.Cells(xRow, 13).Activate
                    Exit Sub
            End If
        Loop Until n = LastRow
Application.ScreenUpdating = True
End Sub
Sub Szukaj_NrSprawy_w_RejestrzeZP_Nowe()
    Dim WksLoop As Worksheet
    Dim WbkLoop As Workbook
    Dim xRow As Long, yRow As Long, LastRow As Long, n As Long, m As Long, o As Long
    Dim ZakresSzuk As Range
    Dim xObj As Object
    Dim NrSprawyXX As String
        Set WbkLoop = Application.ActiveWorkbook
        Set WksLoop = WbkLoop.ActiveSheet
            LastRow = WksLoop.Cells(Rows.Count, 3).End(xlUp).Row
            Set ZakresSzuk = WksLoop.Range(WksLoop.Cells(47, 14), WksLoop.Cells(LastRow, 14))
        NrSprawyXX = SzukajRejestrZP.TextNrSprawy.Value
        m = 0: n = 53: o = 0
        Set xObj = ZakresSzuk.Find(NrSprawyXX, After:=WksLoop.Cells(n, 14), lookat:=xlWhole, MatchCase:=False)
        If xObj Is Nothing Then
            MsgBox "nie znalaz≥em"
            Exit Sub
        Else
            m = ZakresSzuk.Find(NrSprawyXX, After:=WksLoop.Cells(n, 14), lookat:=xlWhole, MatchCase:=False).Row
        End If
        If NastepnyX = 1 Then n = AktRowNext
        Do
            If n = LastRow Then Exit Sub
            If o > 1 Then
                If n = m Then
                    MsgBox "nie znalaz≥em"
                    Exit Sub
                End If
            End If
            If IsError(ZakresSzuk.Find(NrSprawyXX, After:=WksLoop.Cells(n, 14), lookat:=xlWhole, MatchCase:=False).Row) = True Then
                MsgBox "nie znalaz≥em"
                Exit Sub
            Else
                o = o + 1
                xRow = ZakresSzuk.Find(NrSprawyXX, After:=WksLoop.Cells(n, 14), lookat:=xlWhole, MatchCase:=False).Row
                    WksLoop.Cells(xRow, 14).Activate
                    Exit Sub
            End If
        Loop Until n = LastRow
Application.ScreenUpdating = True
End Sub
Sub Szukaj_NrZZ_w_RejestrzeZP()
    Dim WksLoop As Worksheet
    Dim WbkLoop As Workbook
    Dim xRow As Long, yRow As Long, LastRow As Long, n As Long, m As Long, o As Long, xCol As Long
    Dim ZakresSzuk As Range
    Dim xObj As Object
    Dim NrZZXX As String
        Set WbkLoop = Application.ActiveWorkbook
        Set WksLoop = WbkLoop.ActiveSheet
            If WksLoop.Name Like "*2016*" Then xCol = 19 Else xCol = 18
            LastRow = WksLoop.Cells(Rows.Count, 3).End(xlUp).Row
            Set ZakresSzuk = WksLoop.Range(WksLoop.Cells(47, xCol), WksLoop.Cells(LastRow, xCol))
        NrZZXX = SzukajRejestrZP.TextNrSprawy.Value
        m = 0: n = 53: o = 0
        Set xObj = ZakresSzuk.Find(NrZZXX, After:=WksLoop.Cells(n, xCol), lookat:=xlWhole, MatchCase:=False)
        If xObj Is Nothing Then
            MsgBox "nie znalaz≥em"
            Exit Sub
        Else
            m = ZakresSzuk.Find(NrZZXX, After:=WksLoop.Cells(n, xCol), lookat:=xlWhole, MatchCase:=False).Row
        End If
        If NastepnyX = 1 Then n = AktRowNext
        Do
            If n = LastRow Then Exit Sub
            If o > 1 Then
                If n = m Then
                    MsgBox "nie znalaz≥em"
                    Exit Sub
                End If
            End If
            If IsError(ZakresSzuk.Find(NrZZXX, After:=WksLoop.Cells(n, xCol), lookat:=xlWhole, MatchCase:=False).Row) = True Then
                MsgBox "nie znalaz≥em"
                Exit Sub
            Else
                o = o + 1
                xRow = ZakresSzuk.Find(NrZZXX, After:=WksLoop.Cells(n, xCol), lookat:=xlWhole, MatchCase:=False).Row
                    WksLoop.Cells(xRow, xCol).Activate
                    Exit Sub
            End If
        Loop Until n = LastRow
Application.ScreenUpdating = True
End Sub
Sub uzupelnijPuste2()
    Dim Komorka As Range
    Dim poprzedniaWartosc As Variant
    Dim cCol, xCol, FRow, LRow, FCol, LCol As Long
    Dim xLen, ChFind, xSel As Integer
    Dim Adres As String
    Dim rng, rng2 As Range
    For xSel = 1 To Selection.Areas.Count
        cCol = Selection.Areas(xSel).Columns.Count
        Adres = Replace(Selection.Areas(xSel).Address, "$", "")
        ChFind = InStr(1, Adres, ":")
        xLen = Len(Adres)
        Set rng = Application.Range(Adres)
            FCol = rng.Column
            LCol = FCol + cCol - 1
                FRow = rng.Row
                LRow = FRow + (rng.Rows.Count) - 1
            For xCol = FCol To LCol
                Set rng2 = ActiveSheet.Range(Cells(FRow, xCol), Cells(LRow, xCol))
                    poprzedniaWartosc = ""
                For Each Komorka In rng2
                    If Komorka.Value = "" Then Komorka.Formula = poprzedniaWartosc
                    poprzedniaWartosc = Komorka.Value
                Next
                Set rng2 = Nothing
            Next
    Next xSel
End Sub
Sub kolorujDuplikaty2()
Dim pamiec2 As New Collection
Dim Komorka As Range
For Each Komorka In Selection
    If czyWkolekcji(Komorka.Value, pamiec2) Then Komorka.Interior.Color = RGB(255, 255, 0)
    pamiec2.Add Komorka.Value
Next Komorka
End Sub
Sub AutoNumeracja_Zaznaczenie()
    Dim a, b As Long
    Dim Wkb As Workbook
    Dim Wks As Worksheet
    Dim Komorka As Range
        Set Wkb = Application.ActiveWorkbook
        Set Wks = Wkb.ActiveSheet
            a = Selection.Rows.Count
            b = 0
    For Each Komorka In Selection
        b = b + 1
        Komorka = b
    Next
    Selection.Cells.HorizontalAlignment = xlCenter
End Sub
Sub AutoNumeracja_Zaznaczenie_Widoczne()
    Dim a, b As Long
    Dim Wkb As Workbook
    Dim Wks As Worksheet
    Dim Komorka As Range
        Set Wkb = Application.ActiveWorkbook
        Set Wks = Wkb.ActiveSheet
            a = Selection.Rows.Count
            b = 0
    For Each Komorka In Selection
        If Komorka.Rows.Hidden = False Then
            If Komorka.Columns.Hidden = False Then
                b = b + 1
                Komorka = b
            End If
        End If
    Next
    Selection.Cells.HorizontalAlignment = xlCenter
End Sub
Sub Na_tekst()
    Dim xRowNr As Long
    Dim WksLoop As Worksheet
    Dim WbkLoop As Workbook
        Set WbkLoop = Application.ActiveWorkbook
        Set WksLoop = WbkLoop.ActiveSheet
            Selection.Copy
            Selection.PasteSpecial xlPasteValues
            'WbkLoop.Save
End Sub
Function Ostatni_Pelny_Wiersz_W_Kolumnie_dodatek(xCol As Integer, Optional NazwaArk As String) As Long
    Dim WbkA As Workbook
    Dim WksA As Worksheet
        Set WbkA = Application.ActiveWorkbook
        If NazwaArk = "" Then
            Set WksA = WbkA.ActiveSheet
        Else
            Set WksA = WbkA.Worksheets(NazwaArk)
        End If
        Ostatni_Pelny_Wiersz_W_Kolumnie_dodatek = WksA.Cells(Rows.Count, xCol).End(xlUp).Row
End Function
Function Nazwa_Aktywnego_Arkusza() As String
    Dim WbkAkt As Workbook
    Dim WksAkt As Worksheet
        Set WbkAkt = Application.ActiveWorkbook
        Set WksAkt = WbkAkt.ActiveSheet
    Nazwa_Aktywnego_Arkusza = WksAkt.Name
End Function
Sub Na_Duze_Litery()
    Dim xRowNr As Long
    Dim WksLoop As Worksheet
    Dim WbkLoop As Workbook
    Dim Zakr As Range
        Set WbkLoop = Application.ActiveWorkbook
        Set WksLoop = WbkLoop.ActiveSheet
            Set Zakr = Selection
            For Each Kom In Zakr
                Kom.Value = UCase(Kom)
            Next
End Sub
Sub Na_Male_Litery()
    Dim xRowNr As Long
    Dim WksLoop As Worksheet
    Dim WbkLoop As Workbook
    Dim Zakr As Range
        Set WbkLoop = Application.ActiveWorkbook
        Set WksLoop = WbkLoop.ActiveSheet
            Set Zakr = Selection
            For Each Kom In Zakr
                Kom.Value = LCase(Kom)
            Next
End Sub
Sub Stworz_Ark_Zaznaczenie()
    Dim Komorka As Range
    For Each Komorka In Selection
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = Komorka.Value
    Next
End Sub
Sub Arkusze_Zbieram()
    Dim WkbALARM As Workbook
    Dim WksALARM As Worksheet
    Dim WksZbior As Worksheet
    Dim ZakresX As Range
    Dim ZakresCop, ZakresPas As Range
    Dim NazwaArk As String
    Dim oRowALARM As Integer
    Dim oRowZbior As Long
        Set WkbALARM = Application.ActiveWorkbook
        Set WksALARM = WkbALARM.ActiveSheet
            Sheets.Add(Before:=Sheets(1)).Name = "Zbiorczy"
            Set WksZbior = WkbALARM.Worksheets("Zbiorczy")
            Set ZakresX = WksALARM.Range(WksALARM.Cells(1, 1), WksALARM.Cells(1, 10))
            ZakresX.Copy WksZbior.Range(WksZbior.Cells(1, 1), WksZbior.Cells(1, 10))
            Set ZakresX = Nothing
            WksZbior.Cells.WrapText = False
            WksZbior.Columns.AutoFit
            WksZbior.Rows.AutoFit
        For Each WksALARM In WkbALARM.Worksheets
            WksALARM.Activate
            NazwaArk = WksALARM.Name
            If WksALARM.Name = "Zbiorczy" Then GoTo Nast
            If WksALARM.Cells(1, 1) = "" Then GoTo Nast
            oRowALARM = WksALARM.Cells(Rows.Count, 1).End(xlUp).Row
            oRowZbior = WksZbior.Cells(Rows.Count, 1).End(xlUp).Row
                Set ZakresCop = WksALARM.Range(WksALARM.Cells(2, 1), WksALARM.Cells(oRowALARM, 10))
'                    ZakresCop.Copy
                Set ZakresPas = WksZbior.Range(WksZbior.Cells(oRowZbior + 1, 1), WksZbior.Cells(oRowZbior + oRowALARM, 10))
                    ZakresCop.Copy ZakresPas
                    WksZbior.Columns.AutoFit
                    WksZbior.Rows.AutoFit
                    Set ZakresX = WksZbior.Range(WksZbior.Cells(oRowZbior + 1, 9), WksZbior.Cells((oRowZbior + oRowALARM) - 1, 9))
                        ZakresX = NazwaArk
Nast:
        Next
            WksZbior.Cells.WrapText = False
            WksZbior.Columns.AutoFit
            WksZbior.Rows.AutoFit
End Sub
Sub Rozklad_Arkusza_FormatOut()
    Dim WkbALARM As Workbook
    Dim WksALARM As Worksheet
        Set WkbALARM = Application.ActiveWorkbook
        Set WksALARM = WkbALARM.ActiveSheet
        For Each WksALARM In WkbALARM.Worksheets
            WksALARM.Cells.WrapText = False
            WksALARM.Columns.AutoFit
            WksALARM.Rows.AutoFit
        Next
End Sub

Function Zamien_tekst(TekstPrzed As String) As String
    Dim TekstPo As String
    Dim a, b As Long
    Dim TblZnak As New Collection
    Dim a1, a2, a3 As String
        TblZnak.Add "•": TblZnak.Add "∆": TblZnak.Add " ": TblZnak.Add "£": TblZnak.Add "—"
        TblZnak.Add "”": TblZnak.Add "å": TblZnak.Add "è": TblZnak.Add "Ø": TblZnak.Add "."
        TblZnak.Add "/": TblZnak.Add "\": TblZnak.Add ";": TblZnak.Add ":": TblZnak.Add "'"
        TblZnak.Add """": TblZnak.Add "-": TblZnak.Add "+": TblZnak.Add "_": TblZnak.Add " "
        TblZnak.Add ",": TblZnak.Add "**": TblZnak.Add "***": TblZnak.Add "??": TblZnak.Add "*?"
            a1 = "?": a2 = "*": a3 = "*,*"
            TekstPrzed = UCase(TekstPrzed)
            TekstPo = "*" & TekstPrzed & "*"
                For a = 1 To 25
                    If a < 20 Then
                        TekstPo = Replace(TekstPo, TblZnak(a), a1)
                    ElseIf TblZnak(a) = "," Then
                        TekstPo = Replace(TekstPo, TblZnak(a), a1)
                    Else
                        TekstPo = Replace(TekstPo, TblZnak(a), a2)
                    End If
                Next
            TekstPo = LTrim(TekstPo)
            TekstPo = RTrim(TekstPo)
                Zamien_tekst = TekstPo
End Function
Function Miesiac_Slownie(KomorkaD As Range) As String
    Dim ms As Integer
    If IsDate(KomorkaD) = True Then
        If KomorkaD > 0 And KomorkaD < 13 Then
            'Miesiac_Slownie=KomorkaD
        Else
            ms = CInt(Month(KomorkaD.Value))
            If ms = 1 Then Miesiac_Slownie = "styczeÒ": Exit Function
            If ms = 2 Then Miesiac_Slownie = "luty": Exit Function
            If ms = 3 Then Miesiac_Slownie = "marzec": Exit Function
            If ms = 4 Then Miesiac_Slownie = "kwiecieÒ": Exit Function
            If ms = 5 Then Miesiac_Slownie = "maj": Exit Function
            If ms = 6 Then Miesiac_Slownie = "czerwiec": Exit Function
            If ms = 7 Then Miesiac_Slownie = "lipiec": Exit Function
            If ms = 8 Then Miesiac_Slownie = "sierpieÒ": Exit Function
            If ms = 9 Then Miesiac_Slownie = "wrzesieÒ": Exit Function
            If ms = 10 Then Miesiac_Slownie = "paüdziernik": Exit Function
            If ms = 11 Then Miesiac_Slownie = "listopad": Exit Function
            If ms = 12 Then Miesiac_Slownie = "grudzieÒ": Exit Function
        End If
    Else
        Miesiac_Slownie = "Wart. nie jest datπ?!"
    End If
End Function
Function Miesiac_Liczba_Na_Rzymskie(KomorkaD As Range) As String
    Dim ms As Integer
    If IsDate(KomorkaD) = True Then
        If KomorkaD > 0 And KomorkaD < 13 Then
            'Miesiac_Liczba_Na_Rzymskie=KomorkaD
        Else
            ms = CInt(Month(KomorkaD.Value))
            If ms = 1 Then Miesiac_Liczba_Na_Rzymskie = "I": Exit Function
            If ms = 2 Then Miesiac_Liczba_Na_Rzymskie = "II": Exit Function
            If ms = 3 Then Miesiac_Liczba_Na_Rzymskie = "III": Exit Function
            If ms = 4 Then Miesiac_Liczba_Na_Rzymskie = "IV": Exit Function
            If ms = 5 Then Miesiac_Liczba_Na_Rzymskie = "V": Exit Function
            If ms = 6 Then Miesiac_Liczba_Na_Rzymskie = "VI": Exit Function
            If ms = 7 Then Miesiac_Liczba_Na_Rzymskie = "VII": Exit Function
            If ms = 8 Then Miesiac_Liczba_Na_Rzymskie = "VIII": Exit Function
            If ms = 9 Then Miesiac_Liczba_Na_Rzymskie = "IX": Exit Function
            If ms = 10 Then Miesiac_Liczba_Na_Rzymskie = "X": Exit Function
            If ms = 11 Then Miesiac_Liczba_Na_Rzymskie = "XI": Exit Function
            If ms = 12 Then Miesiac_Liczba_Na_Rzymskie = "XII": Exit Function
        End If
    Else
        Miesiac_Liczba_Na_Rzymskie = "Wart. nie jest datπ?!"
    End If
End Function
Sub Usun_Puste_Linie_w_Komorce()
    Dim Komor As Range
    Dim Wkb1 As Workbook
    Dim Wks1 As Worksheet
    Dim Tekst1 As String
    Dim ile As Integer
        Set Wkb1 = Application.ActiveWorkbook
        Set Wks1 = Wkb1.ActiveSheet
            For Each Komor In Selection
                Tekst1 = Komor.Value
                If InStr(1, CStr(Komor.Value), Chr("010") & Chr("010") & Chr("010")) <> 0 Then
                    ile = Len(CStr(Komor.Value)) - Len(Replace(CStr(Komor.Value), Chr("010") & Chr("010") & Chr("010"), ""))
                        Komor.Value = Replace(CStr(Komor.Value), Chr("010") & Chr("010") & Chr("010"), Chr("010"))
                ElseIf InStr(1, CStr(Komor.Value), Chr("010") & Chr("010")) <> 0 Then
                    ile = Len(CStr(Komor.Value)) - Len(Replace(CStr(Komor.Value), Chr("010") & Chr("010"), ""))
                        Komor.Value = Replace(CStr(Komor.Value), Chr("010") & Chr("010"), Chr("010"))
                ElseIf InStr(1, CStr(Komor.Value), Chr("013") & Chr("010")) <> 0 Then
                        Komor.Value = Replace(CStr(Komor.Value), Chr("013") & Chr("010"), Chr("010"))
                ElseIf InStr(1, CStr(Komor.Value), Chr("010") & Chr("013")) <> 0 Then
                        Komor.Value = Replace(CStr(Komor.Value), Chr("010") & Chr("013"), Chr("010"))
                End If
            Next Komor
End Sub
Function Usun_Puste_Linie_w_tekscie(TekstX As String) As String
    Dim Tekst1 As String
    Dim ile As Integer
    While InStr(1, TekstX, Chr("010") & Chr("010")) <> 0
        If InStr(1, TekstX, Chr("010") & Chr("010") & Chr("010")) <> 0 Then
                TekstX = Replace(TekstX, Chr("010") & Chr("010") & Chr("010"), Chr("010"))
        ElseIf InStr(1, TekstX, Chr("010") & Chr("010")) <> 0 Then
                TekstX = Replace(CStr(TekstX), Chr("010") & Chr("010"), Chr("010"))
        End If
    Wend
    Usun_Puste_Linie_w_tekscie = TekstX
End Function
Sub UsunWykresy_wDodatku()
    Dim WkbWykres As Workbook
    Dim WksWykres As Worksheet
    Set WkbWykres = Application.Workbooks("Ribbon2.xlam")
    Set WksWykres = WkbWykres.Worksheets("WykresVBA")
            If WksWykres.ChartObjects.Count > 0 Then WksWykres.ChartObjects.Delete
End Sub
Public Sub OpenUrl()
    Dim lSuccess As Long
        TekstSearchGoogle = LTrim(TekstSearchGoogle)
        TekstSearchGoogle = RTrim(TekstSearchGoogle)
        TekstSearchGoogle = Replace(TekstSearchGoogle, " ", "+")
            lSuccess = ShellExecute(0, "Open", "https://www.google.pl/?gws_rd=ssl#q=" & TekstSearchGoogle)
End Sub
Function Rozbij_Tekst(CustomText As String) As Collection
    Dim ChrPoz As Long
    Dim CollectionStr As New Collection
    ChrPoz = InStr(1, CustomText, ","): CustomText = Replace(CustomText, " ", "")
    CustomText = LTrim(CustomText): CustomText = RTrim(CustomText) '<--
    While ChrPoz <> 0
        CollectionStr.Add Left(CustomText, ChrPoz - 1)
        CustomText = LTrim(CustomText): CustomText = RTrim(CustomText) '<--
        CustomText = Right(CustomText, (Len(CustomText) - ChrPoz))
        If InStr(1, CustomText, Chr("010")) <> 0 Then CustomText = Replace(CustomText, Chr("010"), "")
        ChrPoz = InStr(1, CustomText, ",")
    Wend
    CollectionStr.Add CustomText
    Set Rozbij_Tekst = CollectionStr
End Function
Function Licz_Komorki_wgKoloru(LiczZakres As Range, Optional KomorkaKolor As Range = Nothing) As Long
    Dim KomorkaLicz As Range
    On Error GoTo errExt
    Licz_Komorki_wgKoloru = 0
        For Each KomorkaLicz In LiczZakres
            If KomorkaKolor Is Nothing Then Set KomorkaKolor = ActiveCell
            If KomorkaLicz.Interior.ColorIndex = KomorkaKolor.Interior.ColorIndex Then Licz_Komorki_wgKoloru = Licz_Komorki_wgKoloru + 1
        Next
errExt:
End Function
Sub Grupowanie_wierszy()
    Dim xRow, LastRow, xCnt As Long
    Dim KomRng As Range
    Set SelectRng = Selection
    xCnt = 0: xRow = 0: LastRow = 0
    For Each KomRng In SelectRng
        If KomRng <> "" And xCnt <> 0 Then
            LastRow = KomRng.Row - 1
            Range(Cells(xRow, 1), Cells(LastRow, 1)).Rows.Group
            xRow = 0: LastRow = 0
        ElseIf KomRng = "" And xRow = 0 Then
            xRow = KomRng.Row
        End If
        xCnt = xCnt + 1
    Next
End Sub

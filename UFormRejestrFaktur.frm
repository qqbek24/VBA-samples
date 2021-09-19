VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFormRejestrFaktur 
   Caption         =   "Faktura - pozycje - ZZ - PZ"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   24930
   OleObjectBlob   =   "UFormRejestrFaktur.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFormRejestrFaktur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMetrics _
    Lib "user32.dll" (ByVal nIndex As Long) As Long


'********************************************************************************************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''autor: Jakub Koziorowski''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub CheckBoxMin_Change()
    If UFormRejestrFaktur.CheckBoxMin.Value = True Then
        UFormRejestrFaktur.Height = "75"
        UFormRejestrFaktur.Top = Application.Top + Application.Height - 40 - UFormRejestrFaktur.Height
    ElseIf UFormRejestrFaktur.CheckBoxMin.Value = False Then
        UFormRejestrFaktur.Height = "168"
        UFormRejestrFaktur.Top = Application.Top + Application.Height - 40 - UFormRejestrFaktur.Height
    End If
End Sub
Private Sub PodzielPozB_Click()
    Dim ab As Long, xQty As Long
    Dim CenaNetto As Double
    Dim xOdp As Variant
        If UFormRejestrFaktur.ListBoxPozycjeZZ.ListCount < 1 Then Exit Sub
        ab = UFormRejestrFaktur.ListBoxPozycjeZZ.ListIndex
        xOdp = MsgBox("Czy napewno chcesz podzieliÊ pozycjÍ ?", vbOKCancel)
        Select Case xOdp
            Case vbOK
                GoTo Dodaj
            Case vbCancel
                Exit Sub
        End Select
Dodaj:
        CenaNetto = CDbl(UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 5) / UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 3))
        xQty = Application.InputBox(("iloúÊ do podzielenia: " & UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 3)), _
             "Podziel", , 250, 75, "", , 1)
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 3) = UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 3) - xQty
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 5) = Format(CDbl(UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 3) * CenaNetto), "0.00")
        UFormRejestrFaktur.ListBoxPozycjeZZ.AddItem pvargItem:="", pvargIndex:=ab + 1
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab + 1, 0) = UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 0)
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab + 1, 1) = UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 1)
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab + 1, 2) = UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 2)
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab + 1, 3) = xQty
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab + 1, 4) = UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 4)
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab + 1, 5) = Format(CDbl(xQty * CenaNetto), "0.00")
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab + 1, 6) = UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 6)
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab + 1, 7) = UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 7)
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab + 1, 8) = UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 8)
            UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab + 1, 9) = UFormRejestrFaktur.ListBoxPozycjeZZ.List(ab, 9)
End Sub
Private Sub CzyscB_Click()
    UFormRejestrFaktur.TextBoxNrZZ.Value = ""
    UFormRejestrFaktur.TextBoxROK.Value = ""
    UFormRejestrFaktur.LabelDataFry.Caption = ""
    UFormRejestrFaktur.LabelNazwaDostawcy.Caption = ""
    UFormRejestrFaktur.LabelNrFry.Caption = ""
    UFormRejestrFaktur.LabelPozFry.Caption = ""
    UFormRejestrFaktur.LabelDostawcaAdres.Caption = ""
    UFormRejestrFaktur.LabelWartoscNETTO.Caption = ""
    UFormRejestrFaktur.LabelNRzz.Caption = ""
    UFormRejestrFaktur.LabelNRpz.Caption = ""
    UFormRejestrFaktur.LabelNrSprawy.Caption = ""
    UFormRejestrFaktur.LabelNrSprawy_nowa.Caption = ""
    UFormRejestrFaktur.TekstBranzysty.Caption = ""
    UFormRejestrFaktur.ListBoxPozycjeZZ.Clear
End Sub
Private Sub LabelNrFry_Click()
    Dim ab As Long
    Dim KtoZatwierdzal As String
    Dim UserNameFull As String
    Dim PurchIdxLng As Long
    Dim xRok As Integer
    On Error GoTo error
    If AxApl__ Is Nothing Then loginAX
    If AxApl__ Is Nothing Then GoTo error
    If UFormRejestrFaktur.LabelNrFry.Caption <> "" Then
        If UFormRejestrFaktur.TextBoxNrZZ = 0 Then
            Exit Sub
        ElseIf UFormRejestrFaktur.TextBoxROK = 0 Then
            Exit Sub
        End If
        PurchIdxLng = CLng(UFormRejestrFaktur.TextBoxNrZZ)
        xRok = UFormRejestrFaktur.TextBoxROK
        KtoZatwierdzal = Purch_StanZatwierdzen(PurchIdxLng, xRok)
        UserNameFull = UserId2UserFullName(KtoZatwierdzal)
        MsgBox "Dokument zosta≥ zatwierdzony przez: " & Chr(13) & Chr(13) & KtoZatwierdzal & " - " & UserNameFull
    End If
error:
    Exit Sub
End Sub
Private Sub LabelNRpz_Click()
    Dim xRok, xPZ As Long
    Dim xMag, xPZnr As String
    Dim xOdp As Variant
    On Error GoTo error
    If AxApl__ Is Nothing Then loginAX
    If AxApl__ Is Nothing Then GoTo error
    xPZ = Application.InputBox("Wpisz numer PZ (np. 4851):", _
         "Szukanie nr ZZ po nr PZ", , 250, 75, "", , 1)
    xRok = Application.InputBox("podaj rok (np. 14 lub 15):", _
         "Szukanie nr ZZ po nr PZ", , 250, 75, "", , 1)
    xMag = Application.InputBox("podaj Magazyn (np. 01 lub 10):", _
         "Szukanie nr ZZ po nr PZ", , 250, 75, "", , 2)
    If xPZ = False Then
        Exit Sub
    ElseIf xPZ = 0 Then
        Exit Sub
    Else
        xPZnr = NrPZ_Pelny(CLng(xPZ), CInt(xRok), CStr(xMag))
        UFormRejestrFaktur.LabelNRzz.Caption = Purch_ZZ_DlaPodanego_PZ(CLng(xPZ), CInt(xRok), CStr(xMag))
        UFormRejestrFaktur.LabelNRpz.Caption = xPZnr
    End If
error:
    Exit Sub
End Sub
Private Sub LabelNrSprawy_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    xNrSprawRealizacja = 1
    If UFormRejestrFaktur.CheckBoxNr.Value = "True" Then
        StaraNowa_Numeracja = 3
    Else
        StaraNowa_Numeracja = 1
    End If
    Call MaxNrSprawy_RejZP
    Select Case CzyWstawicNr
        Case 6
            UFormRejestrFaktur.LabelNrSprawy.Caption = xNrMoznaS
            If UFormRejestrFaktur.CheckBoxNr.Value = "True" Then
                If xNrMoznaN <> 0 Then
                    UFormRejestrFaktur.LabelNrSprawy_nowa.Caption = xNrMoznaN
                ElseIf xNrMoznaN = 0 Then
                    UFormRejestrFaktur.LabelNrSprawy_nowa.Caption = ""
                End If
            Else
                UFormRejestrFaktur.LabelNrSprawy_nowa.Caption = ""
            End If
        Case 7
            UFormRejestrFaktur.LabelNrSprawy.Caption = ""
            UFormRejestrFaktur.LabelNrSprawy_nowa.Caption = ""
            Exit Sub
    End Select
End Sub
Private Sub LabelNRzz_Click()
    Dim NrZZLng As Long
    Dim NrZZx, xRok As String
        If UFormRejestrFaktur.LabelNRzz.Caption <> "" Then
                NrZZx = UFormRejestrFaktur.LabelNRzz.Caption
                    NrZZLng = CLng(Mid(NrZZx, CLng(InStr(1, NrZZx, "Z/") + 2)))
                    xRok = CInt(Left(NrZZx, InStr(1, NrZZx, "/") - 1))
            UFormRejestrFaktur.TextBoxNrZZ.text = NrZZLng
            UFormRejestrFaktur.TextBoxROK.text = xRok
        End If
End Sub
Private Sub LabelTYTnrSpr_Click()
    xNrSprawRealizacja = 0
    UFormRejestrFaktur.LabelNrSprawy.Caption = ""
    UFormRejestrFaktur.LabelNrSprawy_nowa.Caption = ""
End Sub
Private Sub LabelTYTzz_Click()
    UFormRejestrFaktur.LabelNRzz.Caption = ""
    UFormRejestrFaktur.LabelNRpz.Caption = ""
End Sub
Private Sub ListBoxPozycjeZZ_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim WkbRejZP As Workbook
    Dim WksRejZP As Worksheet
    Dim AktKomorka As Range
    Dim aRow, xCol As Long
        If UFormRejestrFaktur.CheckBoxZablokowaneDodawanie.Value = True Then
            If UFormRejestrFaktur.CheckBoxCzyUsuwac.Value = True Then
                aListRow = UFormRejestrFaktur.ListBoxPozycjeZZ.ListIndex
                    If aListRow = -1 Then Exit Sub
                    If aListRow = 0 Then Exit Sub
                UFormRejestrFaktur.ListBoxPozycjeZZ.RemoveItem (aListRow)
            End If
            Exit Sub
        End If
            Set WkbRejZP = Application.ActiveWorkbook
            If WkbRejZP.Name Like "Rejestr ZP*" Then
                UFormRejestrFaktur.WstawB.Locked = True
                Set WksRejZP = WkbRejZP.ActiveSheet
                    Set AktKomorka = WksRejZP.Application.Selection
                    aRow = AktKomorka.Row
                    aListRow = UFormRejestrFaktur.ListBoxPozycjeZZ.ListIndex
                    If aListRow = -1 Then Exit Sub
                    If aListRow = 0 Then Exit Sub
                    WksRejZP.Cells(aRow, 9).Value = UFormRejestrFaktur.ListBoxPozycjeZZ.List(aListRow, 1)
                    If WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 24).Value = UFormRejestrFaktur.ListBoxPozycjeZZ.List(aListRow, 2)
                        If Not WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 23).Value = UFormRejestrFaktur.ListBoxPozycjeZZ.List(aListRow, 2)
                    If UFormRejestrFaktur.LabelNrSprawy.Caption <> "" Then
                        If UFormRejestrFaktur.LabelNrSprawy.Caption <> "sprawdzam..." Then
                            WksRejZP.Cells(aRow, 13).Value = UFormRejestrFaktur.LabelNrSprawy.Caption
                            If UFormRejestrFaktur.CheckBoxNr.Value = "True" Then If WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 14).Value = UFormRejestrFaktur.LabelNrSprawy_nowa.Caption 'Nr sprawy(2 numery sprawy)
                    End If
                    If WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 15).Value = UFormRejestrFaktur.ListBoxPozycjeZZ.List(aListRow, 3)
                        If Not WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 14).Value = UFormRejestrFaktur.ListBoxPozycjeZZ.List(aListRow, 3)
                    If WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 16).Value = UFormRejestrFaktur.ListBoxPozycjeZZ.List(aListRow, 4)
                        If Not WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 15).Value = UFormRejestrFaktur.ListBoxPozycjeZZ.List(aListRow, 4)
                    If WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 17).Value = CDbl(UFormRejestrFaktur.ListBoxPozycjeZZ.List(aListRow, 5))
                        If Not WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 16).Value = CDbl(UFormRejestrFaktur.ListBoxPozycjeZZ.List(aListRow, 5))
                    If WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 19).Value = UFormRejestrFaktur.TextBoxNrZZ.text '
                        If Not WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 18).Value = UFormRejestrFaktur.TextBoxNrZZ.text '
                    If WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 22).Value = UFormRejestrFaktur.LabelDostawcaAdres '
                        If Not WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 21).Value = UFormRejestrFaktur.LabelDostawcaAdres
                    If WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 20).Value = UFormRejestrFaktur.LabelDataFry
                        If Not WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 19).Value = UFormRejestrFaktur.LabelDataFry
                    If WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 21).Value = UCase(Format((WksRejZP.Cells(aRow, 20).Value), "mmmm"))
                        If Not WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 20).Value = UCase(Format((WksRejZP.Cells(aRow, 19).Value), "mmmm")) '
                    Call data_Format
                    If WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 25).Value = Now(): WksRejZP.Cells(aRow, 29).Value = Environ("USERNAME")
                        If Not WksRejZP.Name Like "2016" Then WksRejZP.Cells(aRow, 24).Value = Now(): WksRejZP.Cells(aRow, 28).Value = Environ("USERNAME")
                    If UFormRejestrFaktur.CheckBoxCzyUsuwac.Value = True Then UFormRejestrFaktur.ListBoxPozycjeZZ.RemoveItem (aListRow)
            End If
                UFormRejestrFaktur.WstawB.Locked = False
End Sub
Private Sub SzukajAxB_Click()
    Dim NrZZx, NrZZxPelny As String
    Dim xPoz, xPozLista As Integer
    Dim WartNetto, SumWartNetto As Double
    Dim AxQuery As AxaptaCOMConnector.IAxaptaObject
    Dim AxQueryDataSource As AxaptaCOMConnector.IAxaptaObject
    Dim AxQueryRun As AxaptaCOMConnector.IAxaptaObject
    Dim AxQueryRange As AxaptaCOMConnector.IAxaptaObject
    Dim RecorD As AxaptaCOMConnector.IAxaptaRecord
    On Error GoTo error
    If AxApl__ Is Nothing Then loginAX
    If AxApl__ Is Nothing Then GoTo error
    Application.Volatile
    UFormRejestrFaktur.ListBoxPozycjeZZ.Clear
    If UFormRejestrFaktur.TextBoxNrZZ.Value <> "" Then
        If Len(UFormRejestrFaktur.TextBoxROK.Value) > 2 Then
            Exit Sub
        ElseIf UFormRejestrFaktur.TextBoxROK.Value = "" Then
            Exit Sub
        End If
        NrZZx = UFormRejestrFaktur.TextBoxNrZZ.Value
        If (UFormRejestrFaktur.TextBoxROK.Value) <> "" Then
            NrZZxPelny = NrZZ_Pelny(CLng(UFormRejestrFaktur.TextBoxNrZZ.Value), UFormRejestrFaktur.TextBoxROK.Value)
        Else
            NrZZxPelny = NrZZ_Pelny(CLng(UFormRozdzielnik.TextBoxNrZZ.Value))
        End If
        UFormRejestrFaktur.LabelNRzz.Caption = NrZZxPelny
        UFormRejestrFaktur.LabelNRpz.Caption = Purch_PZ_DlaPodanego_ZZ(CStr(NrZZx), UFormRejestrFaktur.TextBoxROK.Value)
        UFormRejestrFaktur.LabelDataFry.Caption = Purch_Dok_Dostawy_Data(CStr(NrZZx), UFormRejestrFaktur.TextBoxROK.Value)
        UFormRejestrFaktur.LabelDostawcaAdres.Caption = Usun_Puste_Linie_w_tekscie(PurchFirma_Nazwa_Adress(CStr(NrZZx), UFormRejestrFaktur.TextBoxROK.Value))
        UFormRejestrFaktur.LabelNazwaDostawcy.Caption = PurchFirma_Nazwa(CStr(NrZZx), UFormRejestrFaktur.TextBoxROK.Value)
        UFormRejestrFaktur.LabelNrFry.Caption = Purch_Dok_Dostawy(CStr(NrZZx), UFormRejestrFaktur.TextBoxROK.Value)
        UFormRejestrFaktur.LabelPozFry.Caption = Purch_Ile_Pozycji_na_ZZ(CStr(NrZZx), UFormRejestrFaktur.TextBoxROK.text)
        UFormRejestrFaktur.TekstBranzysty.Caption = Purch_TekstBranzysty(CStr(NrZZx), UFormRejestrFaktur.TextBoxROK.text)
            xPoz = 1: xPozLista = 0
            UFormRejestrFaktur.ListBoxPozycjeZZ.ColumnWidths = "30;70;300;30;50;50;50;120;50;200"
                With UFormRejestrFaktur.ListBoxPozycjeZZ
                    .AddItem "[Lp]"
                    .List(xPozLista, 1) = "[Indeks]"
                    .List(xPozLista, 2) = "[Nazwa towaru]"
                    .List(xPozLista, 3) = "[ilosc]"
                    .List(xPozLista, 4) = "[Jm]"
                    .List(xPozLista, 5) = "[Wartosc]"
                    .List(xPozLista, 6) = "[Stan il.]"
                    .List(xPozLista, 7) = "[Status]"
                    .List(xPozLista, 8) = "[kom]"
                    .List(xPozLista, 9) = "[Gr.mat-nazwa]"
                End With
                xPozLista = xPozLista + 1
            While xPoz <= CInt(UFormRejestrFaktur.LabelPozFry.Caption)
                If AxApl__ Is Nothing Then loginAX
                Set AxQuery = AxApl__.CreateObject("Query")
                Set AxQueryDataSource = AxQuery.Call("addDataSource", 340)
                    Set AxQueryRange = AxQueryDataSource.Call("addRange", 1)
                        AxQueryRange.Call "value", NrZZxPelny
                        Set AxQueryRun = AxApl__.CreateObject("QueryRun", AxQuery)
                        While AxQueryRun.Call("Next")
                            Set RecorD = AxQueryRun.Call("GetNo", 1)
                        
                            WartNetto = Format((RecorD.field("LineAmount")), "0.00")
                            SumWartNetto = SumWartNetto + WartNetto
                            With UFormRejestrFaktur.ListBoxPozycjeZZ
                                .AddItem xPoz & "."
                                .List(xPozLista, 1) = RecorD.field("itemId")
                                .List(xPozLista, 2) = NazwaTowaru(UFormRejestrFaktur.ListBoxPozycjeZZ.List(xPozLista, 1))
                                .List(xPozLista, 3) = CDbl(RecorD.field("QtyOrdered"))
                                .List(xPozLista, 4) = JednMiary(UFormRejestrFaktur.ListBoxPozycjeZZ.List(xPozLista, 1))
                                .List(xPozLista, 5) = WartNetto
                                .List(xPozLista, 6) = StanNaDzienIndeks(UFormRejestrFaktur.ListBoxPozycjeZZ.List(xPozLista, 1), Now())
                                .List(xPozLista, 7) = Format(CDbl(WartNetto / RecorD.field("QtyOrdered")), "0.00") & " z≥ - " & Purch_Status_ZZ(CStr(NrZZx), UFormRejestrFaktur.TextBoxROK.text)
                                .List(xPozLista, 8) = UCase(RecorD.field(30007))
                                .List(xPozLista, 9) = GrMaterial(UFormRejestrFaktur.ListBoxPozycjeZZ.List(xPozLista, 1)) & " - " & NazwaGrMaterial(GrMaterial(UFormRejestrFaktur.ListBoxPozycjeZZ.List(xPozLista, 1)))
                            End With
                            xPoz = xPoz + 1: xPozLista = xPozLista + 1
                        Wend
            Wend
            UFormRejestrFaktur.LabelWartoscNETTO.Caption = SumWartNetto
            If Application.ActiveWorkbook.Name Like "*Rejestr*ZP*" Then
                If Right(ActiveSheet.Name, 2) <> UFormRejestrFaktur.TextBoxROK.Value Then Exit Sub  'tymczasowo
                    Dim WksZP As Worksheet
                    Dim ZakrSzukZZ As Range
                    Dim xObj As Object
                    Dim LastR As Long
                    Dim WynikSzukZZ, xRow, xCol As Long
                    Dim WynikTF As Boolean
                    Dim SzukZZ As Long
                    xCol = 1
                    xRow = 1
                    Set WksZP = ActiveSheet
                    For xRow = 1 To WksZP.Cells(Rows.Count, 3).End(xlUp).Row
                        If WksZP.Cells(xRow, 1).Value Like "*Lp*" Then xRow = WksZP.Cells(xRow, 19).End(xlDown).Row: Exit For
                    Next
                            LastR = WksZP.Cells(Rows.Count, 3).End(xlUp).Row
                        If Not IsDate(WksZP.Cells(xRow, 19)) Then xCol = 19
                        If IsDate(WksZP.Cells(xRow, 19)) Then xCol = 18
                        Set ZakrSzukZZ = WksZP.Range(WksZP.Cells(50, xCol), WksZP.Cells(LastR, xCol))
                        SzukZZ = UFormRejestrFaktur.TextBoxNrZZ.Value
                        Set xObj = ZakrSzukZZ.Find(SzukZZ, After:=WksZP.Cells(50, xCol), lookat:=xlWhole, MatchCase:=False)
                    If xObj Is Nothing Then
                        GoTo Pomin
                    Else
NastRow:
                        WynikSzukZZ = ZakrSzukZZ.Find(SzukZZ, After:=WksZP.Cells(xRow, xCol), lookat:=xlWhole, MatchCase:=False).Row
                        If WynikSzukZZ <> 0 Then
                            If UFormRejestrFaktur.TextBoxROK.Value = Right(Year(WksZP.Cells(WynikSzukZZ, xCol + 1).Value), 2) Then
                                MsgBox "Wpisany nr ZZ, jest juø w rejestrze, sprawdü czy nie powielasz wpisu"
                            Else
                                xRow = WynikSzukZZ: GoTo NastRow
                            End If
                        End If
                    End If
            End If
Pomin:
    Else
        Exit Sub
    End If
error:
    Exit Sub
End Sub
Private Sub UserForm_Activate()
        Dim ABC, ABC2, CBA, CBA2 As Variant
        With Me
            UFormRejestrFaktur.Move ((.InsideWidth - UFormRejestrFaktur.Width) / 2), ((.InsideHeight - UFormRejestrFaktur.Height) / 2)
            If UFormRejestrFaktur.Width < 1251 Then .ScrollBars = fmScrollBarsHorizontal Else .ScrollBars = fmScrollBarsNone
            .ScrollHeight = .InsideHeight * 2
            .ScrollWidth = .InsideWidth * 1.5
            .Height = "168"
        End With
    If Windows.Application.WindowState = xlMaximized Then
        UFormRejestrFaktur.Width = Application.Width - 70
        UFormRejestrFaktur.StartUpPosition = 0
        UFormRejestrFaktur.Top = Application.Top + Application.Height - 40 - UFormRejestrFaktur.Height
        UFormRejestrFaktur.Left = Application.Left + Application.Width - 30 - UFormRejestrFaktur.Width
        ThisWorkbook.Activate
    ElseIf Windows.Application.WindowState = xlNormal Then
        ABC = GetSystemMetrics(0)
        CBA = GetSystemMetrics(1)
        ABC2 = (ABC / 4)
        CBA2 = (CBA / 4)
        UFormRejestrFaktur.Width = ABC - ABC2 - 16
        UFormRejestrFaktur.StartUpPosition = 0
        UFormRejestrFaktur.Top = CBA - CBA2 - 190
        UFormRejestrFaktur.Left = 10
        ThisWorkbook.Activate
    Else
        ABC = GetSystemMetrics(0)
        CBA = GetSystemMetrics(1)
        ABC2 = (ABC / 4)
        CBA2 = (CBA / 4)
        UFormRejestrFaktur.Width = ABC - ABC2 - 16
        UFormRejestrFaktur.StartUpPosition = 0
        UFormRejestrFaktur.Top = CBA - CBA2 - 190
        UFormRejestrFaktur.Left = 10
        ThisWorkbook.Activate
    End If
End Sub
Private Sub UserForm_Initialize()
    Dim WkbRejZP As Workbook
    Dim WksRejZP As Worksheet
        Set WkbRejZP = Application.ActiveWorkbook
        Set WksRejZP = WkbRejZP.ActiveSheet
        If WksRejZP.Name = "2016" Then UFormRejestrFaktur.CheckBoxNr.Visible = True: UFormRejestrFaktur.LabelTYTNrSpr_nowa.Visible = True: UFormRejestrFaktur.LabelNrSprawy_nowa.Visible = True Else UFormRejestrFaktur.CheckBoxNr.Visible = False: UFormRejestrFaktur.LabelTYTNrSpr_nowa.Visible = False: UFormRejestrFaktur.LabelNrSprawy_nowa.Visible = False
        If UCase(WkbRejZP.Name) Like UCase("*Rejestr ZP*") Then
            UFormRejestrFaktur.WstawB.Visible = False
            UFormRejestrFaktur.WstawB.Locked = True
        ElseIf UCase(WkbRejZP.Name) Like UCase("*Rejestr Faktur*") Then
            UFormRejestrFaktur.WstawB.Visible = True
            UFormRejestrFaktur.WstawB.Locked = False
        Else
            UFormRejestrFaktur.WstawB.Visible = False
            UFormRejestrFaktur.WstawB.Locked = True
        End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Application.StatusBar = ""
End Sub
Private Sub WstawB_Click()
    Dim WkbRej As Workbook
    Dim WksRej As Worksheet
        Set WkbRej = Application.ActiveWorkbook
        Set WksRej = WkbRej.ActiveSheet
        If WksRej.Cells(ActiveCell.Row, 5) <> "" Then
            Exit Sub
        ElseIf WksRej.Cells(ActiveCell.Row, 6) <> "" Then
            Exit Sub
        ElseIf WksRej.Cells(ActiveCell.Row, 4) <> "" Then
            Exit Sub
        ElseIf WksRej.Cells(ActiveCell.Row, 13) <> "" Then
            Exit Sub
        Else
            WksRej.Cells(ActiveCell.Row, 4) = UFormRejestrFaktur.LabelNrFry.Caption
            WksRej.Cells(ActiveCell.Row, 5) = UFormRejestrFaktur.LabelDataFry.Caption
            WksRej.Cells(ActiveCell.Row, 6) = UFormRejestrFaktur.LabelPozFry.Caption
            WksRej.Cells(ActiveCell.Row, 7) = UCase(UFormRejestrFaktur.LabelNazwaDostawcy.Caption)
            WksRej.Cells(ActiveCell.Row, 8) = CDbl(UFormRejestrFaktur.LabelWartoscNETTO.Caption) ', "0.00")
            WksRej.Cells(ActiveCell.Row, 11) = UFormRejestrFaktur.TextBoxNrZZ.Value
        End If
End Sub

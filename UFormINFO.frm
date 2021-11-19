VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFormINFO 
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   24150
   OleObjectBlob   =   "UFormINFO.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFormINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IdxStrPub As String, NazwaStrPub As String, GrMatStrPub As String, DostawStrPub As String, KomOrgStrPub As String, DataUtwStrPub As String
Public IlePobrane As Long

'********************************************************************************************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''autor: Jakub Koziorowski''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ButtonCzysc_Click()
    IdxStrPub = "": NazwaStrPub = "": GrMatStrPub = "": DostawStrPub = "": KomOrgStrPub = "": DataUtwStrPub = ""
    UFormINFO.TextBoxIndeks.Value = "indeks"
    UFormINFO.TextBoxNazwaTowaru.Value = "Nazwa towaru"
    UFormINFO.TextBoxGrupaMaterialowa.Value = "Gr. mat."
    UFormINFO.TextBoxDostawca.Value = "Dostawca"
    UFormINFO.TextBoxKomOrgFiltr.Value = "Komorka Org."
    UFormINFO.TextBoxDataUtworzenia.Value = "Data utworzenia"
    UFormINFO.TextBoxSzukGoogle.Value = ""
    UFormINFO.TextBoxKomOrg.Value = ""
    UFormINFO.TextBoxZakresBezIdx.text = ""
    UFormINFO.TextBoxZakresBezIdx2.text = ""
    UFormINFO.ListBoxSzukajIndeks.Clear
    UFormINFO.TextBoxIndeks.Enabled = True
    UFormINFO.TextBoxNazwaTowaru.Enabled = True
    UFormINFO.TextBoxGrupaMaterialowa.Enabled = True
    UFormINFO.TextBoxDataUtworzenia.Enabled = True
    UFormINFO.TextBoxDostawca.Enabled = True
    UFormINFO.TextBoxKomOrgFiltr.Enabled = True
    UFormINFO.ButtonIdxSzczegoly.Top = UFormINFO.ListBoxSzukajIndeks.Top
    UFormINFO.FrameProgress.Visible = False
    UFormINFO.LabelStatusBar.Visible = True
    UFormINFO.ButtonIdxSzczegoly.Visible = False
    UFormINFO.ComboBoxWyborOpcji.Value = "Ca³a kartoteka"
    UFormINFO.MultiPage2.Pages("PageZuzycie").Caption = "Zu¿ycie w latach"
    UFormINFO.LabelStatusBar.Caption = "Autor dodatku: Jakub Koziorowski ; "
    UFormINFO.LabelKomOrgPobierTYT.Caption = "Komorki org. pobierajace mat."
    UFormINFO.Label1.Caption = "Data ostatniej transakcji"
    UFormINFO.LabelDataPrzychTYT.Caption = "Data PZ"
    UFormINFO.LabelDataRozchTYT.Caption = "Data RW"
    UFormINFO.LabelOgolTYT.Caption = "Data ostatnia ogó³."
    VerSzukajIdx = 0: UFormNr = 0
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ButtonIdxSzczegoly_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    ButtonIdxSzczegoly.BackColor = &HC000&
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ComboBoxWyborOpcji_Change()
    If UFormINFO.ComboBoxWyborOpcji.MatchFound = True Then
        UFormINFO.ListBoxSzukajIndeks.Clear
        If UFormINFO.ComboBoxWyborOpcji.Value = "Stany Awaryjne" Then
            IdxStrPub = UFormINFO.TextBoxIndeks.Value: NazwaStrPub = UFormINFO.TextBoxNazwaTowaru.Value: GrMatStrPub = UFormINFO.TextBoxGrupaMaterialowa.Value
            DostawStrPub = UFormINFO.TextBoxDostawca.Value:   KomOrgStrPub = UFormINFO.TextBoxKomOrgFiltr.Value: DataUtwStrPub = UFormINFO.TextBoxDataUtworzenia.Value
                UFormINFO.LabelJM.Visible = False
                UFormINFO.LabelKomOrgPobierTYT.Visible = False
                    UFormINFO.TextBoxIndeks.Enabled = False
                    UFormINFO.TextBoxNazwaTowaru.Enabled = False
                    UFormINFO.TextBoxGrupaMaterialowa.Enabled = False
                    UFormINFO.TextBoxDostawca.Enabled = False
                    UFormINFO.TextBoxKomOrgFiltr.Enabled = False
                    UFormINFO.TextBoxDataUtworzenia.Enabled = False
                        UFormINFO.TextBoxIndeks.Value = "indeks"
                        UFormINFO.TextBoxNazwaTowaru.Value = "Nazwa towaru"
                        UFormINFO.TextBoxGrupaMaterialowa.Value = "Gr. mat."
                        UFormINFO.TextBoxDostawca.Value = "Dostawca"
                        UFormINFO.TextBoxKomOrgFiltr.Value = "Komorka Org."
                        UFormINFO.TextBoxDataUtworzenia.Value = "Data utworzenia"
                        UFormINFO.TextBoxZakresBezIdx.text = ""
                        UFormINFO.TextBoxZakresBezIdx2.text = ""
        Else
            UFormINFO.LabelJM.Visible = True
            UFormINFO.LabelKomOrgPobierTYT.Visible = True
                UFormINFO.TextBoxIndeks.Enabled = True
                UFormINFO.TextBoxNazwaTowaru.Enabled = True
                UFormINFO.TextBoxGrupaMaterialowa.Enabled = True
                UFormINFO.TextBoxDostawca.Enabled = True
                UFormINFO.TextBoxKomOrgFiltr.Enabled = True
                UFormINFO.TextBoxDataUtworzenia.Enabled = True
                    If IdxStrPub = "" Then UFormINFO.TextBoxIndeks.Value = "indeks" Else UFormINFO.TextBoxIndeks.Value = IdxStrPub
                    If NazwaStrPub = "" Then UFormINFO.TextBoxNazwaTowaru.Value = "Nazwa towaru" Else UFormINFO.TextBoxNazwaTowaru.Value = NazwaStrPub
                    If GrMatStrPub = "" Then UFormINFO.TextBoxGrupaMaterialowa.Value = "Gr. mat." Else UFormINFO.TextBoxGrupaMaterialowa.Value = GrMatStrPub
                    If DostawStrPub = "" Then UFormINFO.TextBoxDostawca.Value = "Dostawca" Else UFormINFO.TextBoxDostawca.Value = DostawStrPub
                    If KomOrgStrPub = "" Then UFormINFO.TextBoxKomOrgFiltr.Value = "Komorka Org." Else UFormINFO.TextBoxKomOrgFiltr.Value = KomOrgStrPub
                    If DataUtwStrPub = "" Then UFormINFO.TextBoxDataUtworzenia.Value = "Data utworzenia" Else UFormINFO.TextBoxDataUtworzenia.Value = DataUtwStrPub
        End If
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If UFormINFO.Label1.Caption = "Data ostatniej transakcji" Then
        UFormINFO.Label1.Caption = "Zuzycie"
        UFormINFO.LabelDataPrzychTYT.Caption = (Year(CDate(Now()))) - 2
        UFormINFO.LabelDataRozchTYT.Caption = (Year(CDate(Now()))) - 1
        UFormINFO.LabelOgolTYT.Caption = Year(CDate(Now()))
    Else
        UFormINFO.Label1.Caption = "Data ostatniej transakcji"
        UFormINFO.LabelDataPrzychTYT.Caption = "Data PZ"
        UFormINFO.LabelDataRozchTYT.Caption = "Data RW"
        UFormINFO.LabelOgolTYT.Caption = "Data ostatnia ogó³."
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LabelKomOrgPobierTYT_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If UFormINFO.LabelKomOrgPobierTYT.Caption = "Komorki org. pobierajace mat." Then UFormINFO.LabelKomOrgPobierTYT.Caption = "Ostatni zakup - firma" Else UFormINFO.LabelKomOrgPobierTYT.Caption = "Komorki org. pobierajace mat."
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    UFormINFO.ButtonIdxSzczegoly.BackColor = &HFFC0C0
End Sub
               
'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ButtonPokazTransIdx_Click()
    Dim IdxPozList As Long
        IdxSzukF = ""
        If FormTransIdxList.Visible = True Then Unload FormTransIdxList
        If (UFormINFO.ListBoxSzukajIndeks.ListIndex) >= 0 Then
            IdxPozList = UFormINFO.ListBoxSzukajIndeks.ListIndex
            IdxSzukF = UFormINFO.ListBoxSzukajIndeks.List(IdxPozList, 1)
            VerSzukajIdx = 1
                FormTransIdxList.Show vbModeless
        End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ButtonPominIdx_Click()
    PominVer = 1: UFormNr = 2
    UFormRefEditPominIdx.Show
End Sub
                  
'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ButtonPominIdx2_Click()
    PominVer = 2: UFormNr = 2
    UFormRefEditPominIdx.Show
End Sub
''===========================================================================================================================================
Private Sub ButtonSzukaj_Click()
    Dim AxQuery As AxaptaCOMConnector.IAxaptaObject
    Dim AxDataSource As AxaptaCOMConnector.IAxaptaObject
    Dim AxRange1 As AxaptaCOMConnector.IAxaptaObject
    Dim AxRange2 As AxaptaCOMConnector.IAxaptaObject
    Dim AxRange3 As AxaptaCOMConnector.IAxaptaObject
    Dim AxRange4 As AxaptaCOMConnector.IAxaptaObject
    Dim AxRange5 As AxaptaCOMConnector.IAxaptaObject
    Dim AxQueryRun As AxaptaCOMConnector.IAxaptaObject
    Dim AxRecord As AxaptaCOMConnector.IAxaptaRecord
    Dim LpIdx, xPozList As Long
    Dim Tabela, SumTr, a As Long
    Dim AxPole1, AxPole2, AxPole3, AxPole4, AxPole5 As Long
    Dim AxIndeksSzuk, AxNazwaSzuk, AxGrMatSzuk, AxDostawcaSzuk, AxDataUtw1, KomOrgFiltr As String
    Dim AxDataUtw2 As Date
    Dim TblCollect As New Collection
    Dim TblCollectIdx As New Collection
    Dim TblCollectionCh As New Collection
    Dim CzyBylo, CzyBylo2 As Boolean
    Dim Komorka As Range
    Dim Data1, Data2, Data3, Data4, iT As Long
    Dim PctDone As Double
    Dim StanNaDz As Variant
    On Error GoTo error
    UFormINFO.MousePointer = fmMousePointerHourGlass
    UFormINFO.FrameProgress.Visible = True
    UFormINFO.LabelStatusBar.Visible = False
    If UFormINFO.ButtonIdxSzczegoly.Visible = True Then UFormINFO.ButtonIdxSzczegoly.Visible = False
    a = 1: iT = 0
    OstWierszIdx1 = UFormINFO.FrameProgress.Width - 10
    Application.ScreenUpdating = False
        PctDone = (a / OstWierszIdx1)
        UFormINFO.LabelProgress.Width = ((UFormINFO.FrameProgress.Width - 10) / OstWierszIdx1) * a
        UFormINFO.FrameProgress.Caption = Format(PctDone, "0.00%")
        DoEvents
        UFormINFO.ListBoxSzukajIndeks.Clear
        UFormINFO.ListBoxSzukajIndeks.ColumnWidths = "25;65;350;55;65;65;65;95;285"
        Call data_Format
        If AxApl__ Is Nothing Then loginAX
        If AxApl__ Is Nothing Then GoTo error
        If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
        KomOrgFiltr = "": SumTr = 0
            DoEvents
        LpIdx = 1: xPozList = 0: n = 1
        If UFormINFO.ComboBoxWyborOpcji.Value = "Stany Awaryjne" Then
            Tabela = 30262
        ElseIf UFormINFO.TextBoxDostawca.Value <> "Dostawca" Then
            Tabela = 340: AxPole1 = 3: AxPole2 = 8
            AxPole4 = 27
        Else
            Tabela = 175: AxPole1 = 2: AxPole2 = 3: AxPole3 = 1
            AxPole5 = 61444
        End If
        Set AxQuery = AxApl__.CreateObject("Query")
        Set AxDataSource = AxQuery.Call("AddDataSource", Tabela)
            If UFormINFO.TextBoxZakresBezIdx.Value <> "" Then
                For Each Komorka In Range(CStr(UFormINFO.TextBoxZakresBezIdx.Value))
                    If Komorka.text Like "*,*" Then Set TblCollectionCh = Rozbij_Tekst(CStr(Komorka.text))
                    If Komorka.text Like "*,*" Then
                        iT = 1
                        For Each Item In TblCollectionCh
                            TblCollectIdx.Add TblCollectionCh.Item(iT)
                            iT = iT + 1
                        Next
                        iT = 0
                    Else
                        TblCollectIdx.Add Komorka.Value
                    End If
                Next
            End If
            If UFormINFO.TextBoxZakresBezIdx2.Value <> "" Then
                For Each Komorka In Range(CStr(UFormINFO.TextBoxZakresBezIdx2.Value))
                    If Komorka.text Like "*,*" Then Set TblCollectionCh = Rozbij_Tekst(CStr(Komorka.text))
                    If Komorka.text Like "*,*" Then
                        iT = 1
                        For Each Item In TblCollectionCh
                            TblCollectIdx.Add TblCollectionCh.Item(iT)
                            iT = iT + 1
                        Next
                        iT = 0
                    Else
                        TblCollectIdx.Add Komorka.Value
                    End If
                Next
            End If
        If UFormINFO.ComboBoxWyborOpcji.Value <> "Stany Awaryjne" Then
            If UFormINFO.TextBoxIndeks.Value <> "indeks" Then
                AxIndeksSzuk = UFormINFO.TextBoxIndeks.Value
                Set AxRange1 = AxDataSource.Call("addRange", AxPole1)
                    AxRange1.Call "value", AxIndeksSzuk
            End If
            If UFormINFO.TextBoxNazwaTowaru.Value <> "Nazwa towaru" Then
                AxNazwaSzuk = UFormINFO.TextBoxNazwaTowaru.Value
                AxNazwaSzuk = Zamien_tekst_PRIV(CStr(AxNazwaSzuk))
                Set AxRange2 = AxDataSource.Call("addRange", AxPole2)
                    AxRange2.Call "value", AxNazwaSzuk
            End If
            If UFormINFO.TextBoxGrupaMaterialowa.Value <> "Gr. mat." Then
                If Tabela = 175 Then
                    AxGrMatSzuk = UFormINFO.TextBoxGrupaMaterialowa.Value
                    Set AxRange3 = AxDataSource.Call("addRange", AxPole3)
                        AxRange3.Call "value", AxGrMatSzuk
                End If
            End If
            If UFormINFO.TextBoxDostawca.Value <> "Dostawca" Then
                AxDostawcaSzuk = UFormINFO.TextBoxDostawca.Value
                Set AxRange4 = AxDataSource.Call("addRange", AxPole4)
                    AxRange4.Call "value", AxDostawcaSzuk
            End If
            If UFormINFO.TextBoxDataUtworzenia.Value <> "Data utworzenia" Then
                If Tabela = 175 Then
                    AxDataUtw1 = UFormINFO.TextBoxDataUtworzenia.Value
                    Set AxRange5 = AxDataSource.Call("addRange", AxPole5)
                        AxRange5.Call "value", AxDataUtw1
                End If
            End If
                Set AxQueryRun = AxApl__.CreateObject("QueryRun", AxQuery)
                    OstWierszIdx1 = AxApl__.CallStaticClassMethod("SysQuery", "countTotal", AxQueryRun)
                    PctDone = (a / OstWierszIdx1)
                    UFormINFO.LabelProgress.Width = ((UFormINFO.FrameProgress.Width - 10) / OstWierszIdx1) * a
                    UFormINFO.FrameProgress.Caption = Format(PctDone, "0.00%")
                    DoEvents
                If UFormINFO.TextBoxDostawca.Value = "Dostawca" Then
                    If UFormINFO.TextBoxKomOrgFiltr <> "Komorka Org." And UFormINFO.TextBoxKomOrgFiltr <> "" Then
                        KomOrgFiltr = UFormINFO.TextBoxKomOrgFiltr.text
                        If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                        a = 1
                        While (AxQueryRun.Call("Next"))
                            Set AxRecord = AxQueryRun.Call("GetNo", 1)
                            If TblCollectIdx.Count <> 0 Then
                                CzyBylo2 = czyWkolekcji(AxRecord.field(2), TblCollectIdx)
                                If CzyBylo2 = False Then
                                    SumTr = SumaTransakcjiDlaIdx(CStr(AxRecord.field(2)), 5, "!*PZ/*", CDate(0), CDate(0), "", KomOrgFiltr)
                                    If SumTr <> 0 Then
                                        With UFormINFO.ListBoxSzukajIndeks
                                            .AddItem LpIdx & "."
                                            .List(xPozList, 1) = AxRecord.field(2)
                                            .List(xPozList, 2) = AxRecord.field(3)
                                            .List(xPozList, 3) = AxRecord.field(1) & " / " & JednMiary(AxRecord.field(2))
                                            Data4 = 0: Data4 = Indeks_Data_Utworzenia(AxRecord.field(2))
                                            If CLng(Data4) < 100 Then .List(xPozList, 4) = "-" Else .List(xPozList, 4) = Format(CDate(Data4), SystemShortDateFormat)
                                            If UFormINFO.Label1.Caption = "Data ostatniej transakcji" Then
                                                Data1 = 0: Data2 = 0: Data3 = 0
                                                        Data1 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), KomOrgFiltr, "*PZ/*")
                                                    If Indeks_Czy_RWPNU(AxRecord.field(2)) = "RW" Then
                                                        Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), KomOrgFiltr, "*/RW/*")
                                                    ElseIf Indeks_Czy_RWPNU(AxRecord.field(2)) = "RWPNU" Then
                                                        Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), KomOrgFiltr, "*/RWPNU/*")
                                                    End If
                                                        Data3 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2))
                                                If CLng(Data1) < 100 Then .List(xPozList, 5) = "-" Else .List(xPozList, 5) = Format(CDate(Data1), SystemShortDateFormat)
                                                If CLng(Data2) < 100 Then .List(xPozList, 6) = "-" Else .List(xPozList, 6) = Format(CDate(Data2), SystemShortDateFormat)
                                                If CLng(Data3) < 100 Then .List(xPozList, 7) = "-" Else .List(xPozList, 7) = Format(CDate(Data3), SystemShortDateFormat)
                                            Else
                                                If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RW" Then
                                                    .List(xPozList, 5) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31), KomOrgFiltr))
                                                    .List(xPozList, 6) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31), KomOrgFiltr))
                                                    .List(xPozList, 7) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31), KomOrgFiltr))
                                                ElseIf Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RWPNU" Then
                                                    .List(xPozList, 5) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31), , KomOrgFiltr))
                                                    .List(xPozList, 6) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31), , KomOrgFiltr))
                                                    .List(xPozList, 7) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31), , KomOrgFiltr))
                                                Else
                                                    .List(xPozList, 5) = 0
                                                    .List(xPozList, 6) = 0
                                                    .List(xPozList, 7) = 0
                                                End If
                                            End If
                                            If UFormINFO.LabelKomOrgPobierTYT.Caption = "Komorki org. pobierajace mat." Then
                                                If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "inne" Then
                                                    .List(xPozList, 8) = "-"
                                                Else
                                                    .List(xPozList, 8) = Indeks_KtoPobieral_KomOrg(.List(xPozList, 1))
                                                End If
                                            ElseIf UFormINFO.LabelKomOrgPobierTYT.Caption = "Ostatni zakup - firma" Then
                                                .List(xPozList, 8) = Indeks_OstatniZakup_Firma(.List(xPozList, 1))
                                            Else
                                                .List(xPozList, 8) = "-"
                                            End If
                                        End With
                                        LpIdx = LpIdx + 1: xPozList = xPozList + 1
                                    End If
                                End If
                            Else
                                SumTr = SumaTransakcjiDlaIdx(CStr(AxRecord.field(2)), 5, "!*PZ/*", CDate(0), CDate(0), "", KomOrgFiltr)
                                If SumTr <> 0 Then
                                    With UFormINFO.ListBoxSzukajIndeks
                                        .AddItem LpIdx & "."
                                        .List(xPozList, 1) = AxRecord.field(2)
                                        .List(xPozList, 2) = AxRecord.field(3)
                                        .List(xPozList, 3) = AxRecord.field(1) & " / " & JednMiary(AxRecord.field(2))
                                        Data4 = 0: Data4 = Indeks_Data_Utworzenia(AxRecord.field(2))
                                        If CLng(Data4) < 100 Then .List(xPozList, 4) = "-" Else .List(xPozList, 4) = Format(CDate(Data4), SystemShortDateFormat)
                                        If UFormINFO.Label1.Caption = "Data ostatniej transakcji" Then
                                            Data1 = 0: Data2 = 0: Data3 = 0
                                                    Data1 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), KomOrgFiltr, "*PZ/*")
                                                If Indeks_Czy_RWPNU(AxRecord.field(2)) = "RW" Then
                                                    Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), KomOrgFiltr, "*/RW/*")
                                                ElseIf Indeks_Czy_RWPNU(AxRecord.field(2)) = "RWPNU" Then
                                                    Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), KomOrgFiltr, "*/RWPNU/*")
                                                End If
                                                    Data3 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2))
                                            If CLng(Data1) < 100 Then .List(xPozList, 5) = "-" Else .List(xPozList, 5) = Format(CDate(Data1), SystemShortDateFormat)
                                            If CLng(Data2) < 100 Then .List(xPozList, 6) = "-" Else .List(xPozList, 6) = Format(CDate(Data2), SystemShortDateFormat)
                                            If CLng(Data3) < 100 Then .List(xPozList, 7) = "-" Else .List(xPozList, 7) = Format(CDate(Data3), SystemShortDateFormat)
                                        Else
                                            If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RW" Then
                                                .List(xPozList, 5) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31), KomOrgFiltr))
                                                .List(xPozList, 6) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31), KomOrgFiltr))
                                                .List(xPozList, 7) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31), KomOrgFiltr))
                                            ElseIf Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RWPNU" Then
                                                .List(xPozList, 5) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31), , KomOrgFiltr))
                                                .List(xPozList, 6) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31), , KomOrgFiltr))
                                                .List(xPozList, 7) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31), , KomOrgFiltr))
                                            Else
                                                .List(xPozList, 5) = 0
                                                .List(xPozList, 6) = 0
                                                .List(xPozList, 7) = 0
                                            End If
                                        End If
                                        If UFormINFO.LabelKomOrgPobierTYT.Caption = "Komorki org. pobierajace mat." Then
                                            If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "inne" Then
                                                .List(xPozList, 8) = "-"
                                            Else
                                                .List(xPozList, 8) = Indeks_KtoPobieral_KomOrg(.List(xPozList, 1))
                                            End If
                                        ElseIf UFormINFO.LabelKomOrgPobierTYT.Caption = "Ostatni zakup - firma" Then
                                            .List(xPozList, 8) = Indeks_OstatniZakup_Firma(.List(xPozList, 1))
                                        Else
                                            .List(xPozList, 8) = "-"
                                        End If
                                    End With
                                    LpIdx = LpIdx + 1: xPozList = xPozList + 1
                                End If
                            End If
                            SumTr = 0: a = a + 1
                            PctDone = (a / OstWierszIdx1)
                            UFormINFO.LabelProgress.Width = ((UFormINFO.FrameProgress.Width - 10) / OstWierszIdx1) * a
                            UFormINFO.FrameProgress.Caption = Format(PctDone, "0.00%")
                            DoEvents
                            If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                        Wend
                    Else
                        a = 1
                        While (AxQueryRun.Call("Next"))
                            If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                            Set AxRecord = AxQueryRun.Call("GetNo", 1)
                            If TblCollectIdx.Count <> 0 Then
                                CzyBylo2 = czyWkolekcji(AxRecord.field(2), TblCollectIdx)
                                If CzyBylo2 = False Then
                                    With UFormINFO.ListBoxSzukajIndeks
                                        .AddItem LpIdx & "."
                                        .List(xPozList, 1) = AxRecord.field(2)
                                        .List(xPozList, 2) = AxRecord.field(3)
                                        .List(xPozList, 3) = AxRecord.field(1) & " / " & JednMiary(AxRecord.field(2))
                                        Data4 = 0: Data4 = Indeks_Data_Utworzenia(AxRecord.field(2))
                                        If CLng(Data4) < 100 Then .List(xPozList, 4) = "-" Else .List(xPozList, 4) = Format(CDate(Data4), SystemShortDateFormat)
                                        If UFormINFO.Label1.Caption = "Data ostatniej transakcji" Then
                                            Data1 = 0: Data2 = 0: Data3 = 0
                                                    Data1 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), , "*/PZ/*")
                                                If Indeks_Czy_RWPNU(AxRecord.field(2)) = "RW" Then
                                                    Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), , "*/RW/*")
                                                ElseIf Indeks_Czy_RWPNU(AxRecord.field(2)) = "RWPNU" Then
                                                    Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), , "*/RWPNU/*")
                                                End If
                                                    Data3 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2))
                                            If CLng(Data1) < 100 Then .List(xPozList, 5) = "-" Else .List(xPozList, 5) = Format(CDate(Data1), SystemShortDateFormat)
                                            If CLng(Data2) < 100 Then .List(xPozList, 6) = "-" Else: .List(xPozList, 6) = Format(CDate(Data2), SystemShortDateFormat)
                                            If CLng(Data3) < 100 Then .List(xPozList, 7) = "-" Else .List(xPozList, 7) = Format(CDate(Data3), SystemShortDateFormat)
                                        Else
                                            If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RW" Then
                                                .List(xPozList, 5) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31)))
                                                .List(xPozList, 6) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31)))
                                                .List(xPozList, 7) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31)))
                                            ElseIf Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RWPNU" Then
                                                .List(xPozList, 5) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31)))
                                                .List(xPozList, 6) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31)))
                                                .List(xPozList, 7) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31)))
                                            Else
                                                .List(xPozList, 5) = 0
                                                .List(xPozList, 6) = 0
                                                .List(xPozList, 7) = 0
                                            End If
                                        End If
                                        If UFormINFO.LabelKomOrgPobierTYT.Caption = "Komorki org. pobierajace mat." Then
                                            If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "inne" Then
                                                .List(xPozList, 8) = "-"
                                            Else
                                                .List(xPozList, 8) = Indeks_KtoPobieral_KomOrg(.List(xPozList, 1))
                                            End If
                                        ElseIf UFormINFO.LabelKomOrgPobierTYT.Caption = "Ostatni zakup - firma" Then
                                            .List(xPozList, 8) = Indeks_OstatniZakup_Firma(.List(xPozList, 1))
                                        Else
                                            .List(xPozList, 8) = "-"
                                        End If
                                    End With
                                    LpIdx = LpIdx + 1: xPozList = xPozList + 1
                                End If
                            Else
                                With UFormINFO.ListBoxSzukajIndeks
                                    .AddItem LpIdx & "."
                                    .List(xPozList, 1) = AxRecord.field(2)
                                    .List(xPozList, 2) = AxRecord.field(3)
                                    .List(xPozList, 3) = AxRecord.field(1) & " / " & JednMiary(AxRecord.field(2))
                                    Data4 = 0: Data4 = Indeks_Data_Utworzenia(AxRecord.field(2))
                                    If CLng(Data4) < 100 Then .List(xPozList, 4) = "-" Else .List(xPozList, 4) = Format(CDate(Data4), SystemShortDateFormat)
                                    If UFormINFO.Label1.Caption = "Data ostatniej transakcji" Then
                                        Data1 = 0: Data2 = 0: Data3 = 0
                                                Data1 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), , "*/PZ/*")
                                            If Indeks_Czy_RWPNU(AxRecord.field(2)) = "RW" Then
                                                Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), , "*/RW/*")
                                            ElseIf Indeks_Czy_RWPNU(AxRecord.field(2)) = "RWPNU" Then
                                                Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2), , "*/RWPNU/*")
                                            End If
                                                Data3 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(AxRecord.field(2))
                                        If CLng(Data1) < 100 Then .List(xPozList, 5) = "-" Else .List(xPozList, 5) = Format(CDate(Data1), SystemShortDateFormat)
                                        If CLng(Data2) < 100 Then .List(xPozList, 6) = "-" Else: .List(xPozList, 6) = Format(CDate(Data2), SystemShortDateFormat)
                                        If CLng(Data3) < 100 Then .List(xPozList, 7) = "-" Else .List(xPozList, 7) = Format(CDate(Data3), SystemShortDateFormat)
                                    Else
                                        If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RW" Then
                                            .List(xPozList, 5) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31)))
                                            .List(xPozList, 6) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31)))
                                            .List(xPozList, 7) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31)))
                                        ElseIf Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RWPNU" Then
                                            .List(xPozList, 5) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31)))
                                            .List(xPozList, 6) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31)))
                                            .List(xPozList, 7) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31)))
                                        Else
                                            .List(xPozList, 5) = 0
                                            .List(xPozList, 6) = 0
                                            .List(xPozList, 7) = 0
                                        End If
                                    End If
                                    If UFormINFO.LabelKomOrgPobierTYT.Caption = "Komorki org. pobierajace mat." Then
                                        If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "inne" Then
                                            .List(xPozList, 8) = "-"
                                        Else
                                            .List(xPozList, 8) = Indeks_KtoPobieral_KomOrg(.List(xPozList, 1))
                                        End If
                                    ElseIf UFormINFO.LabelKomOrgPobierTYT.Caption = "Ostatni zakup - firma" Then
                                        .List(xPozList, 8) = Indeks_OstatniZakup_Firma(.List(xPozList, 1))
                                    Else
                                        .List(xPozList, 8) = "-"
                                    End If
                                End With
                                LpIdx = LpIdx + 1: xPozList = xPozList + 1
                            End If
                            a = a + 1
                            PctDone = (a / OstWierszIdx1)
                            UFormINFO.LabelProgress.Width = ((UFormINFO.FrameProgress.Width - 10) / OstWierszIdx1) * a
                            UFormINFO.FrameProgress.Caption = Format(PctDone, "0.00%")
                            DoEvents
                            If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                        Wend
                    End If
                Else
                    If UFormINFO.TextBoxKomOrgFiltr <> "Komorka Org." And UFormINFO.TextBoxKomOrgFiltr <> "" Then
                        KomOrgFiltr = UFormINFO.TextBoxKomOrgFiltr.text
                        a = 1
                        While (AxQueryRun.Call("Next"))
                            Set AxRecord = AxQueryRun.Call("GetNo", 1)
                                SumTr = SumaTransakcjiDlaIdx(CStr(AxRecord.field(3)), 5, "!*PZ/*", CDate(0), CDate(0), "", KomOrgFiltr)
                                If SumTr <> 0 Then
                                    If TblCollectIdx.Count <> 0 Then
                                        CzyBylo2 = czyWkolekcji(AxRecord.field(3), TblCollectIdx)
                                        If CzyBylo2 = False Then
                                            CzyBylo = czyWkolekcji(AxRecord.field(3), TblCollect)
                                            If CzyBylo = False Then TblCollect.Add AxRecord.field(3)
                                        End If
                                    Else
                                        CzyBylo = czyWkolekcji(AxRecord.field(3), TblCollect)
                                        If CzyBylo = False Then TblCollect.Add AxRecord.field(3)
                                    End If
                                End If
                                SumTr = 0
                            If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                        Wend
                        For lpx = 1 To TblCollect.Count
                            With UFormINFO.ListBoxSzukajIndeks
                                .AddItem LpIdx & "."
                                .List(xPozList, 1) = TblCollect(lpx)
                                .List(xPozList, 2) = NazwaTowaru(CStr(TblCollect(lpx)))
                                .List(xPozList, 3) = GrMaterial(CStr(TblCollect(lpx))) & " / " & JednMiary(CStr(TblCollect(lpx)))
                                Data4 = 0: Data4 = Indeks_Data_Utworzenia(CStr(TblCollect(lpx)))
                                If CLng(Data4) < 100 Then .List(xPozList, 4) = "-" Else .List(xPozList, 4) = Format(CDate(Data4), SystemShortDateFormat)
                                If UFormINFO.Label1.Caption = "Data ostatniej transakcji" Then
                                    Data1 = 0: Data2 = 0: Data3 = 0
                                            Data1 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(TblCollect(lpx), , "*/PZ/*")
                                        If Indeks_Czy_RWPNU(AxRecord.field(2)) = "RW" Then
                                            Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(TblCollect(lpx), , "*/RW/*")
                                        ElseIf Indeks_Czy_RWPNU(AxRecord.field(2)) = "RWPNU" Then
                                            Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(TblCollect(lpx), , "*/RWPNU/*")
                                        End If
                                            Data3 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(TblCollect(lpx))
                                    If CLng(Data1) < 100 Then .List(xPozList, 5) = "-" Else .List(xPozList, 5) = Format(CDate(Data1), SystemShortDateFormat)
                                    If CLng(Data2) < 100 Then .List(xPozList, 6) = "-" Else .List(xPozList, 6) = Format(CDate(Data2), SystemShortDateFormat)
                                    If CLng(Data3) < 100 Then .List(xPozList, 7) = "-" Else .List(xPozList, 7) = Format(CDate(Data3), SystemShortDateFormat)
                                Else
                                    If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RW" Then
                                        .List(xPozList, 5) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31)))
                                        .List(xPozList, 6) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31)))
                                        .List(xPozList, 7) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31)))
                                    ElseIf Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RWPNU" Then
                                        .List(xPozList, 5) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31)))
                                        .List(xPozList, 6) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31)))
                                        .List(xPozList, 7) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31)))
                                    Else
                                        .List(xPozList, 5) = 0
                                        .List(xPozList, 6) = 0
                                        .List(xPozList, 7) = 0
                                    End If
                                End If
                                If UFormINFO.LabelKomOrgPobierTYT.Caption = "Komorki org. pobierajace mat." Then
                                    If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "inne" Then
                                        .List(xPozList, 8) = "-"
                                    Else
                                        .List(xPozList, 8) = Indeks_KtoPobieral_KomOrg(.List(xPozList, 1))
                                    End If
                                ElseIf UFormINFO.LabelKomOrgPobierTYT.Caption = "Ostatni zakup - firma" Then
                                    .List(xPozList, 8) = Indeks_OstatniZakup_Firma(.List(xPozList, 1))
                                Else
                                    .List(xPozList, 8) = "-"
                                End If
                            End With
                            LpIdx = LpIdx + 1: xPozList = xPozList + 1
                            a = a + 1
                            PctDone = (a / OstWierszIdx1)
                            UFormINFO.LabelProgress.Width = ((UFormINFO.FrameProgress.Width - 10) / OstWierszIdx1) * a
                            UFormINFO.FrameProgress.Caption = Format(PctDone, "0.00%")
                            DoEvents
                            If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                        Next
                    Else
                        a = 1
                        While (AxQueryRun.Call("Next"))
                            Set AxRecord = AxQueryRun.Call("GetNo", 1)
                            If TblCollectIdx.Count <> 0 Then
                                CzyBylo2 = czyWkolekcji(AxRecord.field(3), TblCollectIdx)
                                If CzyBylo2 = False Then
                                    CzyBylo = czyWkolekcji(AxRecord.field(3), TblCollect)
                                    If CzyBylo = False Then TblCollect.Add AxRecord.field(3)
                                End If
                            Else
                                CzyBylo = czyWkolekcji(AxRecord.field(3), TblCollect)
                                If CzyBylo = False Then TblCollect.Add AxRecord.field(3)
                            End If
                            If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                        Wend
                        For lpx = 1 To TblCollect.Count
                            With UFormINFO.ListBoxSzukajIndeks
                                .AddItem LpIdx & "."
                                .List(xPozList, 1) = TblCollect(lpx)
                                .List(xPozList, 2) = NazwaTowaru(CStr(TblCollect(lpx)))
                                .List(xPozList, 3) = GrMaterial(CStr(TblCollect(lpx))) & " / " & JednMiary(CStr(TblCollect(lpx)))
                                Data4 = 0: Data4 = Indeks_Data_Utworzenia(CStr(TblCollect(lpx)))
                                If CLng(Data4) < 100 Then .List(xPozList, 4) = "-" Else .List(xPozList, 4) = Format(CDate(Data4), SystemShortDateFormat)
                                If UFormINFO.Label1.Caption = "Data ostatniej transakcji" Then
                                    Data1 = 0: Data2 = 0: Data3 = 0
                                        Data1 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(TblCollect(lpx), , "*/PZ/*")
                                        If Indeks_Czy_RWPNU(AxRecord.field(2)) = "RW" Then
                                            Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(TblCollect(lpx), , "*/RW/*")
                                        ElseIf Indeks_Czy_RWPNU(AxRecord.field(2)) = "RWPNU" Then
                                            Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(TblCollect(lpx), , "*/RWPNU/*")
                                        End If
                                        Data3 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(TblCollect(lpx))
                                    If CLng(Data1) < 100 Then .List(xPozList, 5) = "-" Else .List(xPozList, 5) = Format(CDate(Data1), SystemShortDateFormat)
                                    If CLng(Data2) < 100 Then .List(xPozList, 6) = "-" Else .List(xPozList, 6) = Format(CDate(Data2), SystemShortDateFormat)
                                    If CLng(Data3) < 100 Then .List(xPozList, 7) = "-" Else: .List(xPozList, 7) = Format(CDate(Data3), SystemShortDateFormat)
                                Else
                                    If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RW" Then
                                        .List(xPozList, 5) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31)))
                                        .List(xPozList, 6) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31)))
                                        .List(xPozList, 7) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31)))
                                    ElseIf Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RWPNU" Then
                                        .List(xPozList, 5) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31)))
                                        .List(xPozList, 6) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31)))
                                        .List(xPozList, 7) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31)))
                                    Else
                                        .List(xPozList, 5) = 0
                                        .List(xPozList, 6) = 0
                                        .List(xPozList, 7) = 0
                                    End If
                                End If
                                If UFormINFO.LabelKomOrgPobierTYT.Caption = "Komorki org. pobierajace mat." Then
                                    If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "inne" Then
                                        .List(xPozList, 8) = "-"
                                    Else
                                        .List(xPozList, 8) = Indeks_KtoPobieral_KomOrg(.List(xPozList, 1))
                                    End If
                                ElseIf UFormINFO.LabelKomOrgPobierTYT.Caption = "Ostatni zakup - firma" Then
                                    .List(xPozList, 8) = Indeks_OstatniZakup_Firma(.List(xPozList, 1))
                                Else
                                    .List(xPozList, 8) = "-"
                                End If
                            End With
                            LpIdx = LpIdx + 1: xPozList = xPozList + 1
                            a = a + 1
                            PctDone = (a / OstWierszIdx1)
                            UFormINFO.LabelProgress.Width = ((UFormINFO.FrameProgress.Width - 10) / OstWierszIdx1) * a
                            UFormINFO.FrameProgress.Caption = Format(PctDone, "0.00%")
                            DoEvents
                            If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                        Next
                    End If
                End If
        Else
            xxx = 0
            AxDataSource.Call "addSortField", 30001, 2
            Set AxQueryRun = AxApl__.CreateObject("QueryRun", AxQuery)
                OstWierszIdx1 = AxApl__.CallStaticClassMethod("SysQuery", "countTotal", AxQueryRun)
                PctDone = (a / OstWierszIdx1)
                UFormINFO.LabelProgress.Width = ((UFormINFO.FrameProgress.Width - 10) / OstWierszIdx1) * a
                UFormINFO.FrameProgress.Caption = Format(PctDone, "0.00%")
                DoEvents
                    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                    a = 1
                    While (AxQueryRun.Call("Next"))
                        Set AxRecord = AxQueryRun.Call("GetNo", 1)
                        If AxRecord.field(30001) = "000-000-000-0" Then GoTo NastepnyIdx
                            If TblCollectIdx.Count <> 0 Then
                                CzyBylo2 = czyWkolekcji(AxRecord.field(30001), TblCollectIdx)
                                If CzyBylo2 = False Then
                                    CzyBylo = czyWkolekcji(AxRecord.field(30001), TblCollect)
                                    If CzyBylo = False Then TblCollect.Add AxRecord.field(30001)
                                End If
                            Else
                                CzyBylo = czyWkolekcji(AxRecord.field(30001), TblCollect)
                                If CzyBylo = False Then TblCollect.Add AxRecord.field(30001)
                            End If
                            If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
NastepnyIdx:
                    Wend
                    For lpx = 1 To TblCollect.Count
                            With UFormINFO.ListBoxSzukajIndeks
                                .AddItem LpIdx & "."
                                .List(xPozList, 1) = TblCollect(lpx)
                                .List(xPozList, 2) = NazwaTowaru(CStr(TblCollect(lpx)))
                                .List(xPozList, 3) = GrMaterial(CStr(TblCollect(lpx))) & " / " & JednMiary(CStr(TblCollect(lpx)))
                                .List(xPozList, 4) = StanNaDzienIndeks(CStr(TblCollect(lpx)), Now)
                                If InStr(1, CStr(.List(xPozList, 4)), ".") <> 0 Then .List(xPozList, 4) = Replace(CStr(.List(xPozList, 4)), ".", ",")
                                If UFormINFO.Label1.Caption = "Data ostatniej transakcji" Then
                                    Data1 = 0: Data2 = 0: Data3 = 0
                                        Data1 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(CStr(TblCollect(lpx)), , "*/PZ/*")
                                        If Indeks_Czy_RWPNU(AxRecord.field(2)) = "RW" Then
                                            Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(CStr(TblCollect(lpx)), , "*/RW/*")
                                        ElseIf Indeks_Czy_RWPNU(AxRecord.field(2)) = "RWPNU" Then
                                            Data2 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(CStr(TblCollect(lpx)), , "*/RWPNU/*")
                                        End If
                                        Data3 = DataFizOstZaksiegowanejDlaIndeksuKomOrg(CStr(TblCollect(lpx)))
                                    If CLng(Data1) < 100 Then .List(xPozList, 5) = "-" Else .List(xPozList, 5) = Format(CDate(Data1), SystemShortDateFormat)
                                    If CLng(Data2) < 100 Then .List(xPozList, 6) = "-" Else .List(xPozList, 6) = Format(CDate(Data2), SystemShortDateFormat)
                                    If CLng(Data3) < 100 Then .List(xPozList, 7) = "-" Else .List(xPozList, 7) = Format(CDate(Data3), SystemShortDateFormat)
                                Else
                                    If Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RW" Then
                                        .List(xPozList, 5) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31)))
                                        .List(xPozList, 6) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31)))
                                        .List(xPozList, 7) = Abs(ZuzycieKomOrgIlosc(.List(xPozList, 1), DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31)))
                                    ElseIf Indeks_Czy_RWPNU(.List(xPozList, 1)) = "RWPNU" Then
                                        .List(xPozList, 5) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataPrzychTYT.Caption), 1, 1), DateSerial(CInt(LabelDataPrzychTYT.Caption), 12, 31)))
                                        .List(xPozList, 6) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelDataRozchTYT.Caption), 1, 1), DateSerial(CInt(LabelDataRozchTYT.Caption), 12, 31)))
                                        .List(xPozList, 7) = Abs(SumaTransakcjiDlaIdx(.List(xPozList, 1), 5, "*RWPNU*", DateSerial(CInt(LabelOgolTYT.Caption), 1, 1), DateSerial(CInt(LabelOgolTYT.Caption), 12, 31)))
                                    Else
                                        .List(xPozList, 5) = 0
                                        .List(xPozList, 6) = 0
                                        .List(xPozList, 7) = 0
                                    End If
                                End If
                                .List(xPozList, 8) = (Indeks_StanyAwaryjne_Min(CStr(TblCollect(lpx)))) & " < " & .List(xPozList, 4) & " < " & (Indeks_StanyAwaryjne_Max(CStr(TblCollect(lpx))))
                            End With
                            LpIdx = LpIdx + 1: xPozList = xPozList + 1
                        SumTr = 0: a = a + 1
                        PctDone = (a / OstWierszIdx1)
                        UFormINFO.LabelProgress.Width = ((UFormINFO.FrameProgress.Width - 10) / OstWierszIdx1) * a
                        UFormINFO.FrameProgress.Caption = Format(PctDone, "0.00%")
                        DoEvents
                        If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                    Next
        End If
        SumTr = 0
        IlePobrane = LpIdx - 1
        UFormINFO.FrameProgress.Visible = False
        UFormINFO.LabelStatusBar.Visible = True
        UFormINFO.LabelStatusBar.Caption = "Autor dodatku: Jakub Koziorowski ; " & Chr(13) & _
                                                "£¹cznie: " & IlePobrane & " z " & IlePobrane
    Application.ScreenUpdating = True
    UFormINFO.MousePointer = fmMousePointerDefault
    DoEvents
    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
    If UFormINFO.Label1.Caption = "Zuzycie" Then Call SortListBox_fVAT(UFormINFO.ListBoxSzukajIndeks, 7) Else Call SortListBox_fVAT(UFormINFO.ListBoxSzukajIndeks, 5)
    Exit Sub
error:
    UFormINFO.FrameProgress.Visible = False
    UFormINFO.LabelStatusBar.Visible = True
    UFormINFO.LabelStatusBar.Caption = "Autor dodatku: Jakub Koziorowski ; " & Chr(13) & _
                                                "B³¹d po³¹czenia !!!"
    UFormINFO.MousePointer = fmMousePointerDefault
    DoEvents
    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
    Exit Sub
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ButtonSzukGoogle_Click()
    UFormINFO.MousePointer = fmMousePointerHourGlass
    DoEvents
        TekstSearchGoogle = UFormINFO.TextBoxSzukGoogle.Value
        Call OpenUrl
    UFormINFO.MousePointer = fmMousePointerDefault
    DoEvents
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckBoxMinimal_Click()
    If UFormINFO.CheckBoxMinimal.Value = True Then
        UFormINFO.Height = "96,75"
        UFormINFO.Top = Application.Top + Application.Height - 40 - UFormINFO.Height
        UFormINFO.Left = Application.Left + Application.Width - 30 - UFormINFO.Width
    ElseIf UFormINFO.CheckBoxMinimal.Value = False Then
        UFormINFO.Height = "465"
        UFormINFO.Top = Application.Top + Application.Height - 40 - UFormINFO.Height
        UFormINFO.Left = Application.Left + Application.Width - 30 - UFormINFO.Width
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LabelLpTYT_Click()
    Dim SelAC As Range
        Set SelAC = ActiveCell
        SelAC = UFormINFO.LabelIdxTYT.Caption
        SelAC.Offset(, 1) = UFormINFO.LabelNazwaTYT.Caption
        SelAC.Offset(, 2) = UFormINFO.LabelGrMatTYT.Caption
        SelAC.Offset(, 3) = UFormINFO.LabelJM.Caption
        SelAC.Offset(, 4) = UFormINFO.LabelDataPrzychTYT.Caption
        SelAC.Offset(, 5) = UFormINFO.LabelDataRozchTYT.Caption
        SelAC.Offset(, 6) = UFormINFO.LabelOgolTYT.Caption
        SelAC.Offset(, 7) = UFormINFO.LabelKomOrgPobierTYT.Caption
    ActiveSheet.Columns.AutoFit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ListBoxSzukajIndeks_Click()
    Dim ActPozLi As Long
    Dim a, c As String
    DoEvents
    If UFormINFO.ButtonIdxSzczegoly.Visible = False Then UFormINFO.ButtonIdxSzczegoly.Visible = True
    If UFormINFO.ListBoxSzukajIndeks.ListCount <> 0 Then
        ActPozLi = UFormINFO.ListBoxSzukajIndeks.ListIndex
        UFormINFO.TextBoxKomOrg.Value = UFormINFO.ListBoxSzukajIndeks.List(ActPozLi, 8)
        UFormINFO.TextBoxSzukGoogle.Value = UFormINFO.ListBoxSzukajIndeks.List(ActPozLi, 2)
            a = UFormINFO.ListBoxSzukajIndeks.Top
            c = UFormINFO.ListBoxSzukajIndeks.TopIndex
                UFormINFO.ButtonIdxSzczegoly.Top = a + ((ActPozLi - c) * 9.81)
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ButtonIdxSzczegoly_Click()
    Dim ActPozLi As Long
    Application.Calculation = xlCalculationManual
    UFormINFO.MousePointer = fmMousePointerHourGlass
    DoEvents
        If UFormINFO.ListBoxSzukajIndeks.ListCount <> 0 Then
            Call IndeksCzyscB_Click
            ActPozLi = UFormINFO.ListBoxSzukajIndeks.ListIndex
            If ActPozLi = -1 Then
                MsgBox "Nie zaznaczy³e pozycji na licie !!!"
            Else
                UFormINFO.IndeksTxtBox.Value = UFormINFO.ListBoxSzukajIndeks.List(ActPozLi, 1)
                Call IndeksSzukajB_Click
            End If
        End If
    UFormINFO.MousePointer = fmMousePointerDefault
    DoEvents
    Application.Calculation = xlCalculationAutomatic
End Sub
''================================================================================================================================================
Private Sub ListBoxSzukajIndeks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim PytanieX As Byte
    Dim ActRow, LastRow, yaRow, ActCol As Long
    Dim ListRow1 As Long
    Dim WbkAC As Workbook
    Dim WksAC As Worksheet
        Set WbkAC = Application.ActiveWorkbook
        Set WksAC = WbkAC.ActiveSheet
    ActRow = 0: LastRow = 0: yaRow = 0: ActCol = 0
    If UFormINFO.ListBoxSzukajIndeks.ListCount = 0 Then Exit Sub
    ab = UFormINFO.ListBoxSzukajIndeks.ListIndex
    If UFormINFO.CheckBoxCzyDodawac = False Then
        If UFormINFO.CheckBoxDodCalaLista = True Then
            ActRow = ActiveCell.Row
            ActCol = ActiveCell.Column
            LastRow = UFormINFO.ListBoxSzukajIndeks.ListCount
                For yaRow = ActRow To (ActRow + LastRow)
                    If WksAC.Cells(yaRow, ActCol).Value <> "" Then
                        PytanieX = MsgBox("Nie dodales odpowiedniej liczby pustych wierszy! Czy nadpisac wartosci ?", vbYesNo)
                        Select Case PytanieX
                            Case vbYes
                                Exit For
                            Case vbNo
                                Exit Sub
                        End Select
                    End If
                Next
            PytanieX = 0
            ActRow = 0: LastRow = 0: yaRow = 0: ActCol = 0: ListRow1 = 0
            If ActiveCell <> "" Then
                PytanieX = MsgBox("komórka która jest zaznaczona, nie jest pusta. Czy dodaæ mimo to ?", vbYesNo)
                Select Case PytanieX
                    Case vbYes
                        ActRow = ActiveCell.Row
                        ActCol = ActiveCell.Column
                        LastRow = UFormINFO.ListBoxSzukajIndeks.ListCount
                            For yaRow = ActRow To (ActRow + LastRow) - 1
                                WksAC.Cells(yaRow, ActCol) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 1)
                                If UFormINFO.CheckBoxCzyTylkoIndeks.Value = False Then
                                    WksAC.Cells(yaRow, ActCol).Offset(, 1) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 2)
                                    WksAC.Cells(yaRow, ActCol).Offset(, 2) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 3)
                                    WksAC.Cells(yaRow, ActCol).Offset(, 3) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 4)
                                    WksAC.Cells(yaRow, ActCol).Offset(, 4) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 5)
                                    WksAC.Cells(yaRow, ActCol).Offset(, 5) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 6)
                                    WksAC.Cells(yaRow, ActCol).Offset(, 6) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 7)
                                    WksAC.Cells(yaRow, ActCol).Offset(, 7) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 8)
                                End If
                                ListRow1 = ListRow1 + 1
                            Next
                    Case vbNo
                        Exit Sub
                End Select
            Else
                ActRow = ActiveCell.Row
                ActCol = ActiveCell.Column
                LastRow = UFormINFO.ListBoxSzukajIndeks.ListCount
                    For yaRow = ActRow To (ActRow + LastRow) - 1
                        WksAC.Cells(yaRow, ActCol) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 1)
                            If UFormINFO.CheckBoxCzyTylkoIndeks.Value = False Then
                                WksAC.Cells(yaRow, ActCol).Offset(, 1) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 2)
                                WksAC.Cells(yaRow, ActCol).Offset(, 2) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 3)
                                WksAC.Cells(yaRow, ActCol).Offset(, 3) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 4)
                                WksAC.Cells(yaRow, ActCol).Offset(, 4) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 5)
                                WksAC.Cells(yaRow, ActCol).Offset(, 5) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 6)
                                WksAC.Cells(yaRow, ActCol).Offset(, 6) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 7)
                                WksAC.Cells(yaRow, ActCol).Offset(, 7) = UFormINFO.ListBoxSzukajIndeks.List(ListRow1, 8)
                            End If
                        ListRow1 = ListRow1 + 1
                    Next
            End If
        Else
            ActiveCell = UFormINFO.ListBoxSzukajIndeks.List(ab, 1)
                If UFormINFO.CheckBoxCzyTylkoIndeks.Value = False Then
                    ActiveCell.Offset(, 1) = UFormINFO.ListBoxSzukajIndeks.List(ab, 2)
                    ActiveCell.Offset(, 2) = UFormINFO.ListBoxSzukajIndeks.List(ab, 3)
                    ActiveCell.Offset(, 3) = UFormINFO.ListBoxSzukajIndeks.List(ab, 4)
                    ActiveCell.Offset(, 4) = UFormINFO.ListBoxSzukajIndeks.List(ab, 5)
                    ActiveCell.Offset(, 5) = UFormINFO.ListBoxSzukajIndeks.List(ab, 6)
                    ActiveCell.Offset(, 6) = UFormINFO.ListBoxSzukajIndeks.List(ab, 7)
                    ActiveCell.Offset(, 7) = UFormINFO.ListBoxSzukajIndeks.List(ab, 8)
                End If
            ActRow = ActiveCell.Row
            ActCol = ActiveCell.Column
            WksAC.Cells(ActRow + 1, ActCol).Select
        End If
    End If
    If UFormINFO.CheckBoxCzyUsuwac.Value = True Then
        If UFormINFO.CheckBoxDodCalaLista = True Then
            UFormINFO.ListBoxSzukajIndeks.Clear
            UFormINFO.LabelStatusBar.Caption = "Autor dodatku: Jakub Koziorowski ; " & Chr(13) & _
                                                    "£¹cznie: " & (UFormINFO.ListBoxSzukajIndeks.ListCount) & " z " & IlePobrane
        Else
            UFormINFO.ListBoxSzukajIndeks.RemoveItem (ab)
                xPoz = (UFormINFO.ListBoxSzukajIndeks.ListCount) - 1
                xLp = 1
                For xPoz1 = 0 To xPoz
                    UFormINFO.ListBoxSzukajIndeks.List(xPoz1, 0) = xLp & "."
                    xLp = xLp + 1
                Next
            UFormINFO.LabelStatusBar.Caption = "Autor dodatku: Jakub Koziorowski ; " & Chr(13) & _
                                                    "£¹cznie: " & (UFormINFO.ListBoxSzukajIndeks.ListCount) & " z " & IlePobrane
        End If
    End If
    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
End Sub
''=======================================================================================================================================
Private Sub TextBoxKomOrgFiltr_Enter()
    If UFormINFO.TextBoxKomOrgFiltr.Value = "Komorka Org." Then UFormINFO.TextBoxKomOrgFiltr.Value = ""
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextBoxKomOrgFiltr_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If UFormINFO.TextBoxKomOrgFiltr.Value = "" Then UFormINFO.TextBoxKomOrgFiltr.Value = "Komorka Org."
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextBoxDostawca_Enter()
    If UFormINFO.TextBoxDostawca.Value = "Dostawca" Then
        UFormINFO.TextBoxDostawca.Value = "": UFormNr = 2
        UFormDostawcyKonta.Show
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextBoxDostawca_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If UFormINFO.TextBoxDostawca.Value = "" Then
        UFormINFO.TextBoxDostawca.Value = "Dostawca": UFormNr = 0
            UFormINFO.TextBoxGrupaMaterialowa.Enabled = True
            UFormINFO.TextBoxDataUtworzenia.Enabled = True
    Else
        UFormINFO.TextBoxGrupaMaterialowa.Value = "Gr. mat."
        UFormINFO.TextBoxDataUtworzenia.Value = "Data utworzenia"
            UFormINFO.TextBoxGrupaMaterialowa.Enabled = False
            UFormINFO.TextBoxDataUtworzenia.Enabled = False
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextBoxGrupaMaterialowa_Enter()
    If UFormINFO.TextBoxGrupaMaterialowa.Value = "Gr. mat." Then UFormINFO.TextBoxGrupaMaterialowa.Value = ""
End Sub
Private Sub TextBoxGrupaMaterialowa_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If UFormINFO.TextBoxGrupaMaterialowa.Value = "" Then UFormINFO.TextBoxGrupaMaterialowa.Value = "Gr. mat."
End Sub
Private Sub TextBoxIndeks_Enter()
    If UFormINFO.TextBoxIndeks.Value = "indeks" Then UFormINFO.TextBoxIndeks.Value = ""
End Sub
Private Sub TextBoxIndeks_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If UFormINFO.TextBoxIndeks.Value = "" Then UFormINFO.TextBoxIndeks.Value = "indeks"
End Sub
Private Sub TextBoxNazwaTowaru_Enter()
    If UFormINFO.TextBoxNazwaTowaru.Value = "Nazwa towaru" Then UFormINFO.TextBoxNazwaTowaru.Value = ""
End Sub
Private Sub TextBoxNazwaTowaru_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If UFormINFO.TextBoxNazwaTowaru.Value = "" Then UFormINFO.TextBoxNazwaTowaru.Value = "Nazwa towaru"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextBoxDataUtworzenia_Enter()
    If UFormINFO.TextBoxDataUtworzenia.Value = "Data utworzenia" Then
        UFormINFO.TextBoxDataUtworzenia.Value = ""
        Call data_Format
            CzyDataKalend = True
        Call OtwCal
        If DataUtwo <> 0 Then
            UFormINFO.TextBoxDataUtworzenia.Value = Format(DataUtwo, SystemShortDateFormat)
            DataUtwo = CDate(0)
        Else
            DataUtwo = CDate(0)
            UFormINFO.TextBoxDataUtworzenia.Value = "Data utworzenia"
        End If
        CzyDataKalend = False
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextBoxDataUtworzenia_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If UFormINFO.TextBoxDataUtworzenia.Value = "" Then
        CzyDataKalend = False
        UFormINFO.TextBoxDataUtworzenia.Value = "Data utworzenia"
    End If
    CzyDataKalend = False
End Sub

'-----------------------------------------Dot. Zakladki Indeks------------------------------------------------
'=============================================================================================================
Private Sub IndeksCzyscB_Click()
    UFormINFO.IndeksTxtBox.Value = "podaj indeks"
    UFormINFO.ComboBoxRwCzyPZ.Value = "Zu¿ycie w latach"
    UFormINFO.ComboBoxKomOrg.Value = ""
    UFormINFO.IndeksNazwa.Value = "": UFormINFO.IndeksJM.Value = "": UFormINFO.IndeksGrMat.Value = "": UFormINFO.IndeksNazwaGrMat.Value = ""
    UFormINFO.IndeksAlias.Value = "": UFormINFO.IndeksDataUtw.Value = "": UFormINFO.IndeksKtoUtw.Value = "": UFormINFO.IndeksDataMod.Value = ""
    UFormINFO.IndeksKtoMod.Value = "": UFormINFO.IndeksStanMIN.Value = "": UFormINFO.IndeksStanMAX.Value = ""
    UFormINFO.IndeksListBoxDostawcy.Clear
    UFormINFO.IndeksListBoxZakupy.Clear
    UFormINFO.IndeksListBoxDostawcy.Height = "137,3"
    UFormINFO.IndeksListBoxZakupy.Height = "285"
    UFormINFO.ButtonIdxSzczegoly.Visible = False
        UFormINFO.LabelRok9.Caption = 0: UFormINFO.LabelRok8.Caption = 0: UFormINFO.LabelRok7.Caption = 0: UFormINFO.LabelRok6.Caption = 0
        UFormINFO.LabelRok5.Caption = 0: UFormINFO.LabelRok4.Caption = 0: UFormINFO.LabelRok3.Caption = 0: UFormINFO.LabelRok2.Caption = 0
        UFormINFO.LabelRok1.Caption = 0
            UFormINFO.TextBoxRok1_ilosc.Value = "": UFormINFO.TextBoxRok1_wartosc.Value = "": UFormINFO.TextBoxRok2_ilosc.Value = "": UFormINFO.TextBoxRok2_wartosc.Value = ""
            UFormINFO.TextBoxRok3_ilosc.Value = "": UFormINFO.TextBoxRok3_wartosc.Value = "": UFormINFO.TextBoxRok4_ilosc.Value = "": UFormINFO.TextBoxRok4_wartosc.Value = ""
            UFormINFO.TextBoxRok5_ilosc.Value = "": UFormINFO.TextBoxRok5_wartosc.Value = "": UFormINFO.TextBoxRok6_ilosc.Value = "": UFormINFO.TextBoxRok6_wartosc.Value = ""
            UFormINFO.TextBoxRok7_ilosc.Value = "": UFormINFO.TextBoxRok7_wartosc.Value = "": UFormINFO.TextBoxRok8_ilosc.Value = "": UFormINFO.TextBoxRok8_wartosc.Value = ""
            UFormINFO.TextBoxRok9_ilosc.Value = "": UFormINFO.TextBoxRok9_wartosc.Value = "": UFormINFO.IndeksStanDzisiaj.Value = ""
        If Len(Dir$(Environ("HOMEPATH") & Application.PathSeparator & "PICTURES" & Application.PathSeparator & "JKoziAddInCHART.gif")) > 0 Then
            Kill Environ("HOMEPATH") & Application.PathSeparator & "PICTURES" & Application.PathSeparator & "JKoziAddInCHART.gif"
            UFormINFO.Indeks_ImageWykresZuzycie.Picture = Nothing
        End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ComboBoxRwCzyPZ_Change()
    If UFormINFO.ComboBoxRwCzyPZ.MatchFound = True Then
        If UFormINFO.ComboBoxRwCzyPZ.Value = "Zu¿ycie w latach" Then UFormINFO.ComboBoxKomOrg.Enabled = True Else UFormINFO.ComboBoxKomOrg.Value = "": UFormINFO.ComboBoxKomOrg.Enabled = False
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ButtonOblicz_Click()
    If UFormINFO.ComboBoxRwCzyPZ.MatchFound = True Then
        UFormINFO.MultiPage2.Pages("PageZuzycie").Caption = UFormINFO.ComboBoxRwCzyPZ.Value
        If UFormINFO.IndeksTxtBox.Value <> "" And UFormINFO.IndeksTxtBox.text <> "podaj indeks" Then
            UFormINFO.IndeksListBoxDostawcy.Height = "137,3"
            UFormINFO.IndeksListBoxZakupy.Height = "285"
            UFormINFO.ButtonIdxSzczegoly.Visible = False
            UFormINFO.TextBoxRok1_ilosc.Value = "": UFormINFO.TextBoxRok1_wartosc.Value = "": UFormINFO.TextBoxRok2_ilosc.Value = "": UFormINFO.TextBoxRok2_wartosc.Value = ""
                    UFormINFO.TextBoxRok3_ilosc.Value = "": UFormINFO.TextBoxRok3_wartosc.Value = "": UFormINFO.TextBoxRok4_ilosc.Value = "": UFormINFO.TextBoxRok4_wartosc.Value = ""
                    UFormINFO.TextBoxRok5_ilosc.Value = "": UFormINFO.TextBoxRok5_wartosc.Value = "": UFormINFO.TextBoxRok6_ilosc.Value = "": UFormINFO.TextBoxRok6_wartosc.Value = ""
                    UFormINFO.TextBoxRok7_ilosc.Value = "": UFormINFO.TextBoxRok7_wartosc.Value = "": UFormINFO.TextBoxRok8_ilosc.Value = "": UFormINFO.TextBoxRok8_wartosc.Value = ""
                    UFormINFO.TextBoxRok9_ilosc.Value = "": UFormINFO.TextBoxRok9_wartosc.Value = ""
                    If Len(Dir$(Environ("HOMEPATH") & Application.PathSeparator & "PICTURES" & Application.PathSeparator & "JKoziAddInCHART.gif")) > 0 Then
                        Kill Environ("HOMEPATH") & Application.PathSeparator & "PICTURES" & Application.PathSeparator & "JKoziAddInCHART.gif"
                        UFormINFO.Indeks_ImageWykresZuzycie.Picture = Nothing
                    End If
                Call Zuzycie_wykres_dla_Idx
        End If
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IndeksPokazTransakcjeB_Click()
    Dim IdxPozList As Long
    IdxSzukF = ""
    If FormTransIdxList.Visible = True Then Unload FormTransIdxList
        If UFormINFO.IndeksTxtBox.Value <> "" Or UFormINFO.IndeksTxtBox.Value <> "podaj indeks" Then
            IdxSzukF = UFormINFO.IndeksTxtBox.Value
            VerSzukajIdx = 1
                FormTransIdxList.Show vbModeless
        End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IndeksSzukajB_Click()
    Dim IndeksStr As String
    Dim xRok As Integer, i As Integer
    Dim WkbWykres As Workbook
    Dim WksWykres As Worksheet
    Dim ZakresWykres As Range
    Dim xChart As Object
    Dim SeriesCol_Ilosc As Object, SeriesCol_Wart As Object
    Dim SeriesLab_Ilosc As DataLabel, SeriesLab_Wart As DataLabel
    Dim myTop As Variant, myLeft As Variant, colTop1 As Variant, colTop2 As Variant
        Dim Query As IAxaptaObject
        Dim QueryDataSource As IAxaptaObject
        Dim QueryDatRange As IAxaptaObject
        Dim QueryDatRange2 As IAxaptaObject
        Dim AxaptaQueryRun As IAxaptaObject
        Dim RecorP As IAxaptaRecord
        Dim TblCollect1 As New Collection
        Dim KryteriumDokument As String, NazwaDostawcy As String, xR As String, xR1 As String, xR2 As String, xR5 As String, xR6 As String
        Dim x1 As Long, x2 As Long, xInStr As Long, xLen As Long
        Dim CzyBylo1 As Boolean
        Dim KtoPobKomOrg As String, KomOrgSingle As String
        Dim Pos As Integer, itemNr As Integer
        Dim KomOrgTbl As New Collection
    On Error GoTo error
    UFormINFO.MousePointer = fmMousePointerHourGlass
    DoEvents
    If AxApl__ Is Nothing Then loginAX
    If AxApl__ Is Nothing Then GoTo error
    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    xRok = 0: x1 = 1: x2 = 0: xInStr = 0
    If UFormINFO.IndeksTxtBox.Value <> "podaj indeks" Then
        Call data_Format
            itemNr = 1
            IndeksStr = UFormINFO.IndeksTxtBox.Value
            KtoPobKomOrg = Indeks_KtoPobieral_KomOrg(IndeksStr)
            Do While Len(KtoPobKomOrg) > 1
                Pos = InStr(1, KtoPobKomOrg, ",")
                If Pos = 0 And Len(KtoPobKomOrg) > 1 Then KomOrgSingle = Trim(KtoPobKomOrg) Else KomOrgSingle = Left(KtoPobKomOrg, (Pos - 1))
                KtoPobKomOrg = LTrim(Right(KtoPobKomOrg, (Len(KtoPobKomOrg) - Pos)))
                KomOrgTbl.Add KomOrgSingle: If Pos = 0 Then Exit Do
            Loop
            UFormINFO.ComboBoxKomOrg.Clear
            UFormINFO.ComboBoxKomOrg.AddItem ""
            For itemNr = 1 To KomOrgTbl.Count
                With UFormINFO.ComboBoxKomOrg
                    .AddItem KomOrgTbl.Item(itemNr)
                End With
            Next
            KomOrgSingle = ""
            UFormINFO.ComboBoxKomOrg.AddItem "*"

            UFormINFO.IndeksNazwa.Value = NazwaTowaru(IndeksStr)
            UFormINFO.IndeksJM.Value = JednMiary(IndeksStr)
            UFormINFO.IndeksGrMat.Value = GrMaterial(IndeksStr)
            UFormINFO.IndeksNazwaGrMat.Value = NazwaGrMaterial(CStr(UFormINFO.IndeksGrMat.text))
            UFormINFO.IndeksAlias.Value = Indeks_Alias_Nazwa(IndeksStr)
            UFormINFO.IndeksDataUtw.Value = Format(CDate(Indeks_Data_Utworzenia(IndeksStr)), SystemShortDateFormat)
            UFormINFO.IndeksKtoUtw.Value = Indeks_Kto_Utworzyl(IndeksStr)
            UFormINFO.IndeksDataMod.Value = Format(CDate(Indeks_Data_modyfikacji(IndeksStr)), SystemShortDateFormat)
            UFormINFO.IndeksKtoMod.Value = Indeks_Kto_Modyfikowal(IndeksStr)
            UFormINFO.IndeksStanMIN.Value = Indeks_StanyAwaryjne_Min(IndeksStr)
            UFormINFO.IndeksStanMAX.Value = Indeks_StanyAwaryjne_Max(IndeksStr)
            UFormINFO.IndeksStanDzisiaj.Value = StanNaDzienIndeks(IndeksStr, Now())

            Set Query = AxApl__.CreateObject("Query")
            Set QueryDataSource = Query.Call("addDataSource", 177)
            Set QueryDatRange = QueryDataSource.Call("addRange", 1)
                QueryDatRange.Call "Value", IndeksStr
                Set QueryDatRange2 = QueryDataSource.Call("addRange", 9)
                    KryteriumDokument = "*ZZ*"
                    QueryDatRange2.Call "Value", KryteriumDokument

            Set AxaptaQueryRun = AxApl__.CreateObject("QueryRun", Query)
                UFormINFO.IndeksListBoxDostawcy.Clear
                UFormINFO.IndeksListBoxDostawcy.ColumnWidths = "30;45;285;75;75;75;75;1;1"
                UFormINFO.IndeksListBoxDostawcy.Height = "137,3"
                UFormINFO.IndeksListBoxZakupy.Height = "285"
                While AxaptaQueryRun.Call("Next")
                    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                    Set RecorP = AxaptaQueryRun.Call("GetNo", 1)
                        CzyBylo1 = czyWkolekcji(CStr(RecorP.field(57)), TblCollect1)
                        If CzyBylo1 = False Then
                            TblCollect1.Add CStr(RecorP.field(57))
                            xInStr = InStr(1, CStr(RecorP.TooltipField(57)), ",")
                            xLen = Len(RecorP.TooltipField(57))
                            NazwaDostawcy = Trim(Right(CStr(RecorP.TooltipField(57)), CLng(xLen - xInStr)))
                            With UFormINFO.IndeksListBoxDostawcy
                                .AddItem x1
                                .List(x2, 1) = RecorP.field(57)
                                .List(x2, 2) = NazwaDostawcy
                                .List(x2, 3) = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                        SumaTransakcjiDlaIdx(IndeksStr, 6, "*PZ/*", , , CStr(RecorP.field(57))), _
                                                                        SumaTransakcjiDlaIdx(IndeksStr, 24, "*PZ/*", , , CStr(RecorP.field(57)))))), "#,##0.00")
                                .List(x2, 4) = SumaTransakcjiDlaIdx(IndeksStr, 5, "*PZ/*", , , CStr(RecorP.field(57)))
                                .List(x2, 5) = Format(CDate(RecorP.field(4)), SystemShortDateFormat)
                                .List(x2, 6) = Format(CDate(DataFizOstZaksiegowanejDlaIndeksuKomOrg(IndeksStr, , "*PZ/*", , CStr(RecorP.field(57)))), SystemShortDateFormat)
                            End With
                            x1 = x1 + 1: x2 = x2 + 1: xInStr = 0
                        ElseIf CzyBylo1 = True Then
                        End If
                    DoEvents
                    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                Wend
                xR = "--": xR1 = "----------------------"
                xR2 = "---------------------------------------------------------------------------------------------------------------------------"
                xR3 = "----------------------------": xR4 = "--------------------------": xR5 = "-------------------------": xR6 = "-------------------------"
                For x1 = 1 To 2
                    With UFormINFO.IndeksListBoxDostawcy
                        .AddItem xR
                        .List(x2, 1) = xR1
                        .List(x2, 2) = xR2
                        .List(x2, 3) = xR3
                        .List(x2, 4) = xR4
                        .List(x2, 5) = xR5
                        .List(x2, 6) = xR6
                    End With
                    xR = "R": xR1 = "": xR2 = "ZAKUPY £¥CZNIE: "
                    xR3 = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 6, "*PZ/*"), _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 24, "*PZ/*")))), "#,##0.00")
                    xR4 = SumaTransakcjiDlaIdx(IndeksStr, 5, "*PZ/*")
                    xR5 = "": xR6 = "": x2 = x2 + 1
                    DoEvents
                    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                Next
            
    End If
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    UFormINFO.IndeksListBoxDostawcy.Height = "137,3"
    UFormINFO.IndeksListBoxZakupy.Height = "285"
error:
    IndeksStr = ""
    UFormINFO.MousePointer = fmMousePointerDefault
    DoEvents
    Application.Calculation = xlCalculationAutomatic
    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
    Exit Sub
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Sub Zuzycie_wykres_dla_Idx()
    Dim IndeksStr As String
    Dim xRok As Integer, i As Integer
    Dim WkbWykres As Workbook
    Dim WksWykres As Worksheet
    Dim ZakresWykres As Range
    Dim xChart As Object
    Dim SeriesCol_Ilosc As Object, SeriesCol_Wart As Object
    Dim SeriesLab_Ilosc As DataLabel, SeriesLab_Wart As DataLabel
    Dim myTop As Variant, myLeft As Variant, colTop1 As Variant, colTop2 As Variant
        Dim KtoPobKomOrg As String, KomOrgSingle As String
        Dim Pos As Integer, itemNr As Integer
        Dim KomOrgTbl As New Collection
    On Error GoTo error
    UFormINFO.MousePointer = fmMousePointerHourGlass
    DoEvents
    If AxApl__ Is Nothing Then loginAX
    If AxApl__ Is Nothing Then GoTo error
    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    xRok = 0: x2 = 0: xInStr = 0
    If UFormINFO.IndeksTxtBox.Value <> "podaj indeks" Then
        Call data_Format
        IndeksStr = UFormINFO.IndeksTxtBox.Value
        
        'ZUZYCIE W LATACH
            xRok = Year(CDate(Now()))
            UFormINFO.LabelRok9.Caption = xRok: UFormINFO.LabelRok8.Caption = xRok - 1: UFormINFO.LabelRok7.Caption = xRok - 2: UFormINFO.LabelRok6.Caption = xRok - 3
            UFormINFO.LabelRok5.Caption = xRok - 4: UFormINFO.LabelRok4.Caption = xRok - 5: UFormINFO.LabelRok3.Caption = xRok - 6: UFormINFO.LabelRok2.Caption = xRok - 7
            UFormINFO.LabelRok1.Caption = xRok - 8
                If UFormINFO.ComboBoxRwCzyPZ = "Zu¿ycie w latach" Then
                If UFormINFO.ComboBoxKomOrg.Value = "" Then KomOrgSingle = "*" Else KomOrgSingle = UFormINFO.ComboBoxKomOrg.Value
                    If Indeks_Czy_RWPNU(IndeksStr) = "RW" Then
                        UFormINFO.TextBoxRok1_ilosc.Value = Abs(ZuzycieKomOrgIlosc(IndeksStr, DateSerial(CInt(LabelRok1), 1, 1), DateSerial(CInt(LabelRok1), 12, 31), KomOrgSingle))
                        UFormINFO.TextBoxRok1_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 6, DateSerial(CInt(LabelRok1), 1, 1), DateSerial(CInt(LabelRok1), 12, 31), KomOrgSingle), _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 24, DateSerial(CInt(LabelRok1), 1, 1), DateSerial(CInt(LabelRok1), 12, 31), KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok2_ilosc.Value = Abs(ZuzycieKomOrgIlosc(IndeksStr, DateSerial(CInt(LabelRok2), 1, 1), DateSerial(CInt(LabelRok2), 12, 31), KomOrgSingle))
                        UFormINFO.TextBoxRok2_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 6, DateSerial(CInt(LabelRok2), 1, 1), DateSerial(CInt(LabelRok2), 12, 31), KomOrgSingle), _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 24, DateSerial(CInt(LabelRok2), 1, 1), DateSerial(CInt(LabelRok2), 12, 31), KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok3_ilosc.Value = Abs(ZuzycieKomOrgIlosc(IndeksStr, DateSerial(CInt(LabelRok3), 1, 1), DateSerial(CInt(LabelRok3), 12, 31), KomOrgSingle))
                        UFormINFO.TextBoxRok3_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 6, DateSerial(CInt(LabelRok3), 1, 1), DateSerial(CInt(LabelRok3), 12, 31), KomOrgSingle), _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 24, DateSerial(CInt(LabelRok3), 1, 1), DateSerial(CInt(LabelRok3), 12, 31), KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok4_ilosc.Value = Abs(ZuzycieKomOrgIlosc(IndeksStr, DateSerial(CInt(LabelRok4), 1, 1), DateSerial(CInt(LabelRok4), 12, 31), KomOrgSingle))
                        UFormINFO.TextBoxRok4_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 6, DateSerial(CInt(LabelRok4), 1, 1), DateSerial(CInt(LabelRok4), 12, 31), KomOrgSingle), _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 24, DateSerial(CInt(LabelRok4), 1, 1), DateSerial(CInt(LabelRok4), 12, 31), KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok5_ilosc.Value = Abs(ZuzycieKomOrgIlosc(IndeksStr, DateSerial(CInt(LabelRok5), 1, 1), DateSerial(CInt(LabelRok5), 12, 31), KomOrgSingle))
                        UFormINFO.TextBoxRok5_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 6, DateSerial(CInt(LabelRok5), 1, 1), DateSerial(CInt(LabelRok5), 12, 31), KomOrgSingle), _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 24, DateSerial(CInt(LabelRok5), 1, 1), DateSerial(CInt(LabelRok5), 12, 31), KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok6_ilosc.Value = Abs(ZuzycieKomOrgIlosc(IndeksStr, DateSerial(CInt(LabelRok6), 1, 1), DateSerial(CInt(LabelRok6), 12, 31), KomOrgSingle))
                        UFormINFO.TextBoxRok6_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 6, DateSerial(CInt(LabelRok6), 1, 1), DateSerial(CInt(LabelRok6), 12, 31), KomOrgSingle), _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 24, DateSerial(CInt(LabelRok6), 1, 1), DateSerial(CInt(LabelRok6), 12, 31), KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok7_ilosc.Value = Abs(ZuzycieKomOrgIlosc(IndeksStr, DateSerial(CInt(LabelRok7), 1, 1), DateSerial(CInt(LabelRok7), 12, 31), KomOrgSingle))
                        UFormINFO.TextBoxRok7_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 6, DateSerial(CInt(LabelRok7), 1, 1), DateSerial(CInt(LabelRok7), 12, 31), KomOrgSingle), _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 24, DateSerial(CInt(LabelRok7), 1, 1), DateSerial(CInt(LabelRok7), 12, 31), KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok8_ilosc.Value = Abs(ZuzycieKomOrgIlosc(IndeksStr, DateSerial(CInt(LabelRok8), 1, 1), DateSerial(CInt(LabelRok8), 12, 31), KomOrgSingle))
                        UFormINFO.TextBoxRok8_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 6, DateSerial(CInt(LabelRok8), 1, 1), DateSerial(CInt(LabelRok8), 12, 31), KomOrgSingle), _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 24, DateSerial(CInt(LabelRok8), 1, 1), DateSerial(CInt(LabelRok8), 12, 31), KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok9_ilosc.Value = Abs(ZuzycieKomOrgIlosc(IndeksStr, DateSerial(CInt(LabelRok9), 1, 1), DateSerial(CInt(LabelRok9), 12, 31), KomOrgSingle))
                        UFormINFO.TextBoxRok9_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 6, DateSerial(CInt(LabelRok9), 1, 1), DateSerial(CInt(LabelRok9), 12, 31), KomOrgSingle), _
                                                                    ZuzycieKomOrgWartosc(IndeksStr, 24, DateSerial(CInt(LabelRok9), 1, 1), DateSerial(CInt(LabelRok9), 12, 31), KomOrgSingle)))), "#,##0.00")
                    ElseIf Indeks_Czy_RWPNU(IndeksStr) = "RWPNU" Then
                        UFormINFO.TextBoxRok1_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*RWPNU*", DateSerial(CInt(LabelRok1), 1, 1), DateSerial(CInt(LabelRok1), 12, 31), , KomOrgSingle))
                        UFormINFO.TextBoxRok1_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 6, "*RWPNU*", DateSerial(CInt(LabelRok1), 1, 1), DateSerial(CInt(LabelRok1), 12, 31), , KomOrgSingle), _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 24, "*RWPNU*", DateSerial(CInt(LabelRok1), 1, 1), DateSerial(CInt(LabelRok1), 12, 31), , KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok2_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*RWPNU*", DateSerial(CInt(LabelRok2), 1, 1), DateSerial(CInt(LabelRok2), 12, 31), , KomOrgSingle))
                        UFormINFO.TextBoxRok2_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 6, "*RWPNU*", DateSerial(CInt(LabelRok2), 1, 1), DateSerial(CInt(LabelRok2), 12, 31), , KomOrgSingle), _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 24, "*RWPNU*", DateSerial(CInt(LabelRok2), 1, 1), DateSerial(CInt(LabelRok2), 12, 31), , KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok3_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*RWPNU*", DateSerial(CInt(LabelRok3), 1, 1), DateSerial(CInt(LabelRok3), 12, 31), , KomOrgSingle))
                        UFormINFO.TextBoxRok3_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 6, "*RWPNU*", DateSerial(CInt(LabelRok3), 1, 1), DateSerial(CInt(LabelRok3), 12, 31), , KomOrgSingle), _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 24, "*RWPNU*", DateSerial(CInt(LabelRok3), 1, 1), DateSerial(CInt(LabelRok3), 12, 31), , KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok4_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*RWPNU*", DateSerial(CInt(LabelRok4), 1, 1), DateSerial(CInt(LabelRok4), 12, 31), , KomOrgSingle))
                        UFormINFO.TextBoxRok4_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 6, "*RWPNU*", DateSerial(CInt(LabelRok4), 1, 1), DateSerial(CInt(LabelRok4), 12, 31), , KomOrgSingle), _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 24, "*RWPNU*", DateSerial(CInt(LabelRok4), 1, 1), DateSerial(CInt(LabelRok4), 12, 31), , KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok5_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*RWPNU*", DateSerial(CInt(LabelRok5), 1, 1), DateSerial(CInt(LabelRok5), 12, 31), , KomOrgSingle))
                        UFormINFO.TextBoxRok5_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 6, "*RWPNU*", DateSerial(CInt(LabelRok5), 1, 1), DateSerial(CInt(LabelRok5), 12, 31), , KomOrgSingle), _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 24, "*RWPNU*", DateSerial(CInt(LabelRok5), 1, 1), DateSerial(CInt(LabelRok5), 12, 31), , KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok6_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*RWPNU*", DateSerial(CInt(LabelRok6), 1, 1), DateSerial(CInt(LabelRok6), 12, 31), , KomOrgSingle))
                        UFormINFO.TextBoxRok6_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 6, "*RWPNU*", DateSerial(CInt(LabelRok6), 1, 1), DateSerial(CInt(LabelRok6), 12, 31), , KomOrgSingle), _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 24, "*RWPNU*", DateSerial(CInt(LabelRok6), 1, 1), DateSerial(CInt(LabelRok6), 12, 31), , KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok7_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*RWPNU*", DateSerial(CInt(LabelRok7), 1, 1), DateSerial(CInt(LabelRok7), 12, 31), , KomOrgSingle))
                        UFormINFO.TextBoxRok7_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 6, "*RWPNU*", DateSerial(CInt(LabelRok7), 1, 1), DateSerial(CInt(LabelRok7), 12, 31), , KomOrgSingle), _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 24, "*RWPNU*", DateSerial(CInt(LabelRok7), 1, 1), DateSerial(CInt(LabelRok7), 12, 31), , KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok8_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*RWPNU*", DateSerial(CInt(LabelRok8), 1, 1), DateSerial(CInt(LabelRok8), 12, 31), , KomOrgSingle))
                        UFormINFO.TextBoxRok8_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 6, "*RWPNU*", DateSerial(CInt(LabelRok8), 1, 1), DateSerial(CInt(LabelRok8), 12, 31), , KomOrgSingle), _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 24, "*RWPNU*", DateSerial(CInt(LabelRok8), 1, 1), DateSerial(CInt(LabelRok8), 12, 31), , KomOrgSingle)))), "#,##0.00")
                        UFormINFO.TextBoxRok9_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*RWPNU*", DateSerial(CInt(LabelRok9), 1, 1), DateSerial(CInt(LabelRok9), 12, 31), , KomOrgSingle))
                        UFormINFO.TextBoxRok9_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 6, "*RWPNU*", DateSerial(CInt(LabelRok9), 1, 1), DateSerial(CInt(LabelRok9), 12, 31), , KomOrgSingle), _
                                                                    SumaTransakcjiDlaIdx(IndeksStr, 24, "*RWPNU*", DateSerial(CInt(LabelRok9), 1, 1), DateSerial(CInt(LabelRok9), 12, 31), , KomOrgSingle)))), "#,##0.00")
                    Else
                        UFormINFO.TextBoxRok1_ilosc.Value = "!!!"
                        UFormINFO.TextBoxRok1_wartosc.Value = "!!!"
                        UFormINFO.TextBoxRok2_ilosc.Value = "!!!"
                        UFormINFO.TextBoxRok2_wartosc.Value = "!!!"
                        UFormINFO.TextBoxRok3_ilosc.Value = "!!!"
                        UFormINFO.TextBoxRok3_wartosc.Value = "!!!"
                        UFormINFO.TextBoxRok4_ilosc.Value = "!!!"
                        UFormINFO.TextBoxRok4_wartosc.Value = "!!!"
                        UFormINFO.TextBoxRok5_ilosc.Value = "!!!"
                        UFormINFO.TextBoxRok5_wartosc.Value = "!!!"
                        UFormINFO.TextBoxRok6_ilosc.Value = "!!!"
                        UFormINFO.TextBoxRok6_wartosc.Value = "!!!"
                        UFormINFO.TextBoxRok7_ilosc.Value = "!!!"
                        UFormINFO.TextBoxRok7_wartosc.Value = "!!!"
                        UFormINFO.TextBoxRok8_ilosc.Value = "!!!"
                        UFormINFO.TextBoxRok8_wartosc.Value = "!!!"
                        UFormINFO.TextBoxRok9_ilosc.Value = "!!!"
                        UFormINFO.TextBoxRok9_wartosc.Value = "!!!"
                        Application.Calculation = xlCalculationAutomatic
                        GoTo error
                    End If
                ElseIf UFormINFO.ComboBoxRwCzyPZ = "Zakupy w latach" Then
                    UFormINFO.TextBoxRok1_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*PZ/*", DateSerial(CInt(LabelRok1), 1, 1), DateSerial(CInt(LabelRok1), 12, 31)))
                    UFormINFO.TextBoxRok1_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 6, "*PZ/*", DateSerial(CInt(LabelRok1), 1, 1), DateSerial(CInt(LabelRok1), 12, 31)), _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 24, "*PZ/*", DateSerial(CInt(LabelRok1), 1, 1), DateSerial(CInt(LabelRok1), 12, 31))))), "#,##0.00")
                    UFormINFO.TextBoxRok2_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*PZ/*", DateSerial(CInt(LabelRok2), 1, 1), DateSerial(CInt(LabelRok2), 12, 31)))
                    UFormINFO.TextBoxRok2_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 6, "*PZ/*", DateSerial(CInt(LabelRok2), 1, 1), DateSerial(CInt(LabelRok2), 12, 31)), _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 24, "*PZ/*", DateSerial(CInt(LabelRok2), 1, 1), DateSerial(CInt(LabelRok2), 12, 31))))), "#,##0.00")
                    UFormINFO.TextBoxRok3_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*PZ/*", DateSerial(CInt(LabelRok3), 1, 1), DateSerial(CInt(LabelRok3), 12, 31)))
                    UFormINFO.TextBoxRok3_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 6, "*PZ/*", DateSerial(CInt(LabelRok3), 1, 1), DateSerial(CInt(LabelRok3), 12, 31)), _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 24, "*PZ/*", DateSerial(CInt(LabelRok3), 1, 1), DateSerial(CInt(LabelRok3), 12, 31))))), "#,##0.00")
                    UFormINFO.TextBoxRok4_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*PZ/*", DateSerial(CInt(LabelRok4), 1, 1), DateSerial(CInt(LabelRok4), 12, 31)))
                    UFormINFO.TextBoxRok4_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 6, "*PZ/*", DateSerial(CInt(LabelRok4), 1, 1), DateSerial(CInt(LabelRok4), 12, 31)), _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 24, "*PZ/*", DateSerial(CInt(LabelRok4), 1, 1), DateSerial(CInt(LabelRok4), 12, 31))))), "#,##0.00")
                    UFormINFO.TextBoxRok5_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*PZ/*", DateSerial(CInt(LabelRok5), 1, 1), DateSerial(CInt(LabelRok5), 12, 31)))
                    UFormINFO.TextBoxRok5_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 6, "*PZ/*", DateSerial(CInt(LabelRok5), 1, 1), DateSerial(CInt(LabelRok5), 12, 31)), _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 24, "*PZ/*", DateSerial(CInt(LabelRok5), 1, 1), DateSerial(CInt(LabelRok5), 12, 31))))), "#,##0.00")
                    UFormINFO.TextBoxRok6_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*PZ/*", DateSerial(CInt(LabelRok6), 1, 1), DateSerial(CInt(LabelRok6), 12, 31)))
                    UFormINFO.TextBoxRok6_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 6, "*PZ/*", DateSerial(CInt(LabelRok6), 1, 1), DateSerial(CInt(LabelRok6), 12, 31)), _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 24, "*PZ/*", DateSerial(CInt(LabelRok6), 1, 1), DateSerial(CInt(LabelRok6), 12, 31))))), "#,##0.00")
                    UFormINFO.TextBoxRok7_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*PZ/*", DateSerial(CInt(LabelRok7), 1, 1), DateSerial(CInt(LabelRok7), 12, 31)))
                    UFormINFO.TextBoxRok7_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 6, "*PZ/*", DateSerial(CInt(LabelRok7), 1, 1), DateSerial(CInt(LabelRok7), 12, 31)), _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 24, "*PZ/*", DateSerial(CInt(LabelRok7), 1, 1), DateSerial(CInt(LabelRok7), 12, 31))))), "#,##0.00")
                    UFormINFO.TextBoxRok8_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*PZ/*", DateSerial(CInt(LabelRok8), 1, 1), DateSerial(CInt(LabelRok8), 12, 31)))
                    UFormINFO.TextBoxRok8_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 6, "*PZ/*", DateSerial(CInt(LabelRok8), 1, 1), DateSerial(CInt(LabelRok8), 12, 31)), _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 24, "*PZ/*", DateSerial(CInt(LabelRok8), 1, 1), DateSerial(CInt(LabelRok8), 12, 31))))), "#,##0.00")
                    UFormINFO.TextBoxRok9_ilosc.Value = Abs(SumaTransakcjiDlaIdx(IndeksStr, 5, "*PZ/*", DateSerial(CInt(LabelRok9), 1, 1), DateSerial(CInt(LabelRok9), 12, 31)))
                    UFormINFO.TextBoxRok9_wartosc.Value = Format(Abs(CDbl(Application.WorksheetFunction.Sum( _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 6, "*PZ/*", DateSerial(CInt(LabelRok9), 1, 1), DateSerial(CInt(LabelRok9), 12, 31)), _
                                                                SumaTransakcjiDlaIdx(IndeksStr, 24, "*PZ/*", DateSerial(CInt(LabelRok9), 1, 1), DateSerial(CInt(LabelRok9), 12, 31))))), "#,##0.00")
                Else
                    UFormINFO.TextBoxRok1_ilosc.Value = "!!!"
                    UFormINFO.TextBoxRok1_wartosc.Value = "!!!"
                    UFormINFO.TextBoxRok2_ilosc.Value = "!!!"
                    UFormINFO.TextBoxRok2_wartosc.Value = "!!!"
                    UFormINFO.TextBoxRok3_ilosc.Value = "!!!"
                    UFormINFO.TextBoxRok3_wartosc.Value = "!!!"
                    UFormINFO.TextBoxRok4_ilosc.Value = "!!!"
                    UFormINFO.TextBoxRok4_wartosc.Value = "!!!"
                    UFormINFO.TextBoxRok5_ilosc.Value = "!!!"
                    UFormINFO.TextBoxRok5_wartosc.Value = "!!!"
                    UFormINFO.TextBoxRok6_ilosc.Value = "!!!"
                    UFormINFO.TextBoxRok6_wartosc.Value = "!!!"
                    UFormINFO.TextBoxRok7_ilosc.Value = "!!!"
                    UFormINFO.TextBoxRok7_wartosc.Value = "!!!"
                    UFormINFO.TextBoxRok8_ilosc.Value = "!!!"
                    UFormINFO.TextBoxRok8_wartosc.Value = "!!!"
                    UFormINFO.TextBoxRok9_ilosc.Value = "!!!"
                    UFormINFO.TextBoxRok9_wartosc.Value = "!!!"
                    GoTo error
                End If
            'WYKRES
                Set WkbWykres = Application.Workbooks("Ribbon2.xlam")
                Set WksWykres = WkbWykres.Worksheets("WykresVBA")
                    If WksWykres.ChartObjects.Count > 0 Then WksWykres.ChartObjects.Delete
                Set SeriesCol_Ilosc = Nothing
                Set SeriesCol_Wart = Nothing
                    Set SeriesLab_Ilosc = Nothing
                    Set SeriesLab_Wart = Nothing
                        WksWykres.Cells(2, 1) = "iloæ": WksWykres.Cells(3, 1) = "wartoæ"
                        WksWykres.Cells(1, 2) = UFormINFO.LabelRok1.Caption: WksWykres.Cells(2, 2) = UFormINFO.TextBoxRok1_ilosc.Value: WksWykres.Cells(3, 2) = CDbl(UFormINFO.TextBoxRok1_wartosc.Value)
                        WksWykres.Cells(1, 3) = UFormINFO.LabelRok2.Caption: WksWykres.Cells(2, 3) = UFormINFO.TextBoxRok2_ilosc.Value: WksWykres.Cells(3, 3) = CDbl(UFormINFO.TextBoxRok2_wartosc.Value)
                        WksWykres.Cells(1, 4) = UFormINFO.LabelRok3.Caption: WksWykres.Cells(2, 4) = UFormINFO.TextBoxRok3_ilosc.Value: WksWykres.Cells(3, 4) = CDbl(UFormINFO.TextBoxRok3_wartosc.Value)
                        WksWykres.Cells(1, 5) = UFormINFO.LabelRok4.Caption: WksWykres.Cells(2, 5) = UFormINFO.TextBoxRok4_ilosc.Value: WksWykres.Cells(3, 5) = CDbl(UFormINFO.TextBoxRok4_wartosc.Value)
                        WksWykres.Cells(1, 6) = UFormINFO.LabelRok5.Caption: WksWykres.Cells(2, 6) = UFormINFO.TextBoxRok5_ilosc.Value: WksWykres.Cells(3, 6) = CDbl(UFormINFO.TextBoxRok5_wartosc.Value)
                        WksWykres.Cells(1, 7) = UFormINFO.LabelRok6.Caption: WksWykres.Cells(2, 7) = UFormINFO.TextBoxRok6_ilosc.Value: WksWykres.Cells(3, 7) = CDbl(UFormINFO.TextBoxRok6_wartosc.Value)
                        WksWykres.Cells(1, 8) = UFormINFO.LabelRok7.Caption: WksWykres.Cells(2, 8) = UFormINFO.TextBoxRok7_ilosc.Value: WksWykres.Cells(3, 8) = CDbl(UFormINFO.TextBoxRok7_wartosc.Value)
                        WksWykres.Cells(1, 9) = UFormINFO.LabelRok8.Caption: WksWykres.Cells(2, 9) = UFormINFO.TextBoxRok8_ilosc.Value: WksWykres.Cells(3, 9) = CDbl(UFormINFO.TextBoxRok8_wartosc.Value)
                        WksWykres.Cells(1, 10) = UFormINFO.LabelRok9.Caption: WksWykres.Cells(2, 10) = UFormINFO.TextBoxRok9_ilosc.Value: WksWykres.Cells(3, 10) = CDbl(UFormINFO.TextBoxRok9_wartosc.Value)
                        WksWykres.Range(WksWykres.Cells(3, 2), WksWykres.Cells(3, 10)).NumberFormat = "# ##0.00 z³"
                    Set ZakresWykres = WksWykres.Range(WksWykres.Cells(1, 1), WksWykres.Cells(WksWykres.Cells(Rows.Count, 1).End(xlUp).Row, WksWykres.Cells(1, Columns.Count).End(xlToLeft).Column))
                    Set xChart = WksWykres.Shapes.AddChart
                        xChart.Chart.SetSourceData Source:=ZakresWykres
                        xChart.Chart.ChartType = xlColumnClustered
                        xChart.Chart.SeriesCollection(1).ChartType = xlLineMarkers
                        xChart.Chart.SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
                        xChart.Chart.SeriesCollection(1).AxisGroup = 2
                            xChart.Chart.SeriesCollection(1).ApplyDataLabels
                            xChart.Chart.SeriesCollection(2).ApplyDataLabels
                            xChart.Chart.SeriesCollection(2).DataLabels.Position = xlLabelPositionInsideEnd
                                On Error Resume Next
                                    With xChart.Chart
                                            For i = 1 To 9
                                                Set SeriesCol_Wart = .SeriesCollection(2).Points(i)
                                                    colTop2 = Round(SeriesCol_Wart.Top, 10)
                                                Set SeriesCol_Ilosc = .SeriesCollection(1).Points(i)
                                                    colTop1 = Round(SeriesCol_Ilosc.Top, 10)
                                                    Set SeriesLab_Wart = .SeriesCollection(2).Points(i).DataLabel
                                                    Set SeriesLab_Ilosc = .SeriesCollection(1).Points(i).DataLabel
                                                If Abs((Abs(colTop1) - Abs(colTop2))) < 15 Then
                                                    If (Abs(colTop1) < Abs(colTop2)) Then
                                                        SeriesLab_Ilosc.Position = xlLabelPositionAbove
                                                        SeriesLab_Ilosc.HorizontalAlignment = xlCenter
                                                    Else
                                                        myTop = Round(SeriesLab_Wart.Top - 20, 10)
                                                        myLeft = Round(SeriesCol_Ilosc.Left, 10)
                                                        SeriesLab_Ilosc.Top = myTop
                                                        SeriesLab_Ilosc.Left = myLeft
                                                    End If
                                                Else
                                                    If (Abs(colTop1) < Abs(colTop2)) Then
                                                        SeriesLab_Ilosc.Position = xlLabelPositionAbove
                                                        SeriesLab_Ilosc.HorizontalAlignment = xlCenter
                                                    Else
                                                        SeriesLab_Ilosc.Position = xlLabelPositionBelow
                                                        SeriesLab_Ilosc.HorizontalAlignment = xlCenter
                                                    End If
                                                End If
                                                Set SeriesCol_Ilosc = Nothing
                                                Set SeriesCol_Wart = Nothing
                                                    Set SeriesLab_Ilosc = Nothing
                                                    Set SeriesLab_Wart = Nothing
                                            Next i
                                    End With
                        xChart.Chart.Legend.Position = xlLegendPositionBottom
                        xChart.Chart.Parent.Height = UFormINFO.Indeks_ImageWykresZuzycie.Height
                        xChart.Chart.Parent.Width = UFormINFO.Indeks_ImageWykresZuzycie.Width
                            Fname = Environ("HOMEPATH") & Application.PathSeparator & "PICTURES" & Application.PathSeparator & "JKoziAddInCHART.gif"
                            xChart.Chart.Export Filename:=Fname, FilterName:="GIF"
                                UFormINFO.Indeks_ImageWykresZuzycie.Picture = LoadPicture(Fname)
                                    xChart.Delete
                                    Set WksWykres = ThisWorkbook.Worksheets("WykresVBA")
                                        WksWykres.Range(WksWykres.Cells(1, 1), WksWykres.Cells(10, 20)).Value = ""
                                        If WksWykres.ChartObjects.Count > 0 Then WksWykres.ChartObjects.Delete
                                            DoEvents
                                        If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    UFormINFO.IndeksListBoxDostawcy.Height = "137,3"
    UFormINFO.IndeksListBoxZakupy.Height = "285"
error:
    IndeksStr = ""
    UFormINFO.MousePointer = fmMousePointerDefault
    DoEvents
    Application.Calculation = xlCalculationAutomatic
    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
    Exit Sub
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IndeksListBoxDostawcy_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim IndeksStr As String
    Dim xNrDostawcy, NrZZ_Pelny, KtoZatwierdzal, UserNameFull, xLpChck As String
    Dim i As Integer
        Dim Query As IAxaptaObject
        Dim QueryDataSource As IAxaptaObject
        Dim QueryDatRange As IAxaptaObject
        Dim QueryDatRange2 As IAxaptaObject
        Dim QueryDatRange3 As IAxaptaObject
        Dim AxaptaQueryRun As IAxaptaObject
        Dim RecorP As IAxaptaRecord
        Dim KryteriumDokument As String
        Dim x1, x2, xInStr, xLen As Long
    On Error GoTo error
    If AxApl__ Is Nothing Then loginAX
    If AxApl__ Is Nothing Then GoTo error
    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    x1 = 1: x2 = 0: xInStr = 0
    IndeksStr = UFormINFO.IndeksTxtBox.Value
    xNrDostawcy = UFormINFO.IndeksListBoxDostawcy.List(UFormINFO.IndeksListBoxDostawcy.ListIndex, 1)
    xLpChck = UFormINFO.IndeksListBoxDostawcy.List(UFormINFO.IndeksListBoxDostawcy.ListIndex, 0)
    If xLpChck = "--" Then
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    ElseIf xLpChck = "R" Then
        xNrDostawcy = ""
    End If
    Set Query = AxApl__.CreateObject("Query")
            Set QueryDataSource = Query.Call("addDataSource", 177)
            Set QueryDatRange = QueryDataSource.Call("addRange", 1)
                QueryDatRange.Call "Value", IndeksStr
                Set QueryDatRange2 = QueryDataSource.Call("addRange", 9)
                    KryteriumDokument = "*ZZ*"
                    QueryDatRange2.Call "Value", KryteriumDokument
                    Set QueryDatRange3 = QueryDataSource.Call("addRange", 57)
                        QueryDatRange3.Call "Value", xNrDostawcy
            Set AxaptaQueryRun = AxApl__.CreateObject("QueryRun", Query)
                UFormINFO.IndeksListBoxZakupy.Clear
                UFormINFO.IndeksListBoxZakupy.ColumnWidths = "35;80;80;60;110;50;50"
                While AxaptaQueryRun.Call("Next")
                    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                    Set RecorP = AxaptaQueryRun.Call("GetNo", 1)
                            xInStr = InStr(1, CStr(RecorP.TooltipField(57)), ",")
                            xLen = Len(RecorP.TooltipField(57))
                            NazwaDostawcy = Trim(Right(CStr(RecorP.TooltipField(57)), CLng(xLen - xInStr)))
                            With UFormINFO.IndeksListBoxZakupy
                                .AddItem x1
                                .List(x2, 1) = RecorP.field(9)
                                .List(x2, 2) = RecorP.field(22)
                                .List(x2, 3) = Format(CDate(RecorP.field(4)), SystemShortDateFormat)
                                NrZZ_Pelny = RecorP.field(9)
                                If NrZZ_Pelny <> "" Then
                                    KtoZatwierdzal = Purch_StanZatwierdzen(NrZZ_Pelny)
                                    If Len(KtoZatwierdzal) <> 0 And KtoZatwierdzal <> "B³¹d!!!" Then UserNameFull = UserId2UserFullName(CStr(KtoZatwierdzal))
                                    If Len(UserNameFull) = 0 And Len(KtoZatwierdzal) <> 0 Then
                                        .List(x2, 4) = KtoZatwierdzal
                                    Else
                                        .List(x2, 4) = UserNameFull
                                    End If
                                Else
                                    .List(x2, 4) = ""
                                End If
                                .List(x2, 5) = Format(Abs(CDbl(Application.WorksheetFunction.Sum(RecorP.field(6), RecorP.field(24)))))
                                .List(x2, 6) = StanNaDzienIndeks(IndeksStr, Format(CDate(RecorP.field(61444)) - 1, SystemShortDateFormat))
                            End With
                            x1 = x1 + 1: x2 = x2 + 1: xInStr = 0
                    DoEvents
                    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
                Wend
error:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If UFormINFO.Visible = False Then UFormINFO.Show vbModeless
    Exit Sub
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IndeksListBoxDostawcy_AfterUpdate()
    UFormINFO.IndeksListBoxDostawcy.Height = "137,3"
    UFormINFO.IndeksListBoxZakupy.Height = "285"
End Sub
Private Sub IndeksListBoxZakupy_Click()
    UFormINFO.IndeksListBoxDostawcy.Height = "137,3"
    UFormINFO.IndeksListBoxZakupy.Height = "285"
End Sub
Private Sub IndeksListBoxZakupy_AfterUpdate()
    UFormINFO.IndeksListBoxZakupy.Height = "285"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IndeksSzukajNET_Click()
    UFormINFO.MousePointer = fmMousePointerHourGlass
    DoEvents
        If UFormINFO.IndeksNazwa.Value <> "" Then
            TekstSearchGoogle = UFormINFO.IndeksNazwa.Value
            Call OpenUrl
        Else
            MsgBox "Pole ""NAZWA MATERIA£U"" jest puste !!!"
        End If
    UFormINFO.MousePointer = fmMousePointerDefault
    DoEvents
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IndeksMinimalizuj_Click()
    If UFormINFO.IndeksMinimalizuj.Value = True Then
        UFormINFO.Height = "96,75"
        UFormINFO.Top = Application.Top + Application.Height - 40 - UFormINFO.Height
        UFormINFO.Left = Application.Left + Application.Width - 30 - UFormINFO.Width
    ElseIf UFormINFO.IndeksMinimalizuj.Value = False Then
        UFormINFO.Height = "465"
        UFormINFO.Top = Application.Top + Application.Height - 40 - UFormINFO.Height
        UFormINFO.Left = Application.Left + Application.Width - 30 - UFormINFO.Width
    End If
End Sub
''==== Obsluga textbox filtru =============================================================================================
Private Sub IndeksTxtBox_Enter()
    If UFormINFO.IndeksTxtBox.Value = "podaj indeks" Then UFormINFO.IndeksTxtBox.Value = ""
End Sub
Private Sub IndeksTxtBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If UFormINFO.IndeksTxtBox.Value = "" Then UFormINFO.IndeksTxtBox.Value = "podaj indeks"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserForm_Activate()
    Dim oFormChanger As New CFormChanger
        Set oFormChanger.Form = Me
            oFormChanger.ShowCaption = True
            Me.SpecialEffect = fmSpecialEffectFlat
End Sub

'-----------------------------------------Dot. Zakladki Indeks--------------------------------------------------
'===============================================================================================================
Private Sub UserForm_Initialize()
    UFormINFO.CheckBoxCzyTylkoIndeks.Value = True
    UFormINFO.ButtonIdxSzczegoly.Visible = False
    UFormINFO.FrameProgress.Visible = False
    UFormINFO.LabelStatusBar.Visible = True
    UFormINFO.LabelJM.Visible = True
    UFormINFO.LabelKomOrgPobierTYT.Visible = True
    With UFormINFO.ComboBoxWyborOpcji
        .AddItem "Ca³a kartoteka"
        .AddItem "Stany Awaryjne"
    End With
    UFormINFO.ComboBoxWyborOpcji.Value = "Ca³a kartoteka"
    UFormINFO.MultiPage2.Pages("PageZuzycie").Caption = "Zu¿ycie w latach"
    With UFormINFO.ComboBoxRwCzyPZ
        .AddItem "Zu¿ycie w latach"
        .AddItem "Zakupy w latach"
    End With
    UFormINFO.ComboBoxRwCzyPZ.Value = "Zu¿ycie w latach"
    UFormINFO.ComboBoxKomOrg.Enabled = True
        UFormINFO.TextBoxIndeks.Value = "indeks"
        UFormINFO.TextBoxNazwaTowaru.Value = "Nazwa towaru"
        UFormINFO.TextBoxGrupaMaterialowa.Value = "Gr. mat."
        UFormINFO.TextBoxDostawca.Value = "Dostawca"
        UFormINFO.TextBoxKomOrgFiltr.Value = "Komorka Org."
        UFormINFO.TextBoxDataUtworzenia.Value = "Data utworzenia"
        UFormINFO.TextBoxSzukGoogle.Value = ""
        UFormINFO.TextBoxKomOrg.Value = ""
        UFormINFO.IndeksTxtBox.text = "podaj indeks"
        UFormINFO.TextBoxZakresBezIdx.text = ""
        UFormINFO.TextBoxZakresBezIdx2.text = ""
        UFormNr = 2
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Len(Dir$(Environ("HOMEPATH") & Application.PathSeparator & "PICTURES" & Application.PathSeparator & "JKoziAddInCHART.gif")) > 0 Then
        Kill Environ("HOMEPATH") & Application.PathSeparator & "PICTURES" & Application.PathSeparator & "JKoziAddInCHART.gif"
        UFormINFO.Indeks_ImageWykresZuzycie.Picture = Nothing
    End If
    Application.StatusBar = "": UFormNr = 0
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Function Zamien_tekst_PRIV(TekstPrzed As String) As String
    Dim TekstPo As String
    Dim a, b As Long
    Dim TblZnak As New Collection
    Dim a1, a2, a3 As String
        TblZnak.Add "¥": TblZnak.Add "Æ": TblZnak.Add "Ê": TblZnak.Add "£": TblZnak.Add "Ñ"
        TblZnak.Add "Ó": TblZnak.Add "": TblZnak.Add "": TblZnak.Add "¯": TblZnak.Add "."
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
                        TekstPo = Replace(TekstPo, TblZnak(a), a3)
                    Else
                        TekstPo = Replace(TekstPo, TblZnak(a), a2)
                    End If
                Next
            TekstPo = LTrim(TekstPo)
            TekstPo = RTrim(TekstPo)
                Zamien_tekst_PRIV = TekstPo
End Function

''==== sortowanie listboxa ======================================================================
Sub SortListBox_fVAT(lst As MSForms.ListBox, xCl As Long)
    Dim arrItems As Variant
    Dim arrTemp1, arrTemp2, arrTemp3, arrTemp4, arrTemp5, arrTemp6, arrTemp7, arrTemp8, arrTemp9, arrTemp10 As Variant
    Dim intOuter As Long
    Dim intInner As Long
    Dim xLp As Long
    Dim ArrItA, ArrItB As Variant
    arrItems = lst.List
    For intOuter = LBound(arrItems, 1) To (UBound(arrItems, 1) + 4)
        For intInner = intOuter + 1 To UBound(arrItems, 1)
            If UFormINFO.Label1.Caption = "Zuzycie" Then If InStr(1, CStr(arrItems(intOuter, 5)), ".") <> 0 Then arrItems(intOuter, 5) = Replace(CStr(arrItems(intOuter, 5)), ".", ",")
            If UFormINFO.Label1.Caption = "Zuzycie" Then If InStr(1, CStr(arrItems(intInner, 5)), ".") <> 0 Then arrItems(intInner, 5) = Replace(CStr(arrItems(intInner, 5)), ".", ",")
            If UFormINFO.Label1.Caption = "Zuzycie" Then If InStr(1, CStr(arrItems(intOuter, 6)), ".") <> 0 Then arrItems(intOuter, 6) = Replace(CStr(arrItems(intOuter, 6)), ".", ",")
            If UFormINFO.Label1.Caption = "Zuzycie" Then If InStr(1, CStr(arrItems(intInner, 6)), ".") <> 0 Then arrItems(intInner, 6) = Replace(CStr(arrItems(intInner, 6)), ".", ",")
            If UFormINFO.Label1.Caption = "Zuzycie" Then If InStr(1, CStr(arrItems(intOuter, 7)), ".") <> 0 Then arrItems(intOuter, 7) = Replace(CStr(arrItems(intOuter, 7)), ".", ",")
            If UFormINFO.Label1.Caption = "Zuzycie" Then If InStr(1, CStr(arrItems(intInner, 7)), ".") <> 0 Then arrItems(intInner, 7) = Replace(CStr(arrItems(intInner, 7)), ".", ",")
            If UFormINFO.ComboBoxWyborOpcji.Value <> "Stany Awaryjne" Then
                If UFormINFO.Label1.Caption = "Zuzycie" Then ArrItA = CLng(arrItems(intOuter, xCl)): ArrItB = CLng(arrItems(intInner, xCl)) Else ArrItA = arrItems(intOuter, xCl): ArrItB = arrItems(intInner, xCl)
            ElseIf UFormINFO.ComboBoxWyborOpcji.Value = "Stany Awaryjne" Then
                If UFormINFO.Label1.Caption = "Zuzycie" Then ArrItA = CDbl(arrItems(intOuter, xCl)): ArrItB = CDbl(arrItems(intInner, xCl)) Else ArrItA = arrItems(intOuter, xCl): ArrItB = arrItems(intInner, xCl)
            End If
                If ArrItA < ArrItB Then
                    arrTemp1 = arrItems(intOuter, 0)
                    arrTemp2 = arrItems(intOuter, 1)
                    arrTemp3 = arrItems(intOuter, 2)
                    arrTemp4 = arrItems(intOuter, 3)
                    arrTemp5 = arrItems(intOuter, 4)
                    arrTemp6 = arrItems(intOuter, 5)
                    arrTemp7 = arrItems(intOuter, 6)
                    arrTemp8 = arrItems(intOuter, 7)
                    arrTemp9 = arrItems(intOuter, 8)
                    arrTemp10 = arrItems(intOuter, 9)
                    arrItems(intOuter, 0) = arrItems(intInner, 0)
                    arrItems(intOuter, 1) = arrItems(intInner, 1)
                    arrItems(intOuter, 2) = arrItems(intInner, 2)
                    arrItems(intOuter, 3) = arrItems(intInner, 3)
                    arrItems(intOuter, 4) = arrItems(intInner, 4)
                    arrItems(intOuter, 5) = arrItems(intInner, 5)
                    arrItems(intOuter, 6) = arrItems(intInner, 6)
                    arrItems(intOuter, 7) = arrItems(intInner, 7)
                    arrItems(intOuter, 8) = arrItems(intInner, 8)
                    arrItems(intOuter, 9) = arrItems(intInner, 9)
                    arrItems(intInner, 0) = arrTemp1
                    arrItems(intInner, 1) = arrTemp2
                    arrItems(intInner, 2) = arrTemp3
                    arrItems(intInner, 3) = arrTemp4
                    arrItems(intInner, 4) = arrTemp5
                    arrItems(intInner, 5) = arrTemp6
                    arrItems(intInner, 6) = arrTemp7
                    arrItems(intInner, 7) = arrTemp8
                    arrItems(intInner, 8) = arrTemp9
                    arrItems(intInner, 9) = arrTemp10
                End If
        Next intInner
    Next intOuter
    lst.Clear
    xLp = 1
    For intOuter = LBound(arrItems, 1) To UBound(arrItems, 1)
        If IsNull(arrItems(intOuter, 0)) = False Then lst.AddItem xLp & "."
        If IsNull(arrItems(intOuter, 1)) = False Then lst.List(intOuter, 1) = arrItems(intOuter, 1)
        If IsNull(arrItems(intOuter, 2)) = False Then lst.List(intOuter, 2) = arrItems(intOuter, 2)
        If IsNull(arrItems(intOuter, 3)) = False Then lst.List(intOuter, 3) = arrItems(intOuter, 3)
        If IsNull(arrItems(intOuter, 4)) = False Then lst.List(intOuter, 4) = arrItems(intOuter, 4)
        If IsNull(arrItems(intOuter, 5)) = False Then lst.List(intOuter, 5) = arrItems(intOuter, 5)
        If IsNull(arrItems(intOuter, 6)) = False Then lst.List(intOuter, 6) = arrItems(intOuter, 6)
        If IsNull(arrItems(intOuter, 7)) = False Then lst.List(intOuter, 7) = arrItems(intOuter, 7)
        If IsNull(arrItems(intOuter, 8)) = False Then lst.List(intOuter, 8) = arrItems(intOuter, 8)
        If IsNull(arrItems(intOuter, 9)) = False Then lst.List(intOuter, 9) = arrItems(intOuter, 9)
        xLp = xLp + 1
    Next intOuter
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Private Function czyWkolekcji_UFormInfo(igla As Variant, stog As Collection) As Boolean
Dim element As Variant
czyWkolekcji_UFormInfo = False
For Each element In stog
    If igla = element Then czyWkolekcji_UFormInfo = True
Next element
End Function

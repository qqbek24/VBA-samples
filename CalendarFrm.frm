VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarFrm 
   Caption         =   "Calendar Control"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3960
   OleObjectBlob   =   "CalendarFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CalendarFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************************************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''autor: Jakub Koziorowski''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''utworzone w 2014r''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim ThisDay As Date
    Dim ThisYear, ThisMth As Date
    Dim CreateCal As Boolean
    Dim i, j As Integer
    Public mc As String
    Public dzien, rok As String
    Public swieto As Date
    Public data_swieta As Boolean
    Public PierwRaz As Single
    Public IleRazy As Integer
    Public Mies As String

Private Sub CommandButton2_Click()
    Unload Me
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub D1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'rebuilds the calendar when the month is changed by the user
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 1 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) - 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) - 1, CInt(Mies) - 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) - 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 1 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) - 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) - 1, CInt(Mies) - 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) - 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 1 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) - 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) - 1, CInt(Mies) - 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) - 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D4_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 1 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) - 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) - 1, CInt(Mies) - 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) - 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D5_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 1 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) - 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) - 1, CInt(Mies) - 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) - 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D6_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 1 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) - 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) - 1, CInt(Mies) - 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) - 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D7_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 1 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) - 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) - 1, CInt(Mies) - 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) - 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub D36_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 12 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) + 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) + 1, CInt(Mies) + 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) + 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D37_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 12 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) + 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) + 1, CInt(Mies) + 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) + 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D38_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 12 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) + 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) + 1, CInt(Mies) + 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) + 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D39_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 12 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) + 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) + 1, CInt(Mies) + 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) + 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D40_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 12 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) + 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) + 1, CInt(Mies) + 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) + 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D41_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 12 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) + 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) + 1, CInt(Mies) + 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) + 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub
Private Sub D42_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PierwRaz <> 1 Then
        miesiac (CB_Mth.Value)
        Mies = CInt(mc)
        If Mies = 12 Then
            CB_Yr.Value = CStr(CInt(CB_Yr.Value) + 1)
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value) + 1, CInt(Mies) + 1, 1), "mmmm"))
        Else
            CB_Mth.Value = CStr(Format(DateSerial(CInt(CB_Yr.Value), CInt(Mies) + 1, 1), "mmmm"))
        End If
        Build_Calendar
    End If
    PierwRaz = 2
End Sub

Private Sub UserForm_Initialize()
    Application.EnableEvents = False
    'starts the form on todays date
    PierwRaz = 1
    ThisDay = Date
    ThisMth = Format(ThisDay, "mm")
    ThisYear = Format(ThisDay, "yyyy")
    For i = 1 To 12
        CB_Mth.AddItem Format(DateSerial(Year(Date), Month(Date) + i, 0), "mmmm")
    Next
    CB_Mth.ListIndex = Format(Date, "mm") - Format(Date, "mm")
    For i = -20 To 50
        If i = 1 Then CB_Yr.AddItem Format((ThisDay), "yyyy") Else CB_Yr.AddItem _
            Format((DateAdd("yyyy", (i - 1), ThisDay)), "yyyy")
    Next
    CB_Yr.ListIndex = 21
    'Builds the calendar with todays date
    CalendarFrm.Width = CalendarFrm.Width
    CreateCal = True
    Call Build_Calendar
    Application.EnableEvents = True
    
End Sub
Private Sub CB_Mth_Change()
    'rebuilds the calendar when the month is changed by the user
    Build_Calendar
End Sub
Private Sub CB_Yr_Change()
    'rebuilds the calendar when the year is changed by the user
    Build_Calendar
End Sub
Private Sub Build_Calendar()
    'the routine that actually builds the calendar each time
    If CreateCal = True Then
    CalendarFrm.Caption = " " & CB_Mth.Value & " " & CB_Yr.Value
    'sets the focus for the todays date button
    CommandButton1.SetFocus
    For i = 1 To 42
        If i < Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value)) Then
            Controls("D" & (i)).Caption = Format(DateAdd("d", (i - Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), _
                ((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), "d")
            Controls("D" & (i)).ControlTipText = DateAdd("d", (i - Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), _
                ((CB_Mth.Value) & "/1/" & (CB_Yr.Value)))
        ElseIf i >= Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value)) Then
            Controls("D" & (i)).Caption = Format(DateAdd("d", (i - Weekday((CB_Mth.Value) _
                & "/1/" & (CB_Yr.Value))), ((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), "d")
            Controls("D" & (i)).ControlTipText = DateAdd("d", (i - Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), _
                ((CB_Mth.Value) & "/1/" & (CB_Yr.Value)))
        End If
        If Format(DateAdd("d", (i - Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), _
        ((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), "mmmm") = ((CB_Mth.Value)) Then
            If Controls("D" & (i)).BackColor <> &H80000016 Then Controls("D" & (i)).BackColor = &H80000018  '&H80000010
            Controls("D" & (i)).Font.Bold = True
        If Format(DateAdd("d", (i - Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), _
            ((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), "m/d/yy") = Format(ThisDay, "m/d/yy") Then Controls("D" & (i)).SetFocus
        Else
            If Controls("D" & (i)).BackColor <> &H80000016 Then Controls("D" & (i)).BackColor = &H8000000F
            Controls("D" & (i)).Font.Bold = False
        End If
    Next
    End If
 Call czarny_kolor
 Call Nd_Sb_color
    miesiac (CB_Mth.Value)
    rok = CB_Yr.Value

    For j = 1 To 42
        dzien = Controls("D" & (j)).Caption
            If Controls("D" & (j)).Font.Bold = True Then
                If dzien < 10 Then
                    dzien = "0" & dzien
                    swieto = rok & "-" & mc & "-" & dzien
                    data_swieta = czy_swieto(swieto)
                End If
                If dzien > 9 Then
                    dzien = dzien
                    swieto = rok & "-" & mc & "-" & dzien
                    data_swieta = czy_swieto(swieto)
                End If
                    If data_swieta Then
                       Controls("D" & (j)).ForeColor = RGB(0, 150, 50)
                    End If
            End If
    Next j
End Sub

Private Sub czarny_kolor()
    For j = 1 To 42
        Controls("D" & (j)).ForeColor = RGB(0, 0, 0)
    Next j
End Sub

Private Sub Nd_Sb_color()
'czerwony 255 0 0
'pomaranczowy 255 155 0
'zielony 0 150 50
D1.ForeColor = RGB(255, 0, 0)
D8.ForeColor = RGB(255, 0, 0)
D15.ForeColor = RGB(255, 0, 0)
D22.ForeColor = RGB(255, 0, 0)
D29.ForeColor = RGB(255, 0, 0)
D36.ForeColor = RGB(255, 0, 0)
D7.ForeColor = RGB(255, 155, 0)
D14.ForeColor = RGB(255, 155, 0)
D21.ForeColor = RGB(255, 155, 0)
D28.ForeColor = RGB(255, 155, 0)
D35.ForeColor = RGB(255, 155, 0)
D42.ForeColor = RGB(255, 155, 0)
End Sub

Function miesiac(ms As String) As String
ms = CB_Mth.Value
    If ms = "styczeñ" Then
       mc = "01"
    End If
    If ms = "luty" Then
        mc = "02"
    End If
    If ms = "marzec" Then
        mc = "03"
    End If
    If ms = "kwiecieñ" Then
        mc = "04"
    End If
    If ms = "maj" Then
        mc = "05"
    End If
    If ms = "czerwiec" Then
        mc = "06"
    End If
    If ms = "lipiec" Then
        mc = "07"
    End If
    If ms = "sierpieñ" Then
        mc = "08"
    End If
    If ms = "wrzesieñ" Then
        mc = "09"
    End If
    If ms = "paŸdziernik" Then
        mc = "10"
    End If
    If ms = "listopad" Then
        mc = "11"
    End If
    If ms = "grudzieñ" Then
        mc = "12"
    End If
    
End Function

Public Function czy_swieto(kiedy As Date) As Boolean
Dim wielkanoc As Date
wielkanoc = WorksheetFunction.Floor(DateSerial(Year(kiedy), 5, Day(Minute(Year(kiedy) / 38) / 2 + 56)), 7) - 34
Select Case kiedy
Case DateSerial(Year(kiedy), 1, 1) 'Nowy rok
        czy_swieto = True
Case DateSerial(Year(kiedy), 1, 6) 'Trzech Króli (Objawienie Pañskie)
        czy_swieto = True
Case wielkanoc
        czy_swieto = True
Case wielkanoc + 1 'Poniedzia³ek Wielkanocny
        czy_swieto = True
Case DateSerial(Year(kiedy), 5, 1) 'Œwiêto pracy
        czy_swieto = True
Case DateSerial(Year(kiedy), 5, 3) 'Konstytucja 3 Maja
        czy_swieto = True
Case wielkanoc + 60 'Bo¿e Cia³o
        czy_swieto = True
Case DateSerial(Year(kiedy), 8, 15) 'Wniebowziêcie NMP
        czy_swieto = True
Case DateSerial(Year(kiedy), 11, 1) 'Wszystkich Œwiêtych
        czy_swieto = True
Case DateSerial(Year(kiedy), 11, 11) 'Œwiêto Niepodleg³oœci
        czy_swieto = True
Case DateSerial(Year(kiedy), 12, 24) 'Wigilia
        czy_swieto = True
Case DateSerial(Year(kiedy), 12, 25) 'Bo¿e Narodzenie
        czy_swieto = True
Case DateSerial(Year(kiedy), 12, 26) 'Szczepana
        czy_swieto = True
End Select
End Function

Private Sub D1_Click()
    'this sub and the ones following represent the buttons for days on the form
    'retrieves the current value of the individual controltiptext and
    'places it in the active cell
    If CzyDataKalend = True Then
        DataUtwo = CDate(D1.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D1.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D1.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D1.ControlTipText)
    Else
        ActiveCell.Value = CDate(D1.ControlTipText)
    End If
    Unload Me
    'after unload you can call a different userform to continue data entry
    'uncomment this line and add a userform named UserForm2
    'Userform2.Show
End Sub
Private Sub D2_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D2.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D2.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D2.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D2.ControlTipText)
    Else
        ActiveCell.Value = CDate(D2.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D3_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D3.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D3.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D3.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D3.ControlTipText)
    Else
        ActiveCell.Value = CDate(D3.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D4_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D4.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D4.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D4.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D4.ControlTipText)
    Else
        ActiveCell.Value = CDate(D4.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D5_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D5.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D5.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D5.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D5.ControlTipText)
    Else
        ActiveCell.Value = CDate(D5.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D6_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D6.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D6.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D6.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D6.ControlTipText)
    Else
        ActiveCell.Value = CDate(D6.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D7_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D7.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D7.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D7.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D7.ControlTipText)
    Else
        ActiveCell.Value = CDate(D7.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D8_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D8.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D8.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D8.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D8.ControlTipText)
    Else
        ActiveCell.Value = CDate(D8.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D9_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D9.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D9.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D9.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D9.ControlTipText)
    Else
        ActiveCell.Value = CDate(D9.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D10_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D10.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D10.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D10.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D10.ControlTipText)
    Else
        ActiveCell.Value = CDate(D10.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D11_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D11.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D11.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D11.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D11.ControlTipText)
    Else
        ActiveCell.Value = CDate(D11.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D12_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D12.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D12.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D12.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D12.ControlTipText)
    Else
        ActiveCell.Value = CDate(D12.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D13_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D13.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D13.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D13.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D13.ControlTipText)
    Else
        ActiveCell.Value = CDate(D13.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D14_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D14.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D14.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D14.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D14.ControlTipText)
    Else
        ActiveCell.Value = CDate(D14.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D15_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D15.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D15.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D15.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D15.ControlTipText)
    Else
        ActiveCell.Value = CDate(D15.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D16_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D16.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D16.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D16.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D16.ControlTipText)
    Else
        ActiveCell.Value = CDate(D16.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D17_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D17.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D17.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D17.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D17.ControlTipText)
    Else
        ActiveCell.Value = CDate(D17.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D18_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D18.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D18.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D18.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D18.ControlTipText)
    Else
        ActiveCell.Value = CDate(D18.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D19_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D19.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D19.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D19.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D19.ControlTipText)
    Else
        ActiveCell.Value = CDate(D19.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D20_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D20.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D20.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D20.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D20.ControlTipText)
    Else
        ActiveCell.Value = CDate(D20.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D21_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D21.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D21.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D21.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D21.ControlTipText)
    Else
        ActiveCell.Value = CDate(D21.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D22_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D22.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D22.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D22.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D22.ControlTipText)
    Else
        ActiveCell.Value = CDate(D22.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D23_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D23.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D23.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D23.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D23.ControlTipText)
    Else
        ActiveCell.Value = CDate(D23.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D24_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D24.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D24.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D24.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D24.ControlTipText)
    Else
        ActiveCell.Value = CDate(D24.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D25_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D25.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D25.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D25.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D25.ControlTipText)
    Else
        ActiveCell.Value = CDate(D25.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D26_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D26.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D26.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D26.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D26.ControlTipText)
    Else
        ActiveCell.Value = CDate(D26.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D27_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D27.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D27.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D27.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D27.ControlTipText)
    Else
        ActiveCell.Value = CDate(D27.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D28_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D28.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D28.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D28.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D28.ControlTipText)
    Else
        ActiveCell.Value = CDate(D28.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D29_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D29.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D29.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D29.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D29.ControlTipText)
    Else
        ActiveCell.Value = CDate(D29.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D30_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D30.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D30.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D30.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D30.ControlTipText)
    Else
        ActiveCell.Value = CDate(D30.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D31_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D31.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D31.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D31.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D31.ControlTipText)
    Else
        ActiveCell.Value = CDate(D31.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D32_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D32.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D32.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D32.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D32.ControlTipText)
    Else
        ActiveCell.Value = CDate(D32.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D33_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D33.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D33.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D33.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D33.ControlTipText)
    Else
        ActiveCell.Value = CDate(D33.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D34_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D34.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D34.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D34.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D34.ControlTipText)
    Else
        ActiveCell.Value = CDate(D34.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D35_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D35.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D35.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D35.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D35.ControlTipText)
    Else
        ActiveCell.Value = CDate(D35.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D36_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D36.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D36.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D36.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D36.ControlTipText)
    Else
        ActiveCell.Value = CDate(D36.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D37_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D37.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D37.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D37.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D37.ControlTipText)
    Else
        ActiveCell.Value = CDate(D37.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D38_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D38.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D38.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D38.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D38.ControlTipText)
    Else
        ActiveCell.Value = CDate(D38.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D39_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D39.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D39.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D39.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D39.ControlTipText)
    Else
        ActiveCell.Value = CDate(D39.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D40_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D40.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D40.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D40.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D40.ControlTipText)
    Else
        ActiveCell.Value = CDate(D40.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D41_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D41.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D41.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D41.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D41.ControlTipText)
    Else
        ActiveCell.Value = CDate(D41.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub D42_Click()
    If CzyDataKalend = True Then
        DataUtwo = CDate(D42.ControlTipText)
        Unload Me
        Exit Sub
    End If
    If WyborDatTrans = True Then
        If DatP = 1 Then
            DataPoczT = CDate(D42.ControlTipText)
        ElseIf DatP = 2 Then
            DataKonT = CDate(D42.ControlTipText)
        End If
    ElseIf CzyDataPZfakPrzekaz = True Then
        DataPZfakPrzekaz = CDate(D42.ControlTipText)
    Else
        ActiveCell.Value = CDate(D42.ControlTipText)
    End If
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Application.StatusBar = ""
End Sub

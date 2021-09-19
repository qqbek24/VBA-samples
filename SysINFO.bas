Attribute VB_Name = "SysINFO"
Private Const LOCALE_SSHORTDATE = &H1F

Private Declare Function GetLocaleInfo Lib "KERNEL32" _
Alias "GetLocaleInfoA" (ByVal Locale As Long, _
ByVal LCType As Long, ByVal lpLCData As String, _
ByVal cchData As Long) As Long

Private Declare Function GetThreadLocale Lib "KERNEL32" () As Long
Public SystemShortDateFormat As String

Public wylPOKAZb As Integer

'********************************************************************************************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''autor: Jakub Koziorowski''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub data_Format()
    Dim localeID As Long
    Dim bLen As Long
    Dim bRet As Long
    Dim sBuf As String
    localeID = GetThreadLocale
    bLen = GetLocaleInfo(localeID, LOCALE_SSHORTDATE, sBuf, 0)
        If bLen > 0 Then
            sBuf = Space(bLen)
            bRet = GetLocaleInfo(localeID, LOCALE_SSHORTDATE, _
            sBuf, bLen)
            sBuf = Left(sBuf, bLen - 1)
        End If
    SystemShortDateFormat = sBuf
End Sub
Sub Sys_INFO_show()
    UserWinPokazForm.Show
End Sub
Sub ListProcedures()
    Dim sProc() As String
    Dim lngLine As Long
    Dim VBCodeMod As VBComponent
    Dim sLine As String, sProcName As String, s As String
    Dim vType As Variant
    Dim AllWorkbooks As New Collection
    Dim TblUser() As String
    Dim i As Byte
    SysVerWinPokazProcedures = 1
        For Each adn In AddIns
            If adn.Installed Then AllWorkbooks.Add (adn.Name)
        Next
            For Each Wbk In Workbooks
                AllWorkbooks.Add Wbk.Name
            Next
                ReDim sProc(1 To 2, 1 To 1)
                    r = 1
                        vType = Array("Function", "Public Function", "Private Function")
                For Each Wbk In AllWorkbooks
                    If Wbk <> UCase("analys32.xll") And Wbk <> UCase("MSADDNDR.DLL") And UCase(Wbk) <> UCase("iasads.dll") Then
                        If Workbooks(Wbk).VBProject.Protection = vbext_pp_none Then
                            For iType = LBound(vType) To UBound(vType)
                                For Each VBCodeMod In Workbooks(Wbk).VBProject.VBComponents
                                    With VBCodeMod.CodeModule
                                        lngLine = .CountOfDeclarationLines + 1
                                        Do Until lngLine >= .CountOfLines
                                            sLine = .Lines(lngLine, 1)
                                            If Left(sLine, Len(vType(iType))) = vType(iType) Then
                                                ReDim Preserve sProc(1 To 2, 1 To r)
                                                sProc(1, r) = Right(sLine, Len(sLine) - Len(vType(iType)) - 1)
                                                sProc(2, r) = Wbk
                                                r = r + 1
                                            End If
                                            lngLine = lngLine + 1
                                        Loop
                                    End With
                                Next VBCodeMod
                            Next iType
                        End If
                    End If
                Next Wbk
        ReDim Preserve TblUser(1 To r, 1 To 2)
            TblUser(1, 1) = "Dodatek"
            TblUser(1, 2) = "Funkcja"
            i = 2
        For n = 1 To r - 1
            sLine = sProc(1, n)
                TblUser(i, 1) = " ( " & sProc(2, n) & " ) "
                TblUser(i, 2) = sLine
                i = i + 1
        Next n
        UserWinPokazForm.ListBoxWin.ColumnWidths = "200;500"
        UserWinPokazForm.ListBoxWin.List = TblUser
        If (UserWinPokazForm.PokazB.Visible = True) = True Then UserWinPokazForm.PokazB.Visible = False
            wylPOKAZb = 1
            UserWinPokazForm.Show
End Sub
Sub Check_Fun(FunkcjaStr As String)
    '...
End Sub

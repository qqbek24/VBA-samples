VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserWinPokazForm 
   Caption         =   "Kto - INFO ?"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   23745
   OleObjectBlob   =   "UserWinPokazForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserWinPokazForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public a, b, c, d As Long                  (zadeklarowane w MODULE "NoweIndeksy")
'Public Tbl1(0 To 1000) As Variant          (zadeklarowane w MODULE "NoweIndeksy")
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
    Dim formhandle As Long
    
Option Explicit
Option Compare Text

'********************************************************************************************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''autor: Jakub Koziorowski''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''Dzia³ Zamowien Publicznych'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''stworzone w 2014r''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ListBoxWin_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        Dim aL As Long
        Dim ProcStr As String
            aL = UserWinPokazForm.ListBoxWin.ListIndex
            ProcStr = UserWinPokazForm.ListBoxWin.List(aL, 1)
            If SysVerWinPokazProcedures = 1 Then
                ProcStr = Mid(ProcStr, 1, InStr(1, ProcStr, "(") - 1)
            End If
                ActiveCell = ProcStr
                ActiveSheet.Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
End Sub
Private Sub PokazB_Click()
    Dim i, j As Byte
    Dim n As Long
    Dim TblUser() As String
    Dim LpUs As Long
        UserWinPokazForm.ListBoxWin.ColumnWidths = "40;250"
        UserWinPokazForm.ListBoxWin.Clear
        j = 2
        Do Until Environ(j) = ""
            If Environ(j) = "" Then GoTo DodajEnv
            j = j + 1
        Loop
DodajEnv:
            ReDim TblUser(1 To j, 1 To 2) As String
            TblUser(1, 1) = "Lp"
            TblUser(1, 2) = "Opis"
            i = 2
            Do Until Environ(i) = "" 'Or Environ(i).EOF
                If Environ(i) = "" Then Exit Sub
                LpUs = i - 1
                TblUser(i, 1) = LpUs
                TblUser(i, 2) = Environ(i)
                i = i + 1
            Loop
                UserWinPokazForm.ListBoxWin.List = TblUser
End Sub
Private Sub UserForm_Initialize()
    'Me.ListBoxWin.BackColor = Me.ListBoxWin.Parent.BackColor
    Me.ListBoxWin.BackColor = Me.ListBoxWin.Parent.PokazB.BackStyle
        formhandle = FindWindow(vbNullString, Me.Caption)
        SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
            SetLayeredWindowAttributes formhandle, vbCyan, 0&, LWA_COLORKEY
        Me.BackColor = vbCyan
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Application.StatusBar = ""
End Sub
Private Sub UserWinFormZamknijB_Click()
    SysVerWinPokazProcedures = 0
    UserWinPokazForm.ListBoxWin.Clear
    wylPOKAZb = 2
    UserWinPokazForm.Hide
End Sub
Private Sub UserForm_Activate()
    Dim oFormChanger As New CFormChanger
        Set oFormChanger.Form = Me
            oFormChanger.ShowCaption = False
            Me.SpecialEffect = fmSpecialEffectFlat 'fmSpecialEffectRaised
        If wylPOKAZb <> 1 Then
            If Me.PokazB.Visible = False Then Me.PokazB.Visible = True
        End If
End Sub



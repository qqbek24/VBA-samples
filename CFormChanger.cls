VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
Option Explicit

'********************************************************************************************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''autor: Jakub Koziorowski''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Windows API calls to do all the dirty work!
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

'Lots of window styles for us to play with!
Private Const GWL_STYLE As Long = (-16)           'The offset of a window's style
Private Const GWL_EXSTYLE As Long = (-20)         'The offset of a window's extended style
Private Const WS_CAPTION As Long = &HC00000       'Style to add a titlebar
Private Const WS_SYSMENU As Long = &H80000        'Style to add a system menu
Private Const WS_THICKFRAME As Long = &H40000     'Style to add a sizable frame
Private Const WS_MINIMIZEBOX As Long = &H20000    'Style to add a Minimize box on the title bar
Private Const WS_MAXIMIZEBOX As Long = &H10000    'Style to add a Maximize box to the title bar
Private Const WS_POPUP As Long = &H80000000       'Standard option, cleared when showing a task bar icon
Private Const WS_VISIBLE As Long = &H10000000     'Standard option, cleared when showing a task bar icon

Private Const WS_EX_DLGMODALFRAME As Long = &H1   'Controls if the window has an icon
Private Const WS_EX_APPWINDOW As Long = &H40000   'Application Window: shown on taskbar
Private Const WS_EX_TOOLWINDOW As Long = &H80     'Tool Window: small titlebar

'Constant to identify the Close menu item
Private Const SC_CLOSE As Long = &HF060

'Constants for hide or show a window
Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5

'Constants for Windows messages
Private Const WM_SETICON = &H80

'Variables to store the various selections/options
Dim hWndForm As Long, mbSizeable As Boolean, mbCaption As Boolean, mbIcon As Boolean, miModal As Integer
Dim mbMaximize As Boolean, mbMinimize As Boolean, mbSysMenu As Boolean, mbCloseBtn As Boolean
Dim mbAppWindow As Boolean, mbToolWindow As Boolean, msIconPath As String
Dim moForm As Object

'Set defaults to mimic a standard userform
Private Sub Class_Initialize()
    mbCaption = True
    mbCloseBtn = True
    mbSysMenu = True
    
    mbMinimize = True '< was a comment
'    mbMaximize = True '< was a comment
'    mbSizeable = True '<was a comment
    mbAppWindow = True '<was a comment
'    mbToolWindow = True
End Sub
Public Property Set Form(oForm As Object)
'Get the userform's window handle
    If Val(Application.Version) < 9 Then
        hWndForm = FindWindow("ThunderXFrame", oForm.Caption)  'XL97
    Else
        hWndForm = FindWindow("ThunderDFrame", oForm.Caption)  'XL2000+
    End If
'Remember the form for later
    Set moForm = oForm
'Set the form's style
    SetFormStyle
End Property
Public Property Let ShowCaption(bCaption As Boolean)
    mbCaption = bCaption
    SetFormStyle
End Property
Public Property Get ShowCaption() As Boolean
    ShowCaption = mbCaption
End Property
Public Property Let Modal(bModal As Boolean)
    miModal = Abs(CInt(Not bModal))
'Make the form modal or modeless by enabling/disabling Excel itself
    EnableWindow FindWindow("XLMAIN", Application.Caption), miModal
End Property
'Routine to set the form's window style
Private Sub SetFormStyle()
    Dim iStyle As Long, hMenu As Long, hID As Long, iItems As Integer
'Have we got a form to set?
    If hWndForm = 0 Then Exit Sub
    iStyle = GetWindowLong(hWndForm, GWL_STYLE)
'Build up the basic window style flags for the form
    If mbCaption Then iStyle = iStyle Or WS_CAPTION Else iStyle = iStyle And Not WS_CAPTION
    If mbSysMenu Then iStyle = iStyle Or WS_SYSMENU Else iStyle = iStyle And Not WS_SYSMENU
    If mbSizeable Then iStyle = iStyle Or WS_THICKFRAME Else iStyle = iStyle And Not WS_THICKFRAME
    If mbMinimize Then iStyle = iStyle Or WS_MINIMIZEBOX Else iStyle = iStyle And Not WS_MINIMIZEBOX
    If mbMaximize Then iStyle = iStyle Or WS_MAXIMIZEBOX Else iStyle = iStyle And Not WS_MAXIMIZEBOX
    If mbAppWindow Then iStyle = iStyle And Not WS_VISIBLE And Not WS_POPUP Else iStyle = iStyle Or WS_VISIBLE Or WS_POPUP
'Set the basic window styles
    SetWindowLong hWndForm, GWL_STYLE, iStyle
    iStyle = GetWindowLong(hWndForm, GWL_EXSTYLE)
'Build up and set the extended window style
'If mbIcon Then iStyle = iStyle Or WS_EX_DLGMODALFRAME Else iStyle = iStyle And Not WS_EX_DLGMODALFRAME
    If mbAppWindow Then iStyle = iStyle Or WS_EX_APPWINDOW Else iStyle = iStyle And Not WS_EX_APPWINDOW
    If mbToolWindow Then iStyle = iStyle Or WS_EX_TOOLWINDOW Else iStyle = iStyle And Not WS_EX_TOOLWINDOW
    SetWindowLong hWndForm, GWL_EXSTYLE, iStyle
'Handle the close button differently
    If mbCloseBtn Then
        'We want it, so reset the control menu
        hMenu = GetSystemMenu(hWndForm, 1)
    Else
        'We don't want it, so delete it from the control menu
        hMenu = GetSystemMenu(hWndForm, 0)
        DeleteMenu hMenu, SC_CLOSE, 0&
    End If
'Show the window with the changes
    ShowWindow hWndForm, SW_SHOW
    DrawMenuBar hWndForm
    SetFocus hWndForm
End Sub





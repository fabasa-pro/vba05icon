VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Licenciado sob a licença MIT.
' Copyright (C) 2012 - 2024 @Fabasa-Pro. Todos os direitos reservados.
' Consulte LICENSE.TXT na raiz do projeto para obter informações.

' ==========================================================================
' NOTA: para editar o código-fonte, executar o arquivo com a tecla <Shift>
' pressionada para ignorar todo o VBA e entre no aplicativo Microsoft Word.
' ==========================================================================

Option Explicit

Private Declare PtrSafe Function FindWindow Lib "User32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hWdd As LongPtr, ByVal nIndex As GWL, ByVal dwNewLong As Long) As Long

Private Declare PtrSafe Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal hWdd As LongPtr, ByVal nIndex As GWL) As LongPtr

Private Declare PtrSafe Function ExtractIcon Lib "Shell32.dll" Alias "ExtractIconA" (ByVal hInst As LongPtr, ByVal pszExeFileName As String, ByVal nIconIndex As Long) As Long

Private Declare PtrSafe Function SendMessage Lib "User32.dll" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Enum GWL
    GWL_EXSTYLE = -20     ' Define um novo estilo de janela estendida .
    GWL_HINSTANCE = -6    ' Define um novo identificador de instância do aplicativo.
    GWL_ID = -12          ' Define um novo identificador da janela filho.
    GWL_STYLE = -16       ' Define um novo estilo de janela .
    GWL_USERDATA = -21    ' Define os dados do usuário associados à janela.
    GWL_WNDPROC = -4      ' Define um novo endereço para o procedimento da janela.
End Enum

Private Declare PtrSafe Function ShowWindowAsync Lib "User32.dll" (ByVal hWnd As LongPtr, ByVal nCmdShow As SW) As Boolean

Private Enum SW
    SW_FORCEMINIMIZE = 11     ' Minimiza uma janela, mesmo se o segmento que possui a janela não estiver respondendo.
    SW_HIDE = 0               ' Oculta a janela e ativa outra janela.
    SW_MAXIMIZE = 3           ' Maximiza a janela especificada.
    SW_MINIMIZE = 6           ' Minimiza a janela especificada e ativa a próxima janela de nível superior na ordem Z.
    SW_RESTORE = 9            ' Ativa e exibe a janela.
    SW_SHOW = 5               ' Ativa a janela e a exibe em seu tamanho e posição atuais.
    SW_SHOWDEFAULT = 10       ' Define o estado de exibição.
    SW_SHOWMAXIMIZED = 3      ' Ativa a janela e a exibe como uma janela maximizada.
    SW_SHOWMINIMIZED = 2      ' Ativa a janela e a exibe como uma janela minimizada.
    SW_SHOWMINNOACTIVE = 7    ' Exibe a janela como uma janela minimizada.
    SW_SHOWNA = 8             ' Exibe a janela em seu tamanho e posição atuais.
    SW_SHOWNOACTIVATE = 4     ' Exibe uma janela em seu tamanho e posição mais recentes.
    SW_SHOWNORMAL = 1         ' Ativa e exibe uma janela.
End Enum

Private Sub UserForm_Initialize()

    Dim hWnd As LongPtr
    hWnd = FindWindow(vbNullString, Me.Caption)
        
    Dim FormBorderStyle As Integer
    FormBorderStyle = 4
    
    Select Case FormBorderStyle
        Case 0                                          ' 0-None
            SetWindowLong hWnd, GWL_EXSTYLE, &H50000
            SetWindowLong hWnd, GWL_STYLE, &H6010000
            Me.BackColor = RGB(255, 255, 255)
        Case 1                                          ' 1-FixedSingle
            SetWindowLong hWnd, GWL_EXSTYLE, &H50100
            SetWindowLong hWnd, GWL_STYLE, &H6CB0000
        Case 2                                          ' 2-Fixed3D
            SetWindowLong hWnd, GWL_EXSTYLE, &H50300
            SetWindowLong hWnd, GWL_STYLE, &H6CB0000
        Case 3                                          ' 3-FixedDialog
            SetWindowLong hWnd, GWL_EXSTYLE, &H50101
            SetWindowLong hWnd, GWL_STYLE, &H6CB0000
        Case 4                                          ' 4-Sizable
            SetWindowLong hWnd, GWL_EXSTYLE, &H50100
            SetWindowLong hWnd, GWL_STYLE, &H6CF0000
        Case 5                                          ' 5-FixedToolWindow
            SetWindowLong hWnd, GWL_EXSTYLE, &H50180
            SetWindowLong hWnd, GWL_STYLE, &H6CB0000
        Case 6                                          ' 6-SizableToolWindow
            SetWindowLong hWnd, GWL_EXSTYLE, &H50180
            SetWindowLong hWnd, GWL_STYLE, &H6CF0000
    End Select
    
    Dim WindowState As Integer
    WindowState = 2
    
    Select Case WindowState
        Case 0
            Call ShowWindowAsync(hWnd, SW_SHOWNORMAL)    ' 0-Normal
        Case 1
            Call ShowWindowAsync(hWnd, SW_MINIMIZE)      ' 1-Minimized
        Case 2
            Call ShowWindowAsync(hWnd, SW_MAXIMIZE)      ' 2-Maximized
    End Select

    Dim hInstance As LongPtr
    hInstance = GetWindowLong(hWnd, GWL_HINSTANCE)
    
    Dim hIcon As Long
    hIcon = ExtractIcon(hInstance, Project.ThisDocument.Path & "\Icon.ico", 0&)
    Call SendMessage(hWnd, &H80, 0&, hIcon)

End Sub

Private Sub UserForm_Terminate()

    Project.ThisDocument.Application.Visible = True                                                    ' Ocultar ou mostrar aplicativos.
    Project.ThisDocument.Application.Quit SaveChanges:=wdSaveChanges, OriginalFormat:=wdWordDocument   ' Salvar e fechar tudo.

End Sub

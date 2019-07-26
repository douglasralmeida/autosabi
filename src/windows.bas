Attribute VB_Name = "Windows"
' windows.bas
' Modulo com Funcoes do Windows

' Constantes
Private Const NOERROR = 0

Private Const SHELL32_DLL = "Shell32.dll"

Private Const GW_HWNDNEXT = 2
Private Const MK_LBUTTON = &H1
Private Const SW_MINIMIZE = 6, SW_NORMAL = 1, SW_MAXIMIZE = 3, SW_RESTORE = 9
Private Const WM_CLOSE = &H10
Private Const WM_GETTEXT As Integer = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_SETTEXT As Long = &HC

' Tipos
Private Type InitCommonControlsExStruct
  lngSize As Long
  lngICC As Long
End Type

Private Type POINTAPI
  x As Long
  y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type WINDOWPLACEMENT
  Length As Long
  flags As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type
    
' APIs do Windows
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As String) As Long

' Funcoes do Módulo

' Clica no botão com o manipulador especificado
Public Sub clicarBotao(handle As Long)
  SendMessage handle, BM_CLICK, 0, 0
End Sub

Public Sub clicarControle(handle As Long, pontoClique As Long)
  SendMessage handle, WM_LBUTTONDOWN, MK_LBUTTON, (pontoClique)
  SendMessage handle, WM_LBUTTONUP, MK_LBUTTON, (pontoClique)
End Sub

' Fecha a janela com o manipulador especificado
Public Sub fecharJanelaPorId(handle As Long)
  SendMessage handle, WM_CLOSE, 0, 0
End Sub

' Fecha a janela com o titulo especificado
Public Sub fecharJanelaPorTitulo(titulo As String)
  Dim handle As Long

  handle = 0
  handle = getJanelaPrincipalIDporTitulo(titulo)
  If handle <> 0 Then fecharJanelaPorId handle
End Sub

Public Sub forcarFecharJanelaPorClasse(classe As String, titulo As String)
  janela = pesquisarJanelaSimples(classe, titulo)
  While janela <> 0
    fecharJanelaPorId janela
    espera 200
    DoEvents
    janela = pesquisarJanelaSimples(classe, titulo)
  Wend
End Sub

Public Sub forcarFecharJanelaPorTitulo(titulo As String)
  janela = getJanelaPrincipalIDporTitulo("Carteira de Benefícios")
  While janela <> 0
    fecharJanelaPorId janela
    espera 200
    DoEvents
    janela = getJanelaPrincipalIDporTitulo("Carteira de Benefícios")
  Wend
End Sub

Public Function getControleHabilitacao(handle As Long)
  getControleHabilitacao = Val(GetWindowLong(handle, GWL_STYLE) And WS_DISABLED) = 0
End Function

' Obtem o texto de um controle
Public Function getControleTexto(ByVal handle As Long) As String
  On Error Resume Next
  Dim texto As String
  Dim tamanho As Long
  
  tamanho = SendMessage(handle, WM_GETTEXTLENGTH, 0, 0)
  If tamanho > 0 Then
    tamanho = tamanho + 1
    texto = String(tamanho, vbNullChar)
    tamanho = SendMessage(handle, WM_GETTEXT, tamanho, ByVal texto)
    getControleTexto = Left(texto, tamanho)
  End If
End Function

' Obtem as dimensoes de uma janela
Public Sub getJanelaDimensoes(handle As Long, ByRef RECT As RECT)
  Dim ret As Long
  
  ret = GetWindowRect(janelaDibdip, dimensoesTelaImprimir)
End Sub

' Obtem o titulo de uma janela com o ID especificado
Public Function getJanelaTitulo(handle As Long) As String
  Dim titulo  As String
  
  titulo = Space(256)
  ret = GetWindowText(handle, titulo, Len(titulo))
  getJanelaTitulo = titulo
End Function

' Obtem o handle da janela com o titulo especificado
Public Function getJanelaPrincipalIDporTitulo(ByVal titulo As String) As Long
  Dim handle As Long
  Dim sStr As String
  
  getJanelaPrincipalIDporTitulo = 0
  handle = pesquisarJanelaSimples(vbNullString, vbNullString) 'Janela Pai
  Do While handle <> 0
    sStr = String(GetWindowTextLength(handle) + 1, Chr$(0))
    GetWindowText handle, sStr, Len(sStr)
    sStr = Left$(sStr, Len(sStr) - 1)
    If InStr(1, sStr, titulo) > 0 Then
      getJanelaPrincipalIDporTitulo = handle
      Exit Function
    End If
    handle = GetWindow(handle, GW_HWNDNEXT)
  Loop
  getJanelaPrincipalIDporTitulo = 0
End Function

' Torna ativa a janela com o manipulador especificado
Function janelaTrazerParaFrente(Optional handle As Long, Optional estadoJanela As Long = SW_NORMAL) As Boolean
  On Error Resume Next
  Dim winPlace As WINDOWPLACEMENT

  If handle Then
    winPlace.Length = Len(winPlace)
    
    'Get the windows current placement
    Call GetWindowPlacement(handle, winPlace)
    
    'Set the windows placement
    winPlace.showCmd = estadoJanela
    
    'Change window state
    Call SetWindowPlacement(handle, winPlace)
    
    'Bring to foreground
    janelaTrazerParaFrente = SetForegroundWindow(handle)
  End If
End Function

' Encontra uma janela filha
Public Function pesquisarJanela(handlePai As Long, handleFilhaApos As Long, classe As String, titulo As String) As Long
  pesquisarJanela = FindWindowEx(handlePai, handleFilhaApos, classe, titulo)
End Function

Public Function pesquisarJanelaSimples(classe As String, titulo As String) As Long
  pesquisarJanelaSimples = FindWindow(classe, titulo)
End Function

Public Function pesquisarJanelaInterna(tituloJanela As String) As Long
  On Error Resume Next
  Dim handleJanela As Long
  Dim handleEncontrado As Long
  Dim tituloEncontrado As String
  Dim ret As Long
  Dim handlePai As Long
    
  pesquisarJanelaInterna = 0
  handlePai = FindWindow(vbNullString, vbNullString) 'janela pai
  Do While handlePai <> 0
    handleJanela = pesquisarJanela(handlePai, 0, vbNullString, vbNullString)
    Do Until handleJanela = 0
      If handleJanela > 0 Then
        handleEncontrado = pesquisarJanela(handleJanela, 0, vbNullString, vbNullString)
        Do Until handleEncontrado = 0
          If handleEncontrado > 0 Then
            tituloEncontrado = Space(256)
            ret = GetWindowText(handleEncontrado, tituloEncontrado, Len(tituloEncontrado))
            If InStr(1, tituloEncontrado, tituloJanela) > 0 Then
              pesquisarJanelaInterna = handleEncontrado
              Exit Function
            End If
          End If
          handleEncontrado = pesquisarJanela(handleEncontrado, handleJanela, vbNullString, vbNullString)
        Loop
      End If
      handleJanela = pesquisarJanela(handlePai, handleJanela, vbNullString, vbNullString)
    Loop
    handlePai = GetWindow(handlePai, GW_HWNDNEXT)
  Loop
End Function

Public Sub posicionarJanela(handle As Long, x As Long, y As Long, largura As Long, altura As Long)
  res = SetWindowPos(handle, 0, x, y, largura, altura, 0)
End Sub

' Altera o texto de um controle
Public Sub setControleTexto(handle As Long, texto As String)
  SendMessage handle, WM_SETTEXT, 0, texto
End Sub

' Altera o texto de um controle mascarado
Public Sub setControleMaskTexto(handle As Long, texto As String)
  Dim centroControle As Long
  Dim dimensoes As RECT
  Dim res As Long
  
  res = GetWindowRect(handle, dimensoes)
  centroControle = convlong(dimensoes.Left + (dimensoes.Right - dimensoes.Left) / 2, dimensoes.Top + (dimensoes.Bottom - dimensoes.Top) / 2)
  SendMessage handle, WM_LBUTTONDOWN, MK_LBUTTON, (centroControle)
  SendMessage handle, WM_LBUTTONUP, MK_LBUTTON, (centroControle)
  DoEvents
  espera 100
  simularTeclado "2"
  espera 200
  setControleTexto handle, texto
  DoEvents
  espera 600
End Sub

' Inicializa o suporte a temas
Public Sub Main()
  Dim iccex As InitCommonControlsExStruct, handle As Long
  
  'constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx
  Const ICC_ANIMATE_CLASS As Long = &H80&
  Const ICC_BAR_CLASSES As Long = &H4&
  Const ICC_COOL_CLASSES As Long = &H400&
  Const ICC_DATE_CLASSES As Long = &H100&
  Const ICC_HOTKEY_CLASS As Long = &H40&
  Const ICC_INTERNET_CLASSES As Long = &H800&
  Const ICC_LINK_CLASS As Long = &H8000&
  Const ICC_LISTVIEW_CLASSES As Long = &H1&
  Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&
  Const ICC_PAGESCROLLER_CLASS As Long = &H1000&
  Const ICC_PROGRESS_CLASS As Long = &H20&
  Const ICC_TAB_CLASSES As Long = &H8&
  Const ICC_TREEVIEW_CLASSES As Long = &H2&
  Const ICC_UPDOWN_CLASS As Long = &H10&
  Const ICC_USEREX_CLASSES As Long = &H200&
  Const ICC_STANDARD_CLASSES As Long = &H4000&
  Const ICC_WIN95_CLASSES As Long = &HFF&
  Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all values above

  With iccex
    .lngSize = LenB(iccex)
    .lngICC = ICC_STANDARD_CLASSES ' vb intrinsic controls (buttons, textbox, etc)
    
    ' if using Common Controls; add appropriate ICC_ constants for type of control you are using
    ' example if using CommonControls v5.0 Progress bar:
     ' .lngICC = ICC_STANDARD_CLASSES Or ICC_PROGRESS_CLASS
  End With
  On Error Resume Next ' error? InitCommonControlsEx requires IEv3 or above
    
  handle = LoadLibrary(SHELL32_DLL) ' patch to prevent XP crashes when VB usercontrols present
  InitCommonControlsEx iccex
  If Err Then
    InitCommonControls ' try Win9x version
    Err.Clear
  End If
  On Error GoTo 0
  
  '... exibir o formulário padrão do aplicativo
  formInicial.Show
  If handle Then FreeLibrary handle
End Sub

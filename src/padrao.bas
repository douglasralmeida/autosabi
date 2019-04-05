Attribute VB_Name = "ModuloPadrao"
Public Type POINTAPI
  x As Long
  y As Long
End Type
    
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type DecodiificaAgendamentos
  Horario As Date
  Segurado As String
  Concluida As String
  Ordem As String
  Requerimento As Long
End Type

Public Type Requerimento
  sequencia As String
  Número As String
  Tipo As String
  Status As String
  NIT As String
  Impresso As String
  Segurado As String
  Crítica As String
  CPF As String
End Type
    
Public Type WINDOWPLACEMENT
  Length As Long
  flags As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type

Public Const MF_ENABLED = &H0&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const WM_NCPAINT = &H85
Public Const MK_LBUTTON = &H1
Public Const WM_SETTEXT As Long = &HC
Public Const BM_CLICK = &HF5
Public Const GW_HWNDNEXT = 2
Public Const VK_LBUTTON = &H1
Public Const WM_CLOSE = &H10
Public Const WS_DISABLED As Long = &H8000000
Public Const GWL_STYLE As Long = -16
Public Const SW_MINIMIZE = 6, SW_NORMAL = 1, SW_MAXIMIZE = 3, SW_RESTORE = 9
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_COMMAND = &H111
Public Const BN_CLICKED = 0
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_ABSOLUTE = &H8000
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_GETTEXT As Integer = &HD
Public Const LB_SETSEL = &H185
Public Const CB_SETCURSEL = &H14E
Public Const WM_KEYDOWN = &H100
Public Const VK_RETURN = &HD
Public Const RDW_INVALIDATE = 1
Public Const MF_BYPOSITION = &H400&
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION As Long = 2

Public GlobalAgendamentosConsulta(1000) As DecodiificaAgendamentos
Public GlobalAgendamentosQuandidade As Long
Public GlobalAgendamentosConsultaCabecalho As String
Public GlobalDataEscolhida As Date
Public GlobalDataEscolhida2 As Date
Public GlobalAgenciaEscolhida As String
Public GlobalLinhaPicture As Long
Public GlobalModoSimulado As Boolean
Public GlobalRelatorioPronto As Boolean
Public GlobalRequerimentoMostrado As String
Public GlobalInicio As Long
Public GlobalIDTelaConsultaRequerimentoBenefício As Long
Public GlobalToolbarConsultaRequerimento As Long
Public GlobalToolbarConsultaRequerimentoOCX As Long
Public GlobalTipo As String
Public GlobalAlerta As Boolean
Public GlobalTempodeEspera As Long
Public GlobalPrimeiraVez As Boolean
Public GlobalSeçao As String
Public GlobalImpressãoAutomática As Boolean
Public GlobalUltimoNitInformado As String
Public GlobalhMDIClient As Long
Public GlobalIDRequerimento As Long
Public GlobalHoradeInicio As Date
Public GlobalNomedoRelatorio As String
Public GlobalPróximoNITaserimpresso As String
Public GlobalRequerimentos(1000) As Requerimento
Public GlobalQuantidadedeRequerimentos As Long
Public GlobalIDTelaAtiva As Long
Public GlobalIDControleOperacional As Long
Public GlobalIDTelaImprimirAgendamento As Long
Public GlobalIDTelaExport As Long
Public GlobalIDTelaRequerimentosCrystalReport As Long
Public GlobalIDTelaSalvarComo As Long
Public GlobalTítulodaTelaAtiva As String
Public GlobalDatadosRequerimentos As String
Public GlobalMenuAtualizado As Boolean
Public GlobalUserName As String
Public GlobalModoImprimeRequerimentos As Boolean
Public GlobalIDTelaRTF As Long
Public res As String
Public GlobalEscalaX As Double
Public GlobalEscalay As Double
Public GlobalPastadeTrabalho As String
Public GlobalAreadeTrabalho As String
Public Const NomeAplicacao = "Automatizador do SABI"
    
Public Const SRCCOPY = &HCC0020
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const WM_VSCROLL = &H115
Public Const SB_BOTTOM = 7
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXVSCROLL = 2

Public mvarListBox As ListBox
Public m_lMaxItemWidth As Long

Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    
Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function LBItemFromPt Lib "COMCTL32.DLL" (ByVal hLB As Long, ByVal ptX As Long, ByVal ptY As Long, ByVal bAutoScroll As Long) As Long

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long ' conta milissegundos desde inicio windows
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
   
Public Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Public Declare Sub ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As RECT)
Public Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function lSetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function apiRedrawWindow Lib "user32" Alias "RedrawWindow" (ByVal hWnd As Long, ByVal lprcUpdate As Boolean, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function TextOut Lib "GDI32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Rectangle Lib "GDI32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Sub CopiaTelaCPF(memonumerodorequerimento As String, numtela As Long)
  On Error Resume Next
  Dim hWndActive As Long
  Dim r As Long
  Dim RectActive As RECT
  
  hWndActive = numtela
  r = GetWindowRect(hWndActive, RectActive)
  Set formInicial.pctCopiaPartedaTelaCPF.Picture = CaptureWindow(hWndActive, False, 54, 69, 80, 15)
End Sub
       
Public Function ItemUnderMouse(ByVal list_hWnd As Long, ByVal x As Single, ByVal y As Single)
  Dim pt As POINTAPI

  pt.x = x \ Screen.TwipsPerPixelX
  pt.y = y \ Screen.TwipsPerPixelY
  ClientToScreen list_hWnd, pt
  ItemUnderMouse = LBItemFromPt(list_hWnd, pt.x, pt.y, False)
End Function

Public Sub Espera(tempo As Long)
  Dim memotempo As Long
  Dim segundos As Long
  Dim minutos As Long
  Dim iddatela As Long
  Dim res As String

  memotempo = 0
  iddatela = GetForegroundWindow
  res = UCase(ObtemTextodoControle(iddatela))
  If InStr(1, res, "SALVAR COMO") Then
    res = SetWindowPos(iddatela, 0, 0, 0, 460, 340, 0)
    res = SetWindowPos(iddatela, 1, 0, 0, 0, 0, 3)
    DoEvents
  End If
  While memotempo < tempo
    segundos = Int((GetTickCount - GlobalInicio) / 100)
    minutos = Int(segundos / 600)
    segundos = segundos - minutos * 600
    segundos = segundos / 10
    formInicial.lblRelogio.Caption = " " & minutos & ":" & Format(segundos, "00") & " "
    Sleep 100
    memotempo = memotempo + 100
  Wend
End Sub
     
Public Function ObtemIDdaTelaPrincipalporTitulo(ByVal sCaption As String) As Long
  Dim lhWndP As Long
  Dim sStr As String
  
  ObtemIDdaTelaPrincipalporTitulo = False
  lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
  Do While lhWndP <> 0
    sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
    GetWindowText lhWndP, sStr, Len(sStr)
    sStr = Left$(sStr, Len(sStr) - 1)
    If InStr(1, sStr, sCaption) > 0 Then
      ObtemIDdaTelaPrincipalporTitulo = lhWndP
      Exit Function
    End If
    lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
  Loop
  ObtemIDdaTelaPrincipalporTitulo = 0
End Function

Public Function convlong(xini As Long, yini As Long)
  Dim lParam As Long
  
  lParam = 256 * 64
  lParam = yini * (lParam * 4) + xini
  convlong = lParam
End Function

Public Function CapturaNumero(Deslocamento As Long, Digitos As Long) As String
  Dim TopRequerimento As Long
  Dim LeftRequerimento As Long
  Dim algarismo As Long
  Dim soma As Long
  Dim indice As Long
  Dim digito As Long
  Dim letra As String
  Dim Requerimento As String
  
  Requerimento = ""
  TopRequerimento = 48 - 45
  LeftRequerimento = Deslocamento
  soma = 0
  For indice = 0 To 5
    soma = soma + formInicial.pctCopiaPartedaTela.Point(LeftRequerimento - 2, TopRequerimento + indice)
  Next indice
  If soma > 0 Then
    Exit Function
  End If
  algarismo = 0
  For digito = 0 To Digitos - 1
    soma = 0
    If formInicial.pctCopiaPartedaTela.Point(algarismo + LeftRequerimento, TopRequerimento + 2) <> 0 Then
      soma = 1
    End If
    If formInicial.pctCopiaPartedaTela.Point(algarismo + LeftRequerimento, TopRequerimento + 6) <> 0 Then
      soma = soma + 2
    End If
    If formInicial.pctCopiaPartedaTela.Point(algarismo + LeftRequerimento + 2, TopRequerimento) <> 0 Then
      soma = soma + 4
    End If
    If formInicial.pctCopiaPartedaTela.Point(algarismo + LeftRequerimento + 2, TopRequerimento + 4) <> 0 Then
      soma = soma + 8
    End If
    If formInicial.pctCopiaPartedaTela.Point(algarismo + LeftRequerimento + 2, TopRequerimento + 8) <> 0 Then
      soma = soma + 16
    End If
    If formInicial.pctCopiaPartedaTela.Point(algarismo + LeftRequerimento + 4, TopRequerimento + 2) <> 0 Then
      soma = soma + 32
    End If
    If formInicial.pctCopiaPartedaTela.Point(algarismo + LeftRequerimento + 4, TopRequerimento + 6) <> 0 Then
      soma = soma + 64
    End If
    Select Case soma
      Case 28
        letra = "1"
      Case 52
        letra = "2"
      Case 124
        letra = "3"
      Case 66
        letra = "4"
      Case 85
        letra = "5"
      Case 95
        letra = "6"
      Case 12
        letra = "7"
      Case 127
        letra = "8"
      Case 125
        letra = "9"
      Case 119
        letra = "0"
      Case 90
        letra = "INICIAL"
      Case 44
        letra = "PP"
      Case 108
        letra = "PR"
      Case 75
        letra = "NORMAL"
      Case 23
        letra = "DEFERIDO"
      Case 3
        letra = "INDEFERIDO"
      Case 15
        letra = "PENDENTE"
      Case Else
        letra = ""
    End Select
    Requerimento = Requerimento & letra
    algarismo = algarismo + 6
  Next digito
  CapturaNumero = Requerimento
End Function

Public Sub SimulaSendKeys(palavra As String)
  Dim c As New classeSendKeys
  Dim conta As Integer
  
  If palavra = "<COPIA>" Then
    c.KeyDown vbKeyControl
    c.KeyDown vbKeyA
    c.KeyUp vbKeyA
    c.KeyUp vbKeyControl
    c.KeyDown vbKeyControl
    c.KeyDown vbKeyC
    c.KeyUp vbKeyC
    c.KeyUp vbKeyControl
    Exit Sub
  End If
  If palavra = "<CONTROL>P" Then
    c.KeyDown vbKeyControl
    c.KeyDown vbKeyP
    c.KeyUp vbKeyP
    c.KeyUp vbKeyControl
    Exit Sub
  End If
  If palavra = "<SALVA>" Then
    c.KeyDown vbKeyShift
    c.KeyDown vbKeyControl
    c.KeyDown vbKeyS
    c.KeyUp vbKeyS
    c.KeyUp vbKeyControl
    c.KeyUp vbKeyShift
    Exit Sub
  End If
  If palavra = "<PASTE>" Then
    c.KeyDown vbKeyControl
    c.KeyDown vbKeyA
    c.KeyUp vbKeyA
    c.KeyUp vbKeyControl
    c.KeyDown vbKeyControl
    c.KeyDown vbKeyV
    c.KeyUp vbKeyV
    c.KeyUp vbKeyControl
    Exit Sub
    End If
  If palavra = "<TAB>" Then
    c.KeyDown vbKeyTab
    c.KeyUp vbKeyTab
    Exit Sub
  End If
  If palavra = "<UP>" Then
    c.KeyDown vbKeyUp
    c.KeyUp vbKeyUp
    Exit Sub
  End If
  If palavra = "<END>" Then
    c.KeyDown vbKeyEnd
    c.KeyUp vbKeyEnd
    Exit Sub
  End If
  If palavra = "<ENTER>" Then
    c.KeyDown vbKeyReturn
    c.KeyUp vbKeyReturn
    Exit Sub
  End If
  If GetKeyState(vbKeyCapital) = 1 Then
    For conta = 1 To Len(palavra)
      If Mid(palavra, conta, 1) = ":" Or Mid(palavra, conta, 1) = "@" Then
        c.KeyDown vbKeyShift
        c.KeyDown c.KeyCode(Mid(palavra, conta, 1))
        c.KeyUp c.KeyCode(Mid(palavra, conta, 1))
        c.KeyUp vbKeyShift
      Else
        If Mid(palavra, conta, 1) <> UCase(Mid(palavra, conta, 1)) Then
          c.KeyDown vbKeyShift
          c.KeyDown c.KeyCode(Mid(palavra, conta, 1))
          c.KeyUp c.KeyCode(Mid(palavra, conta, 1))
          c.KeyUp vbKeyShift
        Else
          c.KeyDown c.KeyCode(Mid(palavra, conta, 1))
          c.KeyUp c.KeyCode(Mid(palavra, conta, 1))
        End If
      End If
    Next conta
  Else
    For conta = 1 To Len(palavra)
      If Mid(palavra, conta, 1) = ":" Or Mid(palavra, conta, 1) = "@" Then
        c.KeyDown vbKeyShift
        c.KeyDown c.KeyCode(Mid(palavra, conta, 1))
        c.KeyUp c.KeyCode(Mid(palavra, conta, 1))
        c.KeyUp vbKeyShift
      Else
        If Mid(palavra, conta, 1) = LCase(Mid(palavra, conta, 1)) And Mid(palavra, conta, 1) <> ":" Then
          c.KeyDown c.KeyCode(Mid(palavra, conta, 1))
          c.KeyUp c.KeyCode(Mid(palavra, conta, 1))
        Else
          c.KeyDown vbKeyShift
          c.KeyDown c.KeyCode(Mid(palavra, conta, 1))
          c.KeyUp c.KeyCode(Mid(palavra, conta, 1))
          c.KeyUp vbKeyShift
        End If
      End If
    Next conta
  End If
End Sub
       
Public Function ObtemIDdoRelatórioCrystalReport() As Long
  On Error Resume Next
  Dim iddatela As Long
  Dim h1AfxWnd42 As Long
  Dim h2AfxWnd42 As Long
  Dim hAfxFrameOrView42 As Long

  ObtemIDdoRelatórioCrystalReport = 0
  iddatela = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
  Do While iddatela <> 0
    h1AfxWnd42 = FindWindowEx(iddatela, 0, "AfxWnd42", "")
    h2AfxWnd42 = FindWindowEx(h1AfxWnd42, 0, "AfxWnd42", "")
    hAfxFrameOrView42 = FindWindowEx(h2AfxWnd42, 0, "AfxFrameOrView42", "")
    If hAfxFrameOrView42 <> 0 Then
      ObtemIDdoRelatórioCrystalReport = iddatela
      Exit Function
    End If
    iddatela = GetWindow(iddatela, GW_HWNDNEXT)
  Loop
End Function

Public Sub CopiaTela(memonumerodorequerimento As String)
  On Error Resume Next
  Dim hWndActive As Long
  Dim r As Long
  Dim RectActive As RECT
    
  hWndActive = GlobalIDControleOperacional
  r = GetWindowRect(hWndActive, RectActive)
  Set formInicial.pctCopiaPartedaTela.Picture = CaptureWindow(hWndActive, False, 19, 232 - 30 + 45, 727, 16)  '_
    SavePicture formInicial.pctCopiaPartedaTela.Picture, GlobalPastadeTrabalho & "\" & GlobalDatadosRequerimentos & Format(GlobalIDRequerimento, "000") & memonumerodorequerimento & ".bmp"
End Sub

Public Function ObtemTextodoControle(ByVal hWnd As Long) As String
  On Error Resume Next
  Dim strBuff As String
  Dim lngLen As Long
  
  lngLen = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, 0)
  If lngLen > 0 Then
    lngLen = lngLen + 1
    strBuff = String(lngLen, vbNullChar)
    lngLen = SendMessage(hWnd, WM_GETTEXT, lngLen, ByVal strBuff)
    ObtemTextodoControle = Left(strBuff, lngLen)
  End If
End Function

Public Sub ClickMenu(lAplicativo As Long, lMenu As Long, lItem As Long)
  On Error Resume Next
  Dim lSubMenu  As Long
  Dim lMenuItem As Long
  Dim lIDMenu As Long

  ' This is a bit more interesting
  lIDMenu = GetMenu(lAplicativo)
  lSubMenu = GetSubMenu(lIDMenu, lMenu)
  lMenuItem = GetMenuItemID(lSubMenu, lItem)
  Call PostMessage(lAplicativo, WM_COMMAND, lMenuItem, 0)
  ' sendmessage would hang app until file is selected in open form but
  ' postmessage is asynchronous which is better in this case
End Sub

Public Function ConsultaRequerimento(NumerodoRequerimento As String, ImpressãoAutomática As Boolean) As Requerimento
  On Error Resume Next
  Dim contador As Long
  Dim DimensõesdoCampoRequerimento As RECT
  Dim posiçaodocursor As POINTAPI
  Dim IDTelaPesquisaAvançada As Long
  Dim res As String
  Dim hCampoRequerimentoPesquisaAvançada As Long
  Dim hBotãoOKPesquisaAvançada As Long
  Dim CentrodoCampoRequerimento As Long
  Dim IDTelaCarteiradeBeneficios As Long
  Dim hPrimeira As Long
  Dim NomedoControle As String
  Dim ComprimentodoNomedoControle As Long
  Dim hBotaoSairNIT As Long
  Dim hValorNIT As Long
  Dim IDTelaNITSecundario As Long
  Dim IDTelaConsultaSemCriterio As Long
  Dim hDC As Long
  Dim hW As Long
  Dim memonit As String
  Dim vernumero As String
  
  hW = GetDesktopWindow()
  hDC = GetWindowDC(hW)
  GlobalAlerta = False
  GlobalTipo = ""
        
  Dim IDTelaTempoTranscorrido As Long
  Dim IDCriticadoSABIsobreSegundaVia As Long
  Dim IDAvisoInportante As Long
  Dim IDNãoAvisoImportante As Long
  Dim memoStatus As String
  Dim memoTipo As String
  Dim sequencia As String
        
  If GlobalIDRequerimento < 100 Then
    sequencia = Format(GlobalIDRequerimento, "00")
  Else
    sequencia = Format(GlobalIDRequerimento, "000")
  End If
  formInicial.pctCopiaPartedaTela.Visible = False
  DoEvents

  'valores iniciais do requerimento
  ConsultaRequerimento.Número = NumerodoRequerimento
  ConsultaRequerimento.NIT = ""
  ConsultaRequerimento.Status = ""
  ConsultaRequerimento.Tipo = ""
  ConsultaRequerimento.Crítica = ""

  'fecha a tela de tempo transcorrido se ela ainda estiver aparecendo
  IDTelaTempoTranscorrido = ObtemIDdaTelaPrincipalporTitulo("Carteira de Benefícios")
  contador = 0
  While IDTelaTempoTranscorrido <> 0
    PostMessage IDTelaTempoTranscorrido, WM_CLOSE, 0, 0
    DoEvents
    Espera 200
    contador = contador + 1
    If contador > 25 Then
      ConsultaRequerimento.Crítica = "Tempo de 5 segundos expirados para fechar a tela 'Tempo Transcorrido'"
      Exit Function
    End If
    IDTelaTempoTranscorrido = ObtemIDdaTelaPrincipalporTitulo("Carteira de Benefícios")
  Wend

  'fecha a tela de tempo pesquisa avançada se ainda estiver aparecendo
  IDTelaPesquisaAvançada = ObtemIDdaTelaPrincipalporTitulo("Pesquisa Avançada")
  contador = 0
  While IDTelaPesquisaAvançada <> 0
    PostMessage IDTelaPesquisaAvançada, WM_CLOSE, 0, 0
    DoEvents
    Espera 200
    contador = contador + 1
    If contador > 25 Then
      ConsultaRequerimento.Crítica = "Tempo de 5 segundos expirados para fechar a tela 'Pesquisa Avançada'"
      Exit Function
    End If
    IDTelaPesquisaAvançada = ObtemIDdaTelaPrincipalporTitulo("Pesquisa Avançada")
  Wend
        
  'fecha a tela de impressao segunda via se ainda estiver aparecendo
  contador = 0
  IDFormulárioSegundaViaMarcaçãodeExame = FindWindowEx(GlobalhMDIClient, 0, "ThunderRT6FormDC", "Segunda Via de Marcação de Exame")
  While IDFormulárioSegundaViaMarcaçãodeExame <> 0
    PostMessage IDFormulárioSegundaViaMarcaçãodeExame, WM_CLOSE, 0, 0
    DoEvents
    Espera 200
    contador = contador + 1
    If contador > 25 Then
      ConsultaRequerimento.Crítica = "Tempo de 5 segundos expirados para fechar a tela 'Segunda Via de Marcação de Exame'"
      Exit Function
    End If
    IDFormulárioSegundaViaMarcaçãodeExame = FindWindowEx(GlobalhMDIClient, 0, "ThunderRT6FormDC", "Segunda Via de Marcação de Exame")
  Wend
  SetForegroundWindow (GlobalIDControleOperacional)
  
  'reafirma a posição da tela Carteira de Beneficios
   res = SetWindowPos(GlobalIDTelaConsultaRequerimentoBenefício, 0, 0, 0, 863, 521, 0)
   DoEvents
  
  'clica em <Avançado>
  PostMessage GlobalToolbarConsultaRequerimentoOCX, WM_COMMAND, 100 + 2 - 1, ByVal GlobalToolbarConsultaRequerimento
  DoEvents
        
  'espera a tela ser montada
  Espera 200
        
  'verifica se a tela Pesquiva Avançada apareceu
  contador = 0
  While IDTelaPesquisaAvançada = 0
    IDTelaPesquisaAvançada = ObtemIDdaTelaPrincipalporTitulo("Pesquisa Avançada")
    If IDTelaPesquisaAvançada <> 0 Then
      'coloca tela Pesquisa Avançada no canto superior a direita do grid
      res = SetWindowPos(IDTelaPesquisaAvançada, 0, 863 + 10, 80, 863, 521, 0)
      res = SetWindowPos(IDTelaPesquisaAvançada, 0, 0, 80, 863, 521, 0)
    End If
    Espera 200
    DoEvents
    contador = contador + 1
    If contador > 50 Then
      ConsultaRequerimento.Crítica = "Tempo de 10 segundos expirados para a tela 'Pesquisa Avançada' aparecer"
      Exit Function
    End If
  Wend
  
  'A tela Pesquisa Avançada apareceu
  'encontra campo requerimento
  contador = 0
  hBotãoOKPesquisaAvançada = 0
  hCampoRequerimentoPesquisaAvançada = FindWindowEx(IDTelaPesquisaAvançada, 0, "ThunderRT6CommandButton", "&OK")
  Do While hCampoRequerimentoPesquisaAvançada = 0 Or hBotãoOKPesquisaAvançada = 0
    Espera 200
    contador = contador + 1
    If contador > 25 Then
      ConsultaRequerimento.Crítica = "Tempo de 5 segundos expirado para encontrar o botão 'OK' e o campo 'Requerimento' na tela 'Pesquisa Avançada'"
      GoTo FechaaTelaPesquisaAvançada
    End If
    hBotãoOKPesquisaAvançada = FindWindowEx(IDTelaPesquisaAvançada, 0, "ThunderRT6CommandButton", "&OK")
    hCampoRequerimentoPesquisaAvançada = FindWindowEx(IDTelaPesquisaAvançada, 0, "MSMaskWndClass", "")
  Loop
        
  'campo encontrado, comanda mouse clique no centro
  res = GetWindowRect(hCampoRequerimentoPesquisaAvançada, DimensõesdoCampoRequerimento)
  CentrodoCampoRequerimento = convlong(DimensõesdoCampoRequerimento.Left + (DimensõesdoCampoRequerimento.Right - DimensõesdoCampoRequerimento.Left) / 2, DimensõesdoCampoRequerimento.Top + (DimensõesdoCampoRequerimento.Bottom - DimensõesdoCampoRequerimento.Top) / 2)
  SendMessage hCampoRequerimentoPesquisaAvançada, WM_LBUTTONDOWN, MK_LBUTTON, CentrodoCampoRequerimento
  SendMessage hCampoRequerimentoPesquisaAvançada, WM_LBUTTONUP, MK_LBUTTON, CentrodoCampoRequerimento
  DoEvents
  Espera 100
        
  'registra o valor do requerimento
  SimulaSendKeys "123"
  Espera 200
  SendMessage2 hCampoRequerimentoPesquisaAvançada, WM_SETTEXT, 0, NumerodoRequerimento & Chr$(0)
  DoEvents
  Espera 600
        
  'comanda o fechamento tela de critica do SABI se ainda estiver aberta
  IDCriticadoSABIsobreSegundaVia = FindWindow("#32770", "SABI - Controle Operacional")
  If IDCriticadoSABIsobreSegundaVia <> 0 Then
    PostMessage IDCriticadoSABIsobreSegundaVia, WM_CLOSE, 0, 0
    DoEvents
    Espera 300
  End If
  PostMessage hBotãoOKPesquisaAvançada, BM_CLICK, 0, 0
  DoEvents
  'DoEvents

  'VERIFICA O NUMERO DO REQUERIMENTO
  contador = 0
  vernumero = 0
  While vernumero <> NumerodoRequerimento
    Espera 200
            
    'verifica se apareceu critica de consulta sem criterio
    IDAvisoInportante = FindWindow("#32770", "AVISO IMPORTANTE")
    If IDAvisoInportante <> 0 Then
      'espera a tela ser montada
      Espera 300
      IDNãoAvisoImportante = FindWindowEx(IDAvisoInportante, 0, "Button", "&Não")
      If IDNãoAvisoImportante <> 0 Then
        PostMessage IDNãoAvisoImportante, BM_CLICK, 0, 0
        ConsultaRequerimento.Crítica = "Esta pesquisa está sendo executada sem nenhum critério."
        formInicial.RequerimentonãoEncontrado NumerodoRequerimento, sequencia
        GoTo FechaaTelaPesquisaAvançada
        Exit Function
      End If
    End If
    CopiaTela (NumerodoRequerimento)
    DoEvents
    formInicial.pctCopiaPartedaTela.Top = 0
    formInicial.pctCopiaPartedaTela.Left = 0
    DoEvents
    vernumero = CapturaNumero(3, 9)
    contador = contador + 1
    If contador > 100 Then
      formInicial.RequerimentonãoEncontrado NumerodoRequerimento, sequencia
      ConsultaRequerimento.Crítica = "Tempo expirado de 20 segundos para aparecerem as informações do requerimento"
      GoTo EsperaaTelaPesquisaAvançadaSerFechada
    End If
  Wend
  DoEvents
  Espera 100
  memonit = CapturaNumero(453, 11)
  If memonit = "" Or memonit = "88888888888" Then memonit = ""
  ConsultaRequerimento.NIT = memonit
  memoTipo = CapturaNumero(118, 1)
  If memoTipo = "INICIAL" Or memoTipo = "PP" Or memoTipo = "PR" Then
    ConsultaRequerimento.Tipo = memoTipo
  Else
    ConsultaRequerimento.Tipo = ""
  End If
        
  memoStatus = CapturaNumero(180, 1)
  If memoStatus = "NORMAL" Or memoStatus = "PENDENTE" Or memoStatus = "DEFERIDO" Or memoStatus = "INDEFERIDO" Then
    ConsultaRequerimento.Status = memoStatus
  Else
    ConsultaRequerimento.Status = ""
  End If
  If memoTipo = "INICIAL" And memoStatus = "NORMAL" Then
    formInicial.efeitos True, sequencia
  Else
    formInicial.efeitos False, sequencia
  End If
  GoTo FechaaTelaPesquisaAvançada
  Exit Function
  
'GOTOs da vida...

  'comanda o fechamento da tela Pesquisa Avançada
FechaaTelaPesquisaAvançada:
  PostMessage IDTelaPesquisaAvançada, WM_CLOSE, 0, 0
  DoEvents
  Espera 200

'espera a tela ser fechada
EsperaaTelaPesquisaAvançadaSerFechada:
  contador = 0
  IDTelaPesquisaAvançada = ObtemIDdaTelaPrincipalporTitulo("Pesquisa Avançada")
  Do While IDTelaPesquisaAvançada <> 0
    PostMessage IDTelaPesquisaAvançada, WM_CLOSE, 0, 0
    DoEvents
    Espera 200
    contador = contador + 1
    If contador > 50 Then
      'depois de 10 segundos sai de qualquer jeito
      Exit Function
    End If
    IDTelaPesquisaAvançada = ObtemIDdaTelaPrincipalporTitulo("Pesquisa Avançada")
  Loop
End Function

Public Function ImprimeSegundaViadoRequerimento(NumerodoNIT As String, ImpressãoAutomática As Boolean) As Requerimento
  On Error Resume Next
  Dim NITsemPonto As String
  Dim contador As Long
  Dim hThunderRT6FrameIMPRIME  As Long
  Dim hImMaskWndCIassIMPRIME As Long
  Dim IDFormulárioSegundaViaMarcaçãodeExame As Long
  Dim hThunderRT6CommandButtonVisualizar As Long
  Dim hThunderRT6CommandButtonCancelar As Long
  Dim hThunderRT6CommandButtonImprimir As Long
  Dim DimensõesdaTeclaVizualizar As RECT
  Dim DimensõesdaTeclaCancelar As RECT
  Dim DimensõesdaTeclaImprimir As RECT
  Dim NomedoControle As String
  Dim res As String
  Dim DimensõesdoCampoNIT As RECT
  Dim CentrodoCampoNIT As Long
  Dim PosiçaodoCursorPressionado As POINTAPI
  Dim PosiçaodoCursorLiberado As POINTAPI
  Dim BotãoImprimirHabilitado As Boolean
  Dim TempoBotãoDisponibilizado As Long
  Dim IDCriticadoSABIsobreSegundaVia As Long
  Dim IDTelaImprimindo As Long
  Dim IDRelatórioSegundaVia As Long
  Dim Hstatic As Long
  Dim IDTelaTempoTranscorrido As Long
        
  formInicial.pctCopiaPartedaTela.Visible = False
  DoEvents
  ImprimeSegundaViadoRequerimento.Impresso = "NÃO"

  'fecha a tela de tempo transcorrido se ainda está aparecendo
  IDTelaTempoTranscorrido = ObtemIDdaTelaPrincipalporTitulo("Carteira de Benefícios")
  While IDTelaTempoTranscorrido <> 0
    PostMessage IDTelaTempoTranscorrido, WM_CLOSE, 0, 0
    Espera 200
    DoEvents
    DTelaTempoTranscorrido = ObtemIDdaTelaPrincipalporTitulo("Carteira de Benefícios")
  Wend
        
  'fecha tela de critica do SABI se ainda estiver aberta
  IDCriticadoSABIsobreSegundaVia = FindWindow("#32770", "SABI - Controle Operacional")
  While IDCriticadoSABIsobreSegundaVia <> 0
    PostMessage IDCriticadoSABIsobreSegundaVia, WM_CLOSE, 0, 0
    Espera 200
    DoEvents
    IDCriticadoSABIsobreSegundaVia = FindWindow("#32770", "SABI - Controle Operacional")
  Wend
  ImprimeSegundaViadoRequerimento.NIT = NumerodoNIT

  'abre tela Segunda Via de Marcação de Exame
  ClickMenu GlobalIDControleOperacional, 4, 7
  DoEvents
  Espera 300
  
  'espera a tela aparecer
  contador = 0
  IDFormulárioSegundaViaMarcaçãodeExame = 0
  While IDFormulárioSegundaViaMarcaçãodeExame = 0
    IDFormulárioSegundaViaMarcaçãodeExame = FindWindowEx(GlobalhMDIClient, 0, "ThunderRT6FormDC", "Segunda Via de Marcação de Exame")
    If IDFormulárioSegundaViaMarcaçãodeExame <> 0 Then
      res = SetWindowPos(IDFormulárioSegundaViaMarcaçãodeExame, 0, 0, 0, 750, 150, 0)
    End If
    DoEvents
    Espera 200
    contador = contador + 1
    If contador > 100 Then
      ImprimeSegundaViadoRequerimento.Crítica = "Tempo de 20 segundos expirado para aparecer a tela 'Segunda Via de Marcação de Exame'"
      Exit Function
    End If
  Wend

  'a tela apareceu
  'coloca a tela no canto superior esquerdo
  'res = SetWindowPos(IDFormulárioSegundaViaMarcaçãodeExame, 0, 0, 0, 750, 150, 0)
  DoEvents
       
  'Espera o mouse ser liberado o usuario pode estar arrastando a tela
  While GetAsyncKeyState(VK_LBUTTON) = True
  Wend
  SetForegroundWindow (GlobalIDControleOperacional)
        
  'encontra as teclas Visualizar, Cancelar e Imprimir
  contador = 0
  hThunderRT6CommandButtonVisualizar = 0
  hThunderRT6CommandButtonCancelar = 0
  hThunderRT6CommandButtonImprimir = 0
  While hThunderRT6CommandButtonVisualizar = 0 Or hThunderRT6CommandButtonCancelar = 0 Or hThunderRT6CommandButtonImprimir = 0
    hThunderRT6CommandButtonVisualizar = FindWindowEx(IDFormulárioSegundaViaMarcaçãodeExame, 0, "ThunderRT6CommandButton", "&Visualizar")
    hThunderRT6CommandButtonCancelar = FindWindowEx(IDFormulárioSegundaViaMarcaçãodeExame, 0, "ThunderRT6CommandButton", "&Cancelar")
    hThunderRT6CommandButtonImprimir = FindWindowEx(IDFormulárioSegundaViaMarcaçãodeExame, 0, "ThunderRT6CommandButton", "&Imprimir")
    contador = contador + 1
    If contador > 5000 Then
      ImprimeSegundaViadoRequerimento.Crítica = "Tempo expirado para encontrar as teclas 'Visualizar', 'Imprimir' e 'Cancelar' na tela 'Segunda Via de Marcação de Exame'"
      
      'fecha a tela Segunda Via de Marcação de Exame
      GoTo FechaaTelaSegundaVia
    End If
  Wend
        
  'encontra o recipiente do campo de NIT
  contador = 0
  hThunderRT6FrameIMPRIME = 0
  While hThunderRT6FrameIMPRIME = 0
    hThunderRT6FrameIMPRIME = FindWindowEx(IDFormulárioSegundaViaMarcaçãodeExame, 0, "ThunderRT6Frame", "NIT Requerente")
    contador = contador + 1
    If contador > 5000 Then
      ImprimeSegundaViadoRequerimento.Crítica = "Tempo expirado para encontrar o recipiente do campo 'NIT' na tela 'Segunda Via de Marcação de Exame'"
      
      'fecha a tela Segunda Via de Marcação de Exame
      GoTo FechaaTelaSegundaVia
    End If
  Wend
        
  'encontra o campo de NIT
  contador = 0
  hImMaskWndCIassIMPRIME = 0
  Do While hImMaskWndCIassIMPRIME = 0
    hImMaskWndCIassIMPRIME = FindWindowEx(hThunderRT6FrameIMPRIME, 0, vbNullString, vbNullString)
    If hImMaskWndCIassIMPRIME <> 0 Then
      NomedoControle = Space(100)
      res = GetClassName(hImMaskWndCIassIMPRIME, NomedoControle, 100)
      If Mid(NomedoControle, 1, 14) = "ImMaskWndClass" Then
        Exit Do
      Else
        ImprimeSegundaViadoRequerimento.Crítica = "Tempo expirado pare encontrar o nome do campo 'NIT' na tela 'Segunda Via de Marcação de Exame'"
        
        'fecha a tela Segunda Via de Marcação de Exame
        GoTo FechaaTelaSegundaVia
      End If
    End If
    contador = contador + 1
    If contador > 5000 Then
      ImprimeSegundaViadoRequerimento.Crítica = "Tempo expirado pare encontrar o campo 'NIT' na tela 'Segunda Via de Marcação de Exame'"
                
      'fecha a tela Segunda Via de Marcação de Exame
      GoTo FechaaTelaSegundaVia
    End If
  Loop
        
  'obtem as dimensões do campo NIT
  res = GetWindowRect(hImMaskWndCIassIMPRIME, DimensõesdoCampoNIT)
  CentrodoCampoNIT = convlong(DimensõesdoCampoNIT.Left + (DimensõesdoCampoNIT.Right - DimensõesdoCampoNIT.Left) / 2, DimensõesdoCampoNIT.Top + (DimensõesdoCampoNIT.Bottom - DimensõesdoCampoNIT.Top) / 2)
  contador = 0
  
  'verifica o valor presente no campo NIT
  Do While ObtemTextodoControle(hImMaskWndCIassIMPRIME) <> Mid(NITsemPonto, 1, 10) & "-" & Mid(NITsemPonto, 11, 1)
    'clica no campo NIT para atribuir foco
    SendMessage hImMaskWndCIassIMPRIME, WM_LBUTTONDOWN, MK_LBUTTON, CentrodoCampoNIT
    SendMessage hImMaskWndCIassIMPRIME, WM_LBUTTONUP, MK_LBUTTON, CentrodoCampoNIT
    DoEvents
    Espera 100
            
    'reformata o valor do NIT
    NITsemPonto = ImprimeSegundaViadoRequerimento.NIT
    While InStr(1, NITsemPonto, ".")
      NITsemPonto = Mid(NITsemPonto, 1, InStr(1, NITsemPonto, ".") - 1) & Mid(NITsemPonto, InStr(1, NITsemPonto, ".") + 1)
    Wend
            
    'digita o valor do NIT sem o sinal - e o ultimo algarismo
    SimulaSendKeys Mid(NITsemPonto, 1, Len(NITsemPonto) - 1)
            
    'Digita o último algarismo
    SimulaSendKeys Right$(NITsemPonto, 1)
    DoEvents
    Espera 200
    contador = contador + 1
    If contador > 10 Then
      ImprimeSegundaViadoRequerimento.Crítica = "Não foi possível registrar o NIT '" & ImprimeSegundaViadoRequerimento.NIT & "'"
           
      'fecha a tela Segunda Via de Marcação de Exame
      GoTo FechaaTelaSegundaVia
    End If
  Loop
        
  'Epera os botões serem habilitados
  contador = 0
  BotãoImprimirHabilitado = False
  While BotãoImprimirHabilitado = False
    'verifica se o SABI criticou o NIT
    IDCriticadoSABIsobreSegundaVia = FindWindow("#32770", "SABI - Controle Operacional")
    If IDCriticadoSABIsobreSegundaVia <> 0 Then
      Espera 300
      Hstatic = FindWindowEx(IDCriticadoSABIsobreSegundaVia, 0, "Static", "")
      Hstatic = FindWindowEx(IDCriticadoSABIsobreSegundaVia, Hstatic, vbNullString, vbNullString)
      ImprimeSegundaViadoRequerimento.Crítica = ObtemTextodoControle(Hstatic)
                
      'fecha a tela de crítica do SABI
      SendMessage IDCriticadoSABIsobreSegundaVia, WM_CLOSE, 0, 0
             
      'fecha a tela Segunda Via de Marcação de Exame
      GoTo FechaaTelaSegundaVia
    End If
    BotãoImprimirHabilitado = Val(GetWindowLong(hThunderRT6CommandButtonImprimir, GWL_STYLE) And WS_DISABLED) = 0
    Espera 200
    contador = contador + 1
    If contador > 100 Then
      ImprimeSegundaViadoRequerimento.Crítica = "Tempo de 10 segundos expirado para o SABI habilitar o botão 'Imprimir'"
              
      'fecha a tela Segunda Via de Marcação de Exame
      GoTo FechaaTelaSegundaVia
    End If
  Wend
        
  'clica programaticamente no botão 'Imprimir'
  If GlobalModoSimulado Then
    PostMessage hThunderRT6CommandButtonCancelar, BM_CLICK, 0, 0
  Else
    PostMessage hThunderRT6CommandButtonImprimir, BM_CLICK, 0, 0
  End If
  DoEvents
  ImprimeSegundaViadoRequerimento.Impresso = "SIM"
        
  'espera a tela ser fechada
  GoTo EsperaaTelaSegundaViaSerFechada
  Exit Function
        
'GOTOs da vida...

'comanda o fechamento da tela Segunda Via de Marcação de Exame
FechaaTelaSegundaVia:
  PostMessage IDFormulárioSegundaViaMarcaçãodeExame, WM_CLOSE, 0, 0

'espera a tela ser fechada
EsperaaTelaSegundaViaSerFechada:
  contador = 0
  While IDFormulárioSegundaViaMarcaçãodeExame <> 0
    Espera 200
    PostMessage IDFormulárioSegundaViaMarcaçãodeExame, WM_CLOSE, 0, 0
    IDFormulárioSegundaViaMarcaçãodeExame = FindWindowEx(GlobalhMDIClient, 0, "ThunderRT6FormDC", "Segunda Via de Marcação de Exame")
    DoEvents
        
    'verifica se o SABI criticou o NIT
    IDCriticadoSABIsobreSegundaVia = FindWindow("#32770", "SABI - Controle Operacional")
    If IDCriticadoSABIsobreSegundaVia <> 0 Then
      Espera 300
      Hstatic = FindWindowEx(IDCriticadoSABIsobreSegundaVia, 0, "Static", "")
      Hstatic = FindWindowEx(IDCriticadoSABIsobreSegundaVia, Hstatic, vbNullString, vbNullString)
      ImprimeSegundaViadoRequerimento.Crítica = ObtemTextodoControle(Hstatic)
      ImprimeSegundaViadoRequerimento.Impresso = "NÃO"
      
      'fecha a tela de crítica do SABI
      SendMessage IDCriticadoSABIsobreSegundaVia, WM_CLOSE, 0, 0
               
      'fecha a tela Segunda Via de Marcação de Exame
      GoTo FechaaTelaSegundaVia
    End If
    If contador > 50 Then
      'depois de 10 segundos sai de qualquer jeito
      Exit Function
    End If
  Wend
  'espera tela Imprimindo... abrir
  IDTelaImprimindo = FindWindow("#32770", "Imprimindo...")
  contador = 0
  While IDTelaImprimindo = 0
    Espera 100
    contador = contador + 1
    If contador > 50 Then
      Exit Function
    End If
    IDTelaImprimindo = FindWindow("#32770", "Imprimindo...")
  Wend
  contador = 0
  hThunderRT6CommandButtonImprimir = 0
  While hThunderRT6CommandButtonImprimir = 0
    hThunderRT6CommandButtonImprimir = FindWindowEx(IDTelaImprimindo, 0, "Button", "&Sim")
    Espera 100
    contador = contador + 1
    If contador > 5000 Then
      Exit Function
    End If
  Wend
  PostMessage hThunderRT6CommandButtonImprimir, BM_CLICK, 0, 0
  
  'espera tela Imprimindo... fechar
  IDTelaImprimindo = FindWindow("#32770", "Imprimindo...")
  contador = 0
  While IDTelaImprimindo <> 0
    contador = contador + 1
    Espera 100
    If contador > 50 Then
      Exit Function
    End If
    IDTelaImprimindo = FindWindow("#32770", "Imprimindo...")
  Wend
End Function

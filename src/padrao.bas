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
    
Public Const MF_ENABLED = &H0&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const WM_NCPAINT = &H85
Public Const BM_CLICK = &HF5
Public Const VK_LBUTTON = &H1
Public Const WS_DISABLED As Long = &H8000000
Public Const GWL_STYLE As Long = -16
Public Const WM_COMMAND = &H111
Public Const BN_CLICKED = 0
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_ABSOLUTE = &H8000
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10
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
Public GlobalImpressaoAuto As Boolean
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

Public Sub esperarProcesso(tempo As Long)
  Dim memotempo As Long
  Dim segundos As Long
  Dim minutos As Long
  Dim handleJanela As Long
  Dim res As String

  memotempo = 0
  handleJanela = GetForegroundWindow
  res = UCase(getControleTexto(handleJanela))
  If InStr(1, res, "SALVAR COMO") Then
    res = SetWindowPos(handleJanela, 0, 0, 0, 460, 340, 0)
    res = SetWindowPos(handleJanela, 1, 0, 0, 0, 0, 3)
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

'------------------------------------------------------------------------

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
     
Public Function convlong(xini As Long, yini As Long)
  Dim lParam As Long
  
  lParam = 256 * 64
  lParam = yini * (lParam * 4) + xini
  convlong = lParam
End Function

Public Function CapturaNumero(deslocamento As Long, Digitos As Long) As String
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
  LeftRequerimento = deslocamento
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

Public Sub simularTeclado(palavra As String)
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
  ConsultaRequerimento.nit = ""
  ConsultaRequerimento.Status = ""
  ConsultaRequerimento.Tipo = ""
  ConsultaRequerimento.Crítica = ""

  'fecha a tela de tempo transcorrido se ela ainda estiver aparecendo
  IDTelaTempoTranscorrido = ObtemIDdaTelaPrincipalporTitulo("Carteira de Benefícios")
  contador = 0
  While IDTelaTempoTranscorrido <> 0
    PostMessage IDTelaTempoTranscorrido, WM_CLOSE, 0, 0
    DoEvents
    espera 200
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
    espera 200
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
    espera 200
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
  espera 200
        
  'verifica se a tela Pesquiva Avançada apareceu
  contador = 0
  While IDTelaPesquisaAvançada = 0
    IDTelaPesquisaAvançada = ObtemIDdaTelaPrincipalporTitulo("Pesquisa Avançada")
    If IDTelaPesquisaAvançada <> 0 Then
      'coloca tela Pesquisa Avançada no canto superior a direita do grid
      res = SetWindowPos(IDTelaPesquisaAvançada, 0, 863 + 10, 80, 863, 521, 0)
      res = SetWindowPos(IDTelaPesquisaAvançada, 0, 0, 80, 863, 521, 0)
    End If
    espera 200
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
    espera 200
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
  SendMessage hCampoRequerimentoPesquisaAvançada, WM_LBUTTONDOWN, MK_LBUTTON, (CentrodoCampoRequerimento)
  SendMessage hCampoRequerimentoPesquisaAvançada, WM_LBUTTONUP, MK_LBUTTON, (CentrodoCampoRequerimento)
  DoEvents
  espera 100
        
  'registra o valor do requerimento
  SimulaSendKeys "123"
  espera 200
  SendMessage hCampoRequerimentoPesquisaAvançada, WM_SETTEXT, 0, NumerodoRequerimento & Chr$(0)
  DoEvents
  espera 600
        
  'comanda o fechamento tela de critica do SABI se ainda estiver aberta
  IDCriticadoSABIsobreSegundaVia = FindWindow("#32770", "SABI - Controle Operacional")
  If IDCriticadoSABIsobreSegundaVia <> 0 Then
    PostMessage IDCriticadoSABIsobreSegundaVia, WM_CLOSE, 0, 0
    DoEvents
    espera 300
  End If
  PostMessage hBotãoOKPesquisaAvançada, BM_CLICK, 0, 0
  DoEvents
  'DoEvents

  'VERIFICA O NUMERO DO REQUERIMENTO
  contador = 0
  vernumero = 0
  While vernumero <> NumerodoRequerimento
    espera 200
            
    'verifica se apareceu critica de consulta sem criterio
    IDAvisoInportante = FindWindow("#32770", "AVISO IMPORTANTE")
    If IDAvisoInportante <> 0 Then
      'espera a tela ser montada
      espera 300
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
  espera 100
  memonit = CapturaNumero(453, 11)
  If memonit = "" Or memonit = "88888888888" Then memonit = ""
  ConsultaRequerimento.nit = memonit
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
  espera 200

'espera a tela ser fechada
EsperaaTelaPesquisaAvançadaSerFechada:
  contador = 0
  IDTelaPesquisaAvançada = ObtemIDdaTelaPrincipalporTitulo("Pesquisa Avançada")
  Do While IDTelaPesquisaAvançada <> 0
    PostMessage IDTelaPesquisaAvançada, WM_CLOSE, 0, 0
    DoEvents
    espera 200
    contador = contador + 1
    If contador > 50 Then
      'depois de 10 segundos sai de qualquer jeito
      Exit Function
    End If
    IDTelaPesquisaAvançada = ObtemIDdaTelaPrincipalporTitulo("Pesquisa Avançada")
  Loop
End Function


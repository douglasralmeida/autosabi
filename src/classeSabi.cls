VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classeSabi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' sabi.bas
' Classe para Manipulação do SABI

'Constantes
Private Const NOMEJANELASABI = "SABI - Módulo de Controle Operacional"
Private Const NOMECARTEIRABEN = "Consulta Requerimento/Benefício"
Private Const NOMEIMPRIMIRAGENDA = "Imprimir Agendamento"

Private Const MSG_ERRO_SABILENTO = "Parece que o SABI está muito lento no momento. Tente executar este processamento mais tarde."
Private Const MSG_ERRO_SEGVIALENTO = "A janela 'Segunda Via de Marcação de Exame' demorou mais de 20 segundos para aparecer."
Private Const MSG_ERRO_CONTROLESSEGVIA = "Os botões 'Visualizar', 'Imprimir' e 'Cancelar' na janela 'Segunda Via de Marcação de Exame' demoraram mais de 20 segundos para serem encontradas."

'Variáveis
Public erro As Boolean
Public erroDescricao As String
Private agenciaEscolhida As String
Private botaoConsultaRequerimento As Long
Private cabecalhoAgenda As String
Private caminhoArquivoAgenda As String
Private dataEscolhida As String
Private horadeInicio As Date
Private inicioProcessamento As Long
Private janelaImprmirAgendamento As Long
Private janelaListaAgendamentos As Long
Private janelaCarteiraBen As Long
Private janelaSabi As Long
Private janelaMDIClient As Long
Private nomeSecao As String
Private relatorioNome As String

'Funcoes da classe
Private Sub Class_Initialize()
  erro = False
  erroDescricao = ""
  botaoConsultaRequerimento = 0
  janelaImprmirAgendamento = 0
  janelaCarteiraBen = 0
  janelaMDICliente = 0
  janelaSabi = 0
  checarSabiAberto
End Sub

Private Sub checarSabiAberto()
  janelaSabi = getJanelaPrincipalIDporTitulo(NOMEJANELASABI)
End Sub

Public Function estaAberto() As Boolean
  If janelaSabi <> 0 Then checarSabiAberto
  estaAberto = janelaSabi <> 0
End Function

Private Function janelaMaximizar() As Boolean
  On Error Resume Next
  
  If janelaSabi Then
    If janelaTrazerParaFrente(janelaSabi, SW_MAXIMIZE) Then
      janelaMaximizar = True
    Else
      erro = "Falha ao maximizar a janela do Controle Operacional."
      janelaMaximizar = False
    End If
  Else
    erro = "A janela do Controle Operacional não foi encontrada."
    janelaMaximizar = False
  End If
End Function

Public Function prepararAmbiente() As Boolean
  On Error Resume Next
  Dim res As String
  Dim barraFerramentasPesquisa As Long
  Dim dimensoesClienteMDI As RECT
  Dim dimensoesJanela As RECT
  Dim idJanela As Long
  Dim idJanelaAApagar As Long
  Dim ocxConsultaRequerimento As Long
    
  'limpa todas telas internas do SABI
  fecharJanela "Imprimir Agendamento"
  fecharJanela "Imprimir Escala"
  fecharJanela "Marcação da Avaliação Social"
  fecharJanela "Segunda Via de Carta de Exigência"
  fecharJanela "Pesquisa de Requerente"
  
  'procura pelo manipulador da janela MDI Client do SABI
  janelaMDICliente = pesquisarJanela(janelaSabi, 0, "MDIClient", "")
  If janelaMDICliente = 0 Then
    prepararAmbiente = False
    erro = True
    erroDescricao = "O identificador da janela do Controle Operacional do SABI não foi encontrado."
    End
  End If
  
  ' Fecha todas as janelas MDI filhas
  idJanela = 0
  idJanela = pesquisarJanela(janelaMDICliente, 0, vbNullString, vbNullString)
  While idJanela <> 0
    If idJanela <> 0 Then idJanelaAApagar = idJanela
    DoEvents
    esperarProcesso 300
    idJanela = pesquisarJanela(janelaMDICliente, idJanela, vbNullString, vbNullString)
    If idJanelaAApagar <> 0 Then fecharJanelaPorId idJanelaAApagar
  Wend
    
  'todas as janelas internas foram fechadas
  DoEvents
  esperarProcesso 300
    
  'maximiza a janela do SABI
  If Not janelaMaximizar Then
    prepararAmbiente = False
    erro = True
    End Function
  End If
  esperarProcesso 1000
  prepararAmbiente = True
End Function
  
'abre a janela Imprimir Agendamento
Public Function abrirJanelaAgendamentos() As Boolean
  Dim contador As Long
  Dim res As Long
  
  clicarMenu 4, 0
  DoEvents
  esperarProcesso 300
  janelaImprmirAgendamento = 0
  contador = 0
  While janelaImprmirAgendamento = 0
    janelaImprmirAgendamento = getJanelaPrincipalIDporTitulo(NOMEIMPRIMIRAGENDA)
    esperarProcesso 300
    DoEvents
    contador = contador + 1
    If contador > 400 Then
      erro = True
      abrirJanelaAgendamentos = False
      erroDescricao = MSG_ERRO_SABILENTO
      End
    End If
  Wend

  ' move a janela para esquerda e a diminui de tamanho
  res = GetWindowRect(janelaImprmirAgendamento, dimensoesJanela)
  res = SetWindowPos(janelaImprmirAgendamento, 0, dimensoesJanela.Left + 120, dimensoesJanela.Top, dimensoesJanela.Right - dimensoesJanela.Left - 60, dimensoesJanela.Bottom - dimensoesJanela.Top, 0)
  
  abrirJanelaAgendamentos = True
End Function

' abre a janela Carteira de Benefícios
Public Function abrirJanelaCarteira() As Boolean
  Dim contador As Long
  Dim idJanela As Long
  
  idJanela = 0
  
  ' clica no menu correspondente
  clicarMenu 2, 0    'terceiro menu, primeiro item
  esperarProcesso 3000
  
  ' espera a janela abrir
  contador = 0
  While idJanela = 0
    idJanela = pesquisarJanela(janelaMDICliente, 0, "ThunderRT6FormDC", NOMECARTEIRABEN)
    esperarProcesso 300
    DoEvents
    contador = contador + 1
    If contador > 400 Then
      abrirJanelaCarteira = False
      erro = True
      erroDescricao = MSG_ERRO_SABILENTO
      End
    End If
  Wend
  janelaCarteiraBen = 0
  While janelaCarteiraBen = 0
    esperarProcesso 300
    DoEvents
    janelaCarteiraBen = pesquisarJanelaInterna(NOMECARTEIRABEN)
  Wend
  
  ' procura pelo botão de pesquisa da certeira de benefícios
  barraFerramentasPesquisa = pesquisarJanela(janelaCarteiraBen, 0, "Toolbar20WndClass", "")
  ocxConsultaRequerimento = pesquisarJanela(janelaCarteiraBen, barraFerramentasPesquisa, "Toolbar20WndClass", "")
  esperarProcesso 300
  botaoConsultaRequerimento = pesquisarJanela(ocxConsultaRequerimento, 0, "msvb_lib_toolbar", vbNullString)
    
  'move a janela da careteira para baixo da tela
  res = GetWindowRect(janelaMDICliente, dimensoesClienteMDI)
  res = GetWindowRect(janelaCarteiraBen, dimensoesJanela)
  res = SetWindowPos(janelaCarteiraBen, 0, dimensoesJanela.Left, dimensoesJanela.Top + (dimensoesClienteMDI.Bottom - dimensoesClienteMDI.Top - 100), dimensoesJanela.Right - dimensoesJanela.Left, dimensoesJanela.Bottom - dimensoesJanela.Top, 0)
  DoEvents
  esperarProcesso 300
  
  abrirJanelaCarteira = True
End Function

' Aguarda o usuário digitar as data da agenda e a processa
Public Function processarDataAgenda() As Boolean
  Dim centroControleDataPara As Long
  Dim comboBobxAgencia As Long
  Dim contador As Long
  Dim dataEscolhida As String
  Dim dimensoesControleDataPara As RECT
  Dim dimensoesJanela As RECT
  Dim editDataDe As Long
  Dim editDataPara As Long
  Dim res As String
  Dim textoDataDe As String
  Dim textoDataPara As String
  
  processarDataAgenda = False

  ' procura pelos controles na janela de imprimir agendamento
  comboBobxAgencia = pesquisarJanela(janelaImprmirAgendamento, 0, "ThunderRT6ComboBox", "")
  comboBobxAgencia = pesquisarJanela(janelaImprmirAgendamento, comboBobxAgencia, "ThunderRT6ComboBox", "")
  editDataDe = pesquisarJanela(janelaImprmirAgendamento, 0, "MSMaskWndClass", "")
  editDataPara = pesquisarJanela(janelaImprmirAgendamento, editDataDe, "MSMaskWndClass", "")
  
  ' espera o usuário digitar a data da agenda
  contador = 0
  Do While janelaImprmirAgendamento > 0
    agenciaEscolhida = getControleTexto(comboBobxAgencia)
    textoDataDe = getControleTexto(editDataDe)
    dataEscolhida = "01/01/1900"
    
    ' escolheu uma data
    If Len(textoDataDe) = 10 And textoDataDe <> "  /  /" Then
      If IsDate(textoDataDe) Then
        dataEscolhida = textoDataDe
        
        'data precisa ser igual ou maior que hoje limitada a 180 dias a partir de hoje
        If dataEscolhida >= Date And dataEscolhida < DateAdd("d", Date, 180) Then
        
          'coloca a data final igual a data inicial
          textoDataPara = getControleTexto(editDataPara)
          If Len(textoDataPara) = 10 And textoDataPara <> "  /  /    " Then
            dataEscolhidaPara = textoDataPara
          Else
            dataEscolhidaPara = "01/01/1900"
          End If
          If dataEscolhidaPara <> dataEscolhida Then
            setControleMaskTexto edutDataPara, Format(dataEscolhida, "dd/mm/yyyy") & Chr$(0)
          End If
        Else
          dataEscolhida = "01/01/1900"
        End If
      End If
    End If
    If dataEscolhida = "01/01/1900" Then
      setControleTexto dataEscolhidaPara, "  /  /    " & Chr$(0)
      dataEscolhida = "01/01/1900"
    End If
    
    ' se estiver tudo certo, mostra o botão exibir agenda
    If GlobalAgenciaEscolhida <> "" And dataEscolhida <> "01/01/1900" And textoDataDe <> "01/01/1900" And dataEscolhida = textoDataDe Then
      res = SetWindowPos(janelaImprmirAgendamento, 0, dimensoesJanela.Left + 120, dimensoesJanela.Top, dimensoesJanela.Right - dimensoesJanela.Left, dimensoesJanela.Bottom - dimensoesJanelar.Top, 0)
      Exit Do
    Else
      res = SetWindowPos(janelaImprmirAgendamento, 0, dimensoesJanela.Left + 120, dimensoesJanela.Top, dimensoesJanela.Right - dimensoesJanela.Left - 60, dimensoesJanela.Bottom - dimensoesJanela.Top, 0)
    End If
  
    contador = contador + 1
    If contador > 1000 Then
      erro = True
      erroDescricao = "Tempo esgotado para informar a data do agendamento."
      processarDataAgenda = False
      DoEvents
      esperarProcesso 6000
      End
    End If
    esperarProcesso 100
  Loop
  processarDataAgenda = True
End Function

Public Sub fecharListaAgendamentos()
  fecharJanelaPorId janelaListaAgendamentos
  janelaListaAgendamentos = 0
End Sub

Public Function exibirListaAgendamentos() As Boolean
  On Error Resume Next
  Dim contador As Long
  Dim hDC As Long
  Dim janelaSalvarComo As Long
  Dim janelaCrystalReport As Long
  
  janelaSalvarComo = 0
  nomeSecao = Format(Date, "YYYYMMDD") & Format(Time, "hhmmss")
  relatorioNome = "requerimentos " & nomeSecao
  
  'espera a proxima tela (a tela do crystal report não tem nome)
  While janelaListaAgendamentos = GetForegroundWindow
    esperarProcesso 100
    DoEvents
  Wend
  
  hDC = GetWindowDC(0)
  janelaCrystalReport = getJanelaCrystalReport
  While janelaCrystalReport = 0
    janelaCrystalReport = getJanelaCrystalReport
    esperarProcesso 300
    DoEvents
    contador = contador + 1
    If contador > 200 Then
      erro = True
      erroDescricao = "A janela 'Imprimir Agendamento' não apareceu."
      exibirListaAgendamentos = False
      Exit Function
    End If
  Wend
  exibirListaAgendamentos = True
End Function

Private Function getJanelaCrystalReport() As Long
  On Error Resume Next
  Dim idJanela As Long
  Dim janelaAfxWnd42_1 As Long
  Dim janelaAfxWnd42_2 As Long
  Dim janelaAfxFrameOrView42 As Long

  getJanelaCrystalReport = 0
  idJanela = FindWindow(vbNullString, vbNullString)
  Do While idJanela <> 0
    janelaAfxWnd42_1 = FindWindowEx(idJanela, 0, "AfxWnd42", "")
    janelaAfxWnd42_2 = FindWindowEx(janelaAfxWnd42_1, 0, "AfxWnd42", "")
    janelaAfxFrameOrView42 = FindWindowEx(janelaAfxWnd42_2, 0, "AfxFrameOrView42", "")
    If janelaAfxFrameOrView42 <> 0 Then
      getJanelaCrystalReport = idJanela
      Exit Function
    End If
    idJanela = GetWindow(idJanela, GW_HWNDNEXT)
  Loop
End Function


Private Function setJanelaModoNormal() As Boolean
  On Error Resume Next
  
  If janelaSabi Then
    If janelaTrazerParaFrente(janelaSabi, SW_NORMAL) Then
      setJanelaModoNormal = True
    Else
      erro = "Falha ao maximizar a janela do Controle Operacional."
      setJanelaModoNormal = False
    End If
  Else
    erro = "A janela do Controle Operacional não foi encontrada."
    setJanelaModoNormal = False
  End If
End Function

Public Function exportarAgendamentos() As Boolean
  Dim botaoOK As Long
  Dim janelaDialogoEmail As Long
  Dim janelaSalvar As Long
  Dim res As Long
  Dim tituloJanela As String
  
  'clicar no botão exportar da barra de ferramentas
  res = SetWindowPos(janelaCrystalReport, 0, 0, 0, 800, 460, 0)
  SetForegroundWindow (janelaCrystalReport)
  DoEvents
  esperarProcesso 2000
  MouseClique 262, 44
  DoEvents
  esperarProcesso 500
                
  'aguarda a msg de email
  contador = 0
  janelaDialogoEmail = getJanelaPrincipalIDporTitulo("Email")
  While JanelaIjanelaDialogoEmailDEmail = 0 And contador < 5
    janelaDialogoEmail = getJanelaPrincipalIDporTitulo("Email")
    esperarProcesso 1000
    DoEvents
    contador = contador + 1
    If contador > 10 Then
      erro = True
      erroDescricao = "A caixa de diálogo 'Email' não apareceu."
      exportarAgendamento = False
      Exit Function
    End If
  Wend
  
  'clica em OK
  If janelaDialogoEmail <> 0 Then
    esperarProcesso 1000
    botaoOK = FindWindowEx(janelaDialogoEmail, 0, "Button", "OK")
    DoEvents
    SendMessage botaoOK, BM_CLICK, 0, 0
    DoEvents
  End If
  
  'abrir a janela Export
  tituloJanela = Space(256)
  janelaAtiva = GetForegroundWindow
  res = GetWindowText(janelaAtiva, tituloJanela, Len(tituloJanela))
  tituloJanelaAtiva = tituloJanela
  While Mid(tituloJanelaAtiva, 1, 6) <> "Export"
    janelaAtiva = GetForegroundWindow
    tituloJanela = Space(256)
    res = GetWindowText(janelaAtiva, tituloJanela, Len(tituloJanela))
    tituloJanelaAtiva = janelaAtiva
    DoEvents
    esperarProcesso 300
    If InStr(1, tituloJanelaAtiva, "SABI - Controle Operacional") > 0 Then
      ClickOpen (janelaAtiva)
    Else
      'clicar no botao Exportar
      If Len(Trim(tituloJanelaAtiva)) = 1 Then MouseClique 262, 44
    End If
  Wend
    
  'espera a tela export
  janelaExportar = getJanelaPrincipalIDporTitulo("Export")
  While janelaExportar = 0
    esperarProcesso 300
    DoEvents
    janelaExportar = getJanelaPrincipalIDporTitulo("Export")
  Wend
  If janelaExportar > 0 Then
    lcount = 0
    destinoExportar = 0
    Do While destinoExportar = 0 Or lcount > 10
      botaoOKExportar = FindWindowEx(janelaExportar, 0, "Button", "OK")
      formatoExportar = FindWindowEx(janelaExportar, 0, "ComboBox", "")
      destinoExportar = FindWindowEx(janelaExportar, formatoExport, "ComboBox", "")
      lcount = lcount + 1
      esperarProcesso 300
      DoEvents
    Loop
       
    'o destino deve ser escolhido antes do formato para não gerar erro de e-mail não configurado
    'destino
    SendMessageByLong destinoExportar, CB_SETCURSEL, 1, 0&
    DoEvents
    Sleep 100
    conta = 0
    If getControleTexto(destinoExportar) <> "Disk file" Then
      For conta = 0 To 100
        SendMessageByLong destinoExportar, CB_SETCURSEL, conta, 0&
        DoEvents
        Sleep 100
        If getControleTexto(destinoExportar) = "Disk file" Then Exit For
      Next conta
      DoEvents
    End If
    If conta > 99 Then
      erro = True
      erroDescricao = "A opção exportar para Disk File não foi encontrada na janela Exportar."
      End
    End If
    
    'formato
    SendMessageByLong formatoExportar, CB_SETCURSEL, 22, 0&
    DoEvents
    Sleep 100
    conta = 0
    If getControleTexto(formatoExportar) <> "Tab-separated text" Then
      For conta = 0 To 100
      SendMessageByLong formatoExportar, CB_SETCURSEL, conta, 0&
      DoEvents
      Sleep 100
      If getControleTexto(formatoExportar) = "Tab-separated text" Then Exit For
        Next conta
      End If
      If conta > 99 Then
        erro = True
        erroDescricao = "A opção exportar para o formato 'Tab-separated text' não foi encontrada na janela Exportar."
        exportarAgendamento = False
      End
    End If
    
    'clicar em OK
    esperarProcesso 100
    clicarBotao botaoOKExportar
    DoEvents
    esperarProcesso 300
    DoEvents
    esperarProcesso 1500
    
    'fechar janela Exportar
    fecharJanelaPorId janelaExportar
    DoEvents
  Else
    erro = True
    erroDescricao = "A janela Export não abriu."
    exportarAgendamento = False
    End
  End If
  
  'espera a janela "Choose Export File"
  janelaSalvar = FindWindow("#32770", "Choose Export File")
  contador = 0
  While janelaSalvar = 0
    espera 300
    conta = conta + 1
    If conta > 200 Then
      erro = True
      erroDescricao = "A janela 'Choose Export File' era esperada, mas não apareceu."
      exportarAgendamento = False
      Exit Function
    End If
  Wend
  
  exportarAgendamento = True
End Function

Function definirArquivoAgendamento() As String
  Dim hwndDialog As Long  ' handle to the dialog box
  Dim hwndButton As Long  ' handle to the  button
  Dim retval As Long      ' return value
  Dim txtlen As Long
  Dim txt As String
  Dim memocritica As String
  Dim GlobalBotãoSalvar As Long
  Dim titletmp As String
  
  Dim botaoSalvar As Long
  Dim contador As Long
  Dim comboBox As Long
  Dim editBox As Long
  Dim idFloatNotify As Long
  Dim idUIHandle As Long
  Dim idUIView As Long
  Dim janelaSalvar As Long
  Dim nomeArquivo As String
  Dim res As Long
  Dim tamanhoTexto As Long
  Dim texto As Long
  
  botaoSalvar = 0
  nomeArquivo = GlobalPastadeTrabalho & "\Agendamentos.txt"
  
  'insere o nome do arquivo na janela Salvar
  janelaSalvar = FindWindow("#32770", "Choose Export File")
  idUIView = FindWindowEx(janelaSalvar, 0, "DUIViewWndClassName", vbNullString)
  idUIHandle = FindWindowEx(idUIView, 0, "DirectUIHWND", vbNullString)
  idFloatNotify = FindWindowEx(idUIHandle, 0, "FloatNotifySink", vbNullString)
  comboBox = FindWindowEx(idFloatNotify, 0, "ComboBox", vbNullString)
  editBox = FindWindowEx(comboBox, 0, "Edit", vbNullString)
  res = SendMessage(comboBox, WM_SETTEXT, vbNullString, nomeArquivo)
  tamanhoTexto = SendMessage(editBox, WM_GETTEXTLENGTH, vbNullString, vbNullString)
  tamanhoTexto = tamanhoTexto + 1
  texto = Space$(txtlen)
  Call SendMessage(editBox, WM_GETTEXT, ByVal 260, texto)
  If InStr(1, texto, nomeArquivo) = 0 Then
    erro = True
    erroDesricao = "Não foi possível inserir na janela Salvar o nome do arquivo " & nomeArquivo & "."
    definirArquivoAgendamento = ""
    Exit Function
  End If
  
  'procura pelo botao Salvar
  botaoSalvar = 0 = FindWindowEx(janelaSalvar, 0, "Button", "Sa&lvar")
  contador = 0
  While botaoSalvar = 0
    Sleep 20
    contador = contador + 1
    If contador > 100 Then
      erro = True
      erroDescricao = "Não foi possível encontrar o botão Salvar da janela Salvar."
      definirArquivoAgendamento = ""
      Exit Function
    End If
  Wend
       
  'se o arquivo existe, excluir antes
  excluirArquivo nomeArquivo
       
  'clica no botão Salvar
  clicarBotao botaoSalvar
    
  'espera a janela Salvar ser fechada
  contador = 0
  texto = Space(256)
  GetWindowText janelaSalvar, texto, Len(texto)
  While InStr(1, UCase(texto), "SALVAR COMO") <> 0
    'reafirma comando de salvamento
    clicarBotao botaoSalvar
    Sleep 50
    contador = contador + 1
    If contador > 50 Then
      erro = True
      erroDescricao = "A janela Salvar Como não foi fechada."
      definirArquivoAgendamento = ""
      Exit Function
    End If
    texto = Space(256)
    GetWindowText janelaSalvar, texto, Len(texto)
    DoEvents
  Wend
End Sub

Public Sub fecharJanelaCrystalReport()
  fecharJanelaPorId janelaCrystalReport
  DoEvents
  While ObtemIDdoRelatórioCrystalReport <> 0
    fecharJanelaPorId getJanelaCrystalReport
    DoEvents
    Sleep 300
  Wend
End Sub

Private Function processarDocAgenda() As Boolean
  On Error Resume Next
  
  Dim contador As Long
  Dim arquivoAgenda As Long
  Dim linhaArquivo As String
  Dim numeroArquivo As Long
  Dim fimCabecalho As Boolean
  Dim i As Integer
  Dim pos1, pos2, pos3, pos4 As Long
  Dim temp As String
    
  contador = 0
  fimCabecalho = False
  cabecalhoAgenda = ""
  Open caminhoArquivoAgenda For Input As #descritorArquivo
  Do While Not EOF(descritorArquivo)
    Line Input #descritorArquivo, linhaArquivo
    If InStr(1, linhaArquivo, "Medico") Then fimCabecalho = True
    If fimCabecalho = False Then
      cabecalhoAgenda = cabecalhoAgenda & linhaArquivo & Chr(13) & Chr(10)
    Else
      If InStr(1, linhaArquivo, "Medico") Or InStr(1, linhaArquivo, "Horário") Then
      Else
        pos1 = InStr(1, linhaArquivo, Chr(9))
        If pos1 > 0 Then
          contador = contador + 1
          quantidadeAgendamentos = contador
          
          ' encontra a hora do agendamento
          agendamentos(contador).Horario = Mid(linhaArquivo, 1, pos1 - 1)
          
          ' encontra o nome do periciando
          pos2 = InStr(pos1 + 1, linhaArquivo, Chr(9))
          If pos2 > 0 Then
            temp = Mid(linhaArquivo, pos1 + 2, pos2 - pos1 - 1)
            temp = InStr(1, temp, Chr(34))
            If i > 0 Then
              temp = Mid(temp, 1, i - 1)
            Else
              temp = Mid(linhaArquivo, pos1 + 2, pos2 - pos1 - 1)
            End If
            agendamentos(contador).Segurado = temp
            
            ' encontra se a perícia já foi realizada
            pos3 = InStr(pos2 + 1, linhaArquivo, Chr(9))
            If pos3 > 0 Then
              agendamentos(contador).Concluida = Mid(linhaArquivo, pos2 + 1, pos3 - pos2 - 1)
              
              ' encontra a ordem e o numero de requerimento
              pos4 = InStr(pos3 + 1, linhaArquivo, Chr(9))
              If pos4 > 0 Then
                agendamentos(contador).Ordem = Mid(linhaArquivo, pos3 + 1, pos4 - pos3 - 1)
                agendamentos(contador).Requerimento = Val(Mid(linhaArquivo, pos4 + 1))
              End If
            End If
          End If
        End If
      End If
    End If
  Loop
  Close #descritorArquivo
  
  If Dir(caminhoArquivoAgenda) <> "" Then
    DeleteFile caminhoArquivoAgenda
  End If
  
  pos1 = InStr(1, cabecalhoAgenda, "Local:")
  If pos1 > 0 Then
    temp = Mid(cabecalhoAgenda, pos1 + 8)
    For pos2 = 1 To Len(temp)
      If Mid(temp, pos2, 1) = Chr(13) Or Mid(temp, pos2, 1) = Chr(10) Then
        temp = Mid(temp, 1, pos2 - 1) & Mid(temp, pos2 + 1)
      End If
      If Asc(Mid(temp, pos2, 1)) = 34 Then
        temp = Mid(temp, 1, pos2 - 1) & " " & Mid(temp, pos2 + 1)
      End If
    Next pos2
    temp = Trim(temp)
    lblLocaleData.Caption = temp
    lblLocaleData.Visible = True
    lbNomePrograma.Visible = False
  End If
End Function

MontaListadeRequerimentos (cabecalhoAgenda)

Function exibirAgendamentos() As Boolean
  Dim pos As Long
  Dim linha As String
  Dim indice As Long
  Dim nomedosegurado As String
  Dim GlobalRequerimentosProv(1000) As Requerimento
  Dim LINHA2 As String
  Dim pos3 As Long
  Dim conta2 As Long
  
  Dim contador As Long
  Dim contadorAgenda As Long
  
  indice = 0
  conta = 0
  listaClassificar.Clear
  listaRequerimentos.Clear
  
  If optOrdem(1).Value = True Then
    For contador = 1 To quantidadeAgendamentos
      listaClassificar.AddItem Mid(agendamentos(contador).Horario, 1, 5) & " - " & agendamentos(conta).Segurado
      listaClassificar.ItemData(listaClassificar.NewIndex) = agendamentos(conta).Requerimento
    Next contador
  Else
    For contador = 1 To quantidadeAgendamentos
      listaClassificar.AddItem agendamentos(conta).Segurado
      listaClassificar.ItemData(listaClassificar.NewIndex) = agendamentos(conta).Requerimento
    Next contador
  End If
  listaRequerimentos.AddItem "Seq." & Chr(9) & "Requerim." & Chr(9) & Chr(9) & Chr(9) & "Segurado"
  For contador = 0 To listaClassificar.ListCount - 1
    For contadorAgenda = 1 To quantidadeAgendamentos
      If agendamentos(contadorAgenda).Requerimento = (listaClassificar.ItemData(contador)) Then
        Exit For
      End If
    Next contadorAgenda
    indice = indice + 1
    requerimentos(contador + 1).Número = listaClassificar.ItemData(contador)
    requerimentos(contador + 1).Segurado = listaClassificar.List(contador)
    listaRequerimentos.AddItem Format(indice, "000") & Chr(9) & listaClassificar.ItemData(conta) & Chr(9) & listaClassificar.List(conta)
    
    'valores iniciais
    requerimentos(contador).Tipo = ""
    requerimentos(contador).Status = ""
    requerimentos(contador).nit = ""
    requerimentos(contador).Crítica = ""
  Next conta
  listaRequerimentos.Visible = True
  If QuantidadedeRequerimentos > 0 Then
    LocalCopiar = True
    Me.Top = Screen.Height - 760 - 3000
    Me.Width = 12540 + 1600
    Me.Height = Screen.Height - 560
    SetForegroundWindow (Me.hWnd)
    grupoOrdem.Visible = False
    redimensionarForm -4000, 2500
    fraImprime.Visible = True
    mostralista
    If QuantidadedeRequerimentos > 40 Then
      paracima.Visible = Trueit
      parabaixo.Visible = True
    End If
    txtUltimo.Text = QuantidadedeRequerimentos
    exibirAgendamentos = True
  Else
    exibirAgendamentos = False
    erro = True
    erroDescricao = "Não foi encontrado nenhum agendamento de perícia para esta data" & Chr(13) & Chr(10) & cabecalhoAgenda
    res = SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
    End
  End If
End Function

Sub clicarMenu(menuindex As Long, itemindex As Long)
  menuClicar janelaSabi, menuindex, itemindex
End Sub

Sub fecharJanela()
  clicarMenu 0, 6
End Sub

Public Sub processarRequerimentos()
  Dim botaoDipDibGCont As Long
  Dim botoaFecharJanela As Long
  Dim contador As Long
  Dim dimensoesTelaImprimir As RECT
  Dim requerimentosNaoMarcados As String
  Dim requerimentoAtual As Requerimento
  Dim janelaAtiva As Long
  Dim janelaDetalhesBeneficio As Long
  Dim janelaAtivaTitulo As String
  Dim ret As Long
  
  Dim memo
  Dim pt As POINTAPI
  Dim numCPF As String
  Dim idtelanafrente As Long

  lstMostrarRequerimentos.Enabled = False
  pctCopiaPartedaTela.Visible = False
  DoEvents
  
  ' maximiza CO
  janelaMaximizar
  
  'força um resize
  relatorioPronto = False
  LocalCopiar = True
  Me.Height = 2000
  
  'lê lstMostrarRequerimentos
  requerimentosNaoMarcados = ""
  For contador = 1 To lstMostrarRequerimentos.ListCount - 1
    If lstMostrarRequerimentos.Selected(contador) = False Then
      requerimentosNaoMarcados = requerimentosNaoMarcados & lstMostrarRequerimentos.List(contador)
    End If
  Next contador
  
  horadeInicio = Time
  inicioProcessamento = GetTickCount
  
  ' executa a rotina DIB/DIP/GCONT e extrai o CPF de cada um dos requerimentos
  For requerimentoID = 1 To QuantidadedeRequerimentos
  
    exibirProgresso  'antes: atualizaprogresso
    
    Me.Visible = True
    DoEvents
    exibirTela 'antes mostratela
  
    'so atua nos requerimentos sem marca de impressão
    If GlobalRequerimentos(requerimentoID).Impresso <> "SIM" And GlobalRequerimentos(requerimentoID).Impresso <> "NÃO" Then
      requerimentoAtual = ConsultaRequerimento(GlobalRequerimentos(requerimentoID).Número, GlobalImpressaoAuto)
      If requerimentoAtual.Crítica = "" Then
        GlobalRequerimentos(requerimentoID).nit = Mid(requerimentoAtual.nit, 1, 11)
        GlobalRequerimentos(requerimentoID).Tipo = Mid(requerimentoAtual.Tipo, 1, 7)
        GlobalRequerimentos(requerimentoID).Status = Mid(requerimentoAtual.Status, 1, 10)
        
        AtualizaListadeRequerimentos (requerimentoID)
        
        If IsNumeric(GlobalRequerimentos(requerimentoID).nit) And GlobalRequerimentos(requerimentoID).Status = "NORMAL" Then
          'rotina de clicar em DIB/DIP e Gcont
          Me.Left = 600
          DoEvents
          conta = 0
          
          'ativa a janela do Controle do SABI
          While GetForegroundWindow <> janelaSabi
            SetForegroundWindow (janelaSabi)
            DoEvents
            espera 100
            conta = conta + 1
            If conta > 100 Then Exit Sub
          Wend
          
          'clicar no botão Detalhes
          GetCursorPos pt
          SetCursorPos 800, 320
          MouseClique 800, 320
          SetCursorPos pt.x, pt.y
    
          'espera a janela "Detalhes Requerimento/Benefício" abrir
          janelaAtiva = GetForegroundWindow
          janelaAtivaTitulo = getJanelaTitulo(janelaAtiva)
          While InStr(1, janelaAtivaTitulo, "Detalhes Requerimento/Benefício") = 0
            janelaAtiva = GetForegroundWindow
            janelaDetalhesBeneficio = GetForegroundWindow
            janelaAtivaTitulo = getJanelaTitulo(janelaAtiva)
            DoEvents
            espera 300
            
            exibirTela
          Wend
          
          getJanelaDimensoes(janelaDetalhesBeneficio, dimensoesTelaImprimir)
          espera 300
          exibirTela
    
          ' apos abrir a janela,
          ' clicar na Aba Documentos para indentificar o CPF
          GlobalRequerimentos(requerimentoID).CPF = getRequerenteCPF(janelaDetalhesBeneficio)
                    
         ' clicar no botão DIB/DIP/GCont
          contador = 0
          botaoDipDibGCont = 0
          botaoFecharJanela = 0
          Do While botaoDipDibGCont = 0 Or botaoFecharJanela = 0
            botaoDipDibGCont = pesquisarJanela(janelaDetalhesBeneficio, 0, "ThunderRT6CommandButton", "DIB/DIP e Gcont")
            botaoFecharJanela = pesquisarJanela(janelaDetalhesBeneficio, 0, "ThunderRT6CommandButton", "&Fechar")
            contador = contador + 1
            If contador > 5000 Then
              Exit Do
            End If
          Loop
          If botaoDipDibGCont > 0 Then
            clicarBotao botaoDipDibGCont
            espera 300
            exibirTela
          End If
          If botaoFecharJanela > 0 Then
            clicarBotao botaoFecharJanela
            espera 300
            exibirTela
          Else
            ' fecha com clique no botão fechar da janela
            espera 1000
            exibirTela
            MouseClique 1010, 630
          End If
    
          'espera a janela "Detalhes Requerimento/Benefício" fechar
          janelaAtiva = GetForegroundWindow
          janelaAtivaTitulo = getJanelaTitulo(janelaAtiva)
          While InStr(1, janelaAtivaTitulo, "Detalhes Requerimento/Benefício") > 0
            janelaAtiva = GetForegroundWindow
            janelaAtivaTitulo = getJanelaTitulo(janelaAtiva)
            SetForegroundWindow janelaDetalhesBeneficio
            DoEvents
            espera 300
            exibirTela
          Wend
                            
          'fim da rotina de DIB/DIP e Gcont
          If Val(GlobalRequerimentos(requerimentoID).sequencia) >= Val(editMarcacaoDe.Text) And Val(GlobalRequerimentos(requerimentoID).sequencia) <= Val(editMarcacaoPara.Text) Then
            If abrangencia <> "NENHUM" Then
              If abrangencia = "TODOS" Or GlobalRequerimentos(requerimentoID).Tipo = abrangencia Then
                
                ' imprime a segunda via de marcação
                requerimentoAtual = ImprimirRequerimento(requerimentoAtual.nit, GlobalImpressaoAuto)
                If requerimentoAtual.Crítica <> "" Then requerimentoAtual.Impresso = False
                GlobalRequerimentos(requerimentoID).Crítica = requerimentoAtual.Crítica
                GlobalRequerimentos(requerimentoID).Impresso = requerimentoAtual.Impresso
                
                ' ??
                atualizarStatus
                atualizarRelatorio requerimentoID
                If GlobalRequerimentos(requerimentoID).Crítica <> "" Then
                  marcarComErro
                Else
                  If GlobalRequerimentos(requerimentoID).Impresso Then
                    marcarComImpressora 240, GlobalLinhaPicture - 14
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  Next requerimentoID
  
  pctProgressoFundo.Visible = False
  DoEvents
    
  pctCopiaPartedaTela.Visible = False
  LocalCopiar = False
  relatorioPronto = True
  Me.Top = Screen.Height - 3760
  Me.Height = 2000
  Picture1.Top = 0
End Sub

Private Function getRequerenteCPF()
  Dim janelaAtiva As Long
  Dim imagem As PictureBox
  Dim numCPF As String
  Dim quadroJanela As RECT
  Dim ret As Long
  
  ' clica na aba Detalhes da janela
  janelaAtiva = GetForegroundWindow
  getJanelaDimensoes janelaAtiva, quadroJanela
  processarCliqueMouse quadroJanela.Left + 110, quadroJanela.Top + 40
  espera 300
  exibirTela
  
  ' extrai a imagem sob o controle que exibe o numero do CPF
  ' essa linha pode ser inutil: getJanelaDimensoes(janelaAtiva, quadroJanela)
  Set imagem.Picture = CaptureWindow(janelaAtiva, False, 54, 69, 80, 15)
  
  ' extrai o texto da imagem
  numCPF = detectarTextoDeImagem imagem
  If numCPF = "" Then
    getRequerenteCPF = "-----------"
  Else
    getRequerenteCPF = numCPF
  End If
End Function

Public Function ImprimirRequerimento(nit As String, ImpressaoAuto As Boolean) As Requerimento
  On Error Resume Next
  Dim botaoVisualizar As Long
  Dim botaoCancelar As Long
  Dim botaoImprimir As Long
  Dim botaoImprimirHabilitado As Boolean
  Dim centroCampoNit As Long
  Dim contador As Long
  Dim dimensoesNit As RECT
  Dim editNit As Long
  Dim janela As Long
  Dim janelaSegundaVia As Long
  Dim nit As String
  Dim nomeControle As String
  Dim quadroNit As Long
  
  Dim res As String
  Dim ret As Integer
        
  formInicial.pctCopiaPartedaTela.Visible = False
  DoEvents
  Requerimento.Impresso = False
  Requerimento.nit = nit

  'fecha a tela de tempo transcorrido se ainda está aparecendo
  forcarFecharJanelaPorTitulo "Carteira de Benefícios"
        
  'fecha tela de critica do SABI se ainda estiver aberta
  forcarFecharJanelaPorClasse "#32770", "SABI - Controle Operacional"
  
  'abre tela Segunda Via de Marcação de Exame
  clicarMenu 4, 7
  DoEvents
  espera 300
  
  'espera a tela aparecer
  contador = 0
  janelaSegundaVia = 0
  While janelaSegundaVia = 0
    janelaSegundaVia = pesquisarJanela(janelaMDIClient, 0, "ThunderRT6FormDC", "Segunda Via de Marcação de Exame")
    If janelaSegundaVia <> 0 Then posicionarJanela(janelaSegundaVia, 0, 0, 750, 150)
    DoEvents
    espera 200
    contador = contador + 1
    If contador > 100 Then
      Requerimento.Critica = MSG_ERRO_SEGVIALENTO
      ImprimirRequerimento = Requerimento
      Exit Function
    End If
  Wend

  'trazer a janela do Controle para frente de todas
  janelaTrazerParaFrente janelaSabi
        
  'encontra as teclas Visualizar, Cancelar e Imprimir
  contador = 0
  botaoVisualizar = 0
  botaoCancelar = 0
  botaoImprimir = 0
  While botaoVisualizar = 0 Or botaoCancelar = 0 Or botaoImprimir = 0
    botaoVisualizar = pesquisarJanela(janelaSegundaVia, 0, "ThunderRT6CommandButton", "&Visualizar")
    botaoCancelar = pesquisarJanela(janelaSegundaVia, 0, "ThunderRT6CommandButton", "&Cancelar")
    botaoImprimir = pesquisarJanela(janelaSegundaVia, 0, "ThunderRT6CommandButton", "&Imprimir")
    contador = contador + 1
    If contador > 5000 Then
      Requerimento.Crítica = MSG_ERRO_CONTROLESSEGVIA
      ImprimirRequerimento = Requerimento
      'fecha a tela Segunda Via de Marcação de Exame
      fecharJanelaPorId janelaSegundaVia
      Exit Function
    End If
  Wend
        
  'encontra o quadro onde está o campo de NIT
  contador = 0
  quadroNit = 0
  While quadroNit = 0
    quadroNit = pesquisarJanela(janelaSegundaVia, 0, "ThunderRT6Frame", "NIT Requerente")
    contador = contador + 1
    If contador > 5000 Then
      Requerimento.Crítica = "Tempo expirado para encontrar o quadro onde está o campo 'NIT' na janela 'Segunda Via de Marcação de Exame'."
      ImprimirRequerimento = Requerimento
      'fecha a tela Segunda Via de Marcação de Exame
      fecharJanelaPorId janelaSegundaVia
      Exit Function
    End If
  Wend
        
  'encontra o campo de NIT
  contador = 0
  editNit = 0
  Do While editNit = 0
    editNit = pesquisarJanela(quadroNit, 0, vbNullString, vbNullString)
    If editNit <> 0 Then
      nomeControle = Space(100)
      res = GetClassName(editNit, nomeControle, 100)
      If Mid(nomeControle, 1, 14) = "ImMaskWndClass" Then
        Exit Do
      Else
        Requerimento.Crítica = "Tempo expirado pare encontrar o campo 'NIT' na janela 'Segunda Via de Marcação de Exame'."
        ImprimirRequerimento = Requerimento
        'fecha a tela Segunda Via de Marcação de Exame
        fecharJanelaPorId janelaSegundaVia
        Exit Function
      End If
    End If
    contador = contador + 1
    If contador > 5000 Then
      Requerimento.Crítica = "Tempo expirado pare encontrar o campo 'NIT' na janela 'Segunda Via de Marcação de Exame'."
      ImprimirRequerimento = Requerimento
      'fecha a tela Segunda Via de Marcação de Exame
      fecharJanelaPorId janelaSegundaVia
      Exit Function
    End If
  Loop
        
  'obtem as dimensões do campo NIT
  res = GetWindowRect(editNit, dimensoesNit)
  centroCampoNit = convlong(dimensoesNit.Left + (dimensoesNit.Right - dimensoesNit.Left) / 2, dimensoesNit.Top + (dimensoesNit.Bottom - dimensoesNit.Top) / 2)
  contador = 0
  
  'digita o valor do NIT no campo apropriado
  Do While getControleTexto(editNit) <> Mid(nit, 1, 10) & "-" & Mid(nit, 11, 1)
    'clica no campo NIT para atribuir foco
    clicarControle centroCampoNit
    DoEvents
    espera 100
            
    'reformata o valor do NIT
    nit = Requerimento.nit
    While InStr(1, nit, ".")
      nit = Mid(nit, 1, InStr(1, nit, ".") - 1) & Mid(nit, InStr(1, nit, ".") + 1)
    Wend
            
    'digita o valor do NIT sem o sinal - e o ultimo algarismo
    SimulaSendKeys Mid(nit, 1, Len(nit) - 1)
            
    'Digita o último algarismo
    SimulaSendKeys Right$(nit, 1)
    DoEvents
    espera 200
    
    contador = contador + 1
    If contador > 10 Then
      Requerimento.Critica = "Não foi possível inserir o NIT '" & Requerimento.nit & "' no campo apropriado."
      ImprimirRequerimento = Requerimento
      'fecha a tela Segunda Via de Marcação de Exame
      fecharJanelaPorId janelaSegundaVia
      Exit Function
    End If
  Loop
        
  'Epera os botões serem habilitados
  contador = 0
  botaoImprimirHabilitado = False
  While botaoImprimirHabilitado = False
    'verifica se o SABI exibiu mensagem de erro
    janela = pesquisarJanelaSimples("#32770", "SABI - Controle Operacional")
    If janela <> 0 Then
      espera 300
      ret = pesquisarJanelaSimples(janela, 0, "Static", "")
      ret = pesquisarJanelaSimples(janela, ret, vbNullString, vbNullString)
      Requerimento.Crítica = getControleTexto(ret)
      ImprimirRequerimento = Requerimento

      'fecha a tela de crítica do SABI
      fecharJanelaPorId janela
             
      'fecha a tela Segunda Via de Marcação de Exame
      fecharJanelaPorId janelaSegundaVia
      Exit Function
    End If
    botaoImprimirHabilitado = getControleHabilitacao(botaoImprimir)
    espera 200
    contador = contador + 1
    If contador > 100 Then
      ImprimirRequerimento.Crítica = "Tempo de 10 segundos expirado para o SABI habilitar o botão 'Imprimir'"
      ImprimirRequerimento = Requerimento
             
      'fecha a tela Segunda Via de Marcação de Exame
      fecharJanelaPorId janelaSegundaVia
      Exit Function
    End If
  Wend
        
  'imprimir
  clicarBotao botaoImprimir
  DoEvents
  ImprimirRequerimento.Impresso = True
        
  'espera a tela ser fechada
  contador = 0
  While janelaSegundaVia <> 0
    espera 200
    fecharJanelaPorId janelaSegundaVia
    janelaSegundaVia = pesquisarJanela(janelaMDIClient, 0, "ThunderRT6FormDC", "Segunda Via de Marcação de Exame")
    DoEvents
        
    'verifica se o SABI criticou o NIT
    janela = pesquisarJanelaSimples("#32770", "SABI - Controle Operacional")
    If janela <> 0 Then
      espera 300
      ret = pesquisarJanela(janela, 0, "Static", "")
      ret = pesquisarJanela(janela, ret, vbNullString, vbNullString)
      Requerimento.Crítica = getControleTexto(ret)
      Requerimento.Impresso = False
      ImprimirRequerimento = Requerimento
      
      'fecha a tela de crítica do SABI
      fecharJanelaPorId janela
      
      'fecha a tela Segunda Via de Marcação de Exame
      fecharJanelaPorId janelaSegundaVia
      Exit Function
    End If
    If contador > 50 Then
      'depois de 10 segundos sai de qualquer jeito
      Exit Function
    End If
  Wend
  
  'espera tela Imprimindo... abrir
  janela = pesquisarJanelaSimples("#32770", "Imprimindo...")
  contador = 0
  While janela = 0
    espera 100
    contador = contador + 1
    If contador > 50 Then
      ImprimirRequerimento = Requerimento
      Exit Function
    End If
    janela = pesquisarJanelaSimples("#32770", "Imprimindo...")
  Wend
  
  'manipula a janela Deseja Imprimir mais uma via
  contador = 0
  botaoImprimir = 0
  While botaoImprimir = 0
    botaoImprimir = pesquisarJanela(janela, 0, "Button", "&Sim")
    espera 100
    contador = contador + 1
    If contador > 5000 Then
      ImprimirRequerimento = Requerimento
      Exit Function
    End If
  Wend
  clicarBotao botaoImprimir
  
  'espera tela Imprimindo... fechar
  janela = pesquisarJanelaSimples("#32770", "Imprimindo...")
  contador = 0
  While janela <> 0
    contador = contador + 1
    espera 100
    If contador > 50 Then
      ImprimirRequerimento = Requerimento
      Exit Function
    End If
    janela = pesquisarJanelaSimples("#32770", "Imprimindo...")
  Wend
  
  ImprimirRequerimento = Requerimento
End Function

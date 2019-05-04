Attribute VB_Name = "Tabela"
Private Sub desenhar(handle As Long)
  Dim retorno
  Dim janela As Long
  Dim esquerda, altura, largura, dimensao As Long
  Dim handleLista As Long

  handleLista = GetWindowDC(listaRequerimentos.hWnd)
  esquerda = 900
  altura = 0
  largura = listaRequerimentos.Width
  dimensao = listaRequerimentos.Height
  janela = handle
  cmdContinua.Visible = False
  btoIniciar.Visible = False
  btoFechar.Visible = False
  fraImprime.Visible = False
  listaRequerimentos.Top = 0
  Me.Height = Screen.Height
  Me.Refresh
  listaRequerimentos.Visible = True
  listaRequerimentos.Height = Me.Height - listaRequerimentos.Top - 40
  listaRequerimentos.Refresh
  pctEsteRequerimento.Refresh
  Me.Top = Screen.Height - 3760
  Me.Left = 600
  listaRequerimentos.Visible = True
  DoEvents
  retorno = BitBlt(GetDC(janela), CLng(esquerda), _
  CLng(altura), CLng(largura), CLng(dimensao), handleLista, CLng(0), CLng(0), SRCCOPY)
  Me.Height = 2000
  listaRequerimentos.Visible = False
  listaRequerimentos.Top = 5000
  fraImprime.Left = grupoOrdem.Left
  fraImprime.Top = Me.Height - fraImprime.Height - 120
  grupoOrdem.Visible = False
  redimensionarForm -4000, 2500
  fraImprime.Visible = True
  cmdContinua.Visible = True
  btoIniciar.Visible = True
  btoFechar.Visible = True
  fraImprime.Visible = True
  mtempo2 = 0
  Timer2.Enabled = True
End Sub

Private Sub marcarComImpressora(deslocamentox As Long, deslocamentoy As Long)
  Dim linha, coluna As Long
  
  For linha = 0 To pctImpressora.Height / 15 - 1
    For coluna = 0 To pctImpressora.Width / 15 - 1
      If pctImpressora.Point(coluna, linha) <> 0 Then Picture1.PSet (coluna + deslocamentox, linha + deslocamentoy), pctImpressora.Point(coluna, linha)
    Next coluna
  Next linha
  SavePicture Picture1.Image, GlobalPastadeTrabalho & "\" & GlobalDatadosRequerimentos & GlobalAgenciaEscolhida & "Todos.bmp"
End Sub

Private Sub marcarComErro()
  On Error Resume Next
  Dim linha, coluna As Long
  
  Picture1.Visible = False
  For linha = Picture1.Height / 15 - 16 - 2 To Picture1.Height / 15 - 2
    For coluna = 1 To formInicial.Picture1.Width / 15
      If Picture1.Point(coluna, linha) = 14474460 Then
        Picture1.PSet (coluna, linha), RGB(255, 220, 220)
      End If
    Next coluna
  Next linha
  Picture1.Visible = True
End Sub


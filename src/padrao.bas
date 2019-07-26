Attribute VB_Name = "Padrao"
Public Const NomeAplicacao = "Automatizador do SABI"

' algum tipo de conversão
Public Function convlong(x As Long, y As Long)
  Dim res As Long
  
  res = 256 * 64 * 4
  res = y * (res) + x
  convlong = res
End Function

' Deixa o processo dormindo por x milisegundos
Public Sub dormirThread(tempo As Long)
  Dim relogio As Long
  Dim janela As Long
  Dim res As String

  relogio = 0
  
  ' o que isso faz???
  janela = GetForegroundWindow
  res = UCase(getControleTexto(janela))
  If InStr(1, res, "SALVAR COMO") Then
    res = SetWindowPos(janela, 0, 0, 0, 460, 340, 0)
    res = SetWindowPos(janela, 1, 0, 0, 0, 0, 3)
    DoEvents
  End If
  
  ' dorme...
  While relogio < tempo
    ' thread dorme por 200 milisegundos
    Sleep 200
    relogio = relogio + 200
  Wend
End Sub

' salva uma parte da tela em um arquivo de imagem
Public Sub salvarTela(handle As Long, requerimentoNumero As String)
  On Error Resume Next
  Dim res As Long
  Dim RECT As RECT
  Dim imagem As PictureBox
  Dim nomeArquivo As String

  res = GetWindowRect(handle, RECT)
  Set imagem.Picture = CaptureWindow(handle, False, 19, 232 + 15, 727, 16)
  nomeArquivo = GlobalDatadosRequerimentos & Format(GlobalIDRequerimento, "000") & requerimentoNumero
  SistemaArquivos.salvarImagem imagem.Picture, nomeArquivo
End Sub

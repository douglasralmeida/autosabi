Attribute VB_Name = "OCR"
Public Function detectarTextoDeImagem(imagem As PictureBox) As String
  On Error Resume Next
  
  Dim algarismo As Long
  Dim deslocamento As Long
  Dim digito As Long
  Dim digitosQuantidade As Long
  Dim indice As Long
  Dim letra As String
  Dim posicaoy As Long
  Dim posicaox As Long
  Dim resultado As String
  Dim soma As Long

  detectarTextoDeImagem = ""
  posicaoy = 2
  deslocamento = 1
  posicaox = deslocamento
  soma = 0
  algarismo = 0
  
  digitosQuantidade = 11
  For digito = 0 To digitosQuantidade - 1
    soma = 0
    If imagem.Point(algarismo + posicaox, posicaoy + 2) = 0 Then
      soma = 1
    End If
    If imagem.Point(algarismo + posicaox, posicaoy + 6) = 0 Then
      soma = soma + 2
    End If
    If imagem.Point(algarismo + posicaox + 2, posicaoy) = 0 Then
      soma = soma + 4
    End If
    If imagem.Point(algarismo + posicaox + 2, posicaoy + 4) = 0 Then
      soma = soma + 8
    End If
    If imagem.Point(algarismo + posicaox + 2, posicaoy + 8) = 0 Then
      soma = soma + 16
    End If
    If imagem.Point(algarismo + posicaox + 4, posicaoy + 2) = 0 Then
      soma = soma + 32
    End If
    If imagem.Point(algarismo + posicaox + 4, posicaoy + 6) = 0 Then
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
    resultado = resultado & letra
    algarismo = algarismo + 6
  Next digito
  detectarTextoDeImagem = resultado
End Function

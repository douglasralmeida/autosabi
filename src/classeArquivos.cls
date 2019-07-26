VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classeArquivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' classeArquivo.cls
' Classe para manipulação de arquivos

' Variáveis
Private pastaTrabalho As String

'Funcoes da classe
Private Sub Class_Initialize()
  pastaTrabalho = ""
End Sub

Public Sub excluirTemporarios()
  Dim memo As String
  
  'apaga todos bmp de datas anteriores a atual
  memo = Dir(pastaTrabalho & "\" & "*.bmp")
  While memo <> ""
    If Mid(memo, 1, 8) < Format(Date, "yyyymmdd") Then
      excluirArquivo pastaTrabalho & "\" & memo
    End If
    memo = Dir()
  Wend
  
  'apaga todos txt de datas anteriores a atual
  memo = Dir(pastaTrabalho & "\" & "*.txt")
  While memo <> ""
    If Mid(memo, 1, 8) < Format(Date, "yyyymmdd") Then
      excluirArquivo pastaTrabalho & "\" & memo
    End If
    memo = Dir()
  Wend
  memo = Dir(GlobalPastadeTrabalho & "\*" & ".txt")
  Do While memo <> ""
    If IsDate(Mid(memo, 7, 2) & "/" & Mid(memo, 5, 2) & "/" & Mid(memo, 1, 4)) Then
      Exit Do
    End If
    memo = Dir()
  Loop
End Sub

Public Function getPastaTrabalho()
  getPastaTrabalho = pastaTrabalho
End Function

Public Sub salvarImagem(imagem As Picture, nomeArquivo As String)
  nomeArquivo As String
  
  nomeArquivo = pastaTrabalho & "\" & nomeArquivo & ".bmp"
  SavePicture Picture, nomeArquivo
End Sub

Public Sub setPastaTrabalho(valor As String)
  pastaTrabalho = valor
End Sub


Attribute VB_Name = "FileUtils"
' fileutils.bas
' Modulo com Funcoes para Manipula��o de Arquivos

' APIs do Windows
Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

' Fun��es do M�dulos
Public Sub excluirArquivo(nomeArquivo As String)
  DeleteFile nomeArquivo
End Sub

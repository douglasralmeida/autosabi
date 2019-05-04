Attribute VB_Name = "Relatorio"
Public Type Requerimento
  sequencia As String
  N�mero As String
  Tipo As String
  Status As String
  nit As String
  impresso As Boolean
  Segurado As String
  Cr�tica As String
  CPF As String
End Type

Function acertarLarguraColuna(referencia As String, palavra As String)
  Dim largura As Long
  
  acertarLarguraColuna = palavra & " "
  largura = Len(referencia)
  If Len(palavra) < largura Then acertarLarguraColuna = palavra & Space(largura - Len(palavra)) & " "
  If Len(palavra) > largura Then acertarLarguraColuna = Mid(palavra, 1, largura) & " "
End Function

Function adicionarColuna(texto As String, mascara As String, dado As String)
  adicionarColuna = texto & acertarLarguraColuna(mascara, dado)
End Function

Sub atualizarRelatorio(atual As Long)
  On Error Resume Next
  Dim contador As Long
  Dim documento As String
  Dim linha As String
  Dim docImpresso As String

  documento = ""
  lstMostrarRequerimentos.Visible = False
  lstMostrarRequerimentos.Clear
  lstMostrarRequerimentos.AddItem acertarLarguraColuna("000", "Seq.") & acertarLarguraColuna("123456789", "Requerim.") & acertarLarguraColuna("INICIAL", "Tipo") & acertarLarguraColuna("INDEFERIDO", "Status") & acertarLarguraColuna("12345678901", "NIT") & acertarLarguraColuna("IMPRESSO", "Impresso") & acertarLarguraColuna("WWWW WWWWWWW WW WWWWWW", "Segurado")
  For conta = 1 To GlobalQuantidadedeRequerimentos
    GlobalRequerimentos(conta).sequencia = Format(conta, "000")
    linha = adicionarColuna(linha, "000", GlobalRequerimentos(conta).sequencia)
    linha = adicionarColuna(linha, "123456789", GlobalRequerimentos(conta).N�mero)
    linha = adicionarColuna(linha, "INICIAL", GlobalRequerimentos(conta).Tipo)
    linha = adicionarColuna(linha, "INDEFERIDO", GlobalRequerimentos(conta).Status)
    linha = adicionarColuna(linha, "12345678901", GlobalRequerimentos(conta).nit)
    If conta <= atual And GlobalRequerimentos(conta).impresso = False Then
      docImpresso = "N�o"
    Else
      docImpresso = "Sim"
    End If
    linha = adicionarColuna(linha, "IMPRESSO", docImpresso)
    linha = linha & GlobalRequerimentos(conta).Segurado
    linha = linha & "     " & GlobalRequerimentos(conta).Cr�tica
    lstMostrarRequerimentos.AddItem linha
    linha = GlobalRequerimentos(conta).sequencia & Chr(9)
    linha = linha & GlobalRequerimentos(conta).N�mero & Chr(9)
    linha = linha & GlobalRequerimentos(conta).Tipo & Chr(9)
    linha = linha & GlobalRequerimentos(conta).Status & Chr(9)
    linha = linha & GlobalRequerimentos(conta).nit & Chr(9)
    linha = linha & GlobalRequerimentos(conta).impresso & Chr(9)
    linha = linha & GlobalRequerimentos(conta).Segurado & Chr(9)
    linha = linha & GlobalRequerimentos(conta).Cr�tica
    documento = documento & linha & Chr(13) & Chr(10)
  Next conta
  lstMostrarRequerimentos.ListIndex = atual
  lstMostrarRequerimentos.Visible = True
  DoEvents
  Open GlobalPastadeTrabalho & "\" & GlobalDatadosRequerimentos & ".txt" For Output As #1
  Print #1, sTodos
  Close #1
End Sub

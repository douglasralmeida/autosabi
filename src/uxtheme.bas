Attribute VB_Name = "UxTheme"
' uxtheme.bas
' Modulo com Funcoes do Tema do Windows


' APIs do Windows
Public Declare Function IsThemeActive Lib "UxTheme.dll" () As Boolean

' Funcoes do Módulo
Public Function estaTemaAtivo() As Boolean
  estaTemaAtivo = True
End Function

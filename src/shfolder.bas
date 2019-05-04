Attribute VB_Name = "ShFolder"
' shfolder.bas
' Modulo com Funcoes das Pastas Especiais do Windows


' Constantes
Public Const CSIDL_DESKTOP = &H0        ' Pasta da Area de Trabalho
Public Const CSIDL_LOCAL_APPDATA = &H1C ' Pasta de Dados da Aplicacao (local)

' Tipos
Public Type shiEMID
  cb As Long
  abID As Byte
End Type

Public Type ITEMIDLIST
  mkid As shiEMID
End Type

' APIs do Windows
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

' Funcoes do Módulo
Public Sub abrirArquivo(endereco As String)
  ShellExecute Me.hwnd, "open", endereco, vbNullString, vbNullString, SW_SHOW
End Sub

Public Function getPastaEspecial(CSIDL As Long) As String
  Dim IDL As ITEMIDLIST
  Dim path As String
  Dim result As Long
    
  result = SHGetSpecialFolderLocation(100, CSIDL, IDL)
  If result = NOERROR Then
    path = Space(512)
    result = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal path)
    path = RTrim$(path)
    If Asc(Right(path, 1)) = 0 Then
      path = Left$(path, Len(path) - 1)
    End If
    getPastaEspecial = path
    Exit Function
  End If
  getPastaEspecial = ""
End Function

Attribute VB_Name = "ModuloWindows"
' windows.bas
' Modulo com Funcoes do Windows

Public Const CSIDL_DESKTOP = &H0        ' Pasta da Area de Trabalho
Public Const CSIDL_LOCAL_APPDATA = &H1C ' Pasta de Dados de Aplicacao (local)

Public Const NOERROR = 0

Public Type shiEMID
  cb As Long
  abID As Byte
End Type

Public Type ITEMIDLIST
  mkid As shiEMID
End Type

Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Declare Function IsThemeActive Lib "UxTheme.dll" () As Boolean

Public Function estaTemaAtivo() As Boolean
  estaTemaAtivo = True
End Function

Public Function getSpecialFolder(CSIDL As Long) As String
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
    getSpecialFolder = path
    Exit Function
  End If
  getSpecialFolder = ""
End Function



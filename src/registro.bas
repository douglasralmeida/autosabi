Attribute VB_Name = "Registro"
' registro.bas
' Modulo com Funcoes de Manipulação do Registro

Private Const CHAVE_PROGRAMA = "Software\Automatizador do SABI"

Private Const HKEY_CURRENT_USER = &H80000001    'Acessar HKCU

Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const REG_SZ As Long = 1

Private Const KEY_WOW64_64KEY As Long = &H100&  'app de 32 bits acessa a colméia de 64 bits

Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

'APIs do Windows
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Integer, ByVal lpSubKey As String, ByVal Reserved As Integer, ByVal lpClass As String, ByVal dwOptions As Integer, ByVal samDesired As Integer, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function abrirRegChave(ByRef hkey As Long) As Boolean
  Dim retorno As Long
  Dim secattr As SECURITY_ATTRIBUTES  'atributos de segurança para a chave
  Dim criououabriu As Long      'indica se a chave  foi criada ou aberta
  
  secattr.nLength = Len(secattr)
  secattr.lpSecurityDescriptor = 0
  secattr.bInheritHandle = 1
  
  retorno = RegCreateKeyEx(HKEY_CURRENT_USER, CHAVE_PROGRAMA, 0, "", 0, KEY_WRITE, secattr, hkey, criououabriu)
  abrirRegChave = retorno = 0
End Function

Public Sub fecharChave(ByVal hkey As Long)
  RegCloseKey (hkey)
End Sub

Public Function lerRegValor(ByVal hkey As Long, ByVal nome As String, Default As String) As String
  Dim retorno As Long
  Dim lngType As Long
  Dim strBuffer As String
  Dim lngBufLen As Long
  
  strBuffer = String(255, vbNullChar)
  lngBufLen = Len(strBuffer)
  If StrComp(nome, "default", vbTextCompare) = 0 Then
    retorno = RegQueryValueEx(hkey, "", ByVal 0&, lngType, ByVal strBuffer, lngBufLen)
  Else
    retorno = RegQueryValueEx(hkey, nome, ByVal 0&, lngType, ByVal strBuffer, lngBufLen)
  End If
  If retonro = 0 Then
    If lngType = REG_SZ Then
      If lngBufLen > 0 Then
        valor = Left$(strBuffer, lngBufLen - 1)
        lerRegValor = valor
      Else
        lerRegValor = Default
      End If
    Else
      lerRegValor = Default
    End If
  Else
    lerRegValor = Default
  End If
End Function

Public Function salvarRegValor(ByVal hkey As Long, ByVal nome As String, ByVal valor As String) As Boolean
  Dim retorno As Long
  
  retorno = RegSetValueEx(hkey, nome, 0, REG_SZ, valor & vbNullChar, Len(valor) + 1)
  salvarRegValor = retorno = 0
End Function

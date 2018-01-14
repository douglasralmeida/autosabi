VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classeSendKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer

Private Declare Sub keybd_event Lib "user32" ( _
   ByVal bVk As Byte, ByVal bScan As Byte, _
   ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" ( _
   ByVal cChar As Byte) As Integer
Private Declare Function VkKeyScanW Lib "user32" ( _
   ByVal cChar As Integer) As Integer

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)


Public Sub KeyDown(ByVal vKey As KeyCodeConstants)
   keybd_event vKey, 0, KEYEVENTF_EXTENDEDKEY, 0
End Sub

Public Sub KeyUp(ByVal vKey As KeyCodeConstants)
   keybd_event vKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
End Sub

Public Function KeyCode(ByVal sChar As String) As KeyCodeConstants
Dim bNt As Boolean
Dim iKeyCode As Integer
Dim b() As Byte
Dim iKey As Integer
Dim vKey As KeyCodeConstants
Dim iShift As ShiftConstants

   ' Determina se nós temos suporte Unicode ou não:
   bNt = ((GetVersion() And &H80000000) = 0)
   
   ' Obtém o código de teclado para o caractere
   If (bNt) Then
      b = sChar
      CopyMemory iKey, b(0), 2
      iKeyCode = VkKeyScanW(iKey)
   Else
      b = StrConv(sChar, vbFromUnicode)
      iKeyCode = VkKeyScan(b(0))
   End If
   
   KeyCode = (iKeyCode And &HFF&)

End Function






Attribute VB_Name = "ModuloWindows"
' windows.bas
' Modulo com Funcoes do Windows

Public Const CSIDL_DESKTOP = &H0        ' Pasta da Area de Trabalho
Public Const CSIDL_LOCAL_APPDATA = &H1C ' Pasta de Dados de Aplicacao (local)

Public Const NOERROR = 0

Private Type InitCommonControlsExStruct
  lngSize As Long
  lngICC As Long
End Type

Public Type shiEMID
  cb As Long
  abID As Byte
End Type

Public Type ITEMIDLIST
  mkid As shiEMID
End Type

Private Declare Function LoadLibraryA Lib "kernel32.dll" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

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

Public Sub Main()
  Dim iccex As InitCommonControlsExStruct, hMod As Long
  
  'constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx
  Const ICC_ANIMATE_CLASS As Long = &H80&
  Const ICC_BAR_CLASSES As Long = &H4&
  Const ICC_COOL_CLASSES As Long = &H400&
  Const ICC_DATE_CLASSES As Long = &H100&
  Const ICC_HOTKEY_CLASS As Long = &H40&
  Const ICC_INTERNET_CLASSES As Long = &H800&
  Const ICC_LINK_CLASS As Long = &H8000&
  Const ICC_LISTVIEW_CLASSES As Long = &H1&
  Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&
  Const ICC_PAGESCROLLER_CLASS As Long = &H1000&
  Const ICC_PROGRESS_CLASS As Long = &H20&
  Const ICC_TAB_CLASSES As Long = &H8&
  Const ICC_TREEVIEW_CLASSES As Long = &H2&
  Const ICC_UPDOWN_CLASS As Long = &H10&
  Const ICC_USEREX_CLASSES As Long = &H200&
  Const ICC_STANDARD_CLASSES As Long = &H4000&
  Const ICC_WIN95_CLASSES As Long = &HFF&
  Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all values above

  With iccex
    .lngSize = LenB(iccex)
    .lngICC = ICC_STANDARD_CLASSES ' vb intrinsic controls (buttons, textbox, etc)
    
    ' if using Common Controls; add appropriate ICC_ constants for type of control you are using
    ' example if using CommonControls v5.0 Progress bar:
     ' .lngICC = ICC_STANDARD_CLASSES Or ICC_PROGRESS_CLASS
  End With
  On Error Resume Next ' error? InitCommonControlsEx requires IEv3 or above
    
  hMod = LoadLibraryA("shell32.dll") ' patch to prevent XP crashes when VB usercontrols present
  InitCommonControlsEx iccex
  If Err Then
    InitCommonControls ' try Win9x version
    Err.Clear
  End If
  On Error GoTo 0
  '... show your main form next (i.e., Form1.Show)
  formInicial.Show
  If hMod Then FreeLibrary hMod
End Sub

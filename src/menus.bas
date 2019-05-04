Attribute VB_Name = "Menus"
' menus.bas
' Modulo com Funcoes de Manipulação de Menus

' APIs do Windows
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Funções do módulo

' Simula um clique em um determinado item de menu da janela
Public Sub menuClicar(handle As Long, menuindex As Long, itemindex As Long)
  On Error Resume Next
  
  Dim subMenu  As Long
  Dim menuItem As Long
  Dim menu As Long

  ' o menu da janela
  menu = GetMenu(handle)
  
  ' o submenu do menu
  subMenu = GetSubMenu(menu, menuindex)
  
  ' o item do submenu
  menuItem = GetMenuItemID(subMenu, itemindex)
  
  ' simula o clique
  Call PostMessage(handle, WM_COMMAND, lMenuItem, 0)
  
  ' sendmessage would hang app until file is selected in open form but
  ' postmessage is asynchronous which is better in this case
End Sub

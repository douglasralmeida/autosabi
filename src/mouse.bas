Attribute VB_Name = "Mouse"
Public Sub simularCliqueMouse(x As Long, y As Long)
  Dim ponteiroMouse As POINTAPI
  
  GetCursorPos ponteiroMouse
  SetCursorPos x, y
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  SetCursorPos ponteiroMouse.x, ponteiroMouse.y
End Sub

VERSION 5.00
Begin VB.UserControl XPButton 
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   DefaultCancel   =   -1  'True
   MaskColor       =   &H000000FF&
   Picture         =   "XPButton.ctx":0000
   ScaleHeight     =   1080
   ScaleWidth      =   1890
   ToolboxBitmap   =   "XPButton.ctx":1692
End
Attribute VB_Name = "XPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10

Public Enum xp_btnState
    xp_Normal = 0
    xp_Pressed = 1
    xp_Disabled = 2
    xp_Hovered = 3
    xp_Focused = 4
End Enum

Const m_Def_State = xp_btnState.xp_Normal

Private Type POINT_API
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Event Click()
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Private m_State As xp_btnState
Private m_Font As Font
Private m_Caption As String
Private m_bFocused As Boolean

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
  If Enabled Then
    If m_bFocused Then
      m_State = xp_Focused
    Else
      m_State = xp_Normal
    End If
  Else
    m_State = xp_Disabled
  End If
  If Enabled Then ForeColor = vbBlack Else ForeColor = RGB(161, 161, 146)
  Make_xpButton
  
End Property

Public Property Get State() As xp_btnState
Attribute State.VB_ProcData.VB_Invoke_Property = ";Apariencia"
    State = m_State
End Property

Public Property Let State(ByVal vNewValue As xp_btnState)
    m_State = vNewValue
    PropertyChanged "State"
    Make_xpButton
End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
  RaiseEvent Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  
  If PropertyName = "DisplayAsDefault" Then
    If UserControl.Ambient.DisplayAsDefault Then
      m_State = xp_Focused
    Else
      m_State = xp_Normal
    End If
    Make_xpButton
  End If
  
End Sub

Private Sub UserControl_Click()
  If m_bFocused Then RaiseEvent Click
End Sub

Private Sub UserControl_EnterFocus()
  m_bFocused = True
  m_State = xp_Focused
  Make_xpButton
End Sub

Private Sub UserControl_ExitFocus()
  m_bFocused = False
  m_State = xp_Normal
  Make_xpButton
End Sub

Private Sub UserControl_Initialize()
Call Module1.no_sirve
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
  Make_xpButton
End If
End Sub

Private Sub UserControl_InitProperties()
  m_State = m_Def_State
  Enabled = True
  Caption = Ambient.DisplayName
  Set Font = UserControl.Ambient.Font
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_State = xp_Pressed
  Make_xpButton
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  SetCapture hwnd

  If Hittest(X, Y) Then
    If Button = vbLeftButton And m_bFocused Then
         m_State = xp_Pressed
    Else
       m_State = xp_Hovered
    End If
    Make_xpButton
    RaiseEvent MouseMove(Button, Shift, X, Y)

  Else
    Make_xpButton
    ReleaseCapture
    RaiseEvent MouseOut
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ReleaseCapture
  If m_bFocused Then m_State = xp_Focused Else m_State = xp_Normal
  Make_xpButton
  
  If m_bFocused Then RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_Paint()
  Make_xpButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_State = PropBag.ReadProperty("State", m_Def_State)
  Enabled = PropBag.ReadProperty("Enabled", True)
  m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
  Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)

  Dim Pos As Integer
  
  Pos = InStr(1, m_Caption, "&", vbBinaryCompare)
  If Pos <> 0 And Pos + 1 <= Len(m_Caption) Then
    UserControl.AccessKeys = Mid$(m_Caption, Pos + 1, 1)
  Else
    UserControl.AccessKeys = ""
  End If

End Sub

Private Sub UserControl_Resize()
    Make_xpButton
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("State", m_State, m_Def_State)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
  Call PropBag.WriteProperty("Font", m_Font, UserControl.Ambient.Font)
End Sub

Private Sub Make_xpButton()
'Toma una celda de 17x20 en el siguiente orden:
'normal, pressed, disabled, hovered(Mouse Move), focused(Gotfocus)

  Dim intCell, brx, bry, bw, bh As Integer
  
  ScaleMode = vbPixels 'Draw in pixels
  intCell = m_State 'm_state enum is related to cell position
  
  'Short cuts
  brx = UserControl.ScaleWidth - 3 'right x
  bry = UserControl.ScaleHeight - 3 'right y
  bw = UserControl.ScaleWidth - 6 'border width - corners width
  bh = UserControl.ScaleHeight - 6 'border height - corners height
  
  'Dibuja el button
  'UserControl.PaintPicture Picture, 0, 0, 3, 3, 0, 0, 3, 3
  UserControl.PaintPicture Picture, 0, 0, 3, 3, intCell * 18, 0, 3, 3
  UserControl.PaintPicture Picture, brx, 0, 3, 3, intCell * 18 + 15, 0, 3, 3
  UserControl.PaintPicture Picture, brx, bry, 3, 3, intCell * 18 + 15, 18, 3, 3
  UserControl.PaintPicture Picture, 0, bry, 3, 3, intCell * 18, 18, 3, 3
  
  'btn face without corners
  UserControl.PaintPicture Picture, 3, 0, bw, 3, intCell * 18 + 3, 0, 12, 3
  UserControl.PaintPicture Picture, brx, 3, 3, bh, intCell * 18 + 15, 3, 3, 15
  UserControl.PaintPicture Picture, 0, 3, 3, bh, intCell * 18, 3, 3, 15
  UserControl.PaintPicture Picture, 3, bry, bw, 3, intCell * 18 + 3, 18, 12, 3
  UserControl.PaintPicture Picture, 3, 3, bw, bh, intCell * 18 + 3, 3, 12, 15
  
  'paint corner points or we could replace it with setwindowrgn in usercontrol_resize event
  'cause not every container use vbbuttonface as backcolor
  
  UserControl.PSet (0, 0), vbButtonFace
  UserControl.PSet (ScaleWidth - 1, 0), vbButtonFace
  UserControl.PSet (ScaleWidth - 1, ScaleHeight - 1), vbButtonFace
  UserControl.PSet (0, ScaleHeight - 1), vbButtonFace
  
  DrawCaption
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute Caption.VB_UserMemId = -520
  Caption = m_Caption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
  Dim Pos As Integer
  
  If m_Caption <> vNewCaption Then
    
    Pos = InStr(1, vNewCaption, "&", vbBinaryCompare)
    If Pos <> 0 And Pos + 1 <= Len(vNewCaption) Then
      UserControl.AccessKeys = Mid$(vNewCaption, Pos + 1, 1)
    Else
      UserControl.AccessKeys = ""
    End If
  
    m_Caption = vNewCaption
    PropertyChanged "Caption"
    Make_xpButton
  End If
End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = "StandardFont;Fuente"
Attribute Font.VB_UserMemId = -512
  Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
  Set m_Font = vNewFont
  Set UserControl.Font = vNewFont
  'Set lbl.Font = m_Font
  Call UserControl_Resize
  PropertyChanged "Font"
End Property


Private Function Hittest(X As Single, Y As Single) As Boolean
  Dim pnt As POINT_API
  Dim rBox As RECT
  
  GetWindowRect hwnd, rBox        'Obtiene el rectangulo relativo de la pantalla
  GetCursorPos pnt
  
  If PtInRect(rBox, pnt.X, pnt.Y) Then
    Hittest = True
  Else
    If m_bFocused Then m_State = xp_Focused Else m_State = xp_Normal
  End If
End Function

Private Sub DrawCaption()
  Dim rBox As RECT
  ScaleMode = vbPixels
  SetRect rBox, 0, 0, ScaleWidth, ScaleHeight
  
  'calcula el rectangulo del texto
  DrawText hDC, m_Caption, Len(m_Caption), rBox, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK
  
  'establece el rectangulo
  SetRect rBox, (ScaleWidth - rBox.Right) / 2, _
    (ScaleHeight - rBox.Bottom) / 2, ((ScaleWidth - rBox.Right) / 2) + rBox.Right, _
    ((ScaleHeight - rBox.Bottom) / 2) + rBox.Bottom
  
  'finalmente dibuja el texto
  DrawText hDC, m_Caption, Len(m_Caption), rBox, DT_CENTER Or DT_WORDBREAK

End Sub

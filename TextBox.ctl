VERSION 5.00
Begin VB.UserControl TextBox 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H000000FF&
   ScaleHeight     =   375
   ScaleWidth      =   1695
   ToolboxBitmap   =   "TextBox.ctx":0000
   Begin VB.TextBox TextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6E6E6&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "TextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint _
 As POINTAPI) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd _
 As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" _
 (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Dim KG As Integer, Running As Boolean, Focused As Boolean, SetStandardSize As Boolean
'Eigenschaft-Variablen:
Dim m_HoverEffect As Boolean
Event Click() 'MappingInfo=TextBox,TextBox,-1,Click
Event DblClick() 'MappingInfo=TextBox,TextBox,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=TextBox,TextBox,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=TextBox,TextBox,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=TextBox,TextBox,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=TextBox,TextBox,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=TextBox,TextBox,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=TextBox,TextBox,-1,MouseUp
Event Change() 'MappingInfo=TextBox,TextBox,-1,Change
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Event MouseOver()
Event MouseOut()
Private Sub TextBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)

 If Not Focused Then
  Dim nPoint As POINTAPI
  Dim nHWnd As Long
  
  GetCursorPos nPoint
  nHWnd = WindowFromPoint(nPoint.X, nPoint.Y)
  With TextBox
    If nHWnd = .hwnd Then
      If GetCapture() <> nHWnd Then
        SetCapture nHWnd
        MouseOver
      End If
      Exit Sub
    End If
    ReleaseCapture
    MouseOut
  End With
 End If
End Sub
Private Sub TextBox_GotFocus()
 MouseOver
 Focused = True
End Sub
Private Sub TextBox_LostFocus()
 Focused = False
 MouseOut
End Sub
'Private Sub FrameDesign(TD As Boolean)
' For idx = 0 To 3
'  LLine(idx).Visible = TD
' Next idx
'End Sub
Private Sub MouseOut()
 RaiseEvent MouseOut
 If m_HoverEffect And Not Running Then
  'FrameDesign False
 End If
End Sub
Private Sub MouseOver()
 RaiseEvent MouseOver
 If m_HoverEffect And Not Running Then
'  FrameDesign True
 End If
End Sub
Private Sub UserControl_Resize()
TextBox.Height = UserControl.Height
TextBox.Width = UserControl.Width
'UserControl.Width = 1695
'UserControl.Height = 375
End Sub
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = TextBox.BackColor
'End Property
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    TextBox.BackColor() = New_BackColor
'    PropertyChanged "BackColor"
'End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=TextBox,TextBox,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = TextBox.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    TextBox.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=TextBox,TextBox,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = TextBox.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    TextBox.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get Font() As Font
    Set Font = TextBox.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set TextBox.Font = New_Font
    PropertyChanged "Font"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=TextBox,TextBox,-1,Refresh
Public Sub Refresh()
    TextBox.Refresh
End Sub

Private Sub TextBox_Click()
    RaiseEvent Click
End Sub

Private Sub TextBox_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub TextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub TextBox_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub TextBox_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub TextBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub TextBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub TextBox_Change()
    RaiseEvent Change
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=TextBox,TextBox,-1,Locked
Public Property Get Locked() As Boolean
    Locked = TextBox.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    TextBox.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=TextBox,TextBox,-1,MaxLength
Public Property Get MaxLength() As Long
    MaxLength = TextBox.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    TextBox.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=TextBox,TextBox,-1,MousePointer
Public Property Get MousePointer() As Integer
    MousePointer = TextBox.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    TextBox.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=TextBox,TextBox,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = TextBox.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set TextBox.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=TextBox,TextBox,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = TextBox.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    TextBox.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=TextBox,TextBox,-1,SelStart
Public Property Get SelStart() As Long
    SelStart = TextBox.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    TextBox.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property
Public Property Get SelText() As String
    SelText = TextBox.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    TextBox.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'Devuelve! o establece el texto contenido en el control!
'MappingInfo=TextBox,TextBox,-1,Text
Public Property Get Texto() As String
    Texto = TextBox.Text
End Property

Public Property Let Texto(ByVal New_Texto As String)
    TextBox.Text() = New_Texto
    PropertyChanged "Texto"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=TextBox,TextBox,-1,ToolTipText
Public Property Get ToolTipText() As String
    ToolTipText = TextBox.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    TextBox.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=TextBox,TextBox,-1,WhatsThisHelpID
Public Property Get WhatsThisHelpID() As Long
    WhatsThisHelpID = TextBox.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    TextBox.WhatsThisHelpID() = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property
Private Sub UserControl_InitProperties()
    m_HoverEffect = True
    
    If UserControl.Ambient.UserMode Then
     Running = True
    Else
     Running = False
    End If
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   ' TextBox.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    TextBox.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    TextBox.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    TextBox.Locked = PropBag.ReadProperty("Locked", False)
    TextBox.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    TextBox.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    TextBox.SelLength = PropBag.ReadProperty("SelLength", 0)
    TextBox.SelStart = PropBag.ReadProperty("SelStart", 0)
    TextBox.SelText = PropBag.ReadProperty("SelText", "")
    TextBox.Text = PropBag.ReadProperty("Texto", "")
    TextBox.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    TextBox.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
    'UserControl.BackColor = PropBag.ReadProperty("BorderColor", &H8000000F)
   ' m_HoverEffect = PropBag.ReadProperty("HoverEffect", True)
    
' If m_HoverEffect Then
'  FrameDesign False
' Else
'  FrameDesign True
' End If

End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Call PropBag.WriteProperty("BackColor", TextBox.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", TextBox.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", TextBox.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Locked", TextBox.Locked, False)
    Call PropBag.WriteProperty("MaxLength", TextBox.MaxLength, 0)
    Call PropBag.WriteProperty("MousePointer", TextBox.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("SelLength", TextBox.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", TextBox.SelStart, 0)
    Call PropBag.WriteProperty("SelText", TextBox.SelText, "")
    Call PropBag.WriteProperty("Texto", TextBox.Text, "")
    Call PropBag.WriteProperty("ToolTipText", TextBox.ToolTipText, "")
    Call PropBag.WriteProperty("WhatsThisHelpID", TextBox.WhatsThisHelpID, 0)
    'Call PropBag.WriteProperty("BorderColor", UserControl.BackColor, &H8000000F)
    'Call PropBag.WriteProperty("HoverEffect", m_HoverEffect, True)
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,BackColor
'Public Property Get BorderColor() As OLE_COLOR
'    BorderColor = UserControl.BackColor
'End Property
'
'Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
'    UserControl.BackColor() = New_BorderColor
'    PropertyChanged "BorderColor"
'End Property

'Public Property Get HoverEffect() As Boolean
'    HoverEffect = m_HoverEffect
'End Property
'
'Public Property Let HoverEffect(ByVal New_HoverEffect As Boolean)
'    m_HoverEffect = New_HoverEffect
'    FrameDesign Not m_HoverEffect
'    PropertyChanged "HoverEffect"
'End Property



VERSION 5.00
Begin VB.Form Efecto 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1560
      Top             =   360
   End
End
Attribute VB_Name = "Efecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public from As Object
Public f As Form
Public Sub Ajustes_Formularios(Form As Form)
Set from = Form
  xt = Screen.Height / 50
  wt = Screen.Width / 50
  X = Screen.Height / 2
  Y = Screen.Width / 2
  Form.Top = X
  Form.Left = Y
  Form.Height = 0
  Form.Width = 0
  accion = "abrir"
  'Timer2.Enabled = True
  Abrir_cerrar Form
  End Sub
Sub salir(Form As Form)
'procedimiento de cierre de la aplicación. Se activa el cronometro para
'realizar la salida visual del Formulario en Forma escalonada.
  Form.Show
  accion = "cerrar"
  Timer2.Enabled = True
End Sub
Sub CIERRE_PANTALLA(Form As Form)
'procedimiento de cierre final de la Formulario MDI, en Forma
'escalonada.  Controlado mediante un TIMER
  If Form.Top >= X Then
    Form.Top = X
    Form.Height = X - X
  Else
    Form.Top = Form.Top + xt
    If Form.Top >= (X - xt) Then
      Form.Top = X
      Form.Height = X - X
    Else
      Form.Height = Form.Height - (2 * xt)
    End If
  End If
  If Form.Left >= Y Then
    Form.Left = Y
    Form.Width = Y
  Else
    Form.Left = Form.Left + wt
    If Form.Left >= (Y - wt) Then
      Form.Left = Y
      Form.Width = Y - Y
    Else
      Form.Width = Form.Width - (2 * wt)
    End If
  End If
  If Form.Top = X And Form.Left = Y Then
    accion = "cierre2"
  End If
End Sub
Sub CIERRE_PANTALLA2(Form As Form)
'procedimiento de cierre final de Formulario, en Forma
'escalonada.  Controlado mediante un TIMER
  If Form.Top > 0 Then
    If Form.Top < 0 Then
      Form.Top = 0
    Else
      Form.Top = Form.Top + xt
    End If
  Else
    Form.Top = 0
  End If
  If Form.Left > 0 Then
    If Form.Left < 0 Then
      Form.Left = 0
    Else
      Form.Left = Form.Left + wt
    End If
  Else
    Form.Left = 0
  End If
  If Form.Left > Screen.Width Then
    Unload Me
    
  End If
End Sub
Sub PANTALLA(Form As Form)
'procedimiento de presentación inicial de la formulario MDI, en forma
'escalonada.  Controlado mediante un TIMER
  If Form.Top > 0 Then
    If Form.Top < 0 Then
      Form.Top = 0
    Else
      Form.Top = Form.Top - xt
      Form.Height = Form.Height + (2 * xt)
    End If
  Else
    Form.Top = 0
  End If
  If Form.Left > 0 Then
    If Form.Left < 0 Then
      Form.Left = 0
    Else
      Form.Left = Form.Left - wt
      Form.Width = Form.Width + (2 * wt)
    End If
  Else
    Form.Left = 0
  End If
  If Form.Top = 0 And Form.Left = 0 Then
    Form.Width = Screen.Width
    Form.Height = Screen.Height
    Timer2.Enabled = False
  End If
End Sub

Public Sub Abrir_cerrar(Form As Form)
    If accion = "abrir" Then
    Efecto.PANTALLA Form
  Else
    If accion = "cerrar" Then
      Efecto.CIERRE_PANTALLA Form
    Else
      Efecto.CIERRE_PANTALLA2 Form
    End If
  End If


End Sub

Private Sub timer2_Timer()
'procedimiento en el cual mediante un cronometro se controla la acciones
'de visualización final e inicial del Formulario.
  If accion = "abrir" Then
    PANTALLA from
  Else
    If accion = "cerrar" Then
      CIERRE_PANTALLA from
    Else
      CIERRE_PANTALLA2 from
    End If
  End If


End Sub



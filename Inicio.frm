VERSION 5.00
Begin VB.Form Inicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Juego del ahorcado"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   Icon            =   "Inicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton CmdIniciar 
      Caption         =   "&Iniciar el juego"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox TPalabra 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   960
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "No escribas números ni caracteres especiales"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Inicio.frx":030A
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdIniciar_Click()
If TPalabra.Text = "" Or Len(Trim(TPalabra.Text)) = 1 Then
    MsgBox "Escribe una palabra luego de leer las indicaciones sugeridas", vbCritical, "Palabra invalida"
    Exit Sub
End If
    Palabra = UCase$(TPalabra.Text)
    Longitud = Len(Palabra)
    Unload Me
    Load Juego
    Juego.Show
End Sub
Private Sub CmdSalir_Click()
Unload Me
End Sub
Private Sub TPalabra_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case 13
            CmdIniciar_Click
            Exit Sub
        Case 32
            MsgBox "No puede usar la barra espaciadora", vbInformation, "Información"
            SendKeys "(BS)"
            Exit Sub
        Case 8
            Exit Sub
End Select
        Module1.LetrasM KeyAscii
End Sub

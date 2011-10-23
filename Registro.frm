VERSION 5.00
Begin VB.Form Registro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Arbo.ocx"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tclave 
      Height          =   375
      Index           =   4
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Tclave 
      Height          =   375
      Index           =   3
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Tclave 
      Height          =   375
      Index           =   2
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Tclave 
      Height          =   375
      Index           =   1
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Cmdvolver 
      Caption         =   "&Volver"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdregistrar 
      Caption         =   "&Registrar"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label version 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese la clave:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2003"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"Registro.frx":0000
      ForeColor       =   &H00FF0000&
      Height          =   1185
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   0
      X2              =   6120
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Título de la aplicación"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alejandro Barreto"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4440
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto elaborado por"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4200
      TabIndex        =   0
      Top             =   360
      Width           =   1665
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "Registro.frx":0122
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Registro As Boolean

Private Sub cmdregistrar_Click()
If Tclave(1).Text = "sys98" And Tclave(2).Text = "413" And Tclave(3).Text = "lBoqEAlq" And Tclave(4).Text = "4r80" Then
MsgBox "Arbo.ocx ha sido registrado satisfactoriamente", vbQuestion, "ArboOcx"
Crear_carpeta "C:\Archivos de programa\Archivos comunes\Temp"
Crear_carpeta "C:\Archivos de programa\Archivos comunes\Temp\Registro"
Open "C:\Archivos de programa\Archivos comunes\Temp\Registro\sys98.sys" For Output As #1
For i = 1 To 4
Print #1, Tclave(i)
Next
Close #1
Registro = True
Cmdvolver_Click
Else
MsgBox "Clave de registro incorrecta", vbCritical, "Error"
End If
End Sub
Public Function Crear_carpeta(Ruta As String)
Dim Objeto As Object
On Error GoTo Mal_escrito
Set Objeto = CreateObject("Scripting.FileSystemObject")
Objeto.CreateFolder Ruta
Exit Function
Mal_escrito:
End Function
Private Sub Cmdvolver_Click()
Unload Me
End Sub

Public Sub Abrir()
On Error GoTo No
Open "C:\Archivos de programa\Archivos comunes\Temp\Registro\sys98.sys" For Input As #1
Do While Not EOF(1)
For i = 1 To 4
Line Input #1, a
Tclave(i) = a
Next
Loop
Close #1
If Tclave(1).Text = "sys98" And Tclave(2).Text = "413" And Tclave(3).Text = "lBoqEAlq" And Tclave(4).Text = "4r80" Then Registro = True
For i = 1 To 4
Tclave(i).Text = ""
Next
Registro = True
No:
If Err.Number = 53 Then
Registro = False
End If
End Sub

Private Sub Form_Load()
Color Me, 0, 256
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
Call Abrir
If Registro = True Then version.Caption = "Versión registrada" Else version.Caption = "Versión no registrada"

End Sub

Private Sub Color(Formulario As Object, COLOR1 As String, COLOR2 As String)
   Dim Amount As Single
   Dim i As Integer
   Formulario.AutoRedraw = True
   Amount = (455 / Formulario.ScaleHeight)
   For i = 0 To Formulario.ScaleHeight
Formulario.Line (0, i)-(Formulario.ScaleWidth, i), RGB(COLOR1, COLOR2, Amount * i)
   Next
End Sub

Private Sub Form_Resize()
Color Me, 0, 256

End Sub

Private Sub Tclave_KeyPress(Index As Integer, KeyAscii As Integer)
If Len(Tclave(1)) = 5 Then Tclave(2).SetFocus
If Len(Tclave(2)) = 3 Then Tclave(3).SetFocus
If Len(Tclave(3)) = 8 Then Tclave(4).SetFocus
If Len(Tclave(4)) = 4 Then cmdregistrar.SetFocus
End Sub


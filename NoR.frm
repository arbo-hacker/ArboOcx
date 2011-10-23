VERSION 5.00
Begin VB.Form NoR 
   Caption         =   "ArboOcx"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   Icon            =   "NoR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "http://www.noti-hackers.cjb.net"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3840
      TabIndex        =   9
      Top             =   1080
      Width           =   2265
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "arbo_hacker@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4200
      TabIndex        =   7
      Top             =   840
      Width           =   1905
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   0
      X2              =   6120
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"NoR.frx":A39A
      ForeColor       =   &H00FF0000&
      Height          =   1185
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1410
      Left            =   120
      Picture         =   "NoR.frx":A4BC
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto elaborado por"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   1665
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alejandro Barreto"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4440
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Título de la aplicación"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2003"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label version 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "NoR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Variable As Integer
Public Fecha As String
Public Ahora As String
Public veces As Integer
Public s As Boolean

Private Sub Command1_Click()
Unload Me
End Sub
Public Function Borrar_archivos(Ruta As String, Sino As Boolean)
On Error GoTo Ruta_Equivocada
Dim BA As Object
Set BA = CreateObject("Scripting.FileSystemObject")
BA.DeleteFile Ruta, Sino
'Sino se usa si quieres que se borren tambien los archvos que sean de solo lectura
Ruta_Equivocada:
Exit Function
End Function
Public Sub empezar()
Call crear2
Call crear
Call listo
If s = True Then Exit Sub
Call validar
Fecha = Now()
Fecha = Format(Fecha, "dd")
If Fecha = Ahora Then
Exit Sub
End If
If Ahora < Fecha Then
Open "C:\Archivos de programa\Archivos comunes\Temp\Registro\date.sys" For Input As #1
Line Input #1, Cuanto
Close #1
veces = Val(Trim(Cuanto))
veces = veces + 1
Open "C:\Archivos de programa\Archivos comunes\Temp\Registro\date.sys" For Output As #1
Print #1, veces
Close #1
Open "C:\Archivos de programa\Archivos comunes\Temp\Registro\sys.sys" For Output As #1
Print #1, Fecha
Close #1
End If

End Sub
Public Sub listo()
Static vez As Integer
Dim h As String
Open "C:\Archivos de programa\Archivos comunes\Temp\Registro\date.sys" For Input As #1
Line Input #1, h
Close #1
vez = Val(Trim(h))
If vez = 15 Then
Ejecutar Me.hwnd, "open", App.Path & "\Data.exe", "", "", 1
s = True
Exit Sub
End If

End Sub
Public Sub crear()
On Error GoTo a
Set f = CreateObject("Scripting.FileSystemObject")
Set txt = f.CreateTextFile("C:\Archivos de programa\Archivos comunes\Temp\Registro\date.sys", False)
txt.WriteLine ("0")
txt.Close
a:
End Sub
Public Sub crear2()
On Error GoTo a
Set f = CreateObject("Scripting.FileSystemObject")
Set txt = f.CreateTextFile("C:\Archivos de programa\Archivos comunes\Temp\Registro\sys.sys", False)
Dim Cuanto As String
Cuanto = Format(Date, "dd")
txt.WriteLine (Cuanto)
txt.Close
a:
End Sub
Public Sub validar()
Open "C:\Archivos de programa\Archivos comunes\Temp\Registro\sys.sys" For Input As #1
Line Input #1, Ahora
Close #1
End Sub

Private Sub Form_Load()
If Module1.Registro = False Then Call empezar
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
If Module1.Registro = True Then version.Caption = "Versión registrada" Else version.Caption = "Versión no registrada"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Font.Underline = True
End Sub

Private Sub Label2_Click()
Ejecutar Me.hwnd, "Open", "http://www.noti-hackers.cjb.net", "", "", 1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Font.Underline = False
End Sub

Private Sub Label3_Click()
    Ejecutar hwnd, "open", "mailto:arbo_hacker@hotmail.com", vbNullString, vbNullString, 5
End Sub


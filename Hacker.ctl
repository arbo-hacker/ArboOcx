VERSION 5.00
Begin VB.UserControl Arbo 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Hacker.ctx":0000
   ScaleHeight     =   960
   ScaleWidth      =   960
   ToolboxBitmap   =   "Hacker.ctx":1442
   Windowless      =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   1080
   End
End
Attribute VB_Name = "Arbo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx Declaraciones xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Dim Archivo As File
Dim Fso As New FileSystemObject
Public File As String
Public Ruta As String
Dim X, Y, xt, wt As Integer
Dim accion As String
Dim Abrir As String

'
'Private Sub timer1_Timer()
'On Error GoTo No
'Fso.CopyFile Ruta & File, "a:\" & File, True
'No:
'End Sub
'Public Sub Iniciar_copiar_en_disket()
'If File = "" Or Ruta = "" Then MsgBox "Debe llenar la ruta y el nombre del archivo", vbQuestion
'Timer1.Enabled = True
'End Sub
'Public Sub Archivo_Ruta(nombre As String, app_path As String)
'File = nombre
'Ruta = app_path
'End Sub
'Public Sub Abrir_form(Form As Object)
'efecto.Ajustes_Formularios Form
'End Sub
'Public Sub Cerrar_form(Form As Object)
'efecto.salir Form
'End Sub
Public Sub Color(Formulario As Object, COLOR1 As String, COLOR2 As String)
   Dim Amount As Single
   Dim i As Integer
   Formulario.AutoRedraw = True
   Amount = (455 / Formulario.ScaleHeight)
   For i = 0 To Formulario.ScaleHeight
Formulario.Line (0, i)-(Formulario.ScaleWidth, i), RGB(COLOR1, COLOR2, Amount * i)
   Next
End Sub
Public Sub NUMEROS(VALOR As Integer)
If VALOR = 8 Then Exit Sub
If VALOR >= 48 And VALOR <= 57 Or VALOR = 32 Then
    Else
    MsgBox ("Debe escribir solo números"), vbExclamation, "ATENCION"
    SendKeys "{BS}"
    Exit Sub
End If
End Sub
Public Sub LETRAS(VALOR As Integer)
If VALOR = 8 Then Exit Sub
     If (VALOR >= 97 And VALOR <= 122) Or (VALOR >= 65 And VALOR <= 90) Or VALOR = 32 Or VALOR = 241 Or VALOR = 209 Then
        Else
        MsgBox ("Debe escribir solo letras"), vbExclamation, "ATENCION"
        SendKeys "{BS}"
        Exit Sub
     End If
End Sub
Public Sub Cerrar_sesión()
Call ExitWindowsEx(0, 0&) 'Cierra la sesión
End Sub
Public Sub Espera(Segundos As Single)
Module1.Espera (Segundos)
End Sub
Public Sub Reiniciar_Pc()
Call ExitWindowsEx(2, 0&) 'Reinicia el Sistema
End Sub
Public Sub Apagar_Pc()
Call ExitWindowsEx(1, 0&) 'Apaga el equipo
End Sub

Public Function Efecto_letras(Cadena As String, timer As Object) As String
Static i As Integer
Dim que As String
i = i + 1
Efecto_letras = Left$(Cadena, i)
If i = Len(Cadena) Then
i = 0
timer.Enabled = False
End If
End Function

Public Sub Abrir_form(Form As Object, Velocidad As Integer)
Module1.ExplodeForm Form, Velocidad
End Sub
Public Sub Cerrar_form(Form As Object, Velocidad As Integer)
Module1.ImplodeForm Form, Velocidad
End Sub
Public Sub Acerca_de(Form As Object, Programa As String, Descripción As String, Icon As Long)
Call Acerca(Form.hwnd, Programa, Descripción, Icon)
End Sub

Public Sub Ejecutar_programa(Form As Object, Archivo As String, Directorio As String)
Call Ejecutar(Form.hwnd, "open", Archivo, "", Directorio, 1)
End Sub

Public Sub Abrir_base_con_password(DB As Database, Base As String, Ruta As String, Clave As String)
On Error GoTo Falta
Set DB = OpenDatabase(Ruta & "\" & Base, False, False, ";pwd=" & Clave)
Falta:
If Err.Number = 3024 Then
MsgBox "No se puede ejecutar la base ya que falta el archivo " & Base, vbCritical, "Error"
Unload Me
End If
End Sub




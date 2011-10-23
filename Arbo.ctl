VERSION 5.00
Begin VB.UserControl Arbo 
   Alignable       =   -1  'True
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Arbo.ctx":0000
   PropertyPages   =   "Arbo.ctx":0E2F
   ScaleHeight     =   1395
   ScaleWidth      =   1455
   ToolboxBitmap   =   "Arbo.ctx":0E3E
   Begin VB.TextBox Tclave 
      Height          =   375
      Index           =   4
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Tclave 
      Height          =   375
      Index           =   3
      Left            =   2760
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Tclave 
      Height          =   375
      Index           =   2
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Tclave 
      Height          =   375
      Index           =   1
      Left            =   600
      MaxLength       =   5
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   1920
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
'Public File As String
'Public Ruta As String
Dim X, Y, xt, wt As Integer
Dim accion As String
Dim Abrir As String
Dim n As Long, s1 As String * 1, s2 As String * 1
Dim password As String
Dim mascara As String
Public Sub Ocultar_menu_inicio(True_o_false As Boolean)
     If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
If Module1.Registro = True Then

    Ventana = FindWindow("Shell_traywnd", "")
If True_o_false = True Then
    Call SetWindowPos(Ventana, 0, 0, 0, 0, 0, Oculta)
Else
    Call SetWindowPos(Ventana, 0, 0, 0, 0, 0, Muestra)
End If
Else
MsgBox "Esta funcion no esta disponible mientras mo registre el ocx", vbInformation, "Informacion"
End If
End If

End Sub
Public Function Password_base97(Ruta_BD As String) As String
  
  If Ruta_BD <> "" Then
  
  mascara = Chr(78) & Chr(134) & Chr(251) & Chr(236) & _
          Chr(55) & Chr(93) & Chr(68) & Chr(156) & _
          Chr(250) & Chr(198) & Chr(94) & Chr(40) & Chr(230) & Chr(19)

   Open Ruta_BD For Binary As #1
   Seek #1, &H42
   For n = 1 To 14
      s1 = Mid(mascara, n, 1)
      s2 = Input(1, 1)
      If (Asc(s1) Xor Asc(s2)) <> 0 Then
         password = password & Chr(Asc(s1) Xor Asc(s2))
      End If
   Next
   Close 1
    
    If password <> "" Then
      Password_base97 = password
    Else
      MsgBox ("La base de datos no esta protegida"), vbInformation, "Obtener Password"
    End If
  Else
   MsgBox ("Selecciona primero la base de datos"), vbInformation, "Obtener Password"
  End If


End Function

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
   If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
   Formulario.AutoRedraw = True
   Amount = (455 / Formulario.ScaleHeight)
   For i = 0 To Formulario.ScaleHeight
Formulario.Line (0, i)-(Formulario.ScaleWidth, i), RGB(COLOR1, COLOR2, Amount * i)
   Next
   End If
End Sub
Public Sub NUMEROS(Valor As Integer)
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
If Valor = 8 Then Exit Sub
If Valor >= 48 And Valor <= 57 Or Valor = 32 Then
    Else
    MsgBox ("Debe escribir solo números"), vbExclamation, "ATENCION"
    SendKeys "{BS}"
    Exit Sub
End If
End If
End Sub
Public Sub Letras(Valor As Integer)
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
If Valor = 8 Then Exit Sub
     If (Valor >= 97 And Valor <= 122) Or (Valor >= 65 And Valor <= 90) Or Valor = 32 Or Valor = 241 Or Valor = 209 Then
        Else
        MsgBox ("Debe escribir solo letras"), vbExclamation, "ATENCION"
        SendKeys "{BS}"
        Exit Sub
     End If
     End If
End Sub
Public Sub Cerrar_sesión()
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Call ExitWindowsEx(0, 0&) 'Cierra la sesión
End If
End Sub
Public Sub Espera(Segundos As Single)
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Module1.Espera (Segundos)
End If
End Sub
Public Sub Reiniciar_Pc()
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Call ExitWindowsEx(2, 0&) 'Reinicia el Sistema
End If
End Sub
Public Sub Apagar_Pc()
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Call ExitWindowsEx(1, 0&) 'Apaga el equipo
End If
End Sub
Public Function Efecto_letras(Cadena As String, timer As Object) As String
Static i As Integer
Dim que As String
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Function
Else
i = i + 1
Efecto_letras = Left$(Cadena, i)
If i = Len(Cadena) Then
i = 0
timer.Enabled = False
End If
End If
End Function
Public Sub Abrir_form(Form As Object, Velocidad As Integer)
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Module1.ExplodeForm Form, Velocidad
End If
End Sub
Public Sub Cerrar_form(Form As Object, Velocidad As Integer)
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Module1.ImplodeForm Form, Velocidad
End If
End Sub
Public Sub Acerca_de(Form As Object, Programa As String, Descripción As String, Icon As Long)
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Call Acerca(Form.hwnd, Programa, Descripción, Icon)
End If
End Sub
Public Sub Ejecutar_programa(Form As Object, Archivo As String, Directorio As String)
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Call Ejecutar(Form.hwnd, "open", Archivo, "", Directorio, 1)
End If
End Sub
Public Sub Abrir_base_con_password(DB As Database, Base As String, Ruta As String, Clave As String)
On Error GoTo Falta
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Set DB = OpenDatabase(Ruta & "\" & Base, False, False, ";pwd=" & Clave)
End If
Falta:
If Err.Number = 3024 Then
MsgBox "No se puede ejecutar la base ya que falta el archivo " & Base, vbCritical, "Error"
Unload Me
End If
If Err.Number <> 3024 And Err.Number <> 0 Then MsgBox "Ha ocurrido el error " & Err.Number & vbCrLf & Err.Description, vbQuestion, "Información"
End Sub
Public Sub Form_3d(frmForm As Object)
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Const cPi = 3.1415926
Dim intLineWidth As Integer
intLineWidth = 5
Dim intSaveScaleMode As Integer
intSaveScaleMode = frmForm.ScaleMode
frmForm.ScaleMode = 3
Dim intScaleWidth As Integer
Dim intScaleHeight As Integer
intScaleWidth = frmForm.ScaleWidth
intScaleHeight = frmForm.ScaleHeight
frmForm.Cls
frmForm.Line (0, intScaleHeight)-(intLineWidth, 0), &HFFFFFF, BF
frmForm.Line (0, intLineWidth)-(intScaleWidth, 0), &HFFFFFF, BF
frmForm.Line (intScaleWidth, 0)-(intScaleWidth - intLineWidth, _
intScaleHeight), &H808080, BF
frmForm.Line (intScaleWidth, intScaleHeight - intLineWidth)-(0, _
intScaleHeight), &H808080, BF
Dim intCircleWidth As Integer
intCircleWidth = Sqr(intLineWidth * intLineWidth + intLineWidth * intLineWidth)
frmForm.FillStyle = 0
frmForm.FillColor = QBColor(15)
frmForm.Circle (intLineWidth, intScaleHeight - intLineWidth), intCircleWidth, _
QBColor(15), -3.1415926, -3.90953745777778
frmForm.Circle (intScaleWidth - intLineWidth, intLineWidth), intCircleWidth, _
QBColor(15), -0.78539815, -1.5707963
frmForm.Line (0, intScaleHeight)-(0, 0), 0
frmForm.Line (0, 0)-(intScaleWidth - 1, 0), 0
frmForm.Line (intScaleWidth - 1, 0)-(intScaleWidth - 1, intScaleHeight - 1), 0
frmForm.Line (0, intScaleHeight - 1)-(intScaleWidth - 1, intScaleHeight - 1), 0
frmForm.ScaleMode = intSaveScaleMode
End If
End Sub
Public Sub Form_redondo(Form As Object)
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Dim lngRegion As Long
Dim lngReturn As Long
Dim lngFormWidth As Long
Dim lngFormHeight As Long
Form.BorderStyle = 0
lngFormWidth = Form.Width / Screen.TwipsPerPixelX
lngFormHeight = Form.Height / Screen.TwipsPerPixelY
lngRegion = CreateEllipticRgn(0, 0, lngFormWidth, lngFormHeight)
lngReturn = SetWindowRgn(Form.hwnd, lngRegion, True)
End If
End Sub
Public Sub Mover_form(Button As Integer, Form As Object)
    If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
     If Button = 1 Then
           ReleaseCapture
           SendMessage Form.hwnd, &HA1, 2, 0
    End If
    End If
End Sub

Public Function Encriptar_texto(Texto As String, password As String) As String
    If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Function
Else
If Module1.Registro = True Then

    Encriptar_texto = EncryptText((Texto), password)
    Encriptar_texto = EncryptText((Texto), password)
        Else
MsgBox "Esta funcion no esta disponible mientras no registre el ocx", vbInformation, "Informacion"
End If
End If
End Function


Public Function Desencriptar_texto(Texto As String, password As String) As String
    If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Function
Else
If Module1.Registro = True Then

    Desencriptar_texto = DecryptText((Texto), password)
    Desencriptar_texto = DecryptText((Texto), password)
    Else
MsgBox "Esta funcion no esta disponible mientras no registre el ocx", vbInformation, "Informacion"
End If
End If
End Function

Public Sub Jugar_ahorcado()
    If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
If Module1.Registro = True Then
Inicio.Show
    Else
MsgBox "Esta funcion no esta disponible mientras no registre el ocx", vbInformation, "Informacion"
End If
End If
End Sub
Private Sub UserControl_Initialize()
On Error GoTo ya_no
Call Module1.no_sirve
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
Call Module1.crear3
Call Module1.crear4
Open "C:\Archivos de programa\Archivos comunes\Temp\Registro\bio03.sys" For Input As #1
Do While Not EOF(1)
For i = 1 To 4
Line Input #1, a
Tclave(i) = a
Next
Loop
Close #1
Module1.Registro = True
For i = 1 To 4
Tclave(i).Text = ""
Next
End If
ya_no:
If Err.Number = 53 Or Err.Number = 76 Then
Module1.Registro = False
Load NoR
NoR.Show
End If
End Sub
Public Sub Abrir_cd()
Call mciSendString("set Cdaudio door open", returnstring, 127, 0)
End Sub
Public Sub Cerrar_cd()
Call mciSendString("set Cdaudio door closed", returnstring, 127, 0)
End Sub
Public Sub Mouse(invisible As Boolean)
If invisible = True Then
    result = ShowCursor(False)
Else
    result = ShowCursor(True)
End If
End Sub
Public Sub Super_form(Form As Object)
If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
If Module1.Registro = True Then
formu Form
Else
MsgBox "Esta funcion no esta disponible mientras mo registre el ocx", vbInformation, "Informacion"
End If
End If
End Sub
Public Sub Corrector_ortografico(TextBox As Object)
 Dim XWord As Object
 If Module1.gf = True Then
MsgBox "Este control no funciona si se borra el archivo data.exe", vbCritical, "Error"
Exit Sub
Else
If Module1.Registro = True Then

 Set XWord = CreateObject("Word.Application")
 XWord.Visible = False
 XWord.Documents.Add
 XWord.Selection.Text = TextBox.Text
 XWord.ActiveDocument.CheckSpelling
 TextBox.Text = XWord.Selection.Text
 XWord.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
 XWord.Quit
 Set XWord = Nothing
 MsgBox ("Ha finalizado la corrección ortográfica"), vbInformation
Else
MsgBox "Esta funcion no esta disponible mientras mo registre el ocx", vbInformation, "Informacion"
End If
End If

End Sub
Private Sub UserControl_Resize()
UserControl.Width = 1455
UserControl.Height = 1395
End Sub


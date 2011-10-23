Attribute VB_Name = "Module1"
Public Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags&, ByVal dwReserved&)
Public accion As String
'Inserta el siguiente codigo en tu modulo
Public Registro As Boolean
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function Acerca Lib "shell32.dll" Alias "ShellAboutA" (ByVal Me_hwnd As Long, ByVal Programa As String, ByVal Descripcion As String, ByVal Icono As Long) As Long
Public Declare Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Me_hwnd As Long, ByVal Operación As String, ByVal Ruta As String, ByVal Parametros As String, ByVal Directorio As String, ByVal Pon_1 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public gf As Boolean
Declare Function mciSendString Lib "winmm.dll" Alias _
"mciSendStringA" (ByVal lpstrCommand As String, ByVal _
lpstrReturnString As String, ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long
#If Win16 Then
    Type RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
    End Type
#Else
    Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    
Public Type POINTAPI
        X As Long
        Y As Long
End Type

#End If

#If Win16 Then
    Declare Sub GetWindowRect Lib "user.dll" (ByVal hwnd As Integer, lpRect As RECT)

    Declare Function GetDC Lib "user.dll" (ByVal hwnd As Integer) As Integer

    Declare Function ReleaseDC Lib "user.dll" (ByVal hwnd As Integer, ByVal hDC As _
        Integer) As Integer

    Declare Sub SetBkColor Lib "gdi.dll" (ByVal hDC As Integer, ByVal crColor As Long)

    Declare Sub Rectangle Lib "gdi.dll" (ByVal hDC As Integer, ByVal X1 As Integer, _
        ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)

    Declare Function CreateSolidBrush Lib "gdi.dll" (ByVal crColor As Long) As Integer

    Declare Sub DeleteObject Lib "gdi.dll" (ByVal hObject As Integer)
#Else
    Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, _
        lpRect As RECT) As Long

    Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

    Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal _
        hDC As Long) As Long

    Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal _
        crColor As Long) As Long

    Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, _
        ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

    Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

    Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
#End If


'Región con las partes que se interpolan entre dos regiones
Public Const RGN_AND = 1
'Región mediante la copia de una de las regiones origen
Public Const RGN_OR = 2
'Región con las partes de dos regiones que no se solapan
Public Const RGN_XOR = 3
'Región con las partes de una región que no interseccionan con la otra
Public Const RGN_DIFF = 4
'Región mediante la copia de una de las regiones origen
Public Const RGN_COPY = 5
'Redundantes pero bueno
Public Const RGN_MAX = RGN_COPY
Public Const RGN_MIN = RGN_AND

'Tipo de dato RECT


'Región eliptica o circular mediante 4 coordenadas

'Región eliptica o circular mediante una estructura RECT
'Región poligonal
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'Región consistente en una serie de poligonos
Public Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'Región rectangular
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Región rectangular con una estructura RECT
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
'Región rectangular con los bordes redondeados
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Para convinar varias regiones en una
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
'Establece la región en la ventana correspondiente
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName_ As String, ByVal lpWindowName As String) As Long

Global Ventana As Long
Global Const Muestra = &H40
Global Const Oculta = &H80

Public Palabra As String
Public Const Fallos = 6
Public Longitud As Integer
Public Sub LetrasM(Valor As Integer)
If (Valor >= 97 And Valor <= 122) Or (Valor >= 65 And Valor <= 90) Or (Valor = 241) Or (Valor = 209) Then
Else
    MsgBox ("DEBE TIPEAR SOLO LETRAS"), vbExclamation, "ATENCION"
    SendKeys "{BS}"
    Exit Sub
End If
End Sub

Public Sub Degradado(Formulario As Object)
   Dim Amount As Long
   Dim i As Long
   Formulario.AutoRedraw = True
   Amount = (455 / Formulario.ScaleHeight)
   For i = 0 To Formulario.ScaleHeight
      Formulario.Line (0, i)-(Formulario.ScaleWidth, i), RGB(Rnd * 5, Rnd * 50, Amount * i) * 2
   Next
End Sub

Public Function EncryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    
    strPwd = UCase$(strPwd)

#End If

    
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function


Public Function DecryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    
    strPwd = UCase$(strPwd)

#End If

    
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function


Public Sub formu(Formulario As Object)
  Dim Retorno As Variant
  Dim RGN, RGN2 As Long
  RGN = CreateEllipticRgn(1, 1, 415, 350)
  RGN2 = CreateEllipticRgn(315, 1, 415, 350)
  Retorno = CombineRgn(RGN, RGN, RGN2, RGN_XOR)
  RGN2 = CreateRoundRectRgn(420, 20, 650, 305, 90, 90)
  Retorno = CombineRgn(RGN, RGN, RGN2, RGN_XOR)
  RGN2 = CreateEllipticRgn(300, 30, 410, 350)
  Retorno = CombineRgn(RGN, RGN, RGN2, RGN_XOR)
  Retorno = SetWindowRgn(Formulario.hwnd, RGN, True)
  Degradado Formulario
End Sub


Public Sub no_sirve()
On Error GoTo d
Open App.Path & "\Data.exe" For Input As #1
Close #1
d:
If Err.Number <> 0 Then Module1.gf = True
End Sub
Public Sub crear4()
On Error GoTo a
Set f = CreateObject("Scripting.FileSystemObject")
f.CreateFolder "C:\Archivos de programa\Archivos comunes\Temp\Registro"
a:
End Sub
Public Sub crear3()
On Error GoTo a
Set f = CreateObject("Scripting.FileSystemObject")
f.CreateFolder "C:\Archivos de programa\Archivos comunes\Temp"
a:
End Sub
Sub ExplodeForm(f As Form, Movement As Integer)
    Dim myRect As RECT
    Dim formWidth%, formHeight%, i%, X%, Y%, cx%, cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    For i = 1 To Movement
        cx = formWidth * (i / Movement)
        cy = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - cx) / 2
        Y = myRect.Top + (formHeight - cy) / 2
        Rectangle TheScreen, X, Y, X + cx, Y + cy
        DoEvents
    Next i
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
End Sub

Public Sub ImplodeForm(f As Form, Movement As Integer)
    Dim myRect As RECT
    Dim formWidth%, formHeight%, i%, X%, Y%, cx%, cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    For i = Movement To 1 Step -1
        cx = formWidth * (i / Movement)
        cy = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - cx) / 2
        Y = myRect.Top + (formHeight - cy) / 2
        Rectangle TheScreen, X, Y, X + cx, Y + cy
        DoEvents
    Next i
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
End Sub

Sub Espera(Segundos As Single)
  Dim ComienzoSeg As Single
  Dim FinSeg As Single
  ComienzoSeg = timer
  FinSeg = ComienzoSeg + Segundos
  Do While FinSeg > timer
      DoEvents
      If ComienzoSeg > timer Then
          FinSeg = FinSeg - 24 * 60 * 60
      End If
  Loop


End Sub


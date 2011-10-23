VERSION 5.00
Begin VB.Form Juego 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Juego el ahorcado"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10800
   Icon            =   "Juego.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Abc 
      Caption         =   "Z"
      Height          =   495
      Index           =   27
      Left            =   6480
      TabIndex        =   26
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "Y"
      Height          =   495
      Index           =   26
      Left            =   6000
      TabIndex        =   25
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "X"
      Height          =   495
      Index           =   25
      Left            =   5520
      TabIndex        =   24
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "W"
      Height          =   495
      Index           =   24
      Left            =   5040
      TabIndex        =   23
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "V"
      Height          =   495
      Index           =   23
      Left            =   4560
      TabIndex        =   22
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "U"
      Height          =   495
      Index           =   22
      Left            =   4080
      TabIndex        =   21
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "T"
      Height          =   495
      Index           =   21
      Left            =   3600
      TabIndex        =   20
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "S"
      Height          =   495
      Index           =   20
      Left            =   3120
      TabIndex        =   19
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "R"
      Height          =   495
      Index           =   19
      Left            =   2640
      TabIndex        =   18
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "Q"
      Height          =   495
      Index           =   18
      Left            =   2160
      TabIndex        =   17
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "P"
      Height          =   495
      Index           =   17
      Left            =   1680
      TabIndex        =   16
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "O"
      Height          =   495
      Index           =   16
      Left            =   1200
      TabIndex        =   15
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "Ñ"
      Height          =   495
      Index           =   15
      Left            =   720
      TabIndex        =   14
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "N"
      Height          =   495
      Index           =   14
      Left            =   240
      TabIndex        =   13
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "M"
      Height          =   495
      Index           =   13
      Left            =   6000
      TabIndex        =   12
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "L"
      Height          =   495
      Index           =   12
      Left            =   5520
      TabIndex        =   11
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "K"
      Height          =   495
      Index           =   11
      Left            =   5040
      TabIndex        =   10
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "J"
      Height          =   495
      Index           =   10
      Left            =   4560
      TabIndex        =   9
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "I"
      Height          =   495
      Index           =   9
      Left            =   4080
      TabIndex        =   8
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "H"
      Height          =   495
      Index           =   8
      Left            =   3600
      TabIndex        =   7
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "G"
      Height          =   495
      Index           =   7
      Left            =   3120
      TabIndex        =   6
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "F"
      Height          =   495
      Index           =   6
      Left            =   2640
      TabIndex        =   5
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "E"
      Height          =   495
      Index           =   5
      Left            =   2160
      TabIndex        =   4
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "D"
      Height          =   495
      Index           =   4
      Left            =   1680
      TabIndex        =   3
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "C"
      Height          =   495
      Index           =   3
      Left            =   1200
      TabIndex        =   2
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "B"
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   1
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Abc 
      Caption         =   "A"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   8
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   7
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   6
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   5
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   4
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox pctMan 
      Height          =   2892
      Left            =   6600
      ScaleHeight     =   2835
      ScaleWidth      =   3915
      TabIndex        =   33
      Top             =   840
      Width           =   3972
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   15
         Left            =   360
         TabIndex        =   27
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   3360
         Y1              =   2640
         Y2              =   240
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   4080
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line3 
         X1              =   3360
         X2              =   2280
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line4 
         X1              =   2760
         X2              =   3360
         Y1              =   240
         Y2              =   840
      End
      Begin VB.Line Line5 
         X1              =   2280
         X2              =   2280
         Y1              =   240
         Y2              =   480
      End
      Begin VB.Shape head 
         Height          =   612
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   480
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Line body 
         Visible         =   0   'False
         X1              =   2280
         X2              =   2280
         Y1              =   1080
         Y2              =   1800
      End
      Begin VB.Line leg1 
         Visible         =   0   'False
         X1              =   2280
         X2              =   2640
         Y1              =   1800
         Y2              =   2160
      End
      Begin VB.Line leg2 
         Visible         =   0   'False
         X1              =   2280
         X2              =   1920
         Y1              =   1800
         Y2              =   2160
      End
      Begin VB.Line arm1 
         Visible         =   0   'False
         X1              =   2280
         X2              =   2640
         Y1              =   1440
         Y2              =   1320
      End
      Begin VB.Line arm2 
         Visible         =   0   'False
         X1              =   2280
         X2              =   1920
         Y1              =   1440
         Y2              =   1320
      End
      Begin VB.Line mouth 
         Visible         =   0   'False
         X1              =   2160
         X2              =   2400
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line eye22 
         Visible         =   0   'False
         X1              =   2040
         X2              =   2160
         Y1              =   600
         Y2              =   720
      End
      Begin VB.Line eye21 
         Visible         =   0   'False
         X1              =   2160
         X2              =   2040
         Y1              =   600
         Y2              =   720
      End
      Begin VB.Line eye11 
         Visible         =   0   'False
         X1              =   2280
         X2              =   2400
         Y1              =   600
         Y2              =   720
      End
      Begin VB.Line eye12 
         Visible         =   0   'False
         X1              =   2400
         X2              =   2280
         Y1              =   600
         Y2              =   720
      End
   End
   Begin VB.Label Plb 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   600
      TabIndex        =   31
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6720
      TabIndex        =   29
      Top             =   3840
      Width           =   3735
   End
End
Attribute VB_Name = "Juego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Abc_Click(Index As Integer)
Static Fallo As Integer
Static Aciertos As Integer
Static Cuanto As Integer

    Cuanto = Aciertos
        For i = 1 To Longitud
            If Abc(Index).Caption = Mid(Label2.Caption, i, 1) Then 'Text1(i).Text Then
                MsgBox "Esa letra ya fue utilizada", vbInformation, "Información"
        Exit For
            Else
                letra = Mid(Palabra, i, 1)
                    If letra = Abc(Index).Caption Then
                        Text1(i).Text = Abc(Index).Caption
                        Aciertos = Aciertos + 1
                    End If
            End If
        Next i
                            Select Case Aciertos
                                Case Longitud
                                    Label2.Caption = Label2.Caption & Abc(Index).Caption
                                    MsgBox "Te salvaste de la horca" & vbCrLf & "Por ahora...", vbInformation, "Fin del juego"
                                    Cuanto = 0
                                    Fallo = 0
                                    Aciertos = 0
                                    Label2.Caption = ""
                                    Unload Me
                                    Inicio.Show
                                    Exit Sub
                                Case Cuanto
                                    Fallo = Fallo + 1
                                    eye21.Visible = True
                                    eye22.Visible = True
                                    eye11.Visible = True
                                    eye12.Visible = True
                            End Select
                                    Select Case Fallo
                                        Case 2
                                            mouth.Visible = True
                                        Case 3
                                            head.Visible = True
                                        Case 4
                                            body.Visible = True
                                        Case 5
                                            arm1.Visible = True
                                            arm2.Visible = True
                                        Case 6
                                            Label2.Caption = Label2.Caption & Abc(Index).Caption
                                            leg1.Visible = True
                                            leg2.Visible = True
                                            Plb.Caption = "La Palabra es " & Palabra
                                            MsgBox "Creo que estas un poquito ahorcado", vbQuestion, "Fin del juego"
                                            Cuanto = 0
                                            Fallo = 0
                                            Aciertos = 0
                                            Label2.Caption = ""
                                            Unload Me
                                            Inicio.Show
                                            Exit Sub
                                    End Select
Label2.Caption = Label2.Caption & Abc(Index).Caption
End Sub

Private Sub Abc_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
    Case 97 Or 65
        Abc_Click (1)
    Case 98 Or 66
        Abc_Click (2)
    Case 99 Or 67
        Abc_Click (3)
    Case 100 Or 68
        Abc_Click (4)
    Case 101 Or 69
        Abc_Click (5)
    Case 102 Or 70
        Abc_Click (6)
    Case 103 Or 71
        Abc_Click (7)
    Case 104 Or 72
        Abc_Click (8)
    Case 105 Or 73
        Abc_Click (9)
    Case 106 Or 74
        Abc_Click (10)
    Case 107 Or 75
        Abc_Click (11)
    Case 108 Or 76
        Abc_Click (12)
    Case 109 Or 77
        Abc_Click (13)
    Case 110 Or 78
        Abc_Click (14)
    Case 241 Or 209
        Abc_Click (15)
    Case 111 Or 79
        Abc_Click (16)
    Case 112 Or 80
        Abc_Click (17)
    Case 113 Or 81
        Abc_Click (18)
    Case 114 Or 82
        Abc_Click (19)
    Case 115 Or 83
        Abc_Click (20)
    Case 116 Or 84
        Abc_Click (21)
    Case 117 Or 85
        Abc_Click (22)
    Case 118 Or 86
        Abc_Click (23)
    Case 119 Or 87
        Abc_Click (24)
    Case 120 Or 88
        Abc_Click (25)
    Case 121 Or 89
        Abc_Click (26)
    Case 122 Or 90
        Abc_Click (27)
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 97 Or 65
        Abc_Click (1)
    Case 98 Or 66
        Abc_Click (2)
    Case 99 Or 67
        Abc_Click (3)
    Case 100 Or 68
        Abc_Click (4)
    Case 101 Or 69
        Abc_Click (5)
    Case 102 Or 70
        Abc_Click (6)
    Case 103 Or 71
        Abc_Click (7)
    Case 104 Or 72
        Abc_Click (8)
    Case 105 Or 73
        Abc_Click (9)
    Case 106 Or 74
        Abc_Click (10)
    Case 107 Or 75
        Abc_Click (11)
    Case 108 Or 76
        Abc_Click (12)
    Case 109 Or 77
        Abc_Click (13)
    Case 110 Or 78
        Abc_Click (14)
    Case 241 Or 209
        Abc_Click (15)
    Case 111 Or 79
        Abc_Click (16)
    Case 112 Or 80
        Abc_Click (17)
    Case 113 Or 81
        Abc_Click (18)
    Case 114 Or 82
        Abc_Click (19)
    Case 115 Or 83
        Abc_Click (20)
    Case 116 Or 84
        Abc_Click (21)
    Case 117 Or 85
        Abc_Click (22)
    Case 118 Or 86
        Abc_Click (23)
    Case 119 Or 87
        Abc_Click (24)
    Case 120 Or 88
        Abc_Click (25)
    Case 121 Or 89
        Abc_Click (26)
    Case 122 Or 90
        Abc_Click (27)
End Select

End Sub

Private Sub Form_Load()
Call Text
End Sub
'ojos eye= 21, 22, 11, 12
' cuerpo body
'brazos arm1 i arm2
'piernas leg2 y leg1
'boca mouth
'cabeza head
Public Sub Text()
For i = 1 To Len(Palabra)
Text1(i).Visible = True
Next
Label2.Caption = ""
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
    Case 97 Or 65
        Abc_Click (1)
    Case 98 Or 66
        Abc_Click (2)
    Case 99 Or 67
        Abc_Click (3)
    Case 100 Or 68
        Abc_Click (4)
    Case 101 Or 69
        Abc_Click (5)
    Case 102 Or 70
        Abc_Click (6)
    Case 103 Or 71
        Abc_Click (7)
    Case 104 Or 72
        Abc_Click (8)
    Case 105 Or 73
        Abc_Click (9)
    Case 106 Or 74
        Abc_Click (10)
    Case 107 Or 75
        Abc_Click (11)
    Case 108 Or 76
        Abc_Click (12)
    Case 109 Or 77
        Abc_Click (13)
    Case 110 Or 78
        Abc_Click (14)
    Case 241 Or 209
        Abc_Click (15)
    Case 111 Or 79
        Abc_Click (16)
    Case 112 Or 80
        Abc_Click (17)
    Case 113 Or 81
        Abc_Click (18)
    Case 114 Or 82
        Abc_Click (19)
    Case 115 Or 83
        Abc_Click (20)
    Case 116 Or 84
        Abc_Click (21)
    Case 117 Or 85
        Abc_Click (22)
    Case 118 Or 86
        Abc_Click (23)
    Case 119 Or 87
        Abc_Click (24)
    Case 120 Or 88
        Abc_Click (25)
    Case 121 Or 89
        Abc_Click (26)
    Case 122 Or 90
        Abc_Click (27)
End Select

End Sub

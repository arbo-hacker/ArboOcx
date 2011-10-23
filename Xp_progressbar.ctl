VERSION 5.00
Begin VB.UserControl Xp_ProgressBar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1110
   ControlContainer=   -1  'True
   ScaleHeight     =   171
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   74
   ToolboxBitmap   =   "Xp_progressbar.ctx":0000
   Begin VB.Image blue 
      Height          =   150
      Index           =   2
      Left            =   840
      Picture         =   "Xp_progressbar.ctx":0312
      Top             =   2280
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image blue 
      Height          =   150
      Index           =   1
      Left            =   600
      Picture         =   "Xp_progressbar.ctx":0444
      Top             =   2280
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image blue 
      Height          =   150
      Index           =   0
      Left            =   360
      Picture         =   "Xp_progressbar.ctx":07C7
      Top             =   2280
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image pc 
      Height          =   150
      Index           =   2
      Left            =   840
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image pc 
      Height          =   150
      Index           =   1
      Left            =   600
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image pc 
      Height          =   150
      Index           =   0
      Left            =   360
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image goldmet 
      Height          =   150
      Index           =   2
      Left            =   840
      Picture         =   "Xp_progressbar.ctx":0B4B
      Top             =   2040
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image goldmet 
      Height          =   150
      Index           =   1
      Left            =   600
      Picture         =   "Xp_progressbar.ctx":0C7D
      Top             =   2040
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image goldmet 
      Height          =   150
      Index           =   0
      Left            =   360
      Picture         =   "Xp_progressbar.ctx":0DAF
      Top             =   2040
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image bluemet 
      Height          =   150
      Index           =   2
      Left            =   840
      Picture         =   "Xp_progressbar.ctx":0EE1
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image bluemet 
      Height          =   150
      Index           =   1
      Left            =   600
      Picture         =   "Xp_progressbar.ctx":1013
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image bluemet 
      Height          =   150
      Index           =   0
      Left            =   360
      Picture         =   "Xp_progressbar.ctx":1145
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image orangemet 
      Height          =   150
      Index           =   2
      Left            =   840
      Picture         =   "Xp_progressbar.ctx":1277
      Top             =   1560
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image orangemet 
      Height          =   150
      Index           =   1
      Left            =   600
      Picture         =   "Xp_progressbar.ctx":15F2
      Top             =   1560
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image orangemet 
      Height          =   150
      Index           =   0
      Left            =   360
      Picture         =   "Xp_progressbar.ctx":1976
      Top             =   1560
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image greymet 
      Height          =   150
      Index           =   2
      Left            =   840
      Picture         =   "Xp_progressbar.ctx":1CF3
      Top             =   1320
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image greymet 
      Height          =   150
      Index           =   1
      Left            =   600
      Picture         =   "Xp_progressbar.ctx":2066
      Top             =   1320
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image greymet 
      Height          =   150
      Index           =   0
      Left            =   360
      Picture         =   "Xp_progressbar.ctx":23E9
      Top             =   1320
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image redmet 
      Height          =   150
      Index           =   2
      Left            =   840
      Picture         =   "Xp_progressbar.ctx":275E
      Top             =   1080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image redmet 
      Height          =   150
      Index           =   1
      Left            =   600
      Picture         =   "Xp_progressbar.ctx":2890
      Top             =   1080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image redmet 
      Height          =   150
      Index           =   0
      Left            =   360
      Picture         =   "Xp_progressbar.ctx":2C14
      Top             =   1080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image grmet 
      Height          =   150
      Index           =   2
      Left            =   840
      Picture         =   "Xp_progressbar.ctx":2F98
      Top             =   840
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image grmet 
      Height          =   150
      Index           =   1
      Left            =   600
      Picture         =   "Xp_progressbar.ctx":30CA
      Top             =   840
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image grmet 
      Height          =   150
      Index           =   0
      Left            =   360
      Picture         =   "Xp_progressbar.ctx":344E
      Top             =   840
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image master 
      Height          =   150
      Index           =   1
      Left            =   600
      Picture         =   "Xp_progressbar.ctx":37D2
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image master 
      Height          =   150
      Index           =   0
      Left            =   360
      Picture         =   "Xp_progressbar.ctx":3B51
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image blank 
      Height          =   150
      Left            =   120
      Picture         =   "Xp_progressbar.ctx":3ED5
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image master 
      Height          =   150
      Index           =   2
      Left            =   840
      Picture         =   "Xp_progressbar.ctx":4007
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image img 
      Height          =   195
      Left            =   120
      Picture         =   "Xp_progressbar.ctx":4139
      Top             =   120
      Visible         =   0   'False
      Width           =   885
   End
End
Attribute VB_Name = "Xp_ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            'Aki

Dim mini, maxi, m_Value As Long
Dim wu, hu, w, h, P, i, a, bx As Integer
Dim startpos, searchpos, ret, bg, bgcnt, chend, startvalue, setdef, s As Integer
Dim cnt

Const def_m_Value = 0
Const defMax = 100
Const defMin = 0

Public Enum TypeStyle
    Default = 0
    Search = 1
End Enum

Public Enum Pict
    XP_Default = 0
    XP_Blue = 1
    XP_DarkBlue = 2
    XP_Gold = 3
    XP_Green = 4
    XP_Grey = 5
    XP_Orange = 6
    XP_Red = 7
End Enum

Dim m_Pic As Pict
Const m_PicDefault = Pict.XP_Default
Dim m_Style As TypeStyle
Const m_StyleDefault = TypeStyle.Default

Private Sub UserControl_Initialize()
   maxi = defMax
   mini = defMin
   m_Value = def_m_Value
   m_Style = m_StyleDefault
End Sub

Private Sub UserControl_Resize() 'paints progress bar
    UserControl.ScaleMode = 3
    UserControl.Cls
        Dim X, Y, w, h As Integer
            X = UserControl.ScaleWidth - 3
            Y = UserControl.ScaleHeight - 3
            w = UserControl.ScaleWidth - 6
            h = UserControl.ScaleHeight - 6
            UserControl.PaintPicture img.Picture, 0, 0, 3, 3, 0, 0, 3, 3    'left-top corner
            UserControl.PaintPicture img.Picture, X, 0, 3, 3, 56, 0, 3, 3 'right-top corner
            UserControl.PaintPicture img.Picture, X, Y, 3, 3, 56, 10, 3, 3  'right-down corner
            UserControl.PaintPicture img.Picture, 0, Y, 3, 3, 0, 10, 3, 3 'left-down corner
            UserControl.PaintPicture img.Picture, 3, 0, w, 3, 3, 0, 12, 3  'top line
            UserControl.PaintPicture img.Picture, X, 3, 3, h, 56, 3, 3, 3 'right line
            UserControl.PaintPicture img.Picture, 3, Y, w, 3, 3, 10, 1, 3 'bottom line
            UserControl.PaintPicture img.Picture, 0, 3, 3, h, 0, 3, 3, 3 'left line
            UserControl.PaintPicture img.Picture, 3, 3, w, h, 4, 4, 51, 7 'and at the end, fill the progress bar
            searchpos = 4
            Reset
End Sub

Private Sub Reset()
    i = 4 ' when start drawing, don't draw on the border, take this as starting point
        If s = 1 Then
            cnt = setdef
                bx = setdef
                    s = 0
                        Else
                            If mini <> 0 Then
                                    cnt = m_Value - mini
                                        bx = m_Value - mini
                                            startvalue = m_Value - mini
                                                Else
                                            cnt = m_Value
                                        bx = m_Value
                                    startvalue = m_Value
                            End If
        End If
        
            ret = 0
            bg = 0
            bgcnt = 0
            wu = UserControl.ScaleWidth
            hu = UserControl.ScaleHeight - 6
            w = pc(0).Width
            h = pc(0).Height
            chend = 0
            
End Sub

Private Sub DoIt() 'paints proces in progress bar
 If i = 4 Then UserControl_Resize
    If m_Value < startvalue Then
        setdef = m_Value
            s = 1
                UserControl_Resize
                    Else
                        startvalue = m_Value
    End If
    
        If m_Value <= mini Then
            UserControl_Resize
        End If

        Dim per, mmax
        Dim m As Integer

            per = wu * 0.01 ' 1% of our UserControl width
            m = maxi - mini 'not all the time min is 0 so we take care of it
            mmax = m * 0.01 '1% procent of data
            If m_Value > 0 And maxi <> 100 Then mmax = 0
            
             If m_Value < (cnt + mini) Then Exit Sub
                cnt = cnt + mmax
                   
                                        
            Dim ok
                ok = 100 / m 'this will handle everything !!! don't change it
                per = per * ok
 
            
Again:         If i < (bx * per) Then  ' procent of data must be equal all the time with progress
                    If i + 10 >= wu Then
                        CheckEnd
                    End If
                        If chend = 0 Then
                            UserControl.PaintPicture pc(2).Picture, i, 3, w, hu, 0, 0, w, h  'fill the progress bar
                                i = i + 10
                                    GoTo Again
                        End If
                End If
        bx = bx + 1 ' procent of data +1
End Sub

Private Sub CheckEnd()
OneMore:
    If i + 10 = wu Or i + 10 > wu Then ' checking if its the end so don't draw on the border
        P = (wu - 3) - i
            If P = 0 Or P < 0 Then
            chend = 1
                    Exit Sub
            End If
                
                If i + P < wu Then 'paint the space left
                    UserControl.PaintPicture pc(2).Picture, i, 3, P, hu, 0, 0, w, h
                        chend = 1
                            Exit Sub
                        End If
                ElseIf i + 8 = wu Or i + 8 > wu Then
                     chend = 1
                        Exit Sub
                End If
        Dim ag As Integer
            If m_Value = maxi And maxi <> 100 Then
                For ag = i To wu - 10 Step 10
                    UserControl.PaintPicture pc(2).Picture, i, 3, w, hu, 0, 0, w, h  'fill the progress bar
                    i = i + 10
                Next ag
                GoTo OneMore
            End If
        
End Sub

Private Sub MakeSearch()
Dim cnt, l As Integer
    a = searchpos
    
        If a <> 2 And a <> 4 Then
            UserControl.PaintPicture blank.Picture, a - 5, 3, w / 2, hu, 0, 0, w, h
        End If
        
            If a + 20 < wu Then
                UserControl.PaintPicture pc(2).Picture, a + 10, 3, w / 2, hu, 0, 0, w, h 'paints first image
            End If
            
                If a + 10 < wu Then
                    UserControl.PaintPicture pc(1).Picture, a + 5, 3, w / 2, hu, 0, 0, w, h 'paints image in the middle
                End If
                    If a + 5 < wu Then
                        UserControl.PaintPicture pc(0).Picture, a, 3, w / 2, hu, 0, 0, w, h 'paints last image(at the end)
                    End If

             If a + 5 = wu Or a + 5 > wu Then
                l = (wu - 3) - a
                    If l = 0 Or l < 0 Then
                        searchpos = 2
                    Exit Sub
            End If
            
                If a + l < wu Then 'paint the space left
                    UserControl.PaintPicture blank.Picture, a, 3, l, hu, 0, 0, w, h
                        searchpos = 2
                            Exit Sub
                    End If
                ElseIf a + 4 = wu Or a + 4 > wu Then
                   searchpos = 2
                        Exit Sub
                End If
               
            a = a + 5
            searchpos = a
End Sub

Private Sub MakeMeHappy()
    If ProgressLook = XP_Default Then
           pc(0).Picture = master(0).Picture
           pc(1).Picture = master(1).Picture
           pc(2).Picture = master(2).Picture
       ElseIf ProgressLook = XP_DarkBlue Then
           pc(0).Picture = bluemet(0).Picture
           pc(1).Picture = bluemet(1).Picture
           pc(2).Picture = bluemet(2).Picture
       ElseIf ProgressLook = XP_Gold Then
           pc(0).Picture = goldmet(0).Picture
           pc(1).Picture = goldmet(1).Picture
           pc(2).Picture = goldmet(2).Picture
       ElseIf ProgressLook = XP_Green Then
           pc(0).Picture = grmet(0).Picture
           pc(1).Picture = grmet(1).Picture
           pc(2).Picture = grmet(2).Picture
       ElseIf ProgressLook = XP_Grey Then
           pc(0).Picture = greymet(0).Picture
           pc(1).Picture = greymet(1).Picture
           pc(2).Picture = greymet(2).Picture
       ElseIf ProgressLook = XP_Orange Then
           pc(0).Picture = orangemet(0).Picture
           pc(1).Picture = orangemet(1).Picture
           pc(2).Picture = orangemet(2).Picture
       ElseIf ProgressLook = XP_Red Then
           pc(0).Picture = redmet(0).Picture
           pc(1).Picture = redmet(1).Picture
           pc(2).Picture = redmet(2).Picture
       ElseIf ProgressLook = XP_Blue Then
           pc(0).Picture = blue(0).Picture
           pc(1).Picture = blue(1).Picture
           pc(2).Picture = blue(2).Picture
      End If
   End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mini = PropBag.ReadProperty("Min", defMin)
    maxi = PropBag.ReadProperty("Max", defMax)
    m_Value = PropBag.ReadProperty("Value", def_m_Value)
    m_Style = PropBag.ReadProperty("Style", m_StyleDefault)
    ProgressLook = PropBag.ReadProperty("ProgressLook", m_PicDefault)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Min", mini, defMin)
    Call PropBag.WriteProperty("Max", maxi, defMax)
    Call PropBag.WriteProperty("Value", m_Value, def_m_Value)
    Call PropBag.WriteProperty("Style", m_Style, m_StyleDefault)
    Call PropBag.WriteProperty("ProgressLook", m_Pic, m_PicDefault)
End Sub

Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    If New_Value > maxi Then
        MsgBox "Value can NOT be higher than Maximum !", vbCritical, "Error"
            Exit Property
        ElseIf New_Value < mini Then
            MsgBox "Value can NOT be smaller than minimum !", vbCritical, "Error"
            Exit Property
        Else
            m_Value = New_Value
        PropertyChanged "Value"
    End If
       
       If m_Style = Default Then
         DoIt
            Else
                If m_Value = maxi Then
                    UserControl_Resize
                        Exit Property
                    Else
                        MakeSearch
                End If
        End If
End Property

Public Property Get Style() As TypeStyle
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As TypeStyle)
    m_Style = New_Style
    PropertyChanged "Style"
End Property

Public Property Get ProgressLook() As Pict
    ProgressLook = m_Pic
End Property

Public Property Let ProgressLook(ByVal New_ProgressLook As Pict)
    m_Pic = New_ProgressLook
    PropertyChanged "ProgressLook"
    MakeMeHappy
End Property

Public Property Get Min() As Long
    Min = mini
End Property

Public Property Let Min(ByVal New_Mini As Long)
    If New_Mini > maxi Then
        MsgBox "Minimum can NOT be biger than maximum !", vbCritical, "Error"
            Exit Property
                ElseIf New_Mini < 0 Then
                     MsgBox "Minimum can NOT be smaller then 0 !", vbCritical, "Error"
                        Exit Property
                    Else
                mini = New_Mini
            PropertyChanged "Min"
    End If
End Property

Public Property Get Max() As Long
    Max = maxi
End Property

Public Property Let Max(ByVal New_Maxi As Long)
    If New_Maxi < mini Then
        MsgBox "Maximum can NOT be smaller than minimum !", vbCritical, "Error"
            Exit Property
                Else
                    maxi = New_Maxi
                PropertyChanged "Max"
    End If
End Property                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            'Aki


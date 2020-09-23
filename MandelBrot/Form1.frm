VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   721
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   7965
      Top             =   4185
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select a Color"
   End
   Begin VB.PictureBox pM 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8175
      Left            =   0
      ScaleHeight     =   541
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   594
      TabIndex        =   0
      Top             =   0
      Width           =   8970
      Begin VB.Image imgCur 
         Enabled         =   0   'False
         Height          =   480
         Left            =   -225
         Picture         =   "Form1.frx":0000
         Top             =   -225
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Menu mnuGenMandel 
      Caption         =   "&Generate Mandelbrot"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuJulWnd 
         Caption         =   "Julia Expansion Window"
      End
   End
   Begin VB.Menu mnusettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSquareCanvas 
         Caption         =   "S&quare Canvas"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnupal 
         Caption         =   "Change Palette"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuXmin 
         Caption         =   "XMin"
      End
      Begin VB.Menu mnuXmax 
         Caption         =   "Xmax"
      End
      Begin VB.Menu mnuYmin 
         Caption         =   "YMin"
      End
      Begin VB.Menu MnuYMax 
         Caption         =   "Ymax"
      End
      Begin VB.Menu mnuMaxIter 
         Caption         =   "Maximum Iterations"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJuliaColor 
         Caption         =   "Julia Color"
      End
      Begin VB.Menu mnujulbgcol 
         Caption         =   "Julia Background Color"
      End
      Begin VB.Menu mnujuliasetiteration 
         Caption         =   "Julia set Iteration"
      End
   End
   Begin VB.Menu mnuZoom 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu mnuabt 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function InvertRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT) As Long

Private Const Pi As Double = 3.14159265358979
Private Const m_def_XMin As Double = -2.1
Private Const m_def_XMax As Double = 0.6
Private Const m_def_YMin As Double = -1.25
Private Const m_def_YMax As Double = 1.25

Dim XMin As Double, XMax As Double
Dim YMin As Double, YMax As Double
Dim MaxIter As Double, XRes As Double, YRes As Double, m_ColorJul As Long
Dim m_JulIter As Long
Dim RandLookUp(1000) As Long
Dim m_clrLevel() As Long, m_clr() As Long, m_numClr As Long
Dim ZoomRect As RECT, ZoomSelected As Boolean
Dim bErased As Boolean, bMouseDown As Boolean

Public Function getNumColors() As Long
    getNumColors = m_numClr
End Function

Public Function setNumColors(nc As Long)
    m_numClr = nc
End Function

Public Function getColors(ByRef c() As Long) As Long
    Dim i As Long
    ReDim c(m_numClr) As Long
    For i = 0 To m_numClr - 1
        c(i) = m_clr(i)
    Next i
    getColors = m_numClr
End Function

Public Function setColors(ByRef c() As Long) As Long
    Dim i As Long
    ReDim m_clr(m_numClr) As Long
    For i = 0 To m_numClr - 1
        m_clr(i) = c(i)
    Next i
    setColors = m_numClr
End Function


Private Sub Form_Load()
    Dim i As Long
    Randomize
    For i = 0 To 1000
        RandLookUp(i) = Rnd() * 100
    Next i
    
    mnuSquareCanvas.Checked = False

    XMin = m_def_XMin
    XMax = m_def_XMax
    YMin = m_def_YMin
    YMax = m_def_YMax
    MaxIter = 50
    m_ColorJul = RGB(50, 255, 100)
    m_JulIter = 500
    
    ReDim m_clr(4) As Long
    m_clr(0) = RGB(0, 0, 0)
    m_clr(1) = RGB(0, 255, 0)
    m_clr(2) = RGB(255, 255, 0)
    m_clr(3) = RGB(255, 0, 0)
    m_numClr = 4
    PrepareLevels m_clr, m_numClr, m_clrLevel
    
    ZoomSelected = False
    bErased = False
    bMouseDown = False
    
    Load frmJul
    frmJul.Show
End Sub

Private Sub Form_Resize()
    Dim wid As Long
    Dim xpos As Double, Ypos As Double
    
    If Me.WindowState = vbMinimized Then
        frmJul.Visible = False
        Exit Sub
    Else
        frmJul.Visible = True
    End If
    
    xpos = imgCur.Left / pM.ScaleWidth
    Ypos = imgCur.Top / pM.ScaleHeight
    
    
    If mnuSquareCanvas.Checked = True Then
        pM.Align = vbAlignNone
        wid = IIf(Me.ScaleWidth < Me.ScaleHeight, Me.ScaleWidth, Me.ScaleHeight)
        pM.Move (Me.ScaleWidth - wid) / 2, (Me.ScaleHeight - wid) / 2, wid, wid
    Else
        pM.Align = vbAlignLeft
        pM.Width = Me.ScaleWidth
    End If
    XRes = pM.ScaleWidth
    YRes = pM.ScaleHeight
    
    imgCur.Left = pM.ScaleWidth * xpos
    imgCur.Top = pM.ScaleHeight * Ypos
    
    Me.Show
    Me.Refresh
    pM.Refresh
    
    ZoomRect.Left = 0
    ZoomRect.Top = 0
    ZoomRect.Right = pM.ScaleWidth
    ZoomRect.Bottom = pM.ScaleHeight
    
    mnuGenMandel_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmJul
End Sub

Private Sub imgCur_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = imgCur.Left + (X / Screen.TwipsPerPixelX)
    Y = imgCur.Top + (Y / Screen.TwipsPerPixelY)
    GenJuliaPoint X, Y
End Sub

Private Sub imgCur_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgCur_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgCur_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgCur_MouseDown Button, Shift, X, Y
End Sub

Private Sub mnuabt_Click()
    frmHowTo.Show vbModal, Me
End Sub

Private Sub mnuGenMandel_Click()
    imgCur.Visible = False
    pM.Cls
    XMin = m_def_XMin
    XMax = m_def_XMax
    YMin = m_def_YMin
    YMax = m_def_YMax
    GenMandelBrot XMin, YMin, XMax, YMax
    bErased = True
    pM.Refresh
    GenJuliaPoint 0, 0
    imgCur.Visible = True
End Sub

Private Sub mnujulbgcol_Click()
    On Error GoTo errExt:
    CD.Flags = &HFF&
    CD.Color = frmJul.BackColor
    CD.ShowColor
    frmJul.BackColor = CD.Color
    mnuGenMandel_Click
errExt:
End Sub

Private Sub mnuJuliaColor_Click()
    On Error GoTo errExt:
    CD.Flags = &HFF&
    CD.Color = m_ColorJul
    CD.ShowColor
    m_ColorJul = CD.Color
    GenJuliaPoint 0, 0
errExt:
End Sub

Private Sub mnujuliasetiteration_Click()
    On Error GoTo errExt:
    Dim X As Long
    X = Val(InputBox("Enter how many points to approximate Julia set ?" & vbCrLf & "(default is 1000)", "Value of Julia Iteration", m_JulIter))
    m_JulIter = IIf(X >= 10, X, m_JulIter)
    GenJuliaPoint 0, 0
errExt:
End Sub

Private Sub mnuJulWnd_Click()
    If frmJul.Visible = False Then
        frmJul.Show
    End If
End Sub

Private Sub mnuMaxIter_Click()
    On Error GoTo errExt:
    MaxIter = Val(InputBox("Enter Maximum number of iterations for calculating mandelbrot set ?" & vbCrLf & "(default is 50)", "Value of Max Iteration", MaxIter))
    PrepareLevels m_clr, m_numClr, m_clrLevel
'    mnuGenMandel_Click
    mnuZoom_Click
errExt:
End Sub

Private Sub mnupal_Click()
    frmPalette.Show vbModal, Me
    PrepareLevels m_clr, m_numClr, m_clrLevel
    mnuZoom_Click
End Sub

Private Sub mnuSquareCanvas_Click()
    mnuSquareCanvas.Checked = Not mnuSquareCanvas.Checked
    Form_Resize
End Sub

Private Sub mnuXmax_Click()
    On Error GoTo errExt:
    XMax = Val(InputBox("Enter Xmax ?" & vbCrLf & "(default is 0.6)", "Value of XMAX", XMax))
    mnuGenMandel_Click
errExt:
End Sub

Private Sub mnuXmin_Click()
    On Error GoTo errExt:
    XMin = Val(InputBox("Enter Xmin ?" & vbCrLf & "(default is -2.1)", "Value of XMin", XMin))
    mnuGenMandel_Click
errExt:
End Sub

Private Sub MnuYMax_Click()
    On Error GoTo errExt:
    YMax = Val(InputBox("Enter YMax ?" & vbCrLf & "(default is 1.2)", "Value of YMax", YMax))
    mnuGenMandel_Click
errExt:
End Sub

Private Sub mnuYmin_Click()
    On Error GoTo errExt:
    YMin = Val(InputBox("Enter YMin ?" & vbCrLf & "(default is -1.2)", "Value of YMin", YMin))
    mnuGenMandel_Click
errExt:
End Sub

Private Sub mnuZoom_Click()
    Dim x1 As Double, y1 As Double, X2 As Double, Y2 As Double
    Dim dx As Double, dy As Double
    
    dx = (XMax - XMin) / (XRes - 1)
    dy = (YMax - YMin) / (YRes - 1)
    x1 = XMin + dx * ZoomRect.Left
    y1 = YMin + dy * ZoomRect.Top
    X2 = XMin + dx * ZoomRect.Right
    Y2 = YMin + dy * ZoomRect.Bottom
    
    GenMandelBrot x1, y1, X2, Y2
    GenJuliaPoint 0, 0
    bErased = True
    
    XMin = x1
    YMin = y1
    XMax = X2
    YMax = Y2
    pM.Refresh
End Sub

Private Sub pM_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Step As Long
    
    Step = IIf((Shift And vbCtrlMask) = vbCtrlMask, 5, 1)
    Select Case KeyCode
        Case vbKeyLeft
                imgCur.Left = IIf(imgCur.Left - Step < 0, 0, imgCur.Left - Step)
        Case vbKeyRight
                imgCur.Left = IIf(imgCur.Left + Step > (pM.ScaleWidth - imgCur.Width), (pM.ScaleWidth - imgCur.Width), imgCur.Left + Step)
        Case vbKeyUp
                imgCur.Top = IIf(imgCur.Top - Step < 0, 0, imgCur.Top - Step)
        Case vbKeyDown
                imgCur.Top = IIf(imgCur.Top + Step > (pM.ScaleHeight - imgCur.Height), (pM.ScaleHeight - imgCur.Height), imgCur.Top + Step)
    End Select
    GenJuliaPoint imgCur.Left + 15, imgCur.Top + 15
End Sub

Private Sub pM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errExt
    If bErased = False Then
        InvertRect pM.hdc, ZoomRect
        bErased = True
'        XMin = m_def_XMin
'        XMax = m_def_XMax
'        YMin = m_def_YMin
'        YMax = m_def_YMax
    End If
    bMouseDown = False
    If Button = vbLeftButton Then
        If (Shift And vbCtrlMask) = vbCtrlMask Then
            ZoomRect.Left = X
            ZoomRect.Top = Y
            ZoomRect.Right = X
            ZoomRect.Bottom = Y
            InvertRect pM.hdc, ZoomRect
            bMouseDown = True
        Else
            GenJuliaPoint X, Y
            imgCur.Move X - imgCur.Width / 2 + 1, Y - imgCur.Height / 2 + 1
        End If
    End If
    pM.Refresh
    DoEvents
errExt:
End Sub

Private Sub pM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim wid As Double, hei As Double
    If Button = vbLeftButton Then
        If (Shift And vbCtrlMask) = vbCtrlMask And bMouseDown = True Then
            InvertRect pM.hdc, ZoomRect
            ZoomRect.Right = X
            ZoomRect.Bottom = Y
            If (Shift And vbShiftMask) = vbShiftMask Then
                wid = Abs(ZoomRect.Right - ZoomRect.Left)
                hei = Abs(ZoomRect.Bottom - ZoomRect.Top)
                wid = IIf(wid < hei, wid, hei)
                ZoomRect.Right = ZoomRect.Left + wid
                ZoomRect.Bottom = ZoomRect.Top + wid
            End If
            InvertRect pM.hdc, ZoomRect
            bErased = False
        Else
            GenJuliaPoint X, Y
            imgCur.Move X - imgCur.Width / 2 + 1, Y - imgCur.Height / 2 + 1
        End If
    End If
    pM.Refresh
    DoEvents
End Sub

Private Sub pM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim wid As Double, hei As Double
    If Button = vbLeftButton Then
        If (Shift And vbCtrlMask) = vbCtrlMask And bMouseDown = True Then
            InvertRect pM.hdc, ZoomRect
            ZoomRect.Right = X
            ZoomRect.Bottom = Y
            If (Shift And vbShiftMask) = vbShiftMask Then
                wid = Abs(ZoomRect.Right - ZoomRect.Left)
                hei = Abs(ZoomRect.Bottom - ZoomRect.Top)
                wid = IIf(wid < hei, wid, hei)
                ZoomRect.Right = ZoomRect.Left + wid
                ZoomRect.Bottom = ZoomRect.Top + wid
            End If
            InvertRect pM.hdc, ZoomRect
            bErased = False
        Else
            GenJuliaPoint X, Y
            imgCur.Move X - imgCur.Width / 2 + 1, Y - imgCur.Height / 2 + 1
        End If
    End If
    pM.Refresh
    DoEvents
End Sub

Private Sub GenJuliaPoint(ByVal X As Double, ByVal Y As Double)
    On Error GoTo errExt
    Dim cx As Double, cy As Double, dx As Double, dy As Double
    dx = (XMax - XMin) / (XRes - 1)
    dy = (YMax - YMin) / (YRes - 1)
    cx = XMin + dx * X
    cy = YMin + dy * Y
    CalcJulSet cx, cy
errExt:
End Sub


'*****************************************************************************
'       MANDELBROT Routines
'*****************************************************************************

' This is the iterative routine to calculate the equation Z(n) = Z(n-1)^2 + c
' Where Zn , Z(n-1) and C are all complex numbers
' MIterate() iterates until maxIteration is reached
' or the function blows up beyond certain limit!
Private Function MIterate(ByVal cx As Double, ByVal cy As Double) As Long
    On Error GoTo errExt
    Dim iters As Long, X As Double, Y As Double, X2 As Double, Y2 As Double
    Dim temp As Double
    X = cx
    X2 = X * X
    Y = cy
    Y2 = Y * Y
    iters = 0
    While (iters < MaxIter) And (X2 + Y2 < 4)
        temp = cx + X2 - Y2
        Y = cy + 2 * X * Y
        Y2 = Y * Y
        X = temp
        X2 = X * X
        iters = iters + 1
    Wend
    MIterate = iters
errExt:
End Function

' Draws a Mandelbrotset using the above MIterate function
' Levels of colors designate the residual value of Iters after
' the function had blown up. Center color (0-default) indicates the
' region where equation holds stable upto Maxiter.
Private Function GenMandelBrot(ByVal xMn As Double, ByVal yMn As Double, ByVal xMx As Double, ByVal yMx As Double)
    On Error GoTo errExt
    Dim iX As Long, iY As Long, iters As Long
    Dim cx As Double, cy As Double, dx As Double, dy As Double
    
    Me.Caption = "Mandel Explorer v-1.2    [ Calculating  0% ]"
    
    dx = (xMx - xMn) / (XRes - 1)
    dy = (yMx - yMn) / (YRes - 1)
    
    For iY = 0 To YRes
        cy = yMn + iY * dy
        For iX = 0 To XRes
            cx = xMn + iX * dx
            iters = MIterate(cx, cy)
            If iters = MaxIter Then
                SetPixel pM.hdc, iX, iY, RGB(0, 0, 0)
'                SetPixel pM.hdc, iX, YRes - iY - 1, RGB(0, 0, 0)
            Else
                SetPixel pM.hdc, iX, iY, m_clrLevel(iters)
'                SetPixel pM.hdc, iX, YRes - iY - 1, Level(iters)
            End If
        Next iX
        Me.Caption = "Mandel Explorer v-1.2    [ Calculating  " & CInt(iY * 100 / YRes) & "% ]"
        pM.Refresh
        DoEvents
    Next iY
    pM.Refresh
    Me.Caption = "Mandel Explorer v-1.2    [ Done ]"
    Exit Function
errExt:
    pM.Cls
    pM.Print vbCrLf & " Error: "; Err.Number & vbCrLf & " Description : " & Err.Description
End Function

' This routine Calculates the Julia set for the Specified Mandelbrot value.
Private Function CalcJulSet(ByVal cx As Double, ByVal cy As Double)
    On Error GoTo errExt
    Dim Xp As Long, yP As Long
    Dim dx As Double, dy As Double
    Dim r As Double
    Dim theta As Double
    Dim X As Double, Y As Double
    Dim i As Long
    Dim rX As Long, rY As Long
    
    X = 0
    Y = 0
    rX = frmJul.ScaleWidth
    rY = frmJul.ScaleHeight
    frmJul.Cls
    For i = 0& To m_JulIter
        dx = X - cx
        dy = Y - cy
        If dx > 0 Then
            theta = Atn(dy / dx) * 0.5
        Else
            If dx < 0 Then
                theta = (Pi + Atn(dy / dx)) * 0.5
            Else
                theta = Pi * 0.25
            End If
        End If
        r = Sqr(Sqr(dx * dx + dy * dy))
        If vRandom() < 50 Then
            r = -r
        End If
        X = r * Cos(theta)
        Y = r * Sin(theta)
        Xp = (rX / 2) + CLng(X * (rX / 3.5))
        yP = (rY / 2) + CLng(Y * (rY / 3.5))
        SetPixel frmJul.hdc, Xp, yP, m_ColorJul
    Next i
    frmJul.Refresh
    Exit Function
errExt:
    frmJul.Cls
    frmJul.Print vbCrLf & " Error: "; Err.Number & vbCrLf & " Description : " & Err.Description
End Function

' Function to Blend between two colors
' These Colors can be customized using the Settings menu.
Private Function PrepareLevels(clr() As Long, n As Long, l() As Long, Optional ByVal nLevel As Long = 0)
    On Local Error Resume Next
    Dim r As Double, g As Double, b As Double
    Dim rs As Double, gs As Double, bs As Double
    Dim rr As Double, gg As Double, bb As Double
    Dim i As Long, j As Long, wid As Long
    
    If nLevel <= 0 Then nLevel = MaxIter
    
    Erase l
    ReDim l(nLevel) As Long
    wid = nLevel / (n - 1)
    
    For j = 0 To n - 2
        toRGB clr(j), r, g, b
        toRGB clr(j + 1), rr, gg, bb
        rs = (rr - r) / (wid + 1)
        gs = (gg - g) / (wid + 1)
        bs = (bb - b) / (wid + 1)
        For i = 0 To wid
            l(j * wid + i) = RGB(r, g, b)
            r = r + rs
            g = g + gs
            b = b + bs
        Next i
    Next j
errExt:
End Function

' Function to parse a Long color Windows native value into
'  corresponding Red, Green, and Blue values.
Private Function toRGB(ByVal c As Long, ByRef r As Double, ByRef g As Double, ByRef b As Double)
    On Error GoTo errExt
    r = CLng(c And &HFF&)
    g = CLng((c And &HFF00&) / &H100&)
    b = CLng((c And &HFF0000) / &H10000)
errExt:
End Function

Private Function vRandom() As Long
    On Error GoTo errExt
    Static Index As Long
    Index = IIf(Index < 0, Index = 0, IIf(Index >= 1000, 0, Index + 1))
    vRandom = RandLookUp(Index)
errExt:
End Function

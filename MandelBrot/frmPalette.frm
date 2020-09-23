VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPalette 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Palette Window"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6570
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   2295
      Top             =   3195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Pick a color"
   End
   Begin VB.CommandButton cmdAddcolor 
      Caption         =   "Add"
      Height          =   330
      Left            =   180
      TabIndex        =   6
      Top             =   135
      Width           =   735
   End
   Begin VB.ListBox lstClr 
      Height          =   2550
      IntegralHeight  =   0   'False
      Left            =   135
      TabIndex        =   5
      Top             =   585
      Width           =   3210
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   330
      Left            =   2340
      TabIndex        =   4
      Top             =   135
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Default"
      Height          =   330
      Left            =   1035
      TabIndex        =   3
      Top             =   135
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   5130
      TabIndex        =   2
      Top             =   3240
      Width           =   1275
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   420
      Left            =   3780
      TabIndex        =   1
      Top             =   3240
      Width           =   1275
   End
   Begin VB.PictureBox PP 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   3
      Height          =   3000
      Left            =   3465
      ScaleHeight     =   98
      ScaleMode       =   0  'User
      ScaleWidth      =   98
      TabIndex        =   0
      Top             =   135
      Width           =   3000
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim prevColor As Long
Dim pclr() As Long, clrcnt As Long
Dim bclrs() As Long

Private Sub cmdAddcolor_Click()
    On Error GoTo errExt
    Dim r As Double, g As Double, b As Double
    With CD
        .Flags = &HFF&
        .Color = prevColor
        .ShowColor
        prevColor = .Color
        toRGB .Color, r, g, b
        lstClr.AddItem " " & Hex(.Color) & " [ " & CByte(r) & " , " & CByte(g) & " , " & CByte(b) & " ] "
        clrcnt = clrcnt + 1
        ReDim Preserve pclr(clrcnt) As Long
        pclr(clrcnt - 1) = .Color
    End With
    DrawPreview
errExt:
End Sub

Private Sub DrawPreview()
    Dim i As Long
    PrepareLevels pclr, clrcnt, bclrs, PP.ScaleHeight
    For i = 0 To PP.ScaleHeight
        PP.Line (0, i)-(PP.ScaleWidth, i), bclrs(i)
    Next i
    PP.Refresh
    If clrcnt >= 2 Then cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
    setColors
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    loadcolors
End Sub

Private Sub Command2_Click()
    lstClr.Clear
    clrcnt = 0
    prevColor = 0
    PP.Cls
End Sub

Private Sub Form_Load()
    clrcnt = 0
    prevColor = 0
    ReDim bclrs(PP.ScaleHeight) As Long
    loadcolors
End Sub

Private Sub loadcolors()
    Dim i As Long
    Dim r As Double, g As Double, b As Double
    Command2_Click
    clrcnt = frmMain.getColors(pclr)
    For i = 0 To clrcnt - 1
        toRGB pclr(i), r, g, b
        lstClr.AddItem " " & Hex(pclr(i)) & " [ " & CByte(r) & " , " & CByte(g) & " , " & CByte(b) & " ] "
    Next i
    DrawPreview
End Sub

Private Sub setColors()
    Dim i As Long
    frmMain.setNumColors clrcnt
    frmMain.setColors pclr
    DrawPreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    Erase pclr
    Erase bclrs
End Sub

' Function to Blend between two colors
' These Colors can be customized using the Settings menu.
Private Function PrepareLevels(clr() As Long, n As Long, l() As Long, Optional ByVal nLevel As Long = 0)
    On Local Error Resume Next
    Dim r As Double, g As Double, b As Double
    Dim rs As Double, gs As Double, bs As Double
    Dim rr As Double, gg As Double, bb As Double
    Dim i As Long, j As Long, wid As Long
    
    If nLevel <= 0 Then nLevel = 10
    
    If n < 2 Then Exit Function
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


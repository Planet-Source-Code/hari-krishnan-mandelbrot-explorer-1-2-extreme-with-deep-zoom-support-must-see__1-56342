VERSION 5.00
Begin VB.Form frmHowTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "How to Explore!"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   ControlBox      =   0   'False
   Icon            =   "frmHowTo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   420
      Left            =   8145
      TabIndex        =   1
      Top             =   4860
      Width           =   1275
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   3473
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2003
      Width           =   2625
   End
End
Attribute VB_Name = "frmHowTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txt.Move 0, 0, Me.ScaleWidth, cmdClose.Top - 5 * Screen.TwipsPerPixelY
    
    txt.Text = " Mandelbrot Set realtime Explorer v1.2 with DeepZooming support."
    txt.Text = txt.Text & vbCrLf & "   code by, Hari krishnan G. (aka eXeption)"
    txt.Text = txt.Text & vbCrLf & "             harietr@yahoo.com"
    txt.Text = txt.Text & vbCrLf
    txt.Text = txt.Text & vbCrLf & """ For improved speed Compile it into an EXE and run! """
    txt.Text = txt.Text & vbCrLf
    txt.Text = txt.Text & vbCrLf & " How To Explore "
    txt.Text = txt.Text & vbCrLf & "-------------------------"
    txt.Text = txt.Text & vbCrLf & "1) Click and drag to Explore the Julia set expansion for each point of the Mandelbrot set."
    txt.Text = txt.Text & vbCrLf & "2) Press and Hold ""CTRL"" then drag to make a ""selection""."
    txt.Text = txt.Text & vbCrLf & "3) Pressing down ""SHIFT"" along with ""CTRL"" while selecting will make the selection square."
    txt.Text = txt.Text & vbCrLf & "4) Now to zoom to the selected point cleck on the ""Zoom"" menu."
    txt.Text = txt.Text & vbCrLf & "5) You can zoom again and again by repeating these steps."
    txt.Text = txt.Text & vbCrLf & "6) Clicking on ""Generate Mandelbrot"" menu will reset the mandelbrot set."
    txt.Text = txt.Text & vbCrLf & "7) The ""Settings"" menu gives a comprehensive control over all the values and thus can generate almost any part of the set."
    txt.Text = txt.Text & vbCrLf & "8) Increase the value of ""Maximum Iterations"" if you are going to explore deeper. (""50-100"" - normally, and ""200-500-1000"" - recommended for computers with AMD64 or Pentium4)."
    txt.Text = txt.Text & vbCrLf & """ For improved speed Compile it into an EXE and run! """
    txt.Text = txt.Text & vbCrLf
    txt.Text = txt.Text & vbCrLf & " Please mail me, if you like this code or any improvements."
    txt.Text = txt.Text & vbCrLf & "---------------------------------------------------------------"
    txt.Text = txt.Text & vbCrLf & " Thanks for improvements suggested by following people :-"
    txt.Text = txt.Text & vbCrLf & "         - Roger Gilchrist <roja.gilkrist@gmail.com> ( for Previous version (1.1) )"
    txt.Text = txt.Text & vbCrLf & " MANY THANKS!!!!!!!!!!!"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Load frmMain
    frmMain.Show
End Sub

VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFISplash 
      Left            =   120
      Top             =   840
   End
   Begin VB.Timer tmrFOSplash 
      Left            =   5280
      Top             =   1200
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmSplash.frx":0000
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   5760
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   5760
      X2              =   5760
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright 2004"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Program description here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Program name here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This should provide some decent reference for
'creating a splash screen that fades in slowly
'waits a few seconds, and then fades out slowly.
'I give credit to some VB students that helped me out with
'the loops.



'API Declarations
'*****************************************

Private Declare Function GetWindowLong Lib "user32" Alias _
"GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


Private Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long


Private Declare Function SetLayeredWindowAttributes Lib "user32" _
(ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long

'Variables
'******************************************
'constants that work with API from above
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&
'basic variables
Dim i As Integer
Dim sdelay As String
Dim tdelay As Single


Private Sub Form_Load()
    Dim i As Integer 'counter
    
'Get attributes of Splash Screen
    SetWindowLong frmSplash.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes frmSplash.hwnd, 0, 0, LWA_ALPHA
'set FadeIn timer interval
    tmrFISplash.Interval = 200
'Set heading label caption
    lblHeading.Caption = "Your Program Name Here"
'set version caption
    lblVersion.Caption = "Version: " & " " & App.Major & "." & App.Minor
    
    
End Sub

Public Sub FadeIn()

'Set delay for timer to a very low number - determines load speed
'make smaller number to make form load faster
    sdelay = 0.0000001
'loop for fading in window
    For i = 0 To 255
'sets windows visibility attributes
        SetLayeredWindowAttributes frmSplash.hwnd, 0, i, LWA_ALPHA
'increase timer interval
        tdelay = Timer + Val(sdelay)
'let Windows do it's thing
        While tdelay > Timer: DoEvents: Wend
'increase i by 1
    Next i
'set fade out timer interval
    tmrFOSplash.Interval = 3000
    
End Sub

Public Sub FadeOut()

'Set delay for timer to a very low number - determines load speed
'make smaller number to make form load faster
    sdelay = 0.00001
'loop for fading out window
    For i = 255 To 0 Step -1
'sets window visibility attributes
        SetLayeredWindowAttributes frmSplash.hwnd, 0, i, LWA_ALPHA
'increase timer interval
        tdelay = Timer + Val(sdelay)
'let Windows do its thing
        While tdelay > Timer: DoEvents: Wend
'decrease i by 1
    Next i
'unload splash form from memory
    Unload Me
'frmlogin.show
'enter code to do events you want
    End
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me   'unload from memory
    'you would do this so users can skip the splash screen
    'by clicking it
    
    
End Sub

Private Sub tmrFISplash_Timer()
    Call FadeIn                     'calls Sub
    tmrFISplash.Enabled = False     'disables FadeIn timer
End Sub

Private Sub tmrFOSplash_Timer()
    Call FadeOut                    'calls sub
    tmrFOSplash.Enabled = False     'disables FadeOut timer
End Sub

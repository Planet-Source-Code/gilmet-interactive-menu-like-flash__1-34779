VERSION 5.00
Begin VB.Form frmTestMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interactive Menu Test"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   454
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2640
      TabIndex        =   1
      Text            =   "Selected: NONE"
      Top             =   180
      Width           =   3735
   End
   Begin VB.PictureBox picPaper 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   3555
      Picture         =   "frmTestMenu.frx":0000
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   144
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1965
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Item # 8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Index           =   7
      Left            =   330
      MouseIcon       =   "frmTestMenu.frx":45A0
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3510
      Width           =   1740
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Item # 7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Index           =   6
      Left            =   330
      MouseIcon       =   "frmTestMenu.frx":48AA
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3045
      Width           =   1740
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Item # 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Index           =   5
      Left            =   330
      MouseIcon       =   "frmTestMenu.frx":4BB4
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2580
      Width           =   1740
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Item # 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Index           =   4
      Left            =   330
      MouseIcon       =   "frmTestMenu.frx":4EBE
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2115
      Width           =   1740
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Item # 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Index           =   3
      Left            =   330
      MouseIcon       =   "frmTestMenu.frx":51C8
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1635
      Width           =   1740
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Item # 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Index           =   2
      Left            =   330
      MouseIcon       =   "frmTestMenu.frx":54D2
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1185
      Width           =   1740
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Item # 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Index           =   1
      Left            =   330
      MouseIcon       =   "frmTestMenu.frx":57DC
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   705
      Width           =   1740
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Item # 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Index           =   0
      Left            =   330
      MouseIcon       =   "frmTestMenu.frx":5AE6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   255
      Width           =   1740
   End
   Begin VB.Image imgMenuOn 
      Height          =   465
      Left            =   3510
      Picture         =   "frmTestMenu.frx":5DF0
      Top             =   1305
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgMenu 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   3990
      Left            =   45
      Picture         =   "frmTestMenu.frx":5F6F
      Top             =   30
      Width           =   2250
   End
End
Attribute VB_Name = "frmTestMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function sndPlaySound Lib "WINMM.DLL" _
    Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As Any, _
     ByVal uFlags As Long) As Long

Private Const SND_ASYNC = &H1     ' Play asynchronously
Private Const SND_NODEFAULT = &H2 ' Don't use default sound
Private Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10

Private intMenuOn As Integer

Private Sub Form_Load()
 intMenuOn = -1
 imgMenuOn.Left = imgMenu.Left
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If intMenuOn > -1 Then
       imgMenuOn.Visible = False
       Text1.Text = "Selected: NONE"
       intMenuOn = -1
    End If
End Sub

Private Sub Form_Paint()
 Dim I As Long, J As Long
 With picPaper
    For I = 0 To Me.ScaleWidth Step .Width
        For J = 0 To Me.ScaleHeight Step .Height
            Me.PaintPicture .Picture, I, J
        Next J
    Next I
 End With
End Sub

Private Sub lblMenu_Click(Index As Integer)
    Text1.Text = "Selected: " & Trim(Str(Index + 1))
    BeginPlaySound 102  'Button_Pressed.wav
    Select Case Index
        Case 0: 'Put your code here...
        Case 1: 'Put your code here...
        Case 2: 'Put your code here...
        Case 3: 'Put your code here...
        Case 4: 'Put your code here...
        Case 5: 'Put your code here...
        Case 6: 'Put your code here...
        Case 7: 'Put your code here...
    End Select
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If intMenuOn = Index Then Exit Sub
    Select Case Index
        Case 0: imgMenuOn.Top = imgMenu.Top + (9)
        Case 1: imgMenuOn.Top = imgMenu.Top + (40)
        Case 2: imgMenuOn.Top = imgMenu.Top + (71)
        Case 3: imgMenuOn.Top = imgMenu.Top + (102)
        Case 4: imgMenuOn.Top = imgMenu.Top + (133)
        Case 5: imgMenuOn.Top = imgMenu.Top + (164)
        Case 6: imgMenuOn.Top = imgMenu.Top + (195)
        Case 7: imgMenuOn.Top = imgMenu.Top + (226)
    End Select
    imgMenuOn.Visible = True
    BeginPlaySound 101  'Button_Over.wav
    intMenuOn = Index
End Sub

Private Sub BeginPlaySound(ByVal ResourceId As Integer)
    Dim SoundBuffer As String
    Dim Ret As Variant
    ' Important: The returned string is converted to Unicode
    SoundBuffer = StrConv(LoadResData(ResourceId, "BUTTON_SOUNDS"), vbUnicode)
    Ret = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_NOSTOP)
    ' Important: This function is neccessary for playing sound asynchronously
    DoEvents
End Sub

Private Sub EndPlaySound()
    Dim Ret As Variant
    Ret = sndPlaySound(0&, 0&)
End Sub

VERSION 5.00
Begin VB.Form frmMsgBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2745
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2745
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   BeginProperty Font 
      Name            =   "Century"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picExpression 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   0
      Left            =   1080
      Picture         =   "frmMsgBox.frx":0000
      ScaleHeight     =   570
      ScaleWidth      =   570
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Timer tmrStopAnimation 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -120
      Top             =   120
   End
   Begin VB.Timer tmrStartAnimation 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -120
      Top             =   120
   End
   Begin VB.PictureBox picExpression 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   1080
      Picture         =   "frmMsgBox.frx":117A
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picExpression 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   1
      Left            =   1080
      Picture         =   "frmMsgBox.frx":1F3C
      ScaleHeight     =   525
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2295
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resCurrent As Res
Public waitTime As Single
Dim resStartup As Res
Dim resFinish As Res

Private Function InitMsgBox()
Dim scrRes As Res
scrRes = GetRes()
resStartup.Y = scrRes.Y + Me.Height
resStartup.X = scrRes.X - Me.Width - 20
resFinish.Y = scrRes.Y - Me.Height
PlaceMe resStartup
tmrStartAnimation.Enabled = True
End Function

Private Function PlaceMe(scrPos As Res)
Me.Left = scrPos.X
Me.Top = scrPos.Y
End Function

Private Sub Form_Load()
InitMsgBox
End Sub


Private Sub tmrStartAnimation_Timer()
UpdatePos
If resCurrent.Y + 410 >= resFinish.Y Then
    Me.Top = Me.Top - 100
Else
    Wait waitTime
    tmrStartAnimation.Enabled = False
    tmrStopAnimation.Enabled = True
End If
End Sub
Public Function UpdatePos()
resCurrent.X = Me.Left
resCurrent.Y = Me.Top
End Function

Private Sub tmrStopAnimation_Timer()
UpdatePos
If resCurrent.Y <= resStartup.Y Then
    Me.Top = Me.Top + 100
Else
    tmrStopAnimation.Enabled = False
    Unload Me
End If
End Sub

VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Custom Message Box"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "2"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Show"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1320
      List            =   "Form1.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Message Exp:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Waiting Time:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Combo1.Text = "" Or Text2.Text = "" Then
    ShowMessage "Incomplete Form....." & vbCr & "Please chk it properly....", msgCritical, 2
Else
    ShowMessage Text1.Text, Combo1.ListIndex, Text2.Text
End If
End Sub


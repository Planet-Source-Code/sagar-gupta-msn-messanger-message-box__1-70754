Attribute VB_Name = "modMain"
Public Type Res
    X As Integer
    Y As Integer
End Type
Public Enum MsgType
    msgCritical = 0
    msgQuestion = 1
    msgExclamation = 2
End Enum

Public Function ShowMessage(msg As String, messageType As MsgType, sWaitTime As Single)
Load frmMsgBox
frmMsgBox.lblMsg.Caption = msg
frmMsgBox.picExpression(messageType).Visible = True
frmMsgBox.Show
frmMsgBox.waitTime = sWaitTime
End Function
Public Function Wait(sSecs As Single)
Dim currentTime As Single
currentTime = Timer
While (currentTime + sSecs) > Timer
    DoEvents
Wend
End Function
Public Function GetRes() As Res
GetRes.X = Screen.Width
GetRes.Y = Screen.Height
End Function



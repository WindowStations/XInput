VERSION 5.00
Begin VB.Form frmXInput 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmXInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents TMRPoll          As Timer
Attribute TMRPoll.VB_VarHelpID = -1
Private WithEvents TMRPollBatteries As Timer
Attribute TMRPollBatteries.VB_VarHelpID = -1
Private WithEvents xinputClass      As clsXInput
Attribute xinputClass.VB_VarHelpID = -1
Private isconnected                 As Boolean
Private output                      As String
Private Sub Form_Initialize()
    Set TMRPoll = New Timer
    Set TMRPollBatteries = New Timer
    Set xinputClass = New clsXInput
End Sub
'    Me.MinButton = False
'    Me.ControlBox = False
'    Me.StartUpPosition = 0
Private Sub Form_Load()
    Me.BorderStyle = 0
    Me.WindowState = vbMaximized
    TMRPoll.Interval = 1
    TMRPollBatteries.Interval = 10000
    TMRPoll.Enabled = True
    TMRPollBatteries.Enabled = True
    KeyForward = vbKeyW
    KeyBackward = vbKeyS
    KeyLeft = vbKeyA
    KeyRight = vbKeyD
    KeyJump = vbKeySpace
    KeyCrouch = vbKeyControl
    KeyFire = vbKeyLButton
    KeyUse = vbKeyE
    KeyScreenshot = vbKeyPrint
    xinputClass.Enable
End Sub
Private Sub Form_Unload(Cancel As Integer)
    TMRPoll.Enabled = False
    TMRPollBatteries.Enabled = False
    Set TMRPoll = Nothing
    Set TMRPollBatteries = Nothing
    Set xinputClass = Nothing
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub xinputClass_OnButtonADown()
    Keyb(KeyFire) = True
End Sub
Private Sub xinputClass_OnButtonAUp()
    Keyb(KeyFire) = False
End Sub
Private Sub xinputClass_OnButtonXDown()
    Keyb(KeyUse) = True
End Sub
Private Sub xinputClass_OnButtonXUp()
    Keyb(KeyUse) = False
End Sub
Private Sub xinputClass_OnButtonRSHDown()
    Keyb(KeyJump) = True
End Sub
Private Sub xinputClass_OnButtonRSHUp()
    Keyb(KeyJump) = False
End Sub
Private Sub xinputClass_OnButtonUpDown()
    Keyb(KeyForward) = True
End Sub
Private Sub xinputClass_OnButtonUpUp()
    Keyb(KeyForward) = False
End Sub
Private Sub xinputClass_OnButtonDownDown()
    Keyb(KeyBackward) = True
End Sub
Private Sub xinputClass_OnButtonDownUp()
    Keyb(KeyBackward) = False
End Sub
Private Sub xinputClass_OnButtonLeftDown()
    Keyb(KeyLeft) = True
End Sub
Private Sub xinputClass_OnButtonLeftUp()
    Keyb(KeyLeft) = False
End Sub
Private Sub xinputClass_OnButtonRightDown()
    Keyb(KeyRight) = True
End Sub
Private Sub xinputClass_OnButtonRightUp()
    Keyb(KeyRight) = False
End Sub
Private Sub xinputClass_OnLThumbChange(ByVal x As Double, ByVal y As Double)
    oldis.Gamepad.sThumbLX = x
    oldis.Gamepad.sThumbLY = y
End Sub
Private Sub xinputClass_OnRThumbChange(ByVal x As Double, ByVal y As Double)
    oldis.Gamepad.sThumbRX = x
    oldis.Gamepad.sThumbRY = y
End Sub
Private Sub xinputClass_OnRTriggerChange(ByVal z As Long)
    oldis.Gamepad.bRightTrigger = z
End Sub
Private Sub xinputClass_OnLTriggerChange(ByVal z As Long)
    oldis.Gamepad.bLeftTrigger = z
End Sub
Private Sub xinputClass_OnDeviceConnected()
    isconnected = True
End Sub
Private Sub xinputClass_OnDeviceDisconnected()
    isconnected = False
End Sub
Private Sub TMRPoll_Elapsed()
    'main game rendering logic here
    Dim inp As String
    inp = inp & "Buttons: " & oldis.Gamepad.wButtons & vbCrLf
    inp = inp & "Left Thumb X: " & oldis.Gamepad.sThumbLX & vbCrLf
    inp = inp & "Left Thumb Y: " & oldis.Gamepad.sThumbLY & vbCrLf
    inp = inp & "Right Thumb X: " & oldis.Gamepad.sThumbRX & vbCrLf
    inp = inp & "Right Thumb Y: " & oldis.Gamepad.sThumbRY & vbCrLf
    inp = inp & "Left Trigger: " & oldis.Gamepad.bLeftTrigger & vbCrLf
    inp = inp & "Right Trigger: " & oldis.Gamepad.bRightTrigger & vbCrLf
    If Keyb(KeyFire) = True Then
        inp = inp & "Firing"
    Else
        inp = inp & ""
    End If
    If inp <> output Then
        output = inp
        Me.Cls
        Me.Print inp
    End If
End Sub
Private Sub TMRPollBatteries_Elapsed()
   Dim inp As String
   inp = inp & "Battery: " & xinputClass.BatteryLevel(0) & vbCrLf

End Sub

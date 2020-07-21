Attribute VB_Name = "modXInput"
Option Explicit
Public Const ERROR_DEVICE_NOT_CONNECTED As Long = 1167
Public Const ERROR_SUCCESS              As Long = 0
Public Const ERROR_EMPTY                As Long = 4306
Private Type XINPUT_GAMEPAD
    wButtons As Integer
    bLeftTrigger As Byte
    bRightTrigger As Byte
    sThumbLX As Integer
    sThumbLY As Integer
    sThumbRX As Integer
    sThumbRY As Integer
End Type
Public Type XINPUT_STATE
    PacketNumber As Long
    Gamepad As XINPUT_GAMEPAD
End Type
Public oldis           As XINPUT_STATE
Public Keyb(-3 To 255) As Boolean
Public KeyForward      As Long
Public KeyBackward     As Long
Public KeyLeft         As Long
Public KeyRight        As Long
Public KeyJump         As Long
Public KeyCrouch       As Long
Public KeyFire         As Long
Public KeyUse          As Long
Public KeyScreenshot   As Long

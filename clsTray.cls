VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'[EVENTS]
Event RButton()
Event LButton()
Event AddedToTray()
Event RemovedFromTray()

'[TYPES]
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'[CONTANTS]
Private Const NIM_ADD = &H0  'Add to Tray
Private Const NIM_MODIFY = &H1 'Modify Details
Private Const NIM_DELETE = &H2 'Remove From Tray
Private Const NIF_MESSAGE = &H1 'Message
Private Const NIF_ICON = &H2 'Icon
Private Const NIF_TIP = &H4 'TooTipText
Private Const WM_MOUSEMOVE = &H200 'On Mousemove
Private Const WM_LBUTTONDBLCLK = &H203 'Left Double Click
Private Const WM_RBUTTONDOWN = &H204 'Right Button Down
Private Const WM_RBUTTONUP = &H205 'Right Button Up
Private Const WM_RBUTTONDBLCLK = &H206 'Right Double Click

'[API]
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
 
'[OBJECT VARIABLES]
Private WithEvents picTray As PictureBox
Attribute picTray.VB_VarHelpID = -1
 
'[PRIVATE VARIABLES]
Private TrayIco As NOTIFYICONDATA
Private InTray As Boolean

 

Sub AddToTray(picBoxIcon As PictureBox, sTrayTip As String)
On Error GoTo ERR_HANDLER:
'-------------------------------------------------
' ltray icon will be the forms icon unless
' specified otherwise
'-------------------------------------------------
'VARIABLES:
 
'CODE:
   Set picTray = picBoxIcon
   InTray = True

   'initialize tray info
   With TrayIco
            .cbSize = Len(TrayIco)
            .hwnd = picBoxIcon.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = picTray.Picture
            .szTip = sTrayTip & vbNullChar 'the tray tooltip
   End With
   'add this to tray
   Shell_NotifyIcon NIM_ADD, TrayIco
 
'END CODE:
Exit Sub
ERR_HANDLER:
  Debug.Print Err.Description
End Sub

Sub RemoveFromTray()
On Error GoTo ERR_HANDLER:
'-------------------------------------------------
' remove the icon from tray..either because showing
' form or ending app
'-------------------------------------------------
'VARIABLES:

'CODE:
  'remove the tray icon
   Shell_NotifyIcon NIM_DELETE, TrayIco
   InTray = False
'END CODE:
Exit Sub
ERR_HANDLER:
  Debug.Print Err.Description
End Sub

Sub ModifyTray(Optional sNewToolTip As String, Optional lNewIcon As Long)
On Error GoTo ERR_HANDLER:
'-------------------------------------------------
' change either the tooltip of the icon associated
' with tray icon
'-------------------------------------------------
'VARIABLES:

'CODE:
 With TrayIco
   If lNewIcon <> 0 Then .hIcon = lNewIcon
   If Len(Trim(sNewToolTip)) > 0 Then .szTip = sNewToolTip & vbNullChar
 End With
 'update tray icon with new values
 Shell_NotifyIcon NIM_MODIFY, TrayIco
'END CODE:
Exit Sub
ERR_HANDLER:
  Debug.Print Err.Description
End Sub
 
Private Sub PicTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_HANDLER:
'-------------------------------------------------
'this sub is called when the mouse moves over the tray
'icon because trays callback msg is wm_mousemove
'-=[thanks to LCSBSSRHXXX for the much shortened tray code]=-
'-------------------------------------------------
'VARIABLES:

'CODE:
    Select Case InTray
        Case True
            If Button = 1 Then 'left click
                RaiseEvent LButton
            ElseIf Button = 2 Then 'right click
                RaiseEvent RButton
            End If
        Case False
            Exit Sub
    End Select
'END CODE:
Exit Sub
ERR_HANDLER:
  Debug.Print Err.Description
End Sub
 

Private Sub Class_Terminate()
'-----------------------------
'destroy local object reference
'-----------------------------
 On Error Resume Next
 Set picTray = Nothing
End Sub

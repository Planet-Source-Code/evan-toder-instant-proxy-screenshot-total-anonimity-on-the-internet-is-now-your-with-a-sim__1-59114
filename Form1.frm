VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picTray 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   45
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   0
      Width           =   600
   End
   Begin VB.Menu mnuTop 
      Caption         =   "top"
      Visible         =   0   'False
      Begin VB.Menu mnuSubProxy 
         Caption         =   "SERVER IP ADDRESS><PORT><PROXY TYPE><LOCATION"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuIeDirect 
         Caption         =   "IE &direct connection"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubWhatIsMyIp 
         Caption         =   "www.whatismyip.com"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "close && &exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim WithEvents cTray   As clsTray
Dim WithEvents cNav    As cNavigateTo
Dim creg               As cRegistry
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub UpdateProxyList()
  cNav.CreateProxyDoc
End Sub

Private Sub cNav_OperationTimedOut()
  MsgBox "Attempt to download proxy list timed out." & vbCrLf & _
         "Check to make sure you have a valid connection" & vbCrLf & _
         "to the internet." & vbCrLf & _
         "This program will now unload."
  Unload Me
End Sub
 

Private Sub cNav_ProxyInfo(proxyAddress As String, proxyPort As String, proxyType As String, proxyLocation As String)
   Dim upper As Integer
   upper = mnuSubProxy.UBound
   Load mnuSubProxy(upper + 1)
   mnuSubProxy(upper + 1).Caption = proxyAddress & "><" & proxyPort & "><" & proxyType & "><" & proxyLocation
   mnuSubProxy(upper + 1).Enabled = True
   mnuSubProxy(upper + 1).Visible = True
End Sub

Private Sub cTray_RButton()
  PopupMenu mnuTop
End Sub

Private Sub Form_Load()
  Set cNav = New cNavigateTo
  Set cTray = New clsTray
  Set creg = New cRegistry
  cTray.AddToTray picTray, "Instant Proxy"
  Call UpdateProxyList
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '  set the reg back to default direct surfing
  Call RegSectionProxyEnable(False)
  Set cNav = Nothing
  cTray.RemoveFromTray
  Set cTray = Nothing
  Set creg = Nothing
End Sub
 
Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub RegSectionProxyEnable(bEnable As Boolean)
 '  access this part of registry and set dword value
 '  to 1 to proxyenable and 0 to not
  creg.ClassKey = HKEY_CURRENT_USER
  creg.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
  creg.ValueKey = "ProxyEnable"
  creg.ValueType = REG_DWORD
  
  If bEnable Then
     creg.Value = 1
  Else
     creg.Value = 0
  End If
End Sub

Private Sub RegSectionProxyServer(server_colon_port As String)
  '  specify a proxy server and port in format "server:port"
  '  (without quotes)
  creg.ClassKey = HKEY_CURRENT_USER
  creg.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
  creg.ValueKey = "ProxyServer"
  creg.ValueType = REG_EXPAND_SZ
  creg.Value = server_colon_port
End Sub

Private Sub mnuIeDirect_Click()
  ' sets ie connect back to direct(no proxy) in the registry
  Call RegSectionProxyEnable(False)
  cTray.ModifyTray "I.E. surfing direct (no proxy server)"
End Sub

Private Sub mnuSubProxy_Click(Index As Integer)
  ' set ie connect to a proxy server
  Dim port As String, ip As String, pxytype As String
  Dim pxylocal As String, ipport As String
  ip = Split(mnuSubProxy(Index).Caption, "><")(0)
  port = Split(mnuSubProxy(Index).Caption, "><")(1)
  pxytype = Split(mnuSubProxy(Index).Caption, "><")(2)
  pxylocal = Split(mnuSubProxy(Index).Caption, "><")(3)
  ipport = ip & ":" & port
  Call RegSectionProxyEnable(True)
  Call RegSectionProxyServer(ipport)
  '  show current proxy in tray tooltip
  cTray.ModifyTray "current proxy server  " & ipport & ":" & pxytype & ":" & pxylocal
End Sub

Private Sub mnuSubWhatIsMyIp_Click()
  ShellExecute hwnd, "open", "http://www.whatismyip.com", vbNullString, vbNullString, 1
End Sub

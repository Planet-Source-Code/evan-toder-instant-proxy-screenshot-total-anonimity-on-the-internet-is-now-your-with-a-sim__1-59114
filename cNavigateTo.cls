VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNavigateTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long


Dim odoc               As HTMLDocument
Dim WithEvents newDoc  As HTMLDocument
Attribute newDoc.VB_VarHelpID = -1
 
 
Event OperationTimedOut()
Event ProxyInfo(proxyAddress$, proxyPort$, proxyType$, proxyLocation$)
            

Sub CreateProxyDoc(Optional secondsTimeOut As Long = 6)
 
 ' "download" the webpage invisibly
 '  we dont need a visible interface and without
 '  any interface updating the page downloads much faster
 Set newDoc = New HTMLDocument
 Set odoc = newDoc.createDocumentFromUrl("http://www.proxy4free.com/page1.html", vbNullString)
 
 Dim starttime  As Long
 starttime = GetTickCount
 
 ' we wait for the "complete" event to fire
 ' we want a timeout value so we dont end
 ' up in an endless loop in case the site
 ' is down or there is internet problem
 Do
   DoEvents
   If odoc.readyState = "complete" Then
        Exit Do
   Else
      If (GetTickCount - starttime) > _
      (secondsTimeOut * 1000) Then
          RaiseEvent OperationTimedOut
      End If
   End If
 Loop
 
 
 Dim gen  As HTMLGenericElement
 Dim gen2 As HTMLGenericElement
 Dim gen3 As HTMLGenericElement
 Dim gen4 As HTMLGenericElement
 Dim i    As Integer
 Dim ret  As Integer
 
 ' "TD" is a web element whos type is table cel
 For i = 0 To odoc.getElementsByTagName("TD").length - 1
     Set gen = odoc.getElementsByTagName("TD")(i)
     '  means we have found an ip address ( xxx.xxx.xxx.xxx)
     If UBound(Split(gen.innerText, ".")) = 3 Then
       Set gen2 = odoc.getElementsByTagName("TD")(i + 1)
       Set gen3 = odoc.getElementsByTagName("TD")(i + 2)
       Set gen4 = odoc.getElementsByTagName("TD")(i + 3)
       RaiseEvent ProxyInfo(gen.innerText, gen2.innerText, gen3.innerText, gen4.innerText)
       ret = (ret + 1)
       ' 25 potential proxy servers should be enouph
       If ret > 25 Then Exit For
     End If
 Next
 
 ' clean up code
 Set gen = Nothing
 Set gen2 = Nothing
 Set gen3 = Nothing
 Set odoc = Nothing
 Set newDoc = Nothing
 
End Sub


Attribute VB_Name = "modDownload"
Option Explicit

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3

Private Const scUserAgent = "VBTagEdit"
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
(ByVal hOpen As Long, ByVal sUrl As String, ByVal sHeaders As String, _
ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long
   
Private Declare Function InternetGetConnectedState _
              Lib "wininet.dll" (ByRef lpdwFlags As Long, _
              ByVal dwReserved As Long) As Long
   
Private Const ERROR_SUCCESS As Long = 0

Public Function GetHTMLFromURL(sUrl As String) As String
Dim S                  As String
Dim hOpen              As Long
Dim hOpenUrl           As Long
Dim bDoLoop            As Boolean
Dim bRet               As Boolean
Dim sReadBuffer        As String * 2048
Dim lNumberOfBytesRead As Long

hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

bDoLoop = True

While bDoLoop
    sReadBuffer = vbNullString
    bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
    S = S & Left$(sReadBuffer, lNumberOfBytesRead)
    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
Wend

If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
If hOpen <> 0 Then InternetCloseHandle (hOpen)

GetHTMLFromURL = S

End Function

Public Function DownloadFile(ssourceUrl As String, _
                             sLocalFile As String) As Boolean

   Dim lngRetVal As Long
   
  'if the API returns ERROR_SUCCESS (0),
  'return True from the function
   DownloadFile = URLDownloadToFile(0&, _
                                    ssourceUrl, _
                                    sLocalFile, _
                                    0&, _
                                    0&) = ERROR_SUCCESS
   
End Function

Public Function IsConnected() As Boolean

    On Error GoTo err
    IsConnected = InternetGetConnectedState(0&, 0&)

Exit Function

err:
    IsConnected = True

End Function








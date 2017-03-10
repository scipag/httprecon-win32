Attribute VB_Name = "modHttpConnectivity"
Option Explicit

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Integer
Private Const HTTP_QUERY_STATUS_CODE = 19
Private Const INTERNET_SERVICE_HTTP = 3
Private Const scUserAgent = "http sample"
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000

Public Function SendHttpRequest(ByRef URL As String) As String
     Dim sBuffer         As String * 1024
     Dim lBufferLength   As Long
     Dim hInternetSession As Long
     Dim hInternetConnect As Long
     Dim hHttpOpenRequest As Long
     
     lBufferLength = 1024
     
     'Remove Http if needed
     If LCase(Left$(URL, 7)) = "http://" Then
      URL = Right$(URL, Len(URL) - 7)
     End If
     
     'Open the Internetconnection
     hInternetSession = InternetOpen(application_name, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    
     If CBool(hInternetSession) = False Then
      SendHttpRequest = 0
      Exit Function
     End If
    
     'Connect and get the Status
     hInternetConnect = InternetConnect(hInternetSession, URL, 80, "", "", INTERNET_SERVICE_HTTP, 0, 0)
     hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", "", "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION, 0)
     HttpSendRequest hHttpOpenRequest, vbNullString, 0, vbNullString, 0
     HttpQueryInfo hHttpOpenRequest, HTTP_QUERY_STATUS_CODE, ByVal sBuffer, lBufferLength, 0
    SendHttpRequest = HttpQueryInfo(hHttpOpenRequest, &H80000000, sBuffer, lBufferLength, 0)
'     SendHttpRequest = Val(Left$(sBuffer, lBufferLength))
    
     'Close connections
     InternetCloseHandle (hHttpOpenRequest)
     InternetCloseHandle (hInternetSession)
     InternetCloseHandle (hInternetConnect)
End Function

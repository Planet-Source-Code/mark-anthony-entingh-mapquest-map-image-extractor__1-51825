Attribute VB_Name = "basGlobal"
Public Const scUserAgent = "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT; DigExt)"

Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3

Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_FLAG_EXISITING_CONNECT = &H20000000

Public Const INTERNET_FILE_MOVE_BEGIN = 0
Public Const INTERNET_FILE_MOVE_CURRENT = 1
Public Const INTERNET_FILE_MOVE_END = 3

Public Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Public Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal lpszServerName As String, ByVal nProxyPort As Integer, ByVal lpszUserName As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetSetFilePointer Lib "wininet.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal pReserved As Long, ByVal dwMoveMethod As Long, ByVal dwContext As Long) As Long

Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Long
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Long

Public Sub CrackURL(strURL As String, strFilePath As String)
  Dim intPos As Integer

  If Left$(strURL, 7) = "http://" Then
    strURL = Right$(strURL, Len(strURL) - 7)
  End If

  intPos = InStr(1, strURL, "/")

  If intPos > 0 Then
    strFilePath = Right$(strURL, Len(strURL) - intPos + 1)
    strURL = Left$(strURL, intPos - 1)
  End If
End Sub

Public Function TrimString(s As String) As String
  TrimString = Left$(s, InStr(1, s, Chr$(0)) - 1)
End Function

Public Function GetFilename(txt)
Dim i, X
For X = Len(txt) To 1 Step -1
If Mid(txt, X, 1) = "\" Or Mid(txt, X, 1) = "/" Then Exit For
Next X
If X < Len(txt) Then
GetFilename = Mid(txt, X + 1)
End If
End Function
Public Function GetDirPath(txt)
Dim i, X
For X = Len(txt) To 1 Step -1
If Mid(txt, X, 1) = "\" Or Mid(txt, X, 1) = "/" Then Exit For
Next X
If X < Len(txt) Then
GetDirPath = Mid(txt, 1, X)
End If
End Function

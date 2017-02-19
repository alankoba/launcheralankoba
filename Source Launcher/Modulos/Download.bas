Attribute VB_Name = "Download_INET"
Global lKnownSize As Long
Global gets As Long
' WinInet Declarations
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal dwNumberOfBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternet As Long, ByVal lpszServerName As String, ByVal nServerPort As Long, ByVal lpszUserName As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
' WinInet Constants
Private Const INTERNET_INVALID_PORT_NUMBER = 0
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_SERVICE_GOPHER = 2
Private Const INTERNET_SERVICE_HTTP = 3
Private Const MAX_PATH = 260
Private Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FileTime
    ftLastAccessTime As FileTime
    ftLastWriteTime As FileTime
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Dim bDownloading As Boolean
Dim bInternetAvailable As Boolean
Const m_UserAgent As String = "Downloader Demo"
Dim bCancel As Boolean
Const lBufferSize As Long = 512
Dim lDoneBytes As Long    ' Used for Speed Test
Dim lOldBytes As Long    ' Used For Speed Test

' Constants for Labels
Const iFileSize As Integer = 0
Const iAverage As Integer = 1
Const iTime As Integer = 2
Public Function DownloadAFile(sFileName As String, sURLfrom As String, bShowStatus As Boolean, bProgressBar As Boolean) As Boolean
    Dim hInternetSession As Long
    Dim hInternetOpen As Long
    Dim hInternetConnection As Long
    Dim lNumberOfBytes As Long
    Dim lTotalBytesRead As Long
    Dim bDoLoop As Boolean
    Dim strURL As String
    Dim strFTPSvr As String
    Dim strFTPDir As String
    Dim aryFileData As WIN32_FIND_DATA
    Dim strArray As String
    Dim iFileHandle As Long


    DownloadAFile = False
    bInternetAvailable = False
    bDownloading = True
    ' Check Other Required Parameters

    If Len(sURLfrom) Or Len(sFileName) Then



        If Left$(LCase(sURLfrom), 7) = "http://" Or Left$(LCase(sURLfrom), 6) = "ftp://" Then

            '            ' Open Connection to Internet using Proxy Server Settings from Control panel
            '            If bInternetAvailable = False Then
            '               ' First attempt
            '             '  Debug.Print "Checking for Internet Connection"
            '            End If

            hInternetSession = InternetOpen(m_UserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
            If bCancel = True Then
                GoTo Exit_Jump
            End If


            If hInternetSession Then

                ' ProgressStatus "Checking for Internet Connection - Success"

                bInternetAvailable = True



                ' If Previous File Exists Locally Delete It

                '   If Fexist(sFileName) Then
                '     Kill sFileName
                '      End If

                ' Set URL to FileName

                strURL = sURLfrom

                ' Attempt to open URL

                lKnownSize = gets


                hInternetOpen = InternetOpenUrl(hInternetSession, strURL, vbNullString, 0, INTERNET_FLAG_EXISTING_CONNECT + INTERNET_FLAG_RELOAD, 0)
                If bCancel = True Then
                    GoTo Exit_Jump
                End If
                ' Check File Opened

                If hInternetOpen Then
                    ' ProgressStatus "Attempting to contact Update Server - Success"


                    ' Open Local Version of File

                    iFileHandle = FreeFile
                    Open sFileName For Binary As #iFileHandle
                    bDoLoop = True


                    Do While bDoLoop
                        DoEvents

                        If bCancel = True Then
                            GoTo Exit_Jump
                        End If
                        ' strArray = ""
                        ' Start Timer for speed test

                        strArray = String(lBufferSize, vbNull)
                        bDoLoop = InternetReadFile(hInternetOpen, strArray, Len(strArray), lNumberOfBytes)

                        lTotalBytesRead = lTotalBytesRead + lNumberOfBytes
                        If lNumberOfBytes < lBufferSize Then
                            ' we end of file
                            lKnownSize = lTotalBytesRead
                        End If
                        ' Update Status with Received Byte Count

                        If bProgressBar Then
                            Call ProgressBar(lKnownSize, lTotalBytesRead, sFileName)
                        End If
                        If CBool(lNumberOfBytes) Then
                            Put #iFileHandle, , Left$(strArray, lNumberOfBytes)
                        Else
                            bDoLoop = False
                        End If
                    Loop
                    Close #iFileHandle
                    ' Close Local File

                    If lTotalBytesRead <> lKnownSize And lKnownSize > 0 Then
                        DownloadAFile = False
                    Else
                        DownloadAFile = True
                    End If
                    ' Close Connection to URL

                    InternetCloseHandle hInternetOpen
                End If

                ' Close Connection to Internet

                InternetCloseHandle hInternetSession

            End If
        End If
    End If
    bDownloading = False

    Exit Function

Exit_Jump:
    InternetCloseHandle hInternetOpen
    InternetCloseHandle hInternetSession
    Close #iFileHandle
    bDownloading = False
    DownloadAFile = False
    KillFile sFileName

End Function
Public Function KillFile(strFile As String) As Boolean
    If Fexist(strFile) Then
        On Error Resume Next
        Kill strFile
    End If
End Function

Public Function Fexist(strFile As String) As Boolean
    Dim strDummy As String
    On Error Resume Next
    strDummy = Dir$(strFile, vbNormal)
    Fexist = Not (strDummy = "")
End Function

Public Function ProgressBar(lSize As Long, lDone As Long, sFiles As String)
    Dim iSendPercent As Integer
    Dim X As Integer
    Dim sTemp As String
    If lDone = 0 Then
        Exit Function
    End If
    lDoneBytes = lDone
    If lSize = 0 Then: lSize = 5000
    update.bar.Value = (lDone / lSize) * 100
    sTemp = SetBytes(lDone) + "/" + SetBytes(lSize)
    If update.tamanho.Caption <> sTemp Then
        update.tamanho.Caption = sTemp
    End If
    DoEvents
End Function

Private Function CheckError() As String
    Dim lErrorNumber As Long

    lErrorNumber = Err.LastDllError
    If lErrorNumber > 12000 Then
        CheckError = ""
        CheckError = TranslateErrorCode(lErrorNumber)
    Else
        CheckError = Err.Description
    End If
End Function
Private Function TranslateErrorCode(ByVal lErrorCode As Long) As String
    Select Case lErrorCode
    Case 12000: TranslateErrorCode = ""
    Case 12001: TranslateErrorCode = "No more handles"
    Case 12002: TranslateErrorCode = "Timeout"
    Case 12003: TranslateErrorCode = "Extended error"
    Case 12004: TranslateErrorCode = "Internal error"
    Case 12005: TranslateErrorCode = "URL invalid"
    Case 12006: TranslateErrorCode = "URL unrecognised"
    Case 12007: TranslateErrorCode = "Server not resolved"
    Case 12008: TranslateErrorCode = "Protocol not located"
    Case 12009: TranslateErrorCode = "Invalid Option"
    Case 12010: TranslateErrorCode = "Bad Option Length"
    Case 12011: TranslateErrorCode = "Option Not Settable"
    Case 12012: TranslateErrorCode = "Internet support shutdown"
    Case 12013: TranslateErrorCode = "Invalid username"
    Case 12014: TranslateErrorCode = "Invalid password"
    Case 12015: TranslateErrorCode = "Login failure"
    Case 12016: TranslateErrorCode = "Invalid operation"
    Case 12017: TranslateErrorCode = "Operation cancelled"
    Case 12018: TranslateErrorCode = "Invalid Internet Handle"
    Case 12019: TranslateErrorCode = "Invalid Handle state"
    Case 12020: TranslateErrorCode = "Proxy request invalid"
    Case 12021: TranslateErrorCode = "Registry Value not located"
    Case 12022: TranslateErrorCode = "Registry Value invalid type"
    Case 12023: TranslateErrorCode = "Network access unavailable"
    Case 12024: TranslateErrorCode = "Zero context supplied"
    Case 12025: TranslateErrorCode = "Callback option not set"
    Case 12026: TranslateErrorCode = "Request Pending"
    Case 12027: TranslateErrorCode = "Request Format Invalid"
    Case 12028: TranslateErrorCode = "Item Not Found"
    Case 12029: TranslateErrorCode = "Connection Failed"
    Case 12030: TranslateErrorCode = "Connection Terminated"
    Case 12031: TranslateErrorCode = "Connection Reset"
    Case 12032: TranslateErrorCode = "Force Retry"
    Case 12033: TranslateErrorCode = "Invlaid Proxy Request"
    Case 12034: TranslateErrorCode = "Need UI"

    Case 12036: TranslateErrorCode = "Handle Already Exists"
    Case 12037: TranslateErrorCode = "Certificate Date Invalid"
    Case 12038: TranslateErrorCode = "Certificate Invalid"
    Case 12039: TranslateErrorCode = "HTTP to HTTPS redirection"
    Case 12040: TranslateErrorCode = "HTTPS to HTTP redirection"
    Case 12041: TranslateErrorCode = "Mixed Security"
    Case 12042: TranslateErrorCode = "Post Not Secure"
    Case 12043: TranslateErrorCode = "Post Not Secure"
    Case 12044: TranslateErrorCode = "Client Certificate Needed"
    Case 12045: TranslateErrorCode = "Invalid Client Authority"
    Case 12046: TranslateErrorCode = "Client Authority not Setup"
    Case 12047: TranslateErrorCode = "Async Thread Failure"
    Case 12048: TranslateErrorCode = "Redirect Scheme Changed"
    Case 12049: TranslateErrorCode = "Dialog Pending"
    Case 12050: TranslateErrorCode = "Retry Dialog"

    Case 12052: TranslateErrorCode = "HTTP HTTP Redirection"
    Case 12053: TranslateErrorCode = "Insert CD-Rom"

    Case Else: TranslateErrorCode = "Undefined Error"
    End Select
End Function


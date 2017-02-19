VERSION 5.00
Begin VB.UserControl ctlfile 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ctlfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum IFile_ConnectType
  Preconfig = 0
  proxy = 1
  Direct = 2
End Enum
  
Const m_def_LocalFile = ""
Const m_def_Url = ""
Const m_def_Port = 80
Const m_def_SiteUser = ""
Const m_def_SitePassword = ""
Const m_def_ProxyServer = ""
Const m_def_ProxyPort = 8080
Const m_def_ConnectType = Preconfig

Dim m_FilePosition As Long
Dim m_CancelDownload As Boolean
Dim m_LocalFile As Variant
Dim m_Url As String
Dim m_Port As Integer
Dim m_SiteUser As String
Dim m_SitePassword As String
Dim m_ProxyPort As Integer
Dim m_ProxyServer As String
Dim m_LastModified As String
Dim m_FileSize As Long
Dim m_FileExists As Boolean
Dim m_ConnectType As IFile_ConnectType

Event DownloadComplete()
Event DownloadProgress(lBytesRead As Long)
Event DownloadError(sErrorDescription As String)
Event DownloadCancelled(lPosition As Long)

Private plInternetSession As Long
Private pfInformationRequested As Boolean
Private plFilePosition As Long
Private pfAllowClose As Boolean

Public Property Get FileExists() As Boolean
  FileExists = m_FileExists
End Property


Public Property Get LastModified() As String
  LastModified = m_LastModified
End Property


Public Property Get FileSize() As Long
  FileSize = m_FileSize
End Property


Public Property Get LocalFile() As Variant
  LocalFile = m_LocalFile
End Property

Public Property Let LocalFile(ByVal New_LocalFile As Variant)
  m_LocalFile = New_LocalFile
  PropertyChanged "LocalFile"
End Property


Public Property Get ConnectType() As IFile_ConnectType
  ConnectType = m_ConnectType
End Property

Public Property Let ConnectType(ByVal New_ConnectType As IFile_ConnectType)
  m_ConnectType = New_ConnectType
  PropertyChanged "ConnectType"
End Property


Public Property Get Url() As String
  Url = m_Url
End Property

Public Property Let Url(ByVal New_Url As String)
  m_Url = New_Url
  PropertyChanged "Url"
End Property


Public Property Get Port() As Integer
  Port = m_Port
End Property

Public Property Let Port(ByVal New_Port As Integer)
  m_Port = New_Port
  PropertyChanged "Port"
End Property


Public Property Get SiteUser() As String
  SiteUser = m_SiteUser
End Property

Public Property Let SiteUser(ByVal New_SiteUser As String)
  m_SiteUser = New_SiteUser
  PropertyChanged "SiteUser"
End Property


Public Property Get SitePassword() As String
  SitePassword = m_SitePassword
End Property

Public Property Let SitePassword(ByVal New_SitePassword As String)
  m_SitePassword = New_SitePassword
  PropertyChanged "SitePassword"
End Property


Public Property Get ProxyPort() As Integer
  ProxyPort = m_ProxyPort
End Property

Public Property Let ProxyPort(ByVal New_ProxyPort As Integer)
  m_ProxyPort = New_ProxyPort
  PropertyChanged "ProxyPort"
End Property


Public Property Get ProxyServer() As String
  ProxyServer = m_ProxyServer
End Property

Public Property Let ProxyServer(ByVal New_ProxyServer As String)
  m_ProxyServer = New_ProxyServer
  PropertyChanged "ProxyServer"
End Property


Public Property Get ProxyUser() As String
  ProxyUser = m_ProxyUser
End Property

Public Property Let ProxyUser(ByVal New_ProxyUser As String)
  m_ProxyUser = New_ProxyUser
  PropertyChanged "ProxyUser"
End Property


Public Property Get ProxyPassword() As String
  ProxyPassword = m_ProxyPassword
End Property

Public Property Let ProxyPassword(ByVal New_ProxyPassword As String)
  m_ProxyPassword = New_ProxyPassword
  PropertyChanged "ProxyPassword"
End Property


Public Sub StartDownload()
  ' set position to 0 and start download
  m_FilePosition = 0
  Download
End Sub


Public Sub ResumeDownload()
  On Error Resume Next
  
  ' Try to get the current filesize in bytes
  plFilePosition = FileLen(m_LocalFile) + 1
  If Err <> 0 Then
    ' file does not exists, so set current position to 0
    plFilePosition = 0
  End If
  
  ' Download the file
  Download
End Sub


Public Sub CancelDownload()
  ' set flag for download cancel
  m_CancelDownload = True
End Sub

Private Sub UserControl_Resize()
  Height = 570
  Width = 570
End Sub

Private Sub UserControl_InitProperties()
  ' set default props
  m_LocalFile = m_def_LocalFile
  m_Url = m_def_Url
  m_Port = m_def_Port
  m_SiteUser = m_def_SiteUser
  m_SitePassword = m_def_SitePassword
  m_ProxyPort = m_def_ProxyPort
  m_ProxyServer = m_def_ProxyServer
  m_ConnectType = m_def_ConnectType
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' get props
  m_LocalFile = PropBag.ReadProperty("LocalFile", m_def_LocalFile)
  m_Url = PropBag.ReadProperty("Url", m_def_Url)
  m_Port = PropBag.ReadProperty("Port", m_def_Port)
  m_SiteUser = PropBag.ReadProperty("SiteUser", m_def_SiteUser)
  m_SitePassword = PropBag.ReadProperty("SitePassword", m_def_SitePassword)
  m_ProxyPort = PropBag.ReadProperty("ProxyPort", m_def_ProxyPort)
  m_ProxyServer = PropBag.ReadProperty("ProxyServer", m_def_ProxyServer)
  m_ConnectType = PropBag.ReadProperty("ConnectType", m_def_ConnectType)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  ' set props
  Call PropBag.WriteProperty("LocalFile", m_LocalFile, m_def_LocalFile)
  Call PropBag.WriteProperty("Url", m_Url, m_def_Url)
  Call PropBag.WriteProperty("Port", m_Port, m_def_Port)
  Call PropBag.WriteProperty("SiteUser", m_SiteUser, m_def_SiteUser)
  Call PropBag.WriteProperty("SitePassword", m_SitePassword, m_def_SitePassword)
  Call PropBag.WriteProperty("ProxyPort", m_ProxyPort, m_def_ProxyPort)
  Call PropBag.WriteProperty("ProxyServer", m_ProxyServer, m_def_ProxyServer)
  Call PropBag.WriteProperty("ConnectType", m_ConnectType, m_def_ConnectType)
End Sub

Private Sub Download()
  Dim sReadBuffer As String * 2048, lNumberOfBytesRead  As Long
  Dim lTotalBytesRead As Long, fFlag As Boolean
  Dim lUrlFile As Long, lInternetSession As Long
  Dim iFileNum As Integer, lReturn As Long
  
  ' Set CancelDownload to false
  m_CancelDownload = False
  
  ' Are we connected to server?
  DoEvents
  If plInternetSession = 0 Then
    Connect
  End If
  lInternetSession = plInternetSession
    
  ' We need to request some data, to show the filesize, etc.
  DoEvents
  pfAllowClose = False
  If Not pfInformationRequested Then RequestInformation
  pfAllowClose = True
  
  ' Try to open Url
  DoEvents
  lUrlFile = InternetOpenUrl(lInternetSession, m_Url, vbNullString, 0, INTERNET_FLAG_EXISITING_CONNECT, 0)
  
  ' Resume Download, or StartDownload
  If plFilePosition <> 0 Then
    DoEvents
    lReturn = InternetSetFilePointer(lUrlFile, plFilePosition, 0, INTERNET_FILE_MOVE_CURRENT, 0)
    If lReturn = 0 Then
      RaiseEvent DownloadError("Error, while setting position")
      InternetCloseHandle lUrlFile
      InternetCloseHandle lInternetSession
      plInternetSession = 0
    Else
      lTotalBytesRead = lReturn
    End If
  End If
  
  ' If lUrlFile = 0 then the url doesn't exists
  If CBool(lUrlFile) Then
    ' Open Localfile
    DoEvents
    iFileNum = FreeFile
    Open m_LocalFile For Binary As #iFileNum
    
    fFlag = True
    Do Until fFlag = False
      ' Start reading
      DoEvents
      sReadBuffer = ""
      fFlag = InternetReadFile(lUrlFile, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
      
      ' If error while reading, cancel Download
      If Not CBool(fFlag) Then
        DoEvents
        InternetCloseHandle lUrlFile
        InternetCloseHandle lInternetSession
        plInternetSession = 0
        Close #iFileNum
        RaiseEvent DownloadError("Could not read Url")
        Exit Sub
      End If
      
      ' Raise our Download Progress
      DoEvents
      lTotalBytesRead = lTotalBytesRead + lNumberOfBytesRead
      If CBool(lNumberOfBytesRead) Then
        Put #iFileNum, , sReadBuffer
      Else
        fFlag = False
      End If
      RaiseEvent DownloadProgress(lTotalBytesRead)
      
      ' Check, if User cancelled Download
      DoEvents
      If m_CancelDownload Then
        ' User cancelled Download, so raise error and close all handles
        DoEvents
        RaiseEvent DownloadCancelled(lTotalBytesRead)
        InternetCloseHandle lUrlFile
        InternetCloseHandle lInternetSession
        plInternetSession = 0
        Close #iFileNum
        Exit Sub
      End If
      DoEvents
    Loop
    
    ' Free all handles and close the localfile
    InternetCloseHandle lUrlFile
    InternetCloseHandle lInternetSession
    plInternetSession = 0
    Close #iFileNum
    
    ' Fire Download complete event
    RaiseEvent DownloadComplete
  Else
    ' Could not find url
    RaiseEvent DownloadError("Could not find Url")
  End If
End Sub

Public Sub RequestInformation()
  Dim sBuffer As String * 1024, lBufferLength   As Long
  Dim strObject As String, sFile As String
  Dim lInternetConnect As Long, lRequest As Long
  Dim lPos As Long, sUrl As String
  
  pfInformationRequested = True
  sUrl = m_Url
  Call CrackURL(sUrl, sFile)
  
  ' Are we connected to server?
  DoEvents
  If plInternetSession = 0 Then Connect
  
  ' Connect to Server, using given Port, Servername, Username and Password for Website
  DoEvents
  lInternetConnect = InternetConnect(plInternetSession, sUrl, 0, IIf(Len(m_SiteUser) > 0, m_SiteUser, vbNullString), IIf(Len(m_SiteUser) > 0, m_SitePassword, vbNullString), 3, 0, 0)
  lRequest = HttpOpenRequest(lInternetConnect, "HEAD", sFile, "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
  
  If CBool(lRequest) Then
    DoEvents
    sBuffer = Space(1024)
    lBufferLength = Len(sBuffer)
    iRetVal = HttpSendRequest(lRequest, vbNullString, 0, 0, 0)
    
    DoEvents
    CheckURL = CBool(HttpQueryInfo(lRequest, 22, ByVal sBuffer, lBufferLength, 0))
    If CheckURL Then
      lPos = InStr(1, LCase(sBuffer), "last-modified: ")
      If lPos > 0 Then
        lPos = lPos + Len("Last Modified: ")
        m_LastModified = Mid$(sBuffer, lPos, InStr(lPos, sBuffer, vbCrLf) - lPos)
      Else
        m_LastModified = "Not Available"
      End If
      
      DoEvents
      lPos = InStr(1, LCase(sBuffer), "content-length: ")
      If lPos > 0 Then
        lPos = lPos + Len("Content-Length: ")
        m_FileSize = CLng(Mid$(sBuffer, lPos, InStr(lPos, sBuffer, vbCrLf) - lPos))
      Else
        m_FileSize = -1
      End If
      
      ' if we can request this information, then the file is available
      DoEvents
      If InStr(1, sBuffer, "HTTP/1.1 404 File Not Found") Then
        m_FileExists = False
      Else
        m_FileExists = True
      End If
    Else
      DoEvents
      HttpQueryInfo lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal sBuffer, lBufferLength, 0
      If TrimString(sBuffer) = "200" Then
        m_FileExists = True
      Else
        m_FileExists = False
      End If
      m_LastModified = "Not Available"
      m_FileSize = -1
    End If
  End If
  
  DoEvents
  InternetCloseHandle lRequest
  InternetCloseHandle lInternetConnect
  If pfAllowClose Then
    InternetCloseHandle plInternetSession
    plInternetSession = 0
  End If
End Sub

Private Sub Connect()
  ' Do we have to connect trough a Proxy, or Direct or do we use the Preconfiguration of IE
  Select Case m_ConnectType
  Case 0 ' Preconfig
    plInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  Case 1
    plInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PROXY, m_ProxyServer & ":" & m_ProxyPort, vbNullString, 0)
  Case 2
    plInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
  End Select
End Sub


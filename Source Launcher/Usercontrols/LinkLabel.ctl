VERSION 5.00
Begin VB.UserControl LinkLabel 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1410
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   MousePointer    =   99  'Custom
   PropertyPages   =   "LinkLabel.ctx":0000
   ScaleHeight     =   1575
   ScaleWidth      =   1410
   ToolboxBitmap   =   "LinkLabel.ctx":0035
   Begin VB.Timer tmrCursor 
      Interval        =   10
      Left            =   -360
      Top             =   0
   End
   Begin VB.Label lblVisited 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   -3787
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblClick 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   -3787
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblHover 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   -3787
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "LinkLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event LinkClick()
Public Event LinkChange()
Public Event Change()
Public Event Click()
Public Event DoubleClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseOver(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum BackStyleCts
    [TRANSPARENT] = 0
    [Opaque]
End Enum

Public Enum AlignmentCts
    [AlignLeft] = 0
    [AlignCenter]
    [AlignRight]
End Enum

Public Enum AppearanceCts
    [Flat] = 0
    [3D]
End Enum

Public Enum BorderStyleCts
    [None] = 0
    [Fixed Single]
End Enum

Public Enum FontStyles
    lv_Plain = 0
    lv_Bold = 2
    lv_Italic = 4
    lv_Strikethrough = 1
    lv_Underline = 8
    lv_BoldItalic = 2 Or 4
    lv_BoldUnderline = 2 Or 8
    lv_BoldStrikethrough = 2 Or 1
    lv_ItalicUnderline = 4 Or 8
    lv_ItalicStrikethrough = 4 Or 1
    lv_StrikethroughUnderline = 1 Or 8
    lv_BoldItalicUnderline = 2 Or 4 Or 8
    lv_BoldStrikethroughUnderline = 2 Or 1 Or 8
    lv_ItalicStrikethroughUnderline = 4 Or 1 Or 8
    lv_BoldItalicStrikethroughUnderline = 2 Or 4 Or 1 Or 8
End Enum

Private Const IEXplorer2 As String = "C:\Program Files\IExplorer\IEXPLORE.EXE"
Private Const IExplorer As String = "C:\Program Files\Internet Explorer\IEXPLORE.EXE"

Private laColor As OLE_COLOR
Private lhColor As OLE_COLOR
Private lcColor As OLE_COLOR
Private lvColor As OLE_COLOR

Private alBColor As OLE_COLOR
Private hlBColor As OLE_COLOR
Private clBColor As OLE_COLOR
Private vlBColor As OLE_COLOR

Private alFont As New StdFont
Private hlFont As New StdFont
Private clFont As New StdFont
Private vlFont As New StdFont

Private rvLink As Boolean
Private mClick As Boolean
Private lEnabled As Boolean

Private mShift As Integer
Private mButton As Integer

Private dText As String
Private sLink As String

Private Cursor As clscursor

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpCommand As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowNormal As Long) As Long

Private Function IsFileExist(ByVal Path As String) As Boolean
    If Not Path = vbNullString Then IsFileExist = Not (Dir$(Path, vbArchive + vbHidden + vbReadOnly + vbSystem) = vbNullString)
End Function

Private Sub SetLink(ByVal StrLink As String)
    If Not StrLink = vbNullString Then Call SaveSetting("LinkLabel", "VisitedLink", StrLink, StrLink)
End Sub

Private Function IsLinkExist(ByVal StrLink As String) As Boolean
    If Not StrLink = vbNullString Then IsLinkExist = Not (GetSetting("LinkLabel", "VisitedLink", StrLink, vbNullString) = vbNullString)
End Function

Public Property Get DisplayedText() As String
    DisplayedText = dText
End Property

Public Property Let DisplayedText(ByVal NewText As String)
    If NewText = lblLink Then Exit Property
    dText = NewText
    lblLink = NewText
    RaiseEvent Change
    Call PropertyChanged("DisplayedText")
    Call UserControl_Resize
End Property

Public Property Get Link() As String
    Link = sLink
End Property

Public Property Let Link(NewLink As String)
    If NewLink = sLink Then Exit Property
    sLink = NewLink
    RaiseEvent LinkChange
    Call PropertyChanged("LinkText")
End Property

Public Property Let ActivateLink(ByVal NewVal As Boolean)
    If NewVal = True Then
        If Not Link = vbNullString Then
            Call ShellExecute(hWnd, "open", IIf(IsFileExist(IExplorer), IExplorer, IEXplorer2), Link, vbNullString, 1)
            If RememberVisitedLink Then Call SetLink(Link)
            RaiseEvent LinkClick
            With lblLink
                If Not RememberVisitedLink Then
                    .ForeColor = ActiveLinkColor
                    Set .Font = ActiveLinkFont
                    .BackColor = ActiveLinkBackColor
                Else
                    If IsLinkExist(Link) Then
                        .ForeColor = VisitedLinkColor
                        Set .Font = VisitedLinkFont
                        .BackColor = VisitedLinkBackColor
                    Else
                        .ForeColor = ActiveLinkColor
                        Set .Font = ActiveLinkFont
                        .BackColor = ActiveLinkBackColor
                    End If
                End If
            End With
            Call UserControl_Resize
        End If
    End If
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = lblLink.AutoSize
End Property

Public Property Let AutoSize(ByVal NewVal As Boolean)
    If NewVal = lblLink.AutoSize Then Exit Property
    lblLink.AutoSize = NewVal
    Call PropertyChanged("AutoSize")
    Call UserControl_Resize
End Property

Public Property Get Alignment() As AlignmentCts
    Alignment = lblLink.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentCts)
    If New_Alignment = lblLink.Alignment Then Exit Property
    lblLink.Alignment = New_Alignment
    Call UserControl_Resize
    Call PropertyChanged("Alignment")
End Property

Public Property Get Appearance() As AppearanceCts
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceCts)
    If New_Appearance = UserControl.Appearance Then Exit Property
    UserControl.Appearance = New_Appearance
    Call PropertyChanged("Appearance")
    Call UserControl_Resize
End Property

Public Property Get BackStyle() As BackStyleCts
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackStyleCts)
    If New_BackStyle = UserControl.BackStyle Then Exit Property
    UserControl.BackStyle = New_BackStyle
    lblLink.BackStyle = New_BackStyle
    Call PropertyChanged("BackStyle")
    Call UserControl_Resize
End Property

Public Property Get BorderStyle() As BorderStyleCts
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleCts)
    If New_BorderStyle = UserControl.BorderStyle Then Exit Property
    UserControl.BorderStyle = New_BorderStyle
    Call PropertyChanged("BorderStyle")
    Call UserControl_Resize
End Property

Public Property Get ActiveLinkColor() As OLE_COLOR
    ActiveLinkColor = laColor
End Property

Public Property Let ActiveLinkColor(nColor As OLE_COLOR)
    If nColor = lblLink.ForeColor Then Exit Property
    laColor = nColor
    lblLink.ForeColor = nColor
    Call PropertyChanged("laColor")
End Property

Public Property Get ActiveLinkBackColor() As OLE_COLOR
    ActiveLinkBackColor = alBColor
End Property

Public Property Let ActiveLinkBackColor(nColor As OLE_COLOR)
    If nColor = alBColor Then Exit Property
    alBColor = nColor
    lblLink.BackColor = nColor
    Call PropertyChanged("alBColor")
End Property

Public Property Set ActiveLinkFont(nFont As StdFont)
    With alFont
        .Charset = nFont.Charset
        .Name = nFont.Name
        .Size = nFont.Size
        .Weight = nFont.Weight
        .Bold = nFont.Bold
        .Italic = nFont.Italic
        .Underline = nFont.Underline
        .Strikethrough = nFont.Strikethrough
    End With
    Call PropertyChanged("ActiveLinkFont")
    Set lblLink.Font = nFont
    Call UserControl_Resize
End Property

Public Property Get ActiveLinkFont() As StdFont
    Set ActiveLinkFont = alFont
End Property

Public Property Let ActiveLinkFontStyle(nStyle As FontStyles)
    With alFont
        .Bold = ((nStyle And lv_Bold) = lv_Bold)
        .Italic = ((nStyle And lv_Italic) = lv_Italic)
        .Underline = ((nStyle And lv_Underline) = lv_Underline)
        .Strikethrough = ((nStyle And lv_Strikethrough) = lv_Strikethrough)
        lblLink.Font.Bold = .Bold
        lblLink.Font.Italic = .Italic
        lblLink.Font.Underline = .Underline
        lblLink.Font.Strikethrough = .Strikethrough
    End With
    Call PropertyChanged("ActiveLinkFontStyle")
    Call UserControl_Resize
End Property

Public Property Get ActiveLinkFontStyle() As FontStyles
    Dim nStyle As Integer
    nStyle = nStyle Or Abs(alFont.Bold) * 2
    nStyle = nStyle Or Abs(alFont.Italic) * 4
    nStyle = nStyle Or Abs(alFont.Strikethrough)
    nStyle = nStyle Or Abs(alFont.Underline) * 8
    ActiveLinkFontStyle = nStyle
End Property

Public Property Get HoverLinkColor() As OLE_COLOR
    HoverLinkColor = lhColor
End Property

Public Property Let HoverLinkColor(nColor As OLE_COLOR)
    If nColor = lhColor Then Exit Property
    lhColor = nColor
    Call PropertyChanged("lhColor")
End Property

Public Property Get HoverLinkBackColor() As OLE_COLOR
    HoverLinkBackColor = hlBColor
End Property

Public Property Let HoverLinkBackColor(nColor As OLE_COLOR)
    If nColor = hlBColor Then Exit Property
    hlBColor = nColor
    Call PropertyChanged("hlBColor")
End Property

Public Property Set HoverLinkFont(nFont As StdFont)
    With hlFont
        .Charset = nFont.Charset
        .Name = nFont.Name
        .Size = nFont.Size
        .Weight = nFont.Weight
        .Bold = nFont.Bold
        .Italic = nFont.Italic
        .Underline = nFont.Underline
        .Strikethrough = nFont.Strikethrough
    End With
    Call PropertyChanged("HoverLinkFont")
    Set lblHover.Font = nFont
    Call UserControl_Resize
End Property

Public Property Get HoverLinkFont() As StdFont
    Set HoverLinkFont = hlFont
End Property

Public Property Let HoverLinkFontStyle(nStyle As FontStyles)
    With hlFont
        .Bold = ((nStyle And lv_Bold) = lv_Bold)
        .Italic = ((nStyle And lv_Italic) = lv_Italic)
        .Underline = ((nStyle And lv_Underline) = lv_Underline)
        .Strikethrough = ((nStyle And lv_Strikethrough) = lv_Strikethrough)
        lblHover.Font.Bold = .Bold
        lblHover.Font.Italic = .Italic
        lblHover.Font.Underline = .Underline
        lblHover.Font.Strikethrough = .Strikethrough
    End With
    Call PropertyChanged("HoverLinkFontStyle")
    Call UserControl_Resize
End Property

Public Property Get HoverLinkFontStyle() As FontStyles
    Dim nStyle As Integer
    nStyle = nStyle Or Abs(hlFont.Bold) * 2
    nStyle = nStyle Or Abs(hlFont.Italic) * 4
    nStyle = nStyle Or Abs(hlFont.Strikethrough)
    nStyle = nStyle Or Abs(hlFont.Underline) * 8
    HoverLinkFontStyle = nStyle
End Property

Public Property Get ClickLinkColor() As OLE_COLOR
    ClickLinkColor = lcColor
End Property

Public Property Let ClickLinkColor(nColor As OLE_COLOR)
    If nColor = lcColor Then Exit Property
    lcColor = nColor
    Call PropertyChanged("lcColor")
End Property

Public Property Get ClickLinkBackColor() As OLE_COLOR
    ClickLinkBackColor = clBColor
End Property

Public Property Let ClickLinkBackColor(nColor As OLE_COLOR)
    If nColor = clBColor Then Exit Property
    clBColor = nColor
    Call PropertyChanged("clBColor")
End Property

Public Property Set ClickLinkFont(nFont As StdFont)
    With clFont
        .Charset = nFont.Charset
        .Name = nFont.Name
        .Size = nFont.Size
        .Weight = nFont.Weight
        .Bold = nFont.Bold
        .Italic = nFont.Italic
        .Underline = nFont.Underline
        .Strikethrough = nFont.Strikethrough
    End With
    Call PropertyChanged("ClickLinkFont")
    Set lblClick.Font = nFont
    Call UserControl_Resize
End Property

Public Property Get ClickLinkFont() As StdFont
    Set ClickLinkFont = clFont
End Property

Public Property Let ClickLinkFontStyle(nStyle As FontStyles)
    With clFont
        .Bold = ((nStyle And lv_Bold) = lv_Bold)
        .Italic = ((nStyle And lv_Italic) = lv_Italic)
        .Underline = ((nStyle And lv_Underline) = lv_Underline)
        .Strikethrough = ((nStyle And lv_Strikethrough) = lv_Strikethrough)
        lblClick.Font.Bold = .Bold
        lblClick.Font.Italic = .Italic
        lblClick.Font.Underline = .Underline
        lblClick.Font.Strikethrough = .Strikethrough
    End With
    Call PropertyChanged("ClickLinkFontStyle")
    Call UserControl_Resize
End Property

Public Property Get ClickLinkFontStyle() As FontStyles
    Dim nStyle As Integer
    nStyle = nStyle Or Abs(clFont.Bold) * 2
    nStyle = nStyle Or Abs(clFont.Italic) * 4
    nStyle = nStyle Or Abs(clFont.Strikethrough)
    nStyle = nStyle Or Abs(clFont.Underline) * 8
    ClickLinkFontStyle = nStyle
End Property

Public Property Get VisitedLinkColor() As OLE_COLOR
    VisitedLinkColor = lvColor
End Property

Public Property Let VisitedLinkColor(nColor As OLE_COLOR)
    If nColor = lvColor Then Exit Property
    lvColor = nColor
    Call PropertyChanged("lvColor")
End Property

Public Property Get VisitedLinkBackColor() As OLE_COLOR
    VisitedLinkBackColor = vlBColor
End Property

Public Property Let VisitedLinkBackColor(nColor As OLE_COLOR)
    If nColor = vlBColor Then Exit Property
    vlBColor = nColor
    Call PropertyChanged("vlBColor")
End Property

Public Property Set VisitedLinkFont(nFont As StdFont)
    With vlFont
        .Charset = nFont.Charset
        .Name = nFont.Name
        .Size = nFont.Size
        .Weight = nFont.Weight
        .Bold = nFont.Bold
        .Italic = nFont.Italic
        .Underline = nFont.Underline
        .Strikethrough = nFont.Strikethrough
    End With
    Call PropertyChanged("VisitedLinkFont")
    Set lblVisited.Font = nFont
    Call UserControl_Resize
End Property

Public Property Get VisitedLinkFont() As StdFont
    Set VisitedLinkFont = vlFont
End Property

Public Property Let VisitedLinkFontStyle(nStyle As FontStyles)
    With vlFont
        .Bold = ((nStyle And lv_Bold) = lv_Bold)
        .Italic = ((nStyle And lv_Italic) = lv_Italic)
        .Underline = ((nStyle And lv_Underline) = lv_Underline)
        .Strikethrough = ((nStyle And lv_Strikethrough) = lv_Strikethrough)
        lblVisited.Font.Bold = .Bold
        lblVisited.Font.Italic = .Italic
        lblVisited.Font.Underline = .Underline
        lblVisited.Font.Strikethrough = .Strikethrough
    End With
    Call PropertyChanged("VisitedLinkFontStyle")
    Call UserControl_Resize
End Property

Public Property Get VisitedLinkFontStyle() As FontStyles
    Dim nStyle As Integer
    nStyle = nStyle Or Abs(vlFont.Bold) * 2
    nStyle = nStyle Or Abs(vlFont.Italic) * 4
    nStyle = nStyle Or Abs(vlFont.Strikethrough)
    nStyle = nStyle Or Abs(vlFont.Underline) * 8
    VisitedLinkFontStyle = nStyle
End Property

Public Property Get RememberVisitedLink() As Boolean
    RememberVisitedLink = rvLink
End Property

Public Property Let RememberVisitedLink(ByVal NewVal As Boolean)
    If NewVal = rvLink Then Exit Property
    rvLink = NewVal
    Call PropertyChanged("rvLink")
End Property

Public Property Get Enabled() As Boolean
    Enabled = lEnabled
End Property

Public Property Let Enabled(NewVal As Boolean)
    If NewVal = lEnabled Then Exit Property
    lEnabled = NewVal
    UserControl.Enabled = NewVal
    lblLink.Enabled = NewVal
    Call PropertyChanged("lEnable")
End Property

Public Property Let MousePointer(nPointer As MousePointerConstants)
    If nPointer = UserControl.MousePointer Then Exit Property
    UserControl.MousePointer = nPointer
    Call PropertyChanged("mPointer")
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Set MouseIcon(nIcon As StdPicture)
    On Error GoTo ShowPropertyError
    If Not nIcon Is Nothing Then Exit Property
    Set UserControl.MouseIcon = nIcon
    If Not nIcon Is Nothing Then
        MousePointer = vbCustom
        Call PropertyChanged("mIcon")
    End If
    Exit Property

ShowPropertyError:
    If Ambient.UserMode = False Then Call Err.Raise(Err.Number, , Err.Description)
End Property

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub lblLink_Click()
    Call UserControl_Click
End Sub

Private Sub lblLink_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub lblLink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mButton = Button
    If Button = vbLeftButton Then
        With lblLink
            .ForeColor = ClickLinkColor
            Set .Font = ClickLinkFont
            .BackColor = ClickLinkBackColor
        End With
        Call UserControl_Resize
    End If
    Call UserControl_MouseDown(Button, Shift, X, Y)
    mClick = True
End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblLink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MousePos As POINTAPI
    mButton = Button
    If Button = vbLeftButton Then
        If Not Link = vbNullString Then
            Call ShellExecute(hWnd, "open", IIf(IsFileExist(IExplorer), IExplorer, IEXplorer2), Link, vbNullString, 1)
            If RememberVisitedLink Then Call SetLink(Link)
            RaiseEvent LinkClick
        End If
        Call GetCursorPos(MousePos)
        With lblLink
            If Not RememberVisitedLink Then
                Select Case WindowFromPoint(MousePos.X, MousePos.Y) = hWnd
                    Case True
                        .ForeColor = HoverLinkColor
                        Set .Font = HoverLinkFont
                        .BackColor = HoverLinkBackColor
                        RaiseEvent MouseOver(mButton, mShift, CSng(MousePos.X), CSng(MousePos.Y))
                    Case False
                        .ForeColor = ActiveLinkColor
                        Set .Font = ActiveLinkFont
                        .BackColor = ActiveLinkBackColor
                        RaiseEvent MouseOut(mButton, mShift, CSng(MousePos.X), CSng(MousePos.Y))
                End Select
            Else
                If IsLinkExist(Link) Then
                    .ForeColor = VisitedLinkColor
                    Set .Font = VisitedLinkFont
                    .BackColor = VisitedLinkBackColor
                Else
                    Select Case WindowFromPoint(MousePos.X, MousePos.Y) = hWnd
                        Case True
                            .ForeColor = HoverLinkColor
                            Set .Font = HoverLinkFont
                            .BackColor = HoverLinkBackColor
                            RaiseEvent MouseOver(mButton, mShift, CSng(MousePos.X), CSng(MousePos.Y))
                        Case False
                            .ForeColor = ActiveLinkColor
                            Set .Font = ActiveLinkFont
                            .BackColor = ActiveLinkBackColor
                            RaiseEvent MouseOut(mButton, mShift, CSng(MousePos.X), CSng(MousePos.Y))
                    End Select
                End If
            End If
        End With
        Call UserControl_Resize
    End If
    Call UserControl_MouseUp(Button, Shift, X, Y)
    mClick = False
End Sub

Private Sub lblLink_OLECompleteDrag(Effect As Long)
    Call UserControl_OLECompleteDrag(Effect)
End Sub

Private Sub lblLink_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub lblLink_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Call UserControl_OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub lblLink_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    Call UserControl_OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub lblLink_OLESetData(Data As DataObject, DataFormat As Integer)
    Call UserControl_OLESetData(Data, DataFormat)
End Sub

Private Sub lblLink_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Call UserControl_OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub tmrCursor_Timer()
    Static mPointerTrue As Boolean, mPointerFalse As Boolean
    Dim MousePos As POINTAPI
    If Not mClick Then
        Call GetCursorPos(MousePos)
        With lblLink
            Select Case WindowFromPoint(MousePos.X, MousePos.Y) = hWnd
                Case True
                    If Not mPointerTrue Then
                        .ForeColor = HoverLinkColor
                        Set .Font = HoverLinkFont
                        .BackColor = HoverLinkBackColor
                        mPointerTrue = True
                        mPointerFalse = False
                    End If
                    RaiseEvent MouseOver(mButton, mShift, CSng(MousePos.X), CSng(MousePos.Y))
                Case False
                    If Not mPointerFalse Then
                        mPointerTrue = False
                        mPointerFalse = True
                        If Not RememberVisitedLink Then
                            .ForeColor = ActiveLinkColor
                            Set .Font = ActiveLinkFont
                            .BackColor = ActiveLinkBackColor
                        Else
                            If IsLinkExist(Link) Then
                                .ForeColor = VisitedLinkColor
                                Set .Font = VisitedLinkFont
                                .BackColor = VisitedLinkBackColor
                            Else
                                .ForeColor = ActiveLinkColor
                                Set .Font = ActiveLinkFont
                                .BackColor = ActiveLinkBackColor
                            End If
                        End If
                    End If
                    RaiseEvent MouseOut(mButton, mShift, CSng(MousePos.X), CSng(MousePos.Y))
            End Select
        End With
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DoubleClick
End Sub

Private Sub UserControl_Initialize()
    Set Cursor = New clscursor
    Set UserControl.MouseIcon = Cursor.ShowCursor(IDC_HAND)
End Sub

Private Sub UserControl_InitProperties()
    Set alFont = New StdFont
    Set hlFont = New StdFont
    Set clFont = New StdFont
    Set vlFont = New StdFont
    RememberVisitedLink = True
    ActiveLinkColor = vbBlue
    ActiveLinkFontStyle = lv_Underline
    ActiveLinkFont = lblLink.Font
    ActiveLinkBackColor = vbButtonFace
    HoverLinkColor = vbRed
    HoverLinkFontStyle = lv_Underline
    HoverLinkFont = lblHover.Font
    HoverLinkBackColor = vbButtonFace
    ClickLinkColor = vbBlack
    ClickLinkFontStyle = lv_Underline
    ClickLinkFont = lblClick.Font
    ClickLinkBackColor = vbButtonFace
    VisitedLinkColor = vbMagenta
    VisitedLinkFontStyle = lv_Underline
    VisitedLinkFont = lblVisited.Font
    VisitedLinkBackColor = vbButtonFace
    DisplayedText = "LinkLabel"
    Link = vbNullString
    Enabled = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 16: mShift = 1
        Case 17: mShift = 2
        Case 18: mShift = 4
    End Select
    If KeyCode Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 16: mShift = 1
        Case 17: mShift = 2
        Case 18: mShift = 4
    End Select
    If KeyCode Then RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyLeft Then Call SetCapture(hWnd)
    mShift = Shift
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyLeft Then Call SetCapture(hWnd)
    mShift = Shift
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyLeft Then Call SetCapture(hWnd)
    mShift = Shift
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        RememberVisitedLink = .ReadProperty("rvLink", True)
        ActiveLinkColor = .ReadProperty("laColor", vbBlue)
        HoverLinkColor = .ReadProperty("lhColor", vbRed)
        ClickLinkColor = .ReadProperty("lcColor", vbBlack)
        VisitedLinkColor = .ReadProperty("lvColor", vbMagenta)
        ActiveLinkBackColor = .ReadProperty("alBColor", vbButtonFace)
        HoverLinkBackColor = .ReadProperty("hlBColor", vbButtonFace)
        ClickLinkBackColor = .ReadProperty("clBColor", vbButtonFace)
        VisitedLinkBackColor = .ReadProperty("vlBColor", vbButtonFace)
        Enabled = .ReadProperty("lEnable", True)
        Alignment = .ReadProperty("Alignment", [AlignLeft])
        Appearance = .ReadProperty("Appearance", [3D])
        BackStyle = .ReadProperty("BackStyle", [Opaque])
        BorderStyle = .ReadProperty("BorderStyle", [None])
        AutoSize = .ReadProperty("AutoSize", True)
        ActiveLinkFontStyle = .ReadProperty("ActiveLinkFontStyle", 8)
        Set ActiveLinkFont = .ReadProperty("ActiveLinkFont", lblLink.Font)
        HoverLinkFontStyle = .ReadProperty("HoverLinkFontStyle", 8)
        Set HoverLinkFont = .ReadProperty("HoverLinkFont", lblHover.Font)
        ClickLinkFontStyle = .ReadProperty("ClickLinkFontStyle", 8)
        Set ClickLinkFont = .ReadProperty("ClickLinkFont", lblClick.Font)
        VisitedLinkFontStyle = .ReadProperty("VisitedLinkFontStyle", 8)
        Set VisitedLinkFont = .ReadProperty("VisitedLinkFont", lblVisited.Font)
        Set MouseIcon = .ReadProperty("mIcon", UserControl.MouseIcon)
        MousePointer = .ReadProperty("mPointer", vbCustom)
        DisplayedText = .ReadProperty("DisplayedText", "LinkLabel")
        Link = .ReadProperty("LinkText", vbNullString)
    End With
    Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    lblLink.Top = 0
    lblLink.Left = 0
    If AutoSize Then
        Width = lblLink.Width
        Height = lblLink.Height
    Else
        lblLink.Width = Width
        lblLink.Height = Height
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("rvLink", RememberVisitedLink, True)
        Call .WriteProperty("laColor", ActiveLinkColor, vbBlue)
        Call .WriteProperty("lhColor", HoverLinkColor, vbRed)
        Call .WriteProperty("lcColor", ClickLinkColor, vbBlack)
        Call .WriteProperty("lvColor", VisitedLinkColor, vbMagenta)
        Call .WriteProperty("alBColor", ActiveLinkBackColor, vbButtonFace)
        Call .WriteProperty("hlBColor", HoverLinkBackColor, vbButtonFace)
        Call .WriteProperty("clBColor", ClickLinkBackColor, vbButtonFace)
        Call .WriteProperty("vlBColor", VisitedLinkBackColor, vbButtonFace)
        Call .WriteProperty("lEnable", Enabled, True)
        Call .WriteProperty("Alignment", Alignment, [AlignLeft])
        Call .WriteProperty("Appearance", Appearance, [3D])
        Call .WriteProperty("BackStyle", BackStyle, [Opaque])
        Call .WriteProperty("BorderStyle", BorderStyle, [None])
        Call .WriteProperty("AutoSize", AutoSize, True)
        Call .WriteProperty("ActiveLinkFontStyle", ActiveLinkFontStyle, 8)
        Call .WriteProperty("ActiveLinkFont", ActiveLinkFont, lblLink.Font)
        Call .WriteProperty("HoverLinkFontStyle", HoverLinkFontStyle, 8)
        Call .WriteProperty("HoverLinkFont", HoverLinkFont, lblHover.Font)
        Call .WriteProperty("ClickLinkFontStyle", ClickLinkFontStyle, 8)
        Call .WriteProperty("ClickLinkFont", ClickLinkFont, lblClick.Font)
        Call .WriteProperty("VisitedLinkFontStyle", VisitedLinkFontStyle, 8)
        Call .WriteProperty("VisitedLinkFont", VisitedLinkFont, lblVisited.Font)
        Call .WriteProperty("mIcon", MouseIcon, UserControl.MouseIcon)
        Call .WriteProperty("mPointer", MousePointer, vbCustom)
        Call .WriteProperty("DisplayedText", DisplayedText, "LinkLabel")
        Call .WriteProperty("LinkText", Link, vbNullString)
    End With
End Sub

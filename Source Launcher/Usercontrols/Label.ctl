VERSION 5.00
Begin VB.UserControl Label 
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ################################################# '
' #                                               # '
' # Control Name:  NiceLabel                      # '
' #                                               # '
' # Created:       24/October/2007                # '
' #                                               # '
' # Created by:    Ahmed Alotiabe                 # '
' #                                               # '
' # You Have A Royalty Free Right To Use          # '
' #   Modify And Reproduce And Distribute.        # '
' #                                               # '
' # But I Assume No Warranty, Obligations         # '
' #   Or Liability For The Use Of This Code.      # '
' #                                               # '
' # For Any Damage This Control                   # '
' #           Can Cause To Your Products.         # '
' ################################################# '

'>> Needed For The API Call >> GradientFill >> Backgraound.
Private Type GRADIENT_RECT
    UpperLeft            As Long
    LowerRight           As Long
End Type

Private Type TRIVERTEX
    X                    As Long
    Y                    As Long
    Red                  As Integer
    Green                As Integer
    Blue                 As Integer
    Alpha                As Integer
End Type
'............................................................
'>> Needed For CaptionPostion.
Private Type POINTAPI
    X                    As Long
    Y                    As Long
End Type
'............................................................
'>> Constant Types Used With GradientFill.
Private Const GRADIENT_FILL_RECT_H As Long = &H0
Private Const GRADIENT_FILL_RECT_V As Long = &H1

'............................................................
'>> Constant Types For BroderStyle.
Private Const BDR_RAISEDOUTER As Long = vb3DHighlight
Private Const BDR_SUNKENOUTER1 As Long = vbButtonShadow
Private Const BDR_SUNKENOUTER2 As Long = vb3DDKShadow
'............................................................
'>> Constant Types For CreateFontIndirect.
Private Const LF_FACESIZE As Long = &H32
Private Const LF_FONTREGULAR As Long = &H0
Private Const LF_FONTBOLD As Long = &H700
Private Const LF_ESCAPEMENT_H As Long = 0
Private Const LF_ESCAPEMENT_V As Long = 900
Private Type LOGFONT
    lfHeight             As Long
    lfWidth              As Long
    lfEscapement         As Long
    lfOrientation        As Long
    lfWeight             As Long
    lfItalic             As Byte
    lfUnderline          As Byte
    lfStrikeOut          As Byte
    lfCharSet            As Byte
    lfOutPrecision       As Byte
    lfClipPrecision      As Byte
    lfQuality            As Byte
    lfPitchAndFamily     As Byte
    lfFaceName           As String * LF_FACESIZE
End Type
'............................................................
'>> API Declare Function's.
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOutW Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function TextOutA Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, ByRef pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Integer
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'............................................................
'>> Label Declarations.
'............................................................
Public Enum EnumLabelBroderStyle
    None
    Broder_Classic
    Broder_Sunken
    Broder_Raised
    Broder_3DInside
    Broder_3DOutside
    Broder
    BroderFill_VH
    BroderFill_HV
    BroderFill_CornerIn
    BroderFill_CornerOut
    BroderFill_HorizontalCenter
    BroderFill_HorizontalOut
End Enum
Public Enum EnumLabelBackStyle
    Transparent
    Opaque
    BackFill_HorizontalRight
    BackFill_HorizontalLeft
    BackFill_VerticalDown
    BackFill_VerticalUp
    BackFill_HorizontalCenter
    BackFill_HorizontalOut
    BackFill_VerticalCenter
    BackFill_VerticalOut
End Enum
Public Enum EnumLabelThemeColor
    NoTheme
    XpBlue
    XpOlive
    XPSilver
    Visual2005
    Norton2005
End Enum
Public Enum EnumLabelAlignment
    Left_Justify
    Center_Justify
    Right_Justify
End Enum
Public Enum EnumLabelEffects
    NoEffect
    Raised
    Sunken
    Outline
    Shadow
    RaisedShadow
    SunkenShadow
    OutlineShadow
    Draw1_3D
    Draw2_3D
End Enum
Public Enum EnumLabelFillStyle
    Normal
    TextFill_HorizontalRight
    TextFill_HorizontalLeft
    TextFill_VerticalDown
    TextFill_VerticalUp
    TextFill_HorizontalCenter
    TextFill_HorizontalOut
    TextFill_VerticalCenter
    TextFill_VerticalOut
End Enum
Public Enum EnumLabelRotation
    Horizontal
    Vertical
End Enum
'............................................................
'>> Property Member Variables.
'............................................................
Private mLabelBroderStyle  As EnumLabelBroderStyle    '>> Broder_Classic,Broder_Sunken,Broder_Raised,Broder_3DInside,Broder_3DOutside,BroderFill V,H ....
Private mLabelBackStyle    As EnumLabelBackStyle      '>> Transparent,Opaque,BackFill_HorizontalRight,BackFill_HorizontalLeft,BackFill_VerticalDown,BackFill_VerticalUp,BackFill_HorizontalCenter,BackFill_HorizontalOut,BackFill_VerticalCenter,BackFill_VerticalOut.
Private mLabelThemeColor   As EnumLabelThemeColor     '>> XpBlue,XpOlive,XPSilver,Visual2005 , Norton2005.
Private mLabelAlignment    As EnumLabelAlignment      '>> Left_Justify,Center_Justify,Right_Justify.
Private mLabelEffects      As EnumLabelEffects     '>> Raised,Sunken,Outline,Shadow,RaisedShadow,SunkenShadow,OutlineShadow,Draw1_3D,Draw2_3D.
Private mLabelFillStyle    As EnumLabelFillStyle      '>> TextFill_HorizontalRight,TextFill_HorizontalLeft,TextFill_VerticalDown,TextFill_VerticalUp,TextFill_HorizontalCenter,TextFill_HorizontalOut,TextFill_VerticalCenter,TextFill_VerticalOut.
Private mLabelRotation     As EnumLabelRotation       '>> Horizontal,Vertical.
'............................................................
'>> Property Color Variables
Dim mBackColor1          As OLE_COLOR
Dim mBackColor2          As OLE_COLOR
'............................................................
Dim mBorderColor1        As OLE_COLOR
Dim mBorderColor2        As OLE_COLOR
'............................................................
Dim mForeColor1          As OLE_COLOR
Dim mForeColor2          As OLE_COLOR
'............................................................
Dim mEffectColor         As OLE_COLOR
'............................................................
Dim mCaption             As String
Dim mAutoSize            As Boolean
Dim mRightToLeft         As Boolean
'............................................................
Dim CaptionPos(1)        As POINTAPI '>> Alignment And Postion The Text.
Dim Visual               As POINTAPI '>> Paint Effect's And GradientFill The Text.
Dim Lines                As POINTAPI '>> Draw The Border's.
'............................................................
Dim RotateMe             As LOGFONT
'............................................................
Dim BkRed(0 To 1), BkGreen(0 To 1), BkBlue(0 To 1) As Long '>> GradientFill BackColor.
Dim Vert(1) As TRIVERTEX, GRect As GRADIENT_RECT '>> GradientFill API.
'............................................................
Dim TotalRed(1), TotalGreen(1), TotalBlue(1) As Long '>> GradientFill TextColor.
Dim FrRed, FrGreen, FrBlue As Integer '>> GradientFill TextColor.
Dim R(2), G(2), B(2)     As Long '>> Convert Color's From RGB To Long.
'............................................................
'>> Label Event Declaration's.
'............................................................
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'............................................................
'>> Start Properties.
'...........................................................
Public Sub Refresh()
    UserControl.Refresh
End Sub
Public Property Get BroderStyle() As EnumLabelBroderStyle
    BroderStyle = mLabelBroderStyle
End Property
Public Property Let BroderStyle(ByVal EBS As EnumLabelBroderStyle)
    mLabelBroderStyle = EBS
    PropertyChanged "BroderStyle"
    UserControl_Paint
End Property
Public Property Get BackStyle() As EnumLabelBackStyle
    BackStyle = mLabelBackStyle
End Property
Public Property Let BackStyle(ByVal EBS As EnumLabelBackStyle)
    mLabelBackStyle = EBS
    PropertyChanged "BackStyle"
    UserControl_Paint
End Property
Public Property Get ThemeColor() As EnumLabelThemeColor
    ThemeColor = mLabelThemeColor
End Property
Public Property Let ThemeColor(ByVal ELTC As EnumLabelThemeColor)
    mLabelThemeColor = ELTC
    PropertyChanged "ThemeColor"
    '.............................................................
    '>> LabelTheme,XpBlue,XpOlive,XPSilver,Visual2005,Norton2005.
    '.............................................................
    Select Case mLabelThemeColor
        Case Is = NoTheme
            mBackColor1 = vbButtonFace
            mBackColor2 = vbWindowBackground
            mBorderColor1 = vbDesktop
            mBorderColor2 = vbWindowBackground
            mForeColor1 = vbBlack
            mForeColor2 = vbBlue
        Case Is = XpBlue
            mBackColor1 = &HF1B39A
            mBackColor2 = &HFAE6DD
            mBorderColor1 = &HF8EAE0
            mBorderColor2 = &HD37739
            mForeColor1 = &HD37739
            mForeColor2 = &HF8EAE0
        Case Is = XpOlive
            mBackColor1 = &H3DB4A2
            mBackColor2 = &HB8E7E0
            mBorderColor1 = &HDFEFEB
            mBorderColor2 = &H57AA94
            mForeColor1 = &H57AA94
            mForeColor2 = &HDFEFEB
        Case Is = XPSilver
            mBackColor1 = &HB5A09D
            mBackColor2 = &HECE7E6
            mBorderColor1 = &HE7DEDF
            mBorderColor2 = &H91686B
            mForeColor1 = &H91686B
            mForeColor2 = &HE7DEDF
        Case Is = Visual2005
            mBackColor1 = &HABC2C2
            mBackColor2 = &HE6EDEE
            mBorderColor1 = &HEBF1F1
            mBorderColor2 = &H759B9C
            mForeColor1 = &H759B9C
            mForeColor2 = &HEBF1F1
        Case Is = Norton2005
            mBackColor1 = &H2C5F5
            mBackColor2 = &HAFEFFE
            mBorderColor1 = &H5BD7F5
            mBorderColor2 = &H7657C
            mForeColor1 = &H7657C
            mForeColor2 = &H5BD7F5
    End Select
    UserControl_Paint
End Property
Public Property Get Alignment() As EnumLabelAlignment
    Alignment = mLabelAlignment
End Property
Public Property Let Alignment(ByVal ELA As EnumLabelAlignment)
    mLabelAlignment = ELA
    PropertyChanged "Alignment"
    UserControl_Paint
End Property
Public Property Get Rotation() As EnumLabelRotation
    Rotation = mLabelRotation
End Property
Public Property Let Rotation(ByVal ELR As EnumLabelRotation)
    mLabelRotation = ELR
    PropertyChanged "Rotation"
    With UserControl
        Dim Twirl
        Select Case mLabelRotation
            Case Is = Horizontal '>> Make The Height = Width And The Width = Height.
                If .Width < .Height Or .Width = .Height Then
                    Twirl = .Height
                    .Height = .Width
                    .Width = Twirl
                End If
            Case Is = Vertical '>> Make The Width = Height And The Height = Width.
                If .Width > .Height Or .Width = .Height Then
                    Twirl = .Width
                    .Width = .Height
                    .Height = Twirl
                End If
        End Select
    End With
    UserControl_Paint
End Property
Public Property Get Effects() As EnumLabelEffects
    Effects = mLabelEffects
End Property
Public Property Let Effects(ByVal ELE As EnumLabelEffects)
    mLabelEffects = ELE
    PropertyChanged "Effects"
    With UserControl
    Select Case mLabelEffects '>> Default Font's And Effect's Color's.
       'Case Is = Default
        Case Is = Raised
            mEffectColor = vbHighlightText
            If .FontSize < 12 Then
            .FontSize = 12
            .FontBold = True
             End If
        Case Is = Sunken
            mEffectColor = vbHighlightText
            If .FontSize < 12 Then
            .FontSize = 12
            .FontBold = True
             End If
        Case Is = Outline
            mEffectColor = vbHighlightText
        Case Is = Shadow
            mEffectColor = vb3DDKShadow
        Case Is = RaisedShadow
            mEffectColor = vbHighlightText
            If .FontSize < 12 Then
            .FontSize = 12
            .FontBold = True
             End If
        Case Is = SunkenShadow
            mEffectColor = vbHighlightText
            If .FontSize < 12 Then
            .FontSize = 12
            .FontBold = True
             End If
        Case Is = OutlineShadow
            mEffectColor = vbHighlightText
        Case Is = Draw1_3D
            mEffectColor = vbHighlightText
            If .FontSize < 16 Then
            .FontSize = 16
            .FontBold = True
             End If
        Case Is = Draw2_3D
            mEffectColor = vbHighlightText
            If .FontSize < 16 Then
            .FontSize = 16
            .FontBold = True
             End If
    End Select
     End With
    UserControl_Paint
End Property
Public Property Get FillStyle() As EnumLabelFillStyle
    FillStyle = mLabelFillStyle
End Property
Public Property Let FillStyle(ByVal ELFS As EnumLabelFillStyle)
    mLabelFillStyle = ELFS
    PropertyChanged "FillStyle"
    UserControl_Paint
End Property
Public Property Get BackColor1() As OLE_COLOR
    BackColor1 = mBackColor1
End Property
Public Property Let BackColor1(ByVal New_Color As OLE_COLOR)
    mBackColor1 = New_Color
    PropertyChanged "BackColor1"
    UserControl_Paint
End Property
Public Property Get BackColor2() As OLE_COLOR
    BackColor2 = mBackColor2
End Property
Public Property Let BackColor2(ByVal New_Color As OLE_COLOR)
    mBackColor2 = New_Color
    PropertyChanged "BackColor2"
    UserControl_Paint
End Property
Public Property Get BorderColor1() As OLE_COLOR
    BorderColor1 = mBorderColor1
End Property
Public Property Let BorderColor1(ByVal New_Color As OLE_COLOR)
    mBorderColor1 = New_Color
    PropertyChanged "BorderColor1"
    UserControl_Paint
End Property
Public Property Get BorderColor2() As OLE_COLOR
    BorderColor2 = mBorderColor2
End Property
Public Property Let BorderColor2(ByVal New_Color As OLE_COLOR)
    mBorderColor2 = New_Color
    PropertyChanged "BorderColor2"
    UserControl_Paint
End Property
Public Property Get ForeColor1() As OLE_COLOR
    ForeColor1 = mForeColor1
End Property
Public Property Let ForeColor1(ByVal New_Color As OLE_COLOR)
    mForeColor1 = New_Color
    PropertyChanged "ForeColor1"
    UserControl_Paint
End Property
Public Property Get ForeColor2() As OLE_COLOR
    ForeColor2 = mForeColor2
End Property
Public Property Let ForeColor2(ByVal New_Color As OLE_COLOR)
    mForeColor2 = New_Color
    PropertyChanged "ForeColor2"
    UserControl_Paint
End Property
Public Property Get EffectColor() As OLE_COLOR
    EffectColor = mEffectColor
End Property
Public Property Let EffectColor(ByVal New_Color As OLE_COLOR)
    mEffectColor = New_Color
    PropertyChanged "EffectColor"
    UserControl_Paint
End Property
Public Property Get Caption() As String
    Caption = mCaption
End Property
Public Property Let Caption(ByVal NewCaption As String)
    mCaption = NewCaption
    PropertyChanged "Caption"
    UserControl_Paint
End Property
Public Property Get Autosize() As Boolean
    Autosize = mAutoSize
End Property
Public Property Let Autosize(ByVal New_Autosize As Boolean)
    mAutoSize = New_Autosize
    PropertyChanged "AutoSize"
    With UserControl
        Select Case mLabelRotation
            Case Is = Horizontal '>> Make LabelSize = TextSize In Case Of Horizontal.
                If mAutoSize Then
                    .Width = .TextWidth(mCaption) * Screen.TwipsPerPixelX
                    .Height = .TextHeight(mCaption) * Screen.TwipsPerPixelY
                    mAutoSize = False
                End If
            Case Is = Vertical '>> Make LabelSize = TextSize In Case Of Vertical.
                If mAutoSize Then
                    .Height = .TextWidth(mCaption) * Screen.TwipsPerPixelY
                    .Width = .TextHeight(mCaption) * Screen.TwipsPerPixelX
                    mAutoSize = False
                End If
        End Select
    End With
    UserControl_Paint
End Property
Public Property Get RightToLeft() As Boolean
    RightToLeft = mRightToLeft
End Property
Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    mRightToLeft = New_RightToLeft
    PropertyChanged "RightToLeft"
    With UserControl
        If mRightToLeft Then
            .RightToLeft = True
        Else
            .RightToLeft = False
        End If
    End With
    UserControl_Paint
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl_Paint
End Property
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    With UserControl
    If .FontBold = True Then '>> See If User Use FontBlod.
        RotateMe.lfWeight = LF_FONTBOLD '>> Create FontBlod >> Weight = 700.
    ElseIf .FontBold = False Then
        RotateMe.lfWeight = LF_FONTREGULAR '>> Create FontRegular.
    End If
    End With
    UserControl_Paint
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
    UserControl_Paint
End Property
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon() = New_MouseIcon
    PropertyChanged "MouseIcon"
    UserControl_Paint
End Property
'............................................................
'>> End Properties
'...........................................................

'............................................................
'>> Start UserControl
'............................................................
Private Sub UserControl_InitProperties()
    'On Local Error Resume Next
    '>> Default Properties.
    '................................................................
    mCaption = Ambient.DisplayName
    mLabelAlignment = Left_Justify
    mAutoSize = False
    mRightToLeft = False
    '................................................................
    mBackColor1 = vbButtonFace
    mBackColor2 = vbWindowBackground
    '................................................................
    mLabelBackStyle = Opaque
    mLabelThemeColor = NoTheme
    mBorderColor1 = vbDesktop
    mBorderColor2 = vbWindowBackground
    mLabelBroderStyle = None
    '................................................................
    mLabelEffects = Default
    mLabelFillStyle = Normal
    mEffectColor = vbHighlightText
    '................................................................
    UserControl.Enabled = True
    UserControl.Font = "Tahoma"
    '................................................................
    mForeColor1 = vbBlack
    mForeColor2 = vbBlue
    mLabelRotation = Horizontal
    '................................................................
    UserControl.Width = 1200
    UserControl.Height = 500
    '................................................................
    UserControl_Paint
End Sub
Private Sub UserControl_Paint()
 On Local Error Resume Next '>> Do Not Remove This Line
                            '>> Because Some Colors No Working In
                            '>> Effects = Draw1_3D Or Draw2_3D
                            ''''''''''''''''''''''''''''''''''''''''
                            ' Run-time error '5':                  '
                            '                                      '
                            '                                      '
                            '                                      '
                            '                                      '
                            ' Invalid procedure call or argument   '
                            '                                      '
                            ''''''''''''''''''''''''''''''''''''''''
    With UserControl
        .AutoRedraw = True
        .ScaleMode = 3 '>> Pixel Mode
        .Cls '>> Clear Before Start
        '................................................................
        '1 >> Set BackColor To Label.
        '................................................................
        ConvertRGB mBackColor1, BkRed(0), BkGreen(0), BkBlue(0)
        ConvertRGB mBackColor2, BkRed(1), BkGreen(1), BkBlue(1)
        Vert(0).Red = Val("&H" & Hex(BkRed(0)) & "00")
        Vert(0).Green = Val("&H" & Hex(BkGreen(0)) & "00")
        Vert(0).Blue = Val("&H" & Hex(BkBlue(0)) & "00")
        Vert(0).Alpha = 0&
        Vert(1).Red = Val("&H" & Hex(BkRed(1)) & "00")
        Vert(1).Green = Val("&H" & Hex(BkGreen(1)) & "00")
        Vert(1).Blue = Val("&H" & Hex(BkBlue(1)) & "00")
        Vert(1).Alpha = 0&
        GRect.UpperLeft = 1
        GRect.LowerRight = 0
        '................................................................
        '2 >> BackStyle,Transparent,Opaque
        '   >> BackFill_HorizontalRight,BackFill_HorizontalLeft
        '    >> BackFill_VerticalDown,BackFill_VerticalUp
        '     >> BackFill_HorizontalCenter,BackFill_HorizontalOut
        '      >> BackFill_VerticalCenter,BackFill_VerticalOut.
        '................................................................
        Select Case mLabelBackStyle
            Case Is = Transparent
                .BackStyle = 0 '>> Make Label Trans.
            Case Is = Opaque
                .BackStyle = 1 '>> Make Label Opaque.
                .BackColor = mBackColor1    '>> The Label Is Single Color.
            Case Is = BackFill_HorizontalRight
                .BackStyle = 1 '>> Make Label Opaque.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        Vert(0).X = 0: Vert(0).Y = .ScaleHeight
                        Vert(1).X = .ScaleWidth: Vert(1).Y = 0
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                    Case Is = Vertical
                        Vert(0).X = 0: Vert(0).Y = .ScaleHeight
                        Vert(1).X = .ScaleWidth: Vert(1).Y = 0
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                End Select
            Case Is = BackFill_HorizontalLeft
                .BackStyle = 1 '>> Make Label Opaque.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        Vert(0).X = .ScaleWidth: Vert(0).Y = 0
                        Vert(1).X = 0: Vert(1).Y = .ScaleHeight
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                    Case Is = Vertical
                        Vert(0).X = .ScaleWidth: Vert(0).Y = 0
                        Vert(1).X = 0: Vert(1).Y = .ScaleHeight
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                End Select
            Case Is = BackFill_VerticalDown
                .BackStyle = 1 '>> Make Label Opaque.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        Vert(0).X = 0: Vert(0).Y = .ScaleHeight
                        Vert(1).X = .ScaleWidth: Vert(1).Y = 0
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                    Case Is = Vertical
                        Vert(0).X = .ScaleWidth: Vert(0).Y = .ScaleHeight
                        Vert(1).X = 0: Vert(1).Y = 0
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                End Select
            Case Is = BackFill_VerticalUp
                .BackStyle = 1 '>> Make Label Opaque.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        Vert(0).X = .ScaleWidth: Vert(0).Y = 0
                        Vert(1).X = 0: Vert(1).Y = .ScaleHeight
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                    Case Is = Vertical
                        Vert(0).X = 0: Vert(0).Y = .ScaleHeight
                        Vert(1).X = .ScaleWidth: Vert(1).Y = 0
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                End Select
            Case Is = BackFill_HorizontalCenter
                .BackStyle = 1 '>> Make Label Opaque.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        '>> From Top To Center.
                        Vert(0).X = 0: Vert(0).Y = .ScaleHeight / 2
                        Vert(1).X = .ScaleWidth: Vert(1).Y = 0
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                        '>> From Center To Bottom.
                        Vert(0).X = 0: Vert(0).Y = .ScaleHeight / 2
                        Vert(1).X = .ScaleWidth: Vert(1).Y = .ScaleHeight
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                    Case Is = Vertical
                        '>> From Left To Center.
                        Vert(0).X = .ScaleWidth / 2: Vert(0).Y = 0
                        Vert(1).X = 0: Vert(1).Y = .ScaleHeight
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                        '>> From Center To Right.
                        Vert(0).X = .ScaleWidth / 2: Vert(0).Y = .ScaleHeight
                        Vert(1).X = .ScaleWidth: Vert(1).Y = 0
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                End Select
            Case Is = BackFill_HorizontalOut
                .BackStyle = 1 '>> Make Label Opaque.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        '>> From Top To Center.
                        Vert(0).X = .ScaleWidth: Vert(0).Y = 0
                        Vert(1).X = 0: Vert(1).Y = .ScaleHeight / 2
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                        '>> From Center To Bottom.
                        Vert(0).X = 0: Vert(0).Y = .ScaleHeight
                        Vert(1).X = .ScaleWidth: Vert(1).Y = .ScaleHeight / 2
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                    Case Is = Vertical
                        '>> From Left To Center.
                        Vert(0).X = 0: Vert(0).Y = 0
                        Vert(1).X = .ScaleWidth / 2: Vert(1).Y = .ScaleHeight
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                        '>> From Center To Right.
                        Vert(0).X = .ScaleWidth: Vert(0).Y = .ScaleHeight
                        Vert(1).X = .ScaleWidth / 2: Vert(1).Y = 0
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                End Select
            Case Is = BackFill_VerticalCenter
                .BackStyle = 1 '>> Make Label Opaque.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        '>> From Left To Center.
                        Vert(0).X = .ScaleWidth / 2: Vert(0).Y = 0
                        Vert(1).X = 0: Vert(1).Y = .ScaleHeight
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                        '>> From Center To Right.
                        Vert(0).X = .ScaleWidth / 2: Vert(0).Y = .ScaleHeight
                        Vert(1).X = .ScaleWidth: Vert(1).Y = 0
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                    Case Is = Vertical
                        '>> From Top To Center.
                        Vert(0).X = 0: Vert(0).Y = .ScaleHeight / 2
                        Vert(1).X = .ScaleWidth: Vert(1).Y = 0
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                        '>> From Center To Bottom.
                        Vert(0).X = 0: Vert(0).Y = .ScaleHeight / 2
                        Vert(1).X = .ScaleWidth: Vert(1).Y = .ScaleHeight
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                 End Select
            Case Is = BackFill_VerticalOut
                .BackStyle = 1 '>> Make Label Opaque.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        '>> From Left To Center.
                        Vert(0).X = 0: Vert(0).Y = 0
                        Vert(1).X = .ScaleWidth / 2: Vert(1).Y = .ScaleHeight
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                        '>> From Center To Right.
                        Vert(0).X = .ScaleWidth: Vert(0).Y = .ScaleHeight
                        Vert(1).X = .ScaleWidth / 2: Vert(1).Y = 0
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                    Case Is = Vertical
                        '>> From Top To Center.
                        Vert(0).X = .ScaleWidth: Vert(0).Y = 0
                        Vert(1).X = 0: Vert(1).Y = .ScaleHeight / 2
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                        '>> From Center To Bottom.
                        Vert(0).X = 0: Vert(0).Y = .ScaleHeight
                        Vert(1).X = .ScaleWidth: Vert(1).Y = .ScaleHeight / 2
                        GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                End Select
        End Select
        '................................................................
        '3 >> LabelAlignment,Get Size And Postion Text,Left Text,Center
        '   >> Text,Right Text.
        '................................................................
        '>> Make New Font To Label.
        RotateMe.lfHeight = (.FontSize * -20) / Screen.TwipsPerPixelY
        RotateMe.lfItalic = .FontItalic
        RotateMe.lfUnderline = .FontUnderline
        RotateMe.lfStrikeOut = .FontStrikethru
        RotateMe.lfFaceName = .FontName & Chr$(0)
        DeleteObject CreateFontIndirect(RotateMe)
        '................................................................
        If Len(mCaption) < 0 Then Exit Sub
        Select Case mLabelRotation
            Case Is = Horizontal '>> Alignment Text At Horizontal Font.
                RotateMe.lfEscapement = LF_ESCAPEMENT_H '>> Rotation = 0
                SelectObject .hdc, CreateFontIndirect(RotateMe)
                Select Case mLabelAlignment
                    Case Is = Left_Justify
                        GetTextExtentPoint32 .hdc, mCaption, Len(mCaption), CaptionPos(1)
                        CaptionPos(0).X = (.ScaleWidth - CaptionPos(1).X) / .ScaleWidth + 2
                        CaptionPos(0).Y = (.ScaleHeight - CaptionPos(1).Y) / 2
                    Case Is = Center_Justify
                        GetTextExtentPoint32 .hdc, mCaption, Len(mCaption), CaptionPos(1)
                        CaptionPos(0).X = (.ScaleWidth - CaptionPos(1).X) / 2
                        CaptionPos(0).Y = (.ScaleHeight - CaptionPos(1).Y) / 2
                    Case Is = Right_Justify
                        GetTextExtentPoint32 .hdc, mCaption, Len(mCaption), CaptionPos(1)
                        CaptionPos(0).X = (.ScaleWidth - CaptionPos(1).X) - 2
                        CaptionPos(0).Y = (.ScaleHeight - CaptionPos(1).Y) / 2
                End Select
            Case Is = Vertical '>> Alignment Text At Vertical Font.
                RotateMe.lfEscapement = LF_ESCAPEMENT_V '>> Rotation = 90
                SelectObject .hdc, CreateFontIndirect(RotateMe)
                Select Case mLabelAlignment
                    Case Is = Left_Justify
                        GetTextExtentPoint32 .hdc, mCaption, Len(mCaption), CaptionPos(1)
                        CaptionPos(0).X = (.ScaleWidth - CaptionPos(1).Y) / 2
                        CaptionPos(0).Y = .ScaleHeight - 3
                    Case Is = Center_Justify
                        GetTextExtentPoint32 .hdc, mCaption, Len(mCaption), CaptionPos(1)
                        CaptionPos(0).X = (.ScaleWidth - CaptionPos(1).Y) / 2
                        CaptionPos(0).Y = (.ScaleHeight - CaptionPos(1).X) / 2 + TextWidth(mCaption)
                    Case Is = Right_Justify
                        GetTextExtentPoint32 .hdc, mCaption, Len(mCaption), CaptionPos(1)
                        CaptionPos(0).X = (.ScaleWidth - CaptionPos(1).Y) / 2
                        CaptionPos(0).Y = TextWidth(mCaption) + 2
                End Select
        End Select
        '................................................................
        '4 >> Effects,Default,Raised,Sunken,Outline,Shadow,Draw1_3D
        '   >> Draw2_3D,RaisedShadow,SunkenShadow,OutlineShadow.
        '................................................................
        If Enabled Then
            .MaskColor = vb3DDKShadow '>> Need To Work Raised And Sunken Font.
            .ForeColor = mForeColor1 '>> Converted ForeColor To TextColor.
            Select Case mLabelEffects
                Case Is = Default
                    SetTextColor .hdc, BlendColor(mBackColor1, mForeColor1, 100)
                    TextOutA .hdc, CaptionPos(0).X, CaptionPos(0).Y, mCaption, Len(mCaption)
                Case Is = Raised
                    Visual.X = -1 '>> -1 Pixel From Down To Top
                    Visual.Y = -1 '>> -1 Pixel From Right To Left
                    SetTextColor .hdc, BlendColor(.MaskColor, mForeColor1, 25)
                    TextOutW .hdc, CaptionPos(0).X - Visual.X, CaptionPos(0).Y - Visual.Y, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 75)
                    TextOutW .hdc, CaptionPos(0).X + Visual.X, CaptionPos(0).Y + Visual.Y, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, mForeColor1, 100)
                    TextOutA .hdc, CaptionPos(0).X, CaptionPos(0).Y, mCaption, Len(mCaption)
                Case Is = Sunken
                    Visual.X = 1 '>> +1 Pixel From Top To Down.
                    Visual.Y = 1 '>> +1 Pixel From Left To Right.
                    SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 75)
                    TextOutW .hdc, CaptionPos(0).X + Visual.X, CaptionPos(0).Y + Visual.Y, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(.MaskColor, mForeColor1, 25)
                    TextOutW .hdc, CaptionPos(0).X - Visual.X, CaptionPos(0).Y - Visual.Y, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, mForeColor1, 100)
                    TextOutA .hdc, CaptionPos(0).X, CaptionPos(0).Y, mCaption, Len(mCaption)
                Case Is = Outline
                    For Visual.X = -1 To 1 '>> -1 Pixel And +1 Pixel From Top To Down
                        For Visual.Y = -1 To 1 '>> -1 Pixel And +1 Pixel From Left To Right
                            SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 100)
                            TextOutW .hdc, CaptionPos(0).X + Visual.X, CaptionPos(0).Y + Visual.Y, mCaption, Len(mCaption)
                            SetTextColor .hdc, BlendColor(mBackColor1, mForeColor1, 100)
                            TextOutA .hdc, CaptionPos(0).X, CaptionPos(0).Y, mCaption, Len(mCaption)
                        Next Visual.Y
                    Next Visual.X
                Case Is = Shadow
                    '>> -1 Pixel From Center To Left.
                    SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 50)
                    TextOutA .hdc, CaptionPos(0).X - 1, CaptionPos(0).Y, mCaption, Len(mCaption)
                    '>> 4 To 1 - Pixel From Top To Down.
                    '>> 6 To 1 -  Pixel From Center To Right.
                    SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 10)
                    TextOutA .hdc, CaptionPos(0).X + 6, CaptionPos(0).Y + 4, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 30)
                    TextOutA .hdc, CaptionPos(0).X + 4, CaptionPos(0).Y + 3, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 50)
                    TextOutA .hdc, CaptionPos(0).X + 2, CaptionPos(0).Y + 2, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 70)
                    TextOutA .hdc, CaptionPos(0).X + 1, CaptionPos(0).Y + 1, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, mForeColor1, 100)
                    TextOutA .hdc, CaptionPos(0).X, CaptionPos(0).Y, mCaption, Len(mCaption)
                Case Is = RaisedShadow
                    '>> -1 Pixel From Down To Top And From Right To Left
                    SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 75)
                    TextOutW .hdc, CaptionPos(0).X - 1, CaptionPos(0).Y - 1, mCaption, Len(mCaption)
                    '>> 4 To 1 - Pixel From Top To Down.
                    '>> 6 To 1 -  Pixel From Center To Right.
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 10)
                    TextOutA .hdc, CaptionPos(0).X + 6, CaptionPos(0).Y + 4, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 30)
                    TextOutA .hdc, CaptionPos(0).X + 4, CaptionPos(0).Y + 3, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 50)
                    TextOutA .hdc, CaptionPos(0).X + 2, CaptionPos(0).Y + 2, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 70)
                    TextOutA .hdc, CaptionPos(0).X + 1, CaptionPos(0).Y + 1, mCaption, Len(mCaption)
                    '>> Fore Color And Fore Text.
                    SetTextColor .hdc, BlendColor(mBackColor1, mForeColor1, 100)
                    TextOutA .hdc, CaptionPos(0).X, CaptionPos(0).Y, mCaption, Len(mCaption)
                Case Is = SunkenShadow
                    '>> 4 To 2 - Pixel From Top To Down.
                    '>> 6 To 2 -  Pixel From Center To Right.
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 10)
                    TextOutA .hdc, CaptionPos(0).X + 6, CaptionPos(0).Y + 4, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 30)
                    TextOutA .hdc, CaptionPos(0).X + 4, CaptionPos(0).Y + 3, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 50)
                    TextOutA .hdc, CaptionPos(0).X + 2, CaptionPos(0).Y + 2, mCaption, Len(mCaption)
                    '>> +1 Pixel From Down To Top And From Right To Left
                    SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 75)
                    TextOutW .hdc, CaptionPos(0).X + 1, CaptionPos(0).Y + 1, mCaption, Len(mCaption)
                    '>> -1 Pixel From Down To Top And From Right To Left
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 70)
                    TextOutA .hdc, CaptionPos(0).X - 1, CaptionPos(0).Y - 1, mCaption, Len(mCaption)
                    '>> Fore Color And Fore Text.
                    SetTextColor .hdc, BlendColor(mBackColor1, mForeColor1, 100)
                    TextOutA .hdc, CaptionPos(0).X, CaptionPos(0).Y, mCaption, Len(mCaption)
                Case Is = OutlineShadow
                    '>> 5 To 2 - Pixel From Top To Down.
                    '>> 8 To 2 -  Pixel From Center To Right.
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 10)
                    TextOutA .hdc, CaptionPos(0).X + 8, CaptionPos(0).Y + 5, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 30)
                    TextOutA .hdc, CaptionPos(0).X + 6, CaptionPos(0).Y + 4, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 50)
                    TextOutA .hdc, CaptionPos(0).X + 3, CaptionPos(0).Y + 3, mCaption, Len(mCaption)
                    SetTextColor .hdc, BlendColor(mBackColor1, .MaskColor, 70)
                    TextOutA .hdc, CaptionPos(0).X + 2, CaptionPos(0).Y + 2, mCaption, Len(mCaption)
                    '>> -1 Pixel From Center To Left.
                    SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 50)
                    TextOutA .hdc, CaptionPos(0).X - 2, CaptionPos(0).Y, mCaption, Len(mCaption)
                    For Visual.X = -1 To 1 '>> -1 Pixel And +1 Pixel From Top To Down
                        For Visual.Y = -1 To 1 '>> -1 Pixel And +1 Pixel From Left To Right
                            SetTextColor .hdc, BlendColor(mBackColor1, mEffectColor, 100)
                            TextOutW .hdc, CaptionPos(0).X + Visual.X, CaptionPos(0).Y + Visual.Y, mCaption, Len(mCaption)
                            SetTextColor .hdc, BlendColor(mBackColor1, mForeColor1, 100)
                            TextOutA .hdc, CaptionPos(0).X, CaptionPos(0).Y, mCaption, Len(mCaption)
                        Next Visual.Y
                    Next Visual.X
                    '>> Fore Color And Fore Text.
                    SetTextColor .hdc, BlendColor(mBackColor1, mForeColor1, 100)
                    TextOutA .hdc, CaptionPos(0).X, CaptionPos(0).Y, mCaption, Len(mCaption)
                Case Is = Draw1_3D
                    ConvertRGB mForeColor1, R(0), G(0), B(0)
                    ConvertRGB mEffectColor, R(1), G(1), B(1)
                    TotalRed(0) = (R(0) - R(1)) / 12
                    TotalGreen(0) = (G(0) - G(1)) / 12
                    TotalBlue(0) = (B(0) - B(1)) / 12
                    FrRed = R(1) + TotalRed(1)
                    FrGreen = G(1) + TotalGreen(1)
                    FrBlue = B(1) + TotalBlue(1)
                    For Visual.X = 0 To 12 '>> 12 Pixel From Left-Top To Center.
                        .ForeColor = RGB(FrRed + Visual.X * TotalRed(0), FrGreen + Visual.X * TotalGreen(0), FrBlue + Visual.X * TotalBlue(0))
                        '>> Alignment The Text If Drwing 3D Effect.
                        Select Case mLabelRotation
                            Case Is = Horizontal
                                Select Case mLabelAlignment
                                    Case Is = Left_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X, CaptionPos(0).Y + Visual.X - 6, mCaption, Len(mCaption)
                                    Case Is = Center_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X - 12, CaptionPos(0).Y + Visual.X - 6, mCaption, Len(mCaption)
                                    Case Is = Right_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X - 12, CaptionPos(0).Y + Visual.X - 6, mCaption, Len(mCaption)
                                End Select
                            Case Is = Vertical
                                Select Case mLabelAlignment
                                    Case Is = Left_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X - 6, CaptionPos(0).Y + Visual.X - 12, mCaption, Len(mCaption)
                                    Case Is = Center_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X - 6, CaptionPos(0).Y + Visual.X - 12, mCaption, Len(mCaption)
                                    Case Is = Right_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X - 6, CaptionPos(0).Y + Visual.X, mCaption, Len(mCaption)
                                End Select
                        End Select
                        R(1) = R(1) + TotalRed(1)
                        G(1) = G(1) + TotalGreen(1)
                        B(1) = B(1) + TotalBlue(1)
                    Next Visual.X
                Case Is = Draw2_3D
                    ConvertRGB mForeColor1, R(1), G(1), B(1)
                    ConvertRGB mEffectColor, R(0), G(0), B(0)
                    TotalRed(0) = (R(0) - R(1)) / 12
                    TotalGreen(0) = (G(0) - G(1)) / 12
                    TotalBlue(0) = (B(0) - B(1)) / 12
                    FrRed = R(1) + TotalRed(1)
                    FrGreen = G(1) + TotalGreen(1)
                    FrBlue = B(1) + TotalBlue(1)
                    For Visual.X = 0 To 12 '>> 12 Pixel From Left-Top To Center.
                        .ForeColor = RGB(FrRed + Visual.X * TotalRed(0), FrGreen + Visual.X * TotalGreen(0), FrBlue + Visual.X * TotalBlue(0))
                        '>> Alignment The Text If Drwing 3D Effect.
                        Select Case mLabelRotation
                            Case Is = Horizontal
                                Select Case mLabelAlignment
                                    Case Is = Left_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X, CaptionPos(0).Y + Visual.X - 6, mCaption, Len(mCaption)
                                    Case Is = Center_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X - 12, CaptionPos(0).Y + Visual.X - 6, mCaption, Len(mCaption)
                                    Case Is = Right_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X - 12, CaptionPos(0).Y + Visual.X - 6, mCaption, Len(mCaption)
                                End Select
                            Case Is = Vertical
                                Select Case mLabelAlignment
                                    Case Is = Left_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X - 6, CaptionPos(0).Y + Visual.X - 12, mCaption, Len(mCaption)
                                    Case Is = Center_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X - 6, CaptionPos(0).Y + Visual.X - 12, mCaption, Len(mCaption)
                                    Case Is = Right_Justify
                                        TextOutA .hdc, CaptionPos(0).X + Visual.X - 6, CaptionPos(0).Y + Visual.X, mCaption, Len(mCaption)
                                End Select
                        End Select
                        R(1) = R(1) + TotalRed(1)
                        G(1) = G(1) + TotalGreen(1)
                        B(1) = B(1) + TotalBlue(1)
                    Next Visual.X
            End Select
        Else '>> If Enabled = False Then >> Draw Disabled Text Sunken Look
            .ForeColor = vb3DHighlight
            TextOutW .hdc, CaptionPos(0).X + 1, CaptionPos(0).Y + 1, mCaption, Len(mCaption)
            .ForeColor = vb3DShadow
            TextOutA .hdc, CaptionPos(0).X, CaptionPos(0).Y, mCaption, Len(mCaption)
        End If
        '................................................................
        '5 >> Set ForeColor To LabelText.
        '   >> FillStyle,SingleColour,TextFill_HorizontalRight
        '    >> TextFill_HorizontalLeft,TextFill_VerticalDown
        '     >> TextFill_VerticalUp,TextFill_HorizontalCenter
        '      >> TextFill_HorizontalOut,TextFill_VerticalCenter
        '       >> TextFill_VerticalOut.
        '................................................................
        Select Case mLabelFillStyle
                'Case Is = Normal
                '>> Go To The Event >> mLabelEffects = Default
            Case Is = TextFill_HorizontalRight '>> GradientFill Horizontal Right Inside Text.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                    Case Is = Vertical
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                End Select
            Case Is = TextFill_HorizontalLeft '>> GradientFill Horizontal Left Inside Text.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                    Case Is = Vertical
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                End Select
            Case Is = TextFill_VerticalDown '>> GradientFill Vertical Down Inside Text.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                    Case Is = Vertical
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                End Select
            Case Is = TextFill_VerticalUp '>> GradientFill Vertical Up Inside Text.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                    Case Is = Vertical
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                End Select
            Case Is = TextFill_HorizontalCenter '>> GradientFill Horizontal Center Inside Text.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight / 2 '>> From Top To Center.
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = .ScaleHeight / 2 To .ScaleHeight - 1 '>> From Center To Bottom.
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                    Case Is = Vertical
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = .ScaleWidth / 2 To .ScaleWidth - 1 '>>From Center To Right.
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = 1 To .ScaleWidth / 2 '>> From Left To Center.
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                End Select
            Case Is = TextFill_HorizontalOut '>> GradientFill Horizontal Out Inside Text.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight / 2 '>> From Top To Center.
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = .ScaleHeight / 2 To .ScaleHeight - 1 '>> From Center To Bottom.
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                    Case Is = Vertical
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = .ScaleWidth / 2 To .ScaleWidth - 1 '>> From Center To Right.
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = 1 To .ScaleWidth / 2 '>> From Left To Center.
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                End Select
            Case Is = TextFill_VerticalCenter '>> GradientFill Vertical Center Inside Text.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = .ScaleWidth / 2 To .ScaleWidth - 1 '>> From Center To Right.
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = 1 To .ScaleWidth / 2 '>> From Left To Center.
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                    Case Is = Vertical
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight / 2 '>> From Top To Center.
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = .ScaleHeight / 2 To .ScaleHeight - 1 '>> From Center To Bottom.
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                End Select
            Case Is = TextFill_VerticalOut '>> GradientFill Vertical Out Inside Text.
                Select Case mLabelRotation
                    Case Is = Horizontal
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = .ScaleWidth / 2 To .ScaleWidth - 1 '>> From Center To Right.
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = 1 To .ScaleWidth / 2 '>> From Left To Center.
                            For Visual.Y = 1 To .ScaleHeight - 1
                                TotalRed(0) = (R(0) - R(1)) / .ScaleWidth * Visual.X
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleWidth * Visual.X
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleWidth * Visual.X
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                    Case Is = Vertical
                        ConvertRGB mForeColor1, R(0), G(0), B(0)
                        ConvertRGB mForeColor2, R(1), G(1), B(1)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = 1 To .ScaleHeight / 2 '>> From Top To Center.
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                        ConvertRGB mForeColor1, R(1), G(1), B(1)
                        ConvertRGB mForeColor2, R(0), G(0), B(0)
                        For Visual.X = 1 To .ScaleWidth - 1
                            For Visual.Y = .ScaleHeight / 2 To .ScaleHeight - 1 '>> From Center To Bottom.
                                TotalRed(0) = (R(0) - R(1)) / .ScaleHeight * Visual.Y
                                TotalGreen(0) = (G(0) - G(1)) / .ScaleHeight * Visual.Y
                                TotalBlue(0) = (B(0) - B(1)) / .ScaleHeight * Visual.Y
                                FrRed = R(1) + TotalRed(0)
                                FrGreen = G(1) + TotalGreen(0)
                                FrBlue = B(1) + TotalBlue(0)
                                If GetPixel(.hdc, Visual.X, Visual.Y) = .ForeColor Then
                                    SetPixel .hdc, Visual.X, Visual.Y, RGB(FrRed, FrGreen, FrBlue)
                                    R(1) = R(1) + TotalRed(1)
                                    G(1) = G(1) + TotalGreen(1)
                                    B(1) = B(1) + TotalBlue(1)
                                End If
                            Next Visual.Y
                        Next Visual.X
                End Select
        End Select
        .MaskColor = .BackColor '>> Change BackStyle From Opaque To Transparent
        '................................................................
        '6 >> Draw Line's Style To Label.
        '................................................................
        ConvertRGB mBorderColor1, BkRed(0), BkGreen(0), BkBlue(0)
        ConvertRGB mBorderColor2, BkRed(1), BkGreen(1), BkBlue(1)
        Vert(0).Red = Val("&H" & Hex(BkRed(0)) & "00")
        Vert(0).Green = Val("&H" & Hex(BkGreen(0)) & "00")
        Vert(0).Blue = Val("&H" & Hex(BkBlue(0)) & "00")
        Vert(0).Alpha = 0&
        Vert(1).Red = Val("&H" & Hex(BkRed(1)) & "00")
        Vert(1).Green = Val("&H" & Hex(BkGreen(1)) & "00")
        Vert(1).Blue = Val("&H" & Hex(BkBlue(1)) & "00")
        Vert(1).Alpha = 0&
        GRect.UpperLeft = 1
        GRect.LowerRight = 0
        Select Case mLabelBroderStyle
           'Case Is = None
            Case Is = Broder_Classic '>> Draw Broder Classic.
                .ForeColor = BlendColor(mBackColor1, BDR_RAISEDOUTER, 95)
                MoveToEx .hdc, 0, .ScaleHeight - 1, Lines
                LineTo .hdc, .ScaleWidth - 1, .ScaleHeight - 1
                LineTo .hdc, .ScaleWidth - 1, 0
                .ForeColor = BlendColor(mBackColor1, BDR_SUNKENOUTER1, 95)
                MoveToEx .hdc, .ScaleWidth - 1, 0, Lines
                LineTo .hdc, 0, 0
                LineTo .hdc, 0, .ScaleHeight - 1
                .ForeColor = BlendColor(BDR_SUNKENOUTER2, mBackColor1, 5)
                MoveToEx .hdc, .ScaleWidth - 2, 1, Lines
                LineTo .hdc, 1, 1
                LineTo .hdc, 1, .ScaleHeight - 2
            Case Is = Broder_Sunken '>> Draw Broder Sunken.
                .ForeColor = BlendColor(mBackColor1, BDR_RAISEDOUTER, 95)
                MoveToEx .hdc, 0, .ScaleHeight - 1, Lines
                LineTo .hdc, .ScaleWidth - 1, .ScaleHeight - 1
                LineTo .hdc, .ScaleWidth - 1, 0
                .ForeColor = BlendColor(BDR_SUNKENOUTER2, mBackColor1, 5)
                MoveToEx .hdc, .ScaleWidth - 1, 0, Lines
                LineTo .hdc, 0, 0
                LineTo .hdc, 0, .ScaleHeight - 1
            Case Is = Broder_Raised '>> Draw Broder Raised.
                .ForeColor = BlendColor(BDR_SUNKENOUTER2, mBackColor1, 5)
                MoveToEx .hdc, 0, .ScaleHeight - 1, Lines
                LineTo .hdc, .ScaleWidth - 1, .ScaleHeight - 1
                LineTo .hdc, .ScaleWidth - 1, 0
                .ForeColor = BlendColor(mBackColor1, BDR_RAISEDOUTER, 95)
                MoveToEx .hdc, .ScaleWidth - 1, 0, Lines
                LineTo .hdc, 0, 0
                LineTo .hdc, 0, .ScaleHeight - 1
            Case Is = Broder_3DInside ' >> Draw Broder 3D Inside.
                .ForeColor = BlendColor(BDR_SUNKENOUTER1, mBackColor1, 5)
                MoveToEx .hdc, 0, .ScaleHeight - 1, Lines
                LineTo .hdc, 0, 0
                LineTo .hdc, .ScaleWidth - 1, 0
                MoveToEx .hdc, .ScaleWidth - 2, 1, Lines
                LineTo .hdc, .ScaleWidth - 2, .ScaleHeight - 2
                LineTo .hdc, 2, .ScaleHeight - 2
                .ForeColor = BlendColor(mBackColor1, BDR_RAISEDOUTER, 95)
                MoveToEx .hdc, .ScaleWidth - 1, 1, Lines
                LineTo .hdc, 1, 1
                LineTo .hdc, 1, .ScaleHeight - 1
                LineTo .hdc, .ScaleWidth - 1, .ScaleHeight - 1
                LineTo .hdc, .ScaleWidth - 1, 1
            Case Is = Broder_3DOutside ' >> Draw Broder 3D Outside.
                .ForeColor = BlendColor(mBackColor1, BDR_RAISEDOUTER, 95)
                MoveToEx .hdc, 0, .ScaleHeight - 1, Lines
                LineTo .hdc, 0, 0
                LineTo .hdc, .ScaleWidth - 1, 0
                MoveToEx .hdc, .ScaleWidth - 2, 1, Lines
                LineTo .hdc, .ScaleWidth - 2, .ScaleHeight - 2
                LineTo .hdc, 2, .ScaleHeight - 2
                .ForeColor = BlendColor(BDR_SUNKENOUTER1, mBackColor1, 5)
                MoveToEx .hdc, .ScaleWidth - 1, 1, Lines
                LineTo .hdc, 1, 1
                LineTo .hdc, 1, .ScaleHeight - 1
                LineTo .hdc, .ScaleWidth - 1, .ScaleHeight - 1
                LineTo .hdc, .ScaleWidth - 1, 1
            Case Is = Broder ' >> Draw Single Color Broder.
                .ForeColor = BlendColor(mBackColor1, mBorderColor1, 95)
                LineTo .hdc, 0, .ScaleHeight - 1
                LineTo .hdc, .ScaleWidth - 1, .ScaleHeight - 1
                LineTo .hdc, .ScaleWidth - 1, 0
                LineTo .hdc, 0, 0
             Case Is = BroderFill_VH ' >> Draw Vertical And Horizontal Gradient Fill Broder.
                '>> From Down To Top.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = .ScaleWidth - 2: Vert(1).Y = 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Right To Left.
                Vert(0).X = 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = .ScaleWidth - 1: Vert(1).Y = .ScaleHeight - 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                '>> From Top To Down.
                Vert(0).X = 1: Vert(0).Y = 1
                Vert(1).X = 2: Vert(1).Y = .ScaleHeight - 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Left To Right.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = 1
                Vert(1).X = 1: Vert(1).Y = 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
             Case Is = BroderFill_HV ' >> Draw Horizontal And Vertical Gradient Fill Broder.
                '>> From Top To Down.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = 1
                Vert(1).X = .ScaleWidth - 2: Vert(1).Y = .ScaleHeight - 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Right To Left.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = 1: Vert(1).Y = .ScaleHeight - 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                '>> From Down To Top.
                Vert(0).X = 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = 2: Vert(1).Y = 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Left To Right.
                Vert(0).X = 1: Vert(0).Y = 1
                Vert(1).X = .ScaleWidth - 1: Vert(1).Y = 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
             Case Is = BroderFill_CornerIn
                '>> From Down To Top.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = .ScaleWidth - 2: Vert(1).Y = 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Right To Left.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = 1: Vert(1).Y = .ScaleHeight - 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                '>> From Top To Down.
                Vert(0).X = 1: Vert(0).Y = 1
                Vert(1).X = 2: Vert(1).Y = .ScaleHeight - 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Left To Right.
                Vert(0).X = 1: Vert(0).Y = 1
                Vert(1).X = .ScaleWidth - 1: Vert(1).Y = 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
             Case Is = BroderFill_CornerOut
                '>> From Top To Down.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = 1
                Vert(1).X = .ScaleWidth - 2: Vert(1).Y = .ScaleHeight - 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Left To Right.
                Vert(0).X = 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = .ScaleWidth - 1: Vert(1).Y = .ScaleHeight - 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                '>> From Down To Top.
                Vert(0).X = 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = 2: Vert(1).Y = 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Right To Left.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = 1
                Vert(1).X = 1: Vert(1).Y = 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
              Case Is = BroderFill_HorizontalCenter
                '>> From Top To Center.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = .ScaleHeight / 2
                Vert(1).X = .ScaleWidth - 2: Vert(1).Y = 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Center To Down.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = .ScaleHeight / 2
                Vert(1).X = .ScaleWidth - 2: Vert(1).Y = .ScaleHeight - 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Center To Right.
                Vert(0).X = .ScaleWidth / 2: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = .ScaleWidth - 1: Vert(1).Y = .ScaleHeight - 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                '>> From Center To Left.
                Vert(0).X = .ScaleWidth / 2: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = 1: Vert(1).Y = .ScaleHeight - 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                '>> From Center To Top.
                Vert(0).X = 1: Vert(0).Y = .ScaleHeight / 2
                Vert(1).X = 2: Vert(1).Y = 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Center To Down.
                Vert(0).X = 1: Vert(0).Y = .ScaleHeight / 2
                Vert(1).X = 2: Vert(1).Y = .ScaleHeight - 1
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Center To Right.
                Vert(0).X = .ScaleWidth / 2: Vert(0).Y = 1
                Vert(1).X = .ScaleWidth - 1: Vert(1).Y = 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                '>> From Center To Left.
                Vert(0).X = .ScaleWidth / 2: Vert(0).Y = 1
                Vert(1).X = 1: Vert(1).Y = 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
             Case Is = BroderFill_HorizontalOut
                '>> From Tpo To Center.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = 1
                Vert(1).X = .ScaleWidth - 2: Vert(1).Y = .ScaleHeight / 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Down To Center.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = .ScaleWidth - 2: Vert(1).Y = .ScaleHeight / 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Right To Center.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = .ScaleWidth / 2: Vert(1).Y = .ScaleHeight - 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                '>> From Left To Center.
                Vert(0).X = 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = .ScaleWidth / 2: Vert(1).Y = .ScaleHeight - 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                '>> From Tpo To Center.
                Vert(0).X = 1: Vert(0).Y = 1
                Vert(1).X = 2: Vert(1).Y = .ScaleHeight / 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Down To Center.
                Vert(0).X = 1: Vert(0).Y = .ScaleHeight - 1
                Vert(1).X = 2: Vert(1).Y = .ScaleHeight / 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_V
                '>> From Right To Center.
                Vert(0).X = .ScaleWidth - 1: Vert(0).Y = 1
                Vert(1).X = .ScaleWidth / 2: Vert(1).Y = 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
                '>> From Left To Center.
                Vert(0).X = 1: Vert(0).Y = 1
                Vert(1).X = .ScaleWidth / 2: Vert(1).Y = 2
                GradientFill .hdc, Vert(0), 2, GRect, 1, GRADIENT_FILL_RECT_H
        End Select
        .MaskPicture = .Image '>> Change BackStyle From Transparent To Opaque When.
        .Refresh '>> AutoRedraw = True
    End With
End Sub
Private Sub UserControl_Resize()
    Call UserControl_Paint
End Sub
Private Sub UserControl_Show()
    Call UserControl_Paint
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'On Local Error Resume Next
    With PropBag '>> Load Saved Property Values

        mLabelAlignment = .ReadProperty("Alignment", Left_Justify)
        mAutoSize = .ReadProperty("AutoSize", False)
        mRightToLeft = .ReadProperty("RightToLeft", False)
        '................................................................
        mBackColor1 = .ReadProperty("BackColor1", vbButtonFace)
        mBackColor2 = .ReadProperty("BackColor2", vbWindowBackground)
        '................................................................
        mLabelBackStyle = .ReadProperty("BackStyle", Opaque)
        mLabelThemeColor = .ReadProperty("ThemeColor", NoTheme)
        mBorderColor1 = .ReadProperty("BorderColor1", vbDesktop)
        mBorderColor2 = .ReadProperty("BorderColor2", vbWindowBackground)
        mLabelBroderStyle = .ReadProperty("BroderStyle", None)
        '................................................................
        mCaption = .ReadProperty("Caption", UserControl.Name)
        mLabelEffects = .ReadProperty("Effects", Default)
        mLabelFillStyle = .ReadProperty("FillStyle", Normal)
        mEffectColor = .ReadProperty("EffectColor", vbHighlightText)
        '................................................................
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set Font = .ReadProperty("Font", Ambient.Font)
        '................................................................
        mForeColor1 = .ReadProperty("ForeColor1", vbBlack)
        mForeColor2 = .ReadProperty("ForeColor2", vbBlue)
        '................................................................
        mLabelRotation = .ReadProperty("Rotation", Horizontal)
        '................................................................
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)

    End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'On Local Error Resume Next
    With PropBag '>> Save Property Values

        .WriteProperty "Alignment", mLabelAlignment, Left_Justify
        .WriteProperty "AutoSize", mAutoSize, False
        .WriteProperty "RightToLeft", mRightToLeft, False
        '................................................................
        .WriteProperty "BackColor1", mBackColor1, vbButtonFace
        .WriteProperty "BackColor2", mBackColor2, vbWindowBackground
        '................................................................
        .WriteProperty "BackStyle", mLabelBackStyle, Opaque
        .WriteProperty "ThemeColor", mLabelThemeColor, NoTheme
        .WriteProperty "BorderColor1", mBorderColor1, vbBlack
        .WriteProperty "BorderColor2", mBorderColor2, vbRed
        .WriteProperty "BroderStyle", mLabelBroderStyle, None
        '................................................................
        .WriteProperty "Caption", mCaption, UserControl.Name
        .WriteProperty "Effects", mLabelEffects, Default
        .WriteProperty "FillStyle", mLabelFillStyle, Normal
        .WriteProperty "EffectColor", mEffectColor, vbHighlightText
        '................................................................
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        '................................................................
        .WriteProperty "ForeColor1", mForeColor1, vbBlack
        .WriteProperty "ForeColor2", mForeColor2, vbBlue
        '................................................................
        .WriteProperty "Rotation", mLabelRotation, Horizontal
        '................................................................
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "MousePointer", UserControl.MousePointer, 0
    End With
End Sub
'............................................................
'>> End UserControl.
'............................................................

'............................................................
'>> Convert From RGB TO Long >> VB Color's.
'............................................................
Private Sub ConvertRGB(ByVal Color As Long, R, G, B As Long)
    TranslateColor Color, 0, Color '>> To Accept Windows SyStem Color's.
    R = Color And vbRed
    G = (Color And vbGreen) / 256  '>> VbRed
    B = (Color And vbBlue) / 65536 '>> VbBlack
End Sub
'............................................................
'>> Blend Tow Color's >> Color1 + Color2 Percentage Of 100.
'............................................................
Private Function BlendColor(ByVal Color1 As OLE_COLOR, Color2 As OLE_COLOR, Percentage)
    ConvertRGB Color1, R(0), G(0), B(0)
    ConvertRGB Color2, R(1), G(1), B(1)
    If Percentage < 0 Then Percentage = 0 '>> From 0 To 100 Only.
    If Percentage > 100 Then Percentage = 100 '>> From 100 To 0 only.
    R(2) = R(0) + ((R(1) - R(0)) / 100) * Percentage
    G(2) = G(0) + ((G(1) - G(0)) / 100) * Percentage
    B(2) = B(0) + ((B(1) - B(0)) / 100) * Percentage
    BlendColor = RGB(R(2), G(2), B(2))
End Function

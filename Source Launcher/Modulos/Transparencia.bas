Attribute VB_Name = "Transparencia"
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bDefaut As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE As Long = (-20)
Private Const LWA_COLORKEY As Long = &H1
Private Const LWA_Defaut As Long = &H2
Private Const WS_EX_LAYERED As Long = &H80000

'
Public Function Transparency(ByVal hWnd As Long, Optional ByVal Col As Long = vbBlack, _
                             Optional ByVal PcTransp As Byte = 255, Optional ByVal TrMode As Boolean = True) As Boolean
' Return : True if there is no error.
' hWnd   : hWnd of the window to make transparent
' Col : Color to make transparent if TrMode=False
' PcTransp  : 0 Ã  255 >> 0 = transparent  -:- 255 = Opaque
    Dim DisplayStyle As Long

    VoirStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    If DisplayStyle <> (DisplayStyle Or WS_EX_LAYERED) Then
        DisplayStyle = (DisplayStyle Or WS_EX_LAYERED)
        Call SetWindowLong(hWnd, GWL_EXSTYLE, DisplayStyle)
    End If
    Transparency = (SetLayeredWindowAttributes(hWnd, Col, PcTransp, IIf(TrMode, LWA_COLORKEY Or LWA_Defaut, LWA_COLORKEY)) <> 0)


    If Not Err.Number = 0 Then Err.Clear
End Function

Public Sub ActiveTransparency(M As Form, d As Boolean, F As Boolean, _
                              T_Transparency As Integer, Optional Color As Long)
    Dim B As Boolean
    If d And F Then
        'Makes color (here the background color of the shape) transparent
        'upon value of T_Transparency
        B = Transparency(M.hWnd, Color, T_Transparency, False)
    ElseIf d Then
        'Makes form, including all components, transparent
        'upon value of T_Transparency
        B = Transparency(M.hWnd, 0, T_Transparency, True)
    Else
        'Restores the form opaque.
        B = Transparency(M.hWnd, , 255, True)
    End If
End Sub




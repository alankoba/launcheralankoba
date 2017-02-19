Attribute VB_Name = "Converssões"
Function SetBytes(Bytes) As String

    On Error GoTo hell

    If Bytes >= 1073741824 Then
        SetBytes = Format(Bytes / 1024 / 1024 / 1024, "#0.0") _
                   & " GB"
    ElseIf Bytes >= 1048576 Then
        SetBytes = Format(Bytes / 1024 / 1024, "#0.0") & " MB"
    ElseIf Bytes >= 1024 Then
        SetBytes = Format(Bytes / 1024, "#0.0") & " KB"
    ElseIf Bytes < 1024 Then
        SetBytes = Fix(Bytes) & " B"
    End If

    Exit Function
hell:
    SetBytes = "0 Bytes"
End Function

Function GetFilePath(FileName As String) As String
    Dim I As Long

    For I = Len(FileName) To 1 Step -1
        Select Case Mid$(FileName, I, 1)
        Case ":"
            ' colons are always included in the result
            GetFilePath = Left$(FileName, I)
            Exit For
        Case "\"
            ' backslash aren't included in the result
            GetFilePath = Left$(FileName, I - 1)
            Exit For
        Case "/"
            ' backslash aren't included in the result
            GetFilePath = Left$(FileName, I - 1)
            Exit For
        End Select
    Next
End Function
Public Function ArquivoExiste(ByVal Caminho As String, Optional ByVal SomenteDiretorio As Boolean = False) As Boolean
    On Error Resume Next
    If SomenteDiretorio Then
        ArquivoExiste = GetAttr(Mid(Caminho, 1, InStrRev(Caminho, ""))) And vbDirectory
    Else
        ArquivoExiste = GetAttr(Caminho)
    End If
    On Error GoTo 0
End Function



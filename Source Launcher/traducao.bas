Attribute VB_Name = "traducao"
Function traduzir()
    If lang_type = 1 Then
        With inicia
            .Label16.Caption = "Options"
            .Label3.Caption = "Other options (On/Off)"
            .Label9.Caption = "Efects quality"
            .lblResolucao.Caption = "Resolution"
            .Label5.Caption = "Login"
            .Check1.Caption = "sound efects"
            .Check2.Caption = "Music"

        End With
        cancel_msg = "Discard changes?"
    Else
        cancel_msg = "Tem certeza que deseja cancelar as alterações?"
    End If
End Function

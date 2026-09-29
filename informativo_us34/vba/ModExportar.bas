Option Explicit

' ============================================================================
'  ModExportar - copia o Informativo / a Passagem de Turno como imagem
'  (para colar no e-mail, Teams ou WhatsApp com Ctrl+V)
' ============================================================================

Private Const XL_SCREEN As Long = 1
Private Const XL_PICTURE As Long = -4147

' Botao da aba Informativo
Public Sub CopiarImagem()
    CopiarAreaComoImagem "iAreaImagem", "Informativo"
End Sub

' Botao da aba Passagem de Turno
Public Sub CopiarImagemPassagem()
    CopiarAreaComoImagem "ptAreaImagem", "Passagem de Turno"
End Sub

Private Sub CopiarAreaComoImagem(ByVal nomeArea As String, ByVal titulo As String)
    On Error GoTo Falha
    Nm(nomeArea).CopyPicture Appearance:=XL_SCREEN, Format:=XL_PICTURE
    Aviso titulo & " copiado como imagem." & vbCrLf & vbCrLf & _
          "Abra o e-mail e cole com Ctrl+V.", vbInformation, "Copiar como imagem"
    Exit Sub
Falha:
    Aviso "Não foi possível copiar a imagem: " & Err.Description, vbExclamation, "Copiar como imagem"
End Sub

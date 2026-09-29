Option Explicit

' ============================================================================
'  ModExportar - copia o Resumo / Informativo / Passagem como imagem
'  (para colar no e-mail, Teams ou WhatsApp com Ctrl+V)
'  Usa formato BITMAP: o formato "imagem do Office" (metarquivo) nao cola no
'  Outlook Web, Gmail, Teams e outros programas.
' ============================================================================

Private Const XL_SCREEN As Long = 1
Private Const XL_BITMAP As Long = 2

' Botoes da aba Resumo do Turno (imagem 1 e imagem 2)
Public Sub CopiarImagemResultados()
    AtualizarResumo
    CopiarAreaComoImagem shResumo, "rsAreaResultados", "Resultados do turno (imagem 1)"
End Sub

Public Sub CopiarImagemPassagemResumo()
    AtualizarResumo
    CopiarAreaComoImagem shResumo, "rsAreaPassagem", "Passagem de turno (imagem 2)"
End Sub

' Botao da aba Informativo
Public Sub CopiarImagem()
    CopiarAreaComoImagem shInformativo, "iAreaImagem", "Informativo"
End Sub

' Botao da aba Passagem de Turno
Public Sub CopiarImagemPassagem()
    CopiarAreaComoImagem shPassagem, "ptAreaImagem", "Passagem de Turno"
End Sub

Private Sub CopiarAreaComoImagem(ByVal ws As Worksheet, ByVal nomeArea As String, ByVal titulo As String)
    Dim area As Range, tentativa As Long
    On Error GoTo Falha
    Set area = Nm(nomeArea)
    ws.Activate
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = area.Row
    Application.CutCopyMode = False
    DoEvents
    ' O Excel as vezes recusa a copia na 1a tentativa (area de transferencia ocupada)
    On Error Resume Next
    For tentativa = 1 To 3
        Err.Clear
        area.CopyPicture Appearance:=XL_SCREEN, Format:=XL_BITMAP
        If Err.Number = 0 Then Exit For
        DoEvents
    Next tentativa
    If Err.Number <> 0 Then GoTo Falha
    On Error GoTo Falha
    DoEvents
    Aviso titulo & " copiado como imagem." & vbCrLf & vbCrLf & _
          "Abra o e-mail (ou Teams) e cole com Ctrl+V." & vbCrLf & _
          "Se preferir, use Windows + Shift + S e recorte a tela.", _
          vbInformation, "Copiar como imagem"
    Exit Sub
Falha:
    Aviso "Não foi possível copiar a imagem: " & Err.Description & vbCrLf & _
          "Use Windows + Shift + S e recorte a tela.", vbExclamation, "Copiar como imagem"
End Sub

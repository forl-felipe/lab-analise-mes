Option Explicit

' ============================================================================
'  ModExportar - copia os Resumos / Informativo / Passagem como imagem
'  (para colar no e-mail, Teams ou WhatsApp com Ctrl+V)
'  Usa formato BITMAP: o formato "imagem do Office" (metarquivo) nao cola no
'  Outlook Web, Gmail, Teams e outros programas.
' ============================================================================

Private Const XL_SCREEN As Long = 1
Private Const XL_BITMAP As Long = 2
' Zoom usado so no momento da copia: o bitmap sai na resolucao da tela, entao copiar
' a 200% gera a imagem com o dobro de pixels (letras nitidas ao ampliar no e-mail)
Private Const ZOOM_IMAGEM As Long = 200

' Botoes das abas Resumo Dia / Resumo Noite (imagem 1 e imagem 2)
Public Sub CopiarResultadosDia()
    CopiarDoResumo "rsD", "AreaResultados", "Resultados do turno Dia (imagem 1)"
End Sub

Public Sub CopiarPassagemDia()
    CopiarDoResumo "rsD", "AreaPassagem", "Passagem do turno Dia (imagem 2)"
End Sub

Public Sub CopiarResultadosNoite()
    CopiarDoResumo "rsN", "AreaResultados", "Resultados do turno Noite (imagem 1)"
End Sub

Public Sub CopiarPassagemNoite()
    CopiarDoResumo "rsN", "AreaPassagem", "Passagem do turno Noite (imagem 2)"
End Sub

Private Sub CopiarDoResumo(ByVal pref As String, ByVal area As String, ByVal titulo As String)
    PrepararResumoParaCopia pref
    CopiarAreaComoImagem AbaResumo(pref), pref & area, titulo
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
    Dim area As Range, tentativa As Long, zoomAnt As Variant, msgErro As String
    On Error GoTo Falha
    Set area = Nm(nomeArea)
    ws.Activate
    zoomAnt = ActiveWindow.Zoom
    ActiveWindow.Zoom = ZOOM_IMAGEM
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
    ActiveWindow.Zoom = zoomAnt
    Aviso titulo & " copiado como imagem." & vbCrLf & vbCrLf & _
          "Abra o e-mail (ou Teams) e cole com Ctrl+V." & vbCrLf & _
          "A imagem sai em alta resolução: se ficar grande no e-mail, reduza pelos cantos." & vbCrLf & _
          "Se preferir, use Windows + Shift + S e recorte a tela.", _
          vbInformation, "Copiar como imagem"
    Exit Sub
Falha:
    msgErro = Err.Description
    On Error Resume Next
    If Not IsEmpty(zoomAnt) Then ActiveWindow.Zoom = zoomAnt
    On Error GoTo 0
    Aviso "Não foi possível copiar a imagem: " & msgErro & vbCrLf & _
          "Use Windows + Shift + S e recorte a tela.", vbExclamation, "Copiar como imagem"
End Sub

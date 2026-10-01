Option Explicit

' ============================================================================
'  ModImagem - copia os resumos / resultados como imagem (Ctrl+V no e-mail)
'  e oculta as linhas vazias para a imagem ficar compacta.
'  Formato BITMAP copiado a 200% de zoom: cola no Outlook (inclusive web), Teams
'  e WhatsApp, com letras nitidas ao ampliar.
' ============================================================================

Private Const XL_SCREEN As Long = 1
Private Const XL_BITMAP As Long = 2
Private Const ZOOM_IMAGEM As Long = 200

Public Sub CopiarImagemDia()
    CompactarAba shResumoDia
    CopiarArea shResumoDia, "dArea", "Resumo do turno Dia"
End Sub

Public Sub CopiarImagemNoite()
    CompactarAba shResumoNoite
    CopiarArea shResumoNoite, "nArea", "Resumo do turno Noite"
End Sub

Public Sub CopiarImagemResultados()
    CompactarAba shResultados
    CopiarArea shResultados, "gArea", "Resultados gerais"
End Sub

' Oculta: linhas de texto vazias (chave "T"), analises sem tag no MES (chave "P..")
' e o titulo de grupo que ficar sem analises (chave "G..").
Public Sub CompactarAba(ByVal ws As Worksheet)
    Dim r As Long, rFim As Long, chave As String, rGrupo As Long, visiveis As Long, oculta As Boolean
    On Error Resume Next
    Desproteger ws
    ' garante as formulas (textos migrados) recalculadas, tambem com o Excel em calculo manual
    Application.Calculate
    ws.Calculate
    rFim = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = 1 To rFim + 1
        chave = CStr(ws.Cells(r, 1).Value)
        If r > rFim Or Left$(chave, 1) = "G" Then
            If rGrupo > 0 Then ws.Rows(rGrupo).Hidden = (visiveis = 0)
            rGrupo = r
            visiveis = 0
        ElseIf chave = "T" Then
            ws.Rows(r).Hidden = (Len(Trim$(ws.Cells(r, 3).Text)) = 0)
        ElseIf Len(chave) = 5 And Left$(chave, 1) = "P" Then
            oculta = Not ParamConfigurado(CLng(Mid$(chave, 2, 2)))
            ws.Rows(r).Hidden = oculta
            If Not oculta Then visiveis = visiveis + 1
        End If
    Next r
    Proteger ws
End Sub

Private Sub CopiarArea(ByVal ws As Worksheet, ByVal nomeArea As String, ByVal titulo As String)
    Dim area As Range, tentativa As Long, zoomAnt As Variant, msgErro As String
    On Error GoTo Falha
    Set area = Nm(nomeArea)
    ws.Activate
    zoomAnt = ActiveWindow.Zoom
    ActiveWindow.Zoom = ZOOM_IMAGEM
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
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
    Aviso titulo & " copiado. Cole no e-mail com Ctrl+V.", vbInformation, "Copiar imagem"
    Exit Sub
Falha:
    msgErro = Err.Description
    On Error Resume Next
    If Not IsEmpty(zoomAnt) Then ActiveWindow.Zoom = zoomAnt
    On Error GoTo 0
    Aviso "Não foi possível copiar a imagem (" & msgErro & ")." & vbCrLf & _
          "Use Windows + Shift + S.", vbExclamation, "Copiar imagem"
End Sub

Option Explicit

' ============================================================================
'  ModImagem - copia os resumos / resultados como imagem (Ctrl+V no e-mail)
'  e oculta as linhas vazias para a imagem ficar compacta.
'
'  Alta resolucao: a area e copiada como figura vetorial, ampliada ESCALA vezes
'  dentro de um grafico temporario e exportada como PNG; o PNG e copiado para a
'  area de transferencia. Se algo falhar, usa a copia comum (bitmap da tela).
' ============================================================================

Private Const XL_SCREEN As Long = 1
Private Const XL_BITMAP As Long = 2
Private Const XL_PICTURE As Long = -4147
' Ampliacao da imagem exportada (2 = o dobro de pixels da tela)
Private Const ESCALA As Double = 2.2
' Caracteres por linha nos campos de texto do Resumo (altura automatica)
Private Const CHARS_LINHA As Long = 120

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

' Chaves na coluna A:
'   "T"   linha de lista: oculta quando vazia
'   "W"   texto livre: altura conforme o tamanho do texto
'   "P.." analise: oculta quando nao tem tag no MES
'   "G.." titulo de grupo: oculto quando o grupo fica sem analises
Public Sub CompactarAba(ByVal ws As Worksheet)
    Dim r As Long, rFim As Long, chave As String, rGrupo As Long, visiveis As Long, oculta As Boolean
    On Error Resume Next
    Desproteger ws
    ' Recalcula os textos migrados. Com a consulta do Aspen ativa (formula presente), recalcula so
    ' esta aba, para nao disparar outra consulta ao MES.
    If ConsultaAspenAtiva() Then ws.Calculate Else Application.Calculate
    rFim = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = 1 To rFim + 1
        chave = CStr(ws.Cells(r, 1).Value)
        If r > rFim Or Left$(chave, 1) = "G" Then
            If rGrupo > 0 Then ws.Rows(rGrupo).Hidden = (visiveis = 0)
            rGrupo = r
            visiveis = 0
        ElseIf chave = "T" Then
            ws.Rows(r).Hidden = (Len(Trim$(ws.Cells(r, 3).Text)) = 0)
        ElseIf chave = "W" Then
            ws.Rows(r).RowHeight = AlturaTexto(ws.Cells(r, 3).Text)
        ElseIf Len(chave) = 5 And Left$(chave, 1) = "P" Then
            oculta = Not ParamConfigurado(CLng(Mid$(chave, 2, 2)))
            ws.Rows(r).Hidden = oculta
            If Not oculta Then visiveis = visiveis + 1
        End If
    Next r
    Proteger ws
End Sub

Private Function AlturaTexto(ByVal t As String) As Double
    Dim partes As Variant, i As Long, n As Long
    partes = Split(Replace(t, vbCr, ""), vbLf)
    For i = LBound(partes) To UBound(partes)
        n = n + 1 + (Len(partes(i)) - 1) \ CHARS_LINHA
    Next i
    If n < 1 Then n = 1
    AlturaTexto = 6 + 14 * n
    If AlturaTexto < 20 Then AlturaTexto = 20
    If AlturaTexto > 400 Then AlturaTexto = 400
End Function

Private Sub CopiarArea(ByVal ws As Worksheet, ByVal nomeArea As String, ByVal titulo As String)
    Dim area As Range, tentativa As Long, msgErro As String
    On Error GoTo Falha
    Set area = Nm(nomeArea)
    ws.Activate
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Application.CutCopyMode = False
    DoEvents
    If CopiarAltaResolucao(ws, area) Then
        Aviso titulo & " copiado. Cole no e-mail com Ctrl+V.", vbInformation, "Copiar imagem"
        Exit Sub
    End If
    ' Alternativa: copia comum (bitmap da tela)
    On Error Resume Next
    For tentativa = 1 To 3
        Err.Clear
        area.CopyPicture Appearance:=XL_SCREEN, Format:=XL_BITMAP
        If Err.Number = 0 Then Exit For
        DoEvents
    Next tentativa
    If Err.Number <> 0 Then GoTo Falha
    On Error GoTo Falha
    Aviso titulo & " copiado. Cole no e-mail com Ctrl+V.", vbInformation, "Copiar imagem"
    Exit Sub
Falha:
    msgErro = Err.Description
    Aviso "Não foi possível copiar a imagem (" & msgErro & ")." & vbCrLf & _
          "Use Windows + Shift + S.", vbExclamation, "Copiar imagem"
End Sub

' Copia a area como PNG em alta resolucao. Devolve False se nao conseguir.
Private Function CopiarAltaResolucao(ByVal ws As Worksheet, ByVal area As Range) As Boolean
    Dim co As Object, fig As Object, img As Object, arq As String, tentativa As Long
    CopiarAltaResolucao = False
    On Error GoTo Falha
    arq = Environ$("TEMP") & "\resumo_quimico_" & Format$(Now, "hhnnss") & ".png"
    Desproteger ws
    ' 1) figura vetorial da area (texto continua nitido ao ampliar)
    On Error Resume Next
    For tentativa = 1 To 3
        Err.Clear
        area.CopyPicture Appearance:=XL_SCREEN, Format:=XL_PICTURE
        If Err.Number = 0 Then Exit For
        DoEvents
    Next tentativa
    If Err.Number <> 0 Then GoTo Falha
    On Error GoTo Falha
    ' 2) grafico temporario ampliado, com a figura ocupando todo o espaco
    Set co = ws.ChartObjects.Add(area.Left, area.Top, area.Width * ESCALA, area.Height * ESCALA)
    co.Activate
    co.Chart.Paste
    Set fig = co.Chart.Shapes(co.Chart.Shapes.Count)
    fig.LockAspectRatio = 0
    fig.Left = 0
    fig.Top = 0
    fig.Width = co.Chart.ChartArea.Width
    fig.Height = co.Chart.ChartArea.Height
    co.Chart.ChartArea.Format.Line.Visible = 0
    ' 3) exporta PNG e copia a imagem
    co.Chart.Export arq, "PNG"
    co.Delete
    Set co = Nothing
    Set img = ws.Shapes.AddPicture(arq, False, True, area.Left, area.Top, -1, -1)
    img.Copy
    img.Delete
    Set img = Nothing
    On Error Resume Next
    Kill arq
    ws.Range("A1").Select
    Proteger ws
    CopiarAltaResolucao = True
    Exit Function
Falha:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Not img Is Nothing Then img.Delete
    Kill arq
    Proteger ws
    CopiarAltaResolucao = False
End Function

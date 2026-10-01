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
'   "S1"  faixa de secao / "S2" subtitulo: ocultos (com os espacos "E") quando nada da secao foi preenchido
'   "E"   espaco da secao
'   "T"   linha de lista (texto na coluna B): oculta quando vazia; altura conforme o texto
'   "W"   texto livre (texto na coluna C): oculto quando vazio; altura conforme o texto
'   "F"/"V" linha de campos (formula): F = algum campo preenchido, V = todos vazios (oculta)
'   "Z"   fim das secoes (resultados do MES)
'   "P.." analise: oculta quando nao tem tag no MES
'   "G.." titulo de grupo: oculto quando o grupo fica sem analises
Public Sub CompactarAba(ByVal ws As Worksheet)
    Dim r As Long, rFim As Long, chave As String, rGrupo As Long, visiveis As Long, oculta As Boolean
    Dim r1 As Long, n1 As Long, e1 As String, r2 As Long, n2 As Long, e2 As String, conteudo As Boolean
    On Error Resume Next
    Desproteger ws
    ' Recalcula os textos migrados. Com a consulta do Aspen ativa (formula presente), recalcula so
    ' esta aba, para nao disparar outra consulta ao MES.
    If ConsultaAspenAtiva() Then ws.Calculate Else Application.Calculate
    rFim = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = 1 To rFim + 1
        chave = Trim$(ws.Cells(r, 1).Text)
        conteudo = False
        If r > rFim Or Left$(chave, 1) = "G" Then
            If rGrupo > 0 Then ws.Rows(rGrupo).Hidden = (visiveis = 0)
            rGrupo = r
            visiveis = 0
        End If
        Select Case chave
            Case "S1"
                FecharSecao ws, r2, n2, e2
                FecharSecao ws, r1, n1, e1
                r1 = r: n1 = 0: e1 = ""
            Case "S2"
                FecharSecao ws, r2, n2, e2
                r2 = r: n2 = 0: e2 = ""
            Case "Z"
                FecharSecao ws, r2, n2, e2
                FecharSecao ws, r1, n1, e1
            Case "E"
                If r2 > 0 Then
                    e2 = e2 & r & ","
                ElseIf r1 > 0 Then
                    e1 = e1 & r & ","
                End If
            Case "T"
                oculta = (Len(Trim$(ws.Cells(r, 2).Text)) = 0)
                ws.Rows(r).Hidden = oculta
                If Not oculta Then AjustarLinha ws, r, 17
                conteudo = Not oculta
            Case "W"
                oculta = (Len(Trim$(ws.Cells(r, 3).Text)) = 0)
                ws.Rows(r).Hidden = oculta
                If Not oculta Then AjustarLinha ws, r, 20
                conteudo = Not oculta
            Case "F"
                ws.Rows(r).Hidden = False
                AjustarLinha ws, r, 20
                conteudo = True
            Case "V"
                ws.Rows(r).Hidden = True
            Case Else
                If Len(chave) = 5 And Left$(chave, 1) = "P" Then
                    oculta = Not ParamConfigurado(CLng(Mid$(chave, 2, 2)))
                    ws.Rows(r).Hidden = oculta
                    If Not oculta Then visiveis = visiveis + 1
                End If
        End Select
        If conteudo Then
            n1 = n1 + 1
            n2 = n2 + 1
        End If
    Next r
    FecharSecao ws, r2, n2, e2
    FecharSecao ws, r1, n1, e1
    Proteger ws
End Sub

' Oculta o titulo da secao (e os espacos dela) quando nenhuma linha foi preenchida
Private Sub FecharSecao(ByVal ws As Worksheet, ByRef rTit As Long, ByRef n As Long, ByRef espacos As String)
    Dim lst As Variant, i As Long
    If rTit = 0 Then Exit Sub
    ws.Rows(rTit).Hidden = (n = 0)
    lst = Split(espacos, ",")
    For i = LBound(lst) To UBound(lst)
        If Len(lst(i)) > 0 Then ws.Rows(CLng(lst(i))).Hidden = (n = 0)
    Next i
    rTit = 0
    n = 0
    espacos = ""
End Sub

' ---------------------------------------------------------------- altura das linhas
' Celula mesclada nao tem ajuste automatico de altura no Excel: a altura e calculada
' pela largura da area mesclada e pelo tamanho do texto.

' Aba Preenchimento: chamada pelo evento de alteracao
Public Sub AjustarAlturas(ByVal ws As Worksheet, ByVal alvo As Range)
    Dim lin As Range, n As Long
    For Each lin In alvo.Rows
        n = n + 1
        If n > 60 Then Exit For
        AjustarLinha ws, lin.Row, 19
    Next lin
End Sub

' Ajusta a altura da linha r para mostrar todo o texto das celulas com quebra de texto
' (so linhas de uma altura: caixas de varias linhas ficam como estao)
Public Sub AjustarLinha(ByVal ws As Worksheet, ByVal r As Long, ByVal minimo As Double)
    Dim c As Long, cel As Range, area As Range, h As Double, hMax As Double, achou As Boolean
    On Error Resume Next
    For c = 2 To 16
        Set cel = ws.Cells(r, c)
        Set area = cel.MergeArea
        If area.Cells(1, 1).Address = cel.Address And area.Rows.Count = 1 And cel.WrapText Then
            If Len(cel.Text) > 0 Then
                h = AlturaTexto(cel.Text, area.Width, cel.Font.Size)
                If h > hMax Then hMax = h
            End If
            achou = True
        End If
    Next c
    If Not achou Then Exit Sub
    If hMax < minimo Then hMax = minimo
    If Abs(ws.Rows(r).RowHeight - hMax) > 0.5 Then
        Err.Clear
        ws.Rows(r).RowHeight = hMax
        If Err.Number <> 0 Then
            ' aba protegida sem permissao de formatar linhas
            Err.Clear
            Desproteger ws
            ws.Rows(r).RowHeight = hMax
            Proteger ws
        End If
    End If
End Sub

' Altura (pontos) para o texto caber na largura (pontos), com quebra de linha
Private Function AlturaTexto(ByVal t As String, ByVal largura As Double, ByVal fonte As Double) As Double
    Dim partes As Variant, i As Long, n As Long, porLinha As Long
    If fonte <= 0 Then fonte = 10
    ' largura media de um caractere ~ 0,5 x tamanho da fonte; descontado o recuo
    porLinha = Int((largura - 8) / (fonte * 0.5))
    If porLinha < 10 Then porLinha = 10
    partes = Split(Replace(t, vbCr, ""), vbLf)
    For i = LBound(partes) To UBound(partes)
        n = n + 1 + (Len(partes(i)) - 1) \ porLinha
    Next i
    If n < 1 Then n = 1
    AlturaTexto = 5 + fonte * 1.4 * n
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

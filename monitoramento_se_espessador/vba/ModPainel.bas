Option Explicit

' ============================================================================
'  ModPainel - graficos do Painel e envio do reporte (mesmo padrao do Comparativo Germano x Ubu)
'  - AjustarGraficos: series com os dias do Painel (rDias) + a coluna "Média" (barras escuras e
'    losango da diferenca media), escala das barras a partir de um pouco abaixo do menor valor
'  - Copiar imagem (PNG em alta resolucao) / Copiar resumo (texto curto para o corpo do e-mail)
' ============================================================================

Private Const XL_SCREEN As Long = 1
Private Const XL_BITMAP As Long = 2
Private Const XL_PICTURE As Long = -4147
Private Const XL_VALUE As Long = 2
Private Const XL_PRIMARY As Long = 1
' Ampliacao da imagem exportada (pixels = ESCALA x tela)
Private Const ESCALA As Double = 2.5
' Dias e data final do ultimo ajuste (so mexe nos graficos quando mudam)
Private mAssinGraf As String

Public Sub AtualizarPainel()
    Recalcular
    AjustarGraficos True
End Sub

Public Sub IrPainel()
    On Error Resume Next
    AtualizarPainel
    shPainel.Activate
    shPainel.Range("A1").Select
End Sub

Public Sub IrUltimoDia()
    Dim d As Long
    d = UltimoDiaHist()
    If d = 0 Then Exit Sub
    ModoCarga True
    Desproteger shPainel
    Nm("rFim").Value = CDate(d)
    Proteger shPainel
    ModoCarga False
    AtualizarPainel
End Sub

' Coluna (1-based) da chave na aba _Graf
Private Function GCol(ByVal chave As String) As Long
    Dim l As Variant, i As Long
    l = Split(G_CHAVES, "|")
    For i = LBound(l) To UBound(l)
        If StrComp(CStr(l(i)), chave, vbTextCompare) = 0 Then
            GCol = i - LBound(l) + 1
            Exit Function
        End If
    Next i
End Function

Public Function DiasPainel() As Long
    DiasPainel = CLng(CfgNum("rDias", 15))
    If DiasPainel < 2 Then DiasPainel = 2
    If DiasPainel > NMAXD Then DiasPainel = NMAXD
End Function

' Menor valor das colunas de barras (linhas 2..m+1 da aba _Graf); -1 = nenhum
Private Function MenorBarras(ByVal chaves As Variant, ByVal nb As Long, ByVal m As Long) As Double
    Dim s As Long, r As Long, v As Variant, achou As Boolean
    MenorBarras = -1
    For s = 0 To nb - 1
        For r = 2 To m + 1
            v = shGraf.Cells(r, GCol(CStr(chaves(s)))).Value
            If Not IsError(v) Then
                If ENumero(v) Then
                    If Not achou Or CDbl(v) < MenorBarras Then MenorBarras = CDbl(v)
                    achou = True
                End If
            End If
        Next r
    Next s
End Function

Public Sub AjustarGraficos(Optional ByVal forcar As Boolean = False)
    Dim n As Long, m As Long, g As Variant, nbs As Variant, i As Long, s As Long, p As Long, ch As Object
    Dim ser As Object, chaves As Variant, c As Long, nb As Long, cats As Range, corBase As Variant, corMed As Variant
    Dim assin As String, mn As Double, eixo As Object
    If gTesteLO Then Exit Sub
    On Error Resume Next
    n = DiasPainel()
    m = n + 1                                    ' dias + a coluna "Média"
    assin = n & "|" & CStr(Nm("rFim").Value)
    If Not forcar And assin = mAssinGraf Then Exit Sub
    mAssinGraf = assin
    Application.ScreenUpdating = False
    Recalcular
    Desproteger shPainel
    Set cats = shGraf.Range(shGraf.Cells(2, GCol("rot")), shGraf.Cells(m + 1, GCol("rot")))
    g = Split(GRAF_SERIES, "|")
    nbs = Split(GRAF_NBARRAS, "|")
    corBase = Split(COR_BASE, "|")
    corMed = Split(COR_MEDIA, "|")
    For i = LBound(g) To UBound(g)
        Set ch = Nothing
        Set ch = shPainel.ChartObjects(i - LBound(g) + 1).Chart
        If Not ch Is Nothing Then
            chaves = Split(CStr(g(i)), ",")
            nb = CLng(nbs(i))
            For s = 1 To ch.SeriesCollection.Count
                If s - 1 <= UBound(chaves) Then
                    c = GCol(CStr(chaves(s - 1)))
                    Set ser = ch.SeriesCollection(s)
                    ser.Values = shGraf.Range(shGraf.Cells(2, c), shGraf.Cells(m + 1, c))
                    ser.XValues = cats
                    ' barras: a ultima (Média) mais escura, como no Comparativo
                    If s <= nb Then
                        For p = 1 To m
                            ser.Points(p).Format.Fill.ForeColor.RGB = IIf(p = m, CLng(corMed(s - 1)), CLng(corBase(s - 1)))
                        Next p
                    End If
                End If
            Next s
            Set eixo = Nothing
            Set eixo = ch.Axes(XL_VALUE, XL_PRIMARY)
            If Not eixo Is Nothing Then
                mn = MenorBarras(chaves, nb, m)
                If mn > 0 Then
                    eixo.MinimumScale = Application.WorksheetFunction.Max(0, Int((mn - 150) / 100) * 100)
                Else
                    eixo.MinimumScaleIsAuto = True
                End If
                eixo.MaximumScaleIsAuto = True
            End If
        End If
    Next i
    Proteger shPainel
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------- imagem
Public Sub CopiarImagemPainel()
    AtualizarPainel
    shPainel.Activate
    If CopiarAreaImagem(shPainel, Nm("pArea")) Then
        Aviso "Painel copiado. Cole no e-mail ou no Teams com Ctrl+V.", vbInformation, "Copiar imagem"
    Else
        Aviso "Não foi possível copiar a imagem. Use Windows + Shift + S.", vbExclamation, "Copiar imagem"
    End If
End Sub

' Copia uma area como imagem (PNG em alta resolucao; se falhar, bitmap da tela)
Public Function CopiarAreaImagem(ByVal ws As Worksheet, ByVal area As Range) As Boolean
    Dim arq As String, tentativa As Long
    arq = ExportarAreaPNG(ws, area, "")
    If arq <> "" Then
        CopiarAreaImagem = CopiarArquivoImagem(arq, ws)
        On Error Resume Next
        Kill arq
        Exit Function
    End If
    On Error Resume Next
    For tentativa = 1 To 3
        Err.Clear
        area.CopyPicture Appearance:=XL_SCREEN, Format:=XL_BITMAP
        If Err.Number = 0 Then Exit For
        DoEvents
    Next tentativa
    CopiarAreaImagem = (Err.Number = 0)
End Function

' Coloca o PNG na area de transferencia (via figura temporaria na aba)
Private Function CopiarArquivoImagem(ByVal arq As String, ByVal ws As Worksheet) As Boolean
    Dim img As Object
    On Error GoTo Falha
    Desproteger ws
    Set img = ws.Shapes.AddPicture(arq, False, True, 0, 0, -1, -1)
    img.Copy
    img.Delete
    Proteger ws
    CopiarArquivoImagem = True
    Exit Function
Falha:
    On Error Resume Next
    If Not img Is Nothing Then img.Delete
    Proteger ws
End Function

' Exporta a area como PNG em alta resolucao. Devolve o caminho ("" se falhar).
Public Function ExportarAreaPNG(ByVal ws As Worksheet, ByVal area As Range, ByVal destino As String) As String
    Dim co As Object, fig As Object, tentativa As Long, arq As String
    On Error GoTo Falha
    If destino = "" Then
        arq = Environ$("TEMP") & "\Painel_SE_" & Format$(Now, "yyyymmdd_hhnnss") & ".png"
    Else
        arq = destino
    End If
    ws.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Application.CutCopyMode = False
    Desproteger ws
    ' 1) figura vetorial da area (inclui os graficos) - texto continua nitido ao ampliar
    On Error Resume Next
    For tentativa = 1 To 3
        Err.Clear
        area.CopyPicture Appearance:=XL_SCREEN, Format:=XL_PICTURE
        If Err.Number = 0 Then Exit For
        DoEvents
    Next tentativa
    If Err.Number <> 0 Then GoTo Falha
    On Error GoTo Falha
    ' 2) grafico temporario ampliado com a figura ocupando todo o espaco
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
    co.Chart.Export arq, "PNG"
    co.Delete
    area.Cells(1, 1).Select
    Proteger ws
    ExportarAreaPNG = arq
    Exit Function
Falha:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    Proteger ws
    ExportarAreaPNG = ""
End Function

' ---------------------------------------------------------------- resumo para o e-mail
Public Function TextoResumo() As String
    Recalcular
    TextoResumo = Replace(CStr(Nm("gResumo").Value), vbLf, vbCrLf)
End Function

Public Sub CopiarResumo()
    If CopiarTexto(TextoResumo()) Then
        Aviso "Resumo copiado. Cole no corpo do e-mail com Ctrl+V.", vbInformation, "Copiar resumo"
    Else
        ' sem acesso a area de transferencia: mostra o texto para copiar a mao
        InputBox "Não foi possível copiar sozinho. Selecione o texto abaixo e copie com Ctrl+C:", _
                 "Copiar resumo", Replace(TextoResumo(), vbCrLf, " / ")
    End If
End Sub

Private Function CopiarTexto(ByVal t As String) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = CreateObject("htmlfile")
    o.ParentWindow.ClipboardData.SetData "text", t
    If Err.Number = 0 Then
        CopiarTexto = True
        Exit Function
    End If
    Err.Clear
    Set o = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    o.SetText t
    o.PutInClipboard
    CopiarTexto = (Err.Number = 0)
End Function

Option Explicit

' ============================================================================
'  ModReport - graficos do Report e copia da imagem para o e-mail
'  - AjustarGraficos: series com o numero de dias do Report (rDias), a partir da aba _Graf
'  - Copiar imagem: PNG em alta resolucao da area do report (pArea) na area de transferencia
' ============================================================================

Private Const XL_SCREEN As Long = 1
Private Const XL_BITMAP As Long = 2
Private Const XL_PICTURE As Long = -4147
' Ampliacao da imagem exportada (pixels = ESCALA x tela)
Private Const ESCALA As Double = 2.5
' Series de cada grafico (ordem do SeriesCollection: colunas primeiro, depois as linhas)
Private Const GRAF_SERIES As String = "difK,espP,eb6|difJ,eb4P,eb6|eb6,tk01,tk02,pctLama"
' Dias do ultimo ajuste (so mexe nos graficos quando muda)
Private mDiasGraf As Long

Public Sub AtualizarReport()
    Recalcular
    AjustarGraficos True
End Sub

Public Sub IrUltimoDia()
    Dim d As Long
    d = UltimoDiaHist()
    If d = 0 Then Exit Sub
    ModoCarga True
    Desproteger shReport
    Nm("rFim").Value = CDate(d)
    Proteger shReport
    ModoCarga False
    AtualizarReport
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

Public Function DiasReport() As Long
    DiasReport = CLng(CfgNum("rDias", 15))
    If DiasReport < 2 Then DiasReport = 2
    If DiasReport > NMAXD Then DiasReport = NMAXD
End Function

Public Sub AjustarGraficos(Optional ByVal forcar As Boolean = False)
    Dim n As Long, g As Variant, i As Long, s As Long, ch As Object, ser As Object, chaves As Variant, c As Long
    Dim cats As Range
    If gTesteLO Then Exit Sub
    On Error Resume Next
    n = DiasReport()
    If Not forcar And n = mDiasGraf Then Exit Sub
    mDiasGraf = n
    Application.ScreenUpdating = False
    Desproteger shReport
    Set cats = shGraf.Range(shGraf.Cells(2, GCol("rot")), shGraf.Cells(n + 1, GCol("rot")))
    g = Split(GRAF_SERIES, "|")
    For i = LBound(g) To UBound(g)
        Set ch = Nothing
        Set ch = shReport.ChartObjects(i - LBound(g) + 1).Chart
        If Not ch Is Nothing Then
            chaves = Split(CStr(g(i)), ",")
            For s = 1 To ch.SeriesCollection.Count
                If s - 1 <= UBound(chaves) Then
                    c = GCol(CStr(chaves(s - 1)))
                    Set ser = ch.SeriesCollection(s)
                    ser.Values = shGraf.Range(shGraf.Cells(2, c), shGraf.Cells(n + 1, c))
                    ser.XValues = cats
                End If
            Next s
        End If
    Next i
    Proteger shReport
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------- imagem
Public Sub CopiarImagemReport()
    AtualizarReport
    shReport.Activate
    If CopiarAreaImagem(shReport, Nm("pArea")) Then
        Aviso "Report copiado. Cole no e-mail ou no Teams com Ctrl+V.", vbInformation, "Copiar imagem"
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
        arq = Environ$("TEMP") & "\Report_SE_" & Format$(Now, "yyyymmdd_hhnnss") & ".png"
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

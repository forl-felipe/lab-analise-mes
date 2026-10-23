Option Explicit

' ============================================================================
'  ModPainel - graficos do Painel e envio do reporte
'  - AjustarGraficos: series do tamanho do batch (pares Germano/Ubu + "Media"),
'    barra da media mais escura (como no painel HTML)
'  - Copiar imagem / Salvar PNG (alta resolucao) / PDF / e-mail no Outlook
' ============================================================================

Private Const XL_SCREEN As Long = 1
Private Const XL_BITMAP As Long = 2
Private Const XL_PICTURE As Long = -4147
' Ampliacao da imagem exportada (pixels = ESCALA x tela)
Private Const ESCALA As Double = 2.5
' Numero de pares de cada grafico no ultimo ajuste
Private mAssinGraf As String

' ---------------------------------------------------------------- graficos
Public Sub AtualizarPainel()
    Recalcular
    AjustarGraficos True
    shPainel.Activate
End Sub

' So mexe nos graficos quando o numero de pares muda (ajustar as 30 series a toda hora deixava a planilha lenta)
Public Sub AjustarGraficos(Optional ByVal forcar As Boolean = False)
    Dim ordem As Variant, i As Long, k As String, p As Long, n As Long, c0 As Long, ult As Long
    Dim co As Object, ch As Object, s As Long, ser As Object, g As String, assin As String
    If gTesteLO Then Exit Sub
    On Error Resume Next
    ordem = Split(GRAF_ORDEM, "|")
    assin = ""
    For i = LBound(ordem) To UBound(ordem)
        assin = assin & shGraf.Cells(G_ROW_ST + PIndice(CStr(ordem(i))) - 1, G_COL_NP).Value & "|"
    Next i
    If Not forcar And assin = mAssinGraf Then Exit Sub
    mAssinGraf = assin
    Application.ScreenUpdating = False
    g = "'" & shGraf.Name & "'!"
    Desproteger shPainel
    For i = LBound(ordem) To UBound(ordem)
        k = CStr(ordem(i))
        p = PIndice(k)
        n = 0
        If ENumero(shGraf.Cells(G_ROW_ST + p - 1, G_COL_NP).Value) Then n = CLng(shGraf.Cells(G_ROW_ST + p - 1, G_COL_NP).Value)
        c0 = (p - 1) * G_BLOCO + 1                 ' coluna do rotulo
        ult = n + 2                                ' linha da "Media" (dados a partir da linha 2)
        Set co = Nothing
        Set co = shPainel.ChartObjects(i - LBound(ordem) + 1)
        If Not co Is Nothing Then
            Set ch = co.Chart
            For s = 1 To ch.SeriesCollection.Count
                Set ser = ch.SeriesCollection(s)
                ser.XValues = "=" & g & Faixa(c0, ult)
                ser.Values = "=" & g & Faixa(c0 + s, ult)
                If s <= 2 Then PintarBarras ser, s, n + 1
            Next s
        End If
    Next i
    Application.ScreenUpdating = True
    Proteger shPainel
End Sub

Private Function Faixa(ByVal c As Long, ByVal ult As Long) As String
    Faixa = shGraf.Range(shGraf.Cells(2, c), shGraf.Cells(ult, c)).Address(True, True)
End Function

' Barras do laboratorio (s = 1 A/Germano, 2 B/Ubu); a ultima (Media) mais escura
Private Sub PintarBarras(ByVal ser As Object, ByVal s As Long, ByVal nPts As Long)
    Dim corN As Long, corM As Long, j As Long
    On Error Resume Next
    If s = 1 Then
        corN = RGB(&HAE, &HCB, &HEE)
        corM = RGB(&H5F, &H8A, &HC8)
    Else
        corN = RGB(&HE0, &HE4, &HEA)
        corM = RGB(&HA7, &HB0, &HBB)
    End If
    ser.Format.Fill.ForeColor.RGB = corN
    For j = 1 To nPts
        ser.Points(j).Format.Fill.ForeColor.RGB = IIf(j = nPts, corM, corN)
    Next j
End Sub

' ---------------------------------------------------------------- imagem
Public Sub CopiarImagemPainel()
    Dim arq As String
    AtualizarPainel
    arq = ExportarPNG()
    If arq = "" Then
        If CopiarBitmap() Then
            Aviso "Painel copiado. Cole no e-mail com Ctrl+V.", vbInformation, "Copiar imagem"
        Else
            Aviso "Não foi possível copiar a imagem. Use Windows + Shift + S.", vbExclamation, "Copiar imagem"
        End If
        Exit Sub
    End If
    If CopiarArquivoImagem(arq) Then
        Aviso "Painel copiado em alta resolução. Cole no e-mail ou no Teams com Ctrl+V.", vbInformation, "Copiar imagem"
    Else
        Aviso "Não foi possível copiar a imagem. Use Salvar PNG.", vbExclamation, "Copiar imagem"
    End If
    On Error Resume Next
    Kill arq
End Sub

Private Function CopiarBitmap() As Boolean
    Dim tentativa As Long
    On Error Resume Next
    For tentativa = 1 To 3
        Err.Clear
        Nm("pArea").CopyPicture Appearance:=XL_SCREEN, Format:=XL_BITMAP
        If Err.Number = 0 Then Exit For
        DoEvents
    Next tentativa
    CopiarBitmap = (Err.Number = 0)
End Function

' Coloca o PNG na area de transferencia (via figura temporaria na aba)
Private Function CopiarArquivoImagem(ByVal arq As String) As Boolean
    Dim img As Object
    On Error GoTo Falha
    Desproteger shPainel
    Set img = shPainel.Shapes.AddPicture(arq, False, True, 0, 0, -1, -1)
    img.Copy
    img.Delete
    Proteger shPainel
    CopiarArquivoImagem = True
    Exit Function
Falha:
    On Error Resume Next
    If Not img Is Nothing Then img.Delete
    Proteger shPainel
End Function

' Exporta a area do report como PNG em alta resolucao. Devolve o caminho ("" se falhar).
Public Function ExportarPNG(Optional ByVal destino As String = "") As String
    Dim area As Range, co As Object, fig As Object, tentativa As Long, arq As String
    On Error GoTo Falha
    Set area = Nm("pArea")
    If destino = "" Then
        arq = Environ$("TEMP") & "\Report_Batch_" & CfgTxt("bId") & "_" & Format$(Now, "hhnnss") & ".png"
    Else
        arq = destino
    End If
    shPainel.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Application.CutCopyMode = False
    Desproteger shPainel
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
    Set co = shPainel.ChartObjects.Add(area.Left, area.Top, area.Width * ESCALA, area.Height * ESCALA)
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
    shPainel.Range("A1").Select
    Proteger shPainel
    ExportarPNG = arq
    Exit Function
Falha:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    Proteger shPainel
    ExportarPNG = ""
End Function





' ---------------------------------------------------------------- resumo e e-mail
Public Function TextoResumo() As String
    Recalcular
    TextoResumo = Replace(CStr(Nm("eResumo").Value), vbLf, vbCrLf)
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





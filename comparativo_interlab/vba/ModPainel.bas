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

' ---------------------------------------------------------------- graficos
Public Sub AtualizarPainel()
    Recalcular
    AjustarGraficos
    shPainel.Activate
End Sub

Public Sub AjustarGraficos()
    Dim ordem As Variant, i As Long, k As String, p As Long, n As Long, c0 As Long, ult As Long
    Dim co As Object, ch As Object, s As Long, ser As Object, g As String
    If gTesteLO Then Exit Sub
    On Error Resume Next
    ordem = Split(GRAF_ORDEM, "|")
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

Private Function PastaPadrao() As String
    Dim p As String
    p = ThisWorkbook.Path
    If p = "" Or LCase$(Left$(p, 4)) = "http" Then p = Environ$("USERPROFILE") & "\Desktop"
    PastaPadrao = p
End Function

Private Function NomeArquivo(ByVal ext As String) As String
    Dim d As String
    If EData(Nm("bData").Value) Then d = "_" & Format$(CDate(Nm("bData").Value), "yyyymmdd")
    NomeArquivo = "Report_Batch_" & CfgTxt("bId") & d & "." & ext
End Function

Public Sub SalvarPNG()
    Dim arq As Variant, r As String
    AtualizarPainel
    arq = Application.GetSaveAsFilename(PastaPadrao() & "\" & NomeArquivo("png"), "Imagem PNG (*.png), *.png", , _
                                        "Salvar painel como PNG")
    If VarType(arq) = vbBoolean Then Exit Sub
    On Error Resume Next
    If Dir$(CStr(arq)) <> "" Then Kill CStr(arq)
    On Error GoTo 0
    r = ExportarPNG(CStr(arq))
    If r = "" Then
        Aviso "Não foi possível gerar a imagem.", vbExclamation, "Salvar PNG"
    Else
        Aviso "PNG salvo:" & vbCrLf & r, vbInformation, "Salvar PNG"
    End If
End Sub

Public Sub SalvarPDF()
    Dim arq As Variant
    AtualizarPainel
    arq = Application.GetSaveAsFilename(PastaPadrao() & "\" & NomeArquivo("pdf"), "PDF (*.pdf), *.pdf", , _
                                        "Salvar painel como PDF")
    If VarType(arq) = vbBoolean Then Exit Sub
    On Error GoTo Falha
    shPainel.ExportAsFixedFormat Type:=0, Filename:=CStr(arq), Quality:=0, IncludeDocProperties:=False, _
                                 IgnorePrintAreas:=False, OpenAfterPublish:=True
    Exit Sub
Falha:
    Aviso "Não foi possível salvar o PDF: " & Err.Description, vbExclamation, "Salvar PDF"
End Sub

' ---------------------------------------------------------------- resumo e e-mail
Public Function TextoResumo() As String
    Recalcular
    TextoResumo = Replace(CStr(Nm("eResumo").Value), vbLf, vbCrLf)
End Function

Public Sub CopiarResumo()
    If CopiarTexto(TextoResumo()) Then
        Aviso "Resumo copiado. Cole no corpo do e-mail com Ctrl+V.", vbInformation, "Copiar resumo"
    Else
        Aviso "Não foi possível copiar. Selecione o texto da aba Enviar e copie com Ctrl+C.", vbExclamation, _
              "Copiar resumo"
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

Private Function Html(ByVal t As String) As String
    t = Replace(t, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, vbCrLf, "<br>")
    t = Replace(t, vbLf, "<br>")
    Html = t
End Function

Public Sub CriarEmail()
    Dim ol As Object, mail As Object, arq As String, nome As String, corpo As String, assunto As String
    AtualizarPainel
    assunto = CfgTxt("eAssunto")
    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        ' sem Outlook: abre o e-mail padrao com assunto e resumo; a imagem vai pela area de transferencia
        CopiarImagemPainel
        On Error Resume Next
        ThisWorkbook.FollowHyperlink "mailto:?subject=" & UrlTexto(assunto) & "&body=" & UrlTexto(TextoResumo())
        Exit Sub
    End If
    arq = ExportarPNG()
    On Error GoTo Falha
    Set mail = ol.CreateItem(0)
    nome = Mid$(arq, InStrRev(arq, "\") + 1)
    corpo = "<div style=""font-family:Segoe UI,Calibri,Arial;font-size:10.5pt;color:#1a2744"">" & _
            Html(TextoResumo()) & "</div>"
    If arq <> "" Then
        mail.Attachments.Add arq, 1, 0
        On Error Resume Next
        mail.Attachments(1).PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x3712001F", nome
        On Error GoTo Falha
        corpo = corpo & "<p><img src=""cid:" & nome & """ style=""max-width:100%""></p>"
    End If
    mail.Subject = assunto
    mail.HTMLBody = corpo & mail.HTMLBody
    mail.Display
    On Error Resume Next
    If arq <> "" Then Kill arq
    Exit Sub
Falha:
    Aviso "Não foi possível criar o e-mail no Outlook: " & Err.Description, vbExclamation, "E-mail"
End Sub

Private Function UrlTexto(ByVal t As String) As String
    Dim i As Long, ch As String, r As String, b() As Byte, j As Long
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "[A-Za-z0-9.-]" Then
            r = r & ch
        ElseIf ch = " " Then
            r = r & "%20"
        ElseIf AscW(ch) < 128 And AscW(ch) >= 0 Then
            r = r & "%" & Right$("0" & Hex$(AscW(ch)), 2)
        Else
            r = r & Utf8Pct(AscW(ch))
        End If
    Next i
    UrlTexto = r
End Function

Private Function Utf8Pct(ByVal cp As Long) As String
    If cp < 0 Then cp = cp + 65536
    If cp < &H800& Then
        Utf8Pct = "%" & Hex$(&HC0 Or (cp \ &H40)) & "%" & Hex$(&H80 Or (cp And &H3F))
    Else
        Utf8Pct = "%" & Hex$(&HE0 Or (cp \ &H1000)) & "%" & Hex$(&H80 Or ((cp \ &H40) And &H3F)) & _
                  "%" & Hex$(&H80 Or (cp And &H3F))
    End If
End Function

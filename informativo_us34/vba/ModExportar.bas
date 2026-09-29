Option Explicit

' ============================================================================
'  ModExportar - PDF, imagem, e-mail (Outlook) e logo
' ============================================================================

Public Function PastaPDF() As String
    Dim p As String
    p = CfgTxt("cfgPastaPDF")
    If p = "" Then
        p = ThisWorkbook.Path
        If p = "" Or LCase$(Left$(p, 4)) = "http" Then p = Environ$("USERPROFILE") & "\Documents"
        p = p & "\Informativos PDF"
    End If
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    On Error Resume Next
    If Dir(p, vbDirectory) = "" Then MkDir p
    On Error GoTo 0
    PastaPDF = p
End Function

Public Function ExportarPDF(ByVal dt As Date, ByVal turno As String, ByVal turma As String, _
                            ByVal abrir As Boolean) As String
    Dim f As String, nomeTurno As String
    If EhDia(turno) Then nomeTurno = "Dia" Else nomeTurno = "Noite"
    f = PastaPDF() & "\Informativo_US3-US4_" & Format$(dt, "yyyy-mm-dd") & "_" & nomeTurno
    If turma <> "" Then f = f & "_Turma" & turma
    f = f & ".pdf"

    On Error GoTo Falha
    ThisWorkbook.Worksheets(Array(shInformativo.Name, shPassagem.Name)).Select
    ActiveSheet.ExportAsFixedFormat Type:=0, Filename:=f, Quality:=0, IncludeDocProperties:=False, _
                                   IgnorePrintAreas:=False, OpenAfterPublish:=abrir
    shPainel.Select
    ExportarPDF = f
    Exit Function
Falha:
    On Error Resume Next
    shPainel.Select
    Aviso "Não foi possível gerar o PDF em:" & vbCrLf & f & vbCrLf & vbCrLf & Err.Description & vbCrLf & _
           "Ajuste a pasta de PDF na aba Configurações.", vbExclamation, "PDF"
    ExportarPDF = ""
End Function

' Botao da aba Informativo
Public Sub ExportarPDFManual()
    Dim dt As Date, turno As String, turma As String, resp As String, f As String
    PrepararEdicao
    If ENumero(Nm("iData").Value) Or VarType(Nm("iData").Value) = vbDate Then
        dt = CDate(Nm("iData").Value)
        turno = CfgTxt("iTurno")
        turma = CfgTxt("iTurma")
    ElseIf Not LerTurnoPainel(dt, turno, turma, resp, True) Then
        Exit Sub
    End If
    f = ExportarPDF(dt, turno, turma, True)
    IrInformativo
    If f <> "" Then Aviso "PDF gerado:" & vbCrLf & f, vbInformation, "PDF"
End Sub

' Botao da aba Informativo: copia o relatorio como imagem (colar no e-mail / Teams / WhatsApp)
Public Sub CopiarImagem()
    On Error GoTo Falha
    Nm("iAreaImagem").CopyPicture Appearance:=1, Format:=-4147
    Aviso "Imagem do Informativo copiada!" & vbCrLf & "Cole (Ctrl+V) no e-mail, Teams ou WhatsApp.", _
           vbInformation, "Imagem"
    Exit Sub
Falha:
    Aviso "Não foi possível copiar a imagem: " & Err.Description, vbExclamation, "Imagem"
End Sub

Public Sub CriarEmail(ByVal dt As Date, ByVal turno As String, ByVal turma As String, _
                      ByVal resp As String, ByVal pdf As String)
    Dim ol As Object, m As Object, corpo As String
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo Falha
    If ol Is Nothing Then
        Aviso "Outlook não disponível. O PDF foi salvo em:" & vbCrLf & pdf, vbInformation, "E-mail"
        Exit Sub
    End If
    corpo = CorpoEmail(dt, turno, turma, resp)
    Set m = ol.CreateItem(0)
    m.Display
    m.To = CfgTxt("cfgEmailPara")
    m.CC = CfgTxt("cfgEmailCC")
    m.Subject = "Informativo de Qualidade US3/US4 - " & Format$(dt, "dd\/mm\/yyyy") & " - " & turno & _
                " - Turma " & turma
    m.HTMLBody = corpo & m.HTMLBody
    If pdf <> "" Then
        If Dir(pdf) <> "" Then m.Attachments.Add pdf
    End If
    Exit Sub
Falha:
    Aviso "Não foi possível criar o e-mail: " & Err.Description, vbExclamation, "E-mail"
End Sub

Private Function CorpoEmail(ByVal dt As Date, ByVal turno As String, ByVal turma As String, _
                            ByVal resp As String) As String
    Dim h As String, p As Long, lin As Long, c As Range, nok As String
    Const TD As String = "<td style='padding:3px 8px;border-bottom:1px solid #B5BEC4'>"

    h = "<div style='font-family:Segoe UI,Arial;font-size:10pt;color:#002643'>"
    h = h & "<p style='font-size:13pt;font-weight:bold;color:#00335A;margin:0'>Informativo de Qualidade - Usinas 3 e 4</p>"
    h = h & "<p style='margin:2px 0 10px 0'>" & Format$(dt, "dd\/mm\/yyyy") & " &middot; " & HtmlEsc(turno) & _
            " &middot; Turma " & HtmlEsc(turma) & " &middot; " & HtmlEsc(resp) & "</p>"
    If InformativoConfere(dt, turno) Then
        h = h & "<table style='border-collapse:collapse;font-size:9pt'>"
        h = h & "<tr style='background:#00335A;color:#FFFFFF'><th style='padding:4px 8px;text-align:left'>Parâmetro</th>" & _
                "<th style='padding:4px 8px'>Unid.</th><th style='padding:4px 8px'>US3</th><th style='padding:4px 8px'>US4</th></tr>"
        For p = 1 To NPARAM
            lin = LinhaInformativo(p, 1)
            h = h & "<tr>" & TD & HtmlEsc(CStr(CfgCel(p, CFG_COL_PARAM))) & "</td>" & TD & _
                HtmlEsc(CStr(CfgCel(p, CFG_COL_UNID))) & "</td>" & TD & _
                HtmlEsc(shInformativo.Cells(lin, INF_COL_RES).Text) & "</td>" & TD & _
                HtmlEsc(shInformativo.Cells(LinhaInformativo(p, 2), INF_COL_RES).Text) & "</td></tr>"
        Next p
        h = h & "</table>"
    End If
    For Each c In Nm("ptStatus").Cells
        If CStr(c.Value) = "NÃO OK" Then
            nok = nok & "<li>" & HtmlEsc(CStr(shPassagem.Cells(c.Row, 3).Value)) & " - " & _
                  HtmlEsc(CStr(shPassagem.Cells(c.Row, c.Column + 2).Value)) & "</li>"
        End If
    Next c
    If nok <> "" Then h = h & "<p style='margin:10px 0 2px 0;font-weight:bold;color:#F37021'>Equipamentos NÃO OK</p><ul>" & nok & "</ul>"
    If Not Vazio(Nm("ptPendencias").Value) Then
        h = h & "<p style='margin:10px 0 2px 0;font-weight:bold'>Pendências para o próximo turno</p><p style='margin:0'>" & _
                HtmlEsc(CStr(Nm("ptPendencias").Value)) & "</p>"
    End If
    CorpoEmail = h & "<p style='color:#61889B;font-size:8pt;margin-top:12px'>Gerado automaticamente pelo Informativo US3/US4.</p></div><br>"
End Function

' Botao da aba Configuracoes: escolhe a imagem do logo e aplica nos cabecalhos
Public Sub InserirLogo()
    Dim arq As Variant, alvo As Variant, ws As Worksheet, rng As Range, shp As Object, nomes As Variant, i As Long

    arq = Application.GetOpenFilename("Imagens (*.png;*.jpg;*.jpeg;*.gif;*.bmp),*.png;*.jpg;*.jpeg;*.gif;*.bmp", , _
                                      "Selecione o arquivo do logo")
    If VarType(arq) = vbBoolean Then Exit Sub
    PrepararEdicao
    nomes = Array("logoPainel", "logoInformativo", "logoPassagem", "logoHistorico", "logoRegistro", "logoConfig")
    For i = LBound(nomes) To UBound(nomes)
        Set rng = Nm(CStr(nomes(i)))
        Set ws = rng.Worksheet
        On Error Resume Next
        ws.Shapes("LogoSamarco").Delete
        On Error GoTo 0
        rng.MergeArea.ClearContents
        Set shp = ws.Shapes.AddPicture(CStr(arq), 0, -1, rng.Left + 2, rng.Top + 2, -1, -1)
        shp.Name = "LogoSamarco"
        shp.LockAspectRatio = -1
        shp.Height = rng.MergeArea.Height - 6
        If shp.Width > rng.MergeArea.Width - 4 Then shp.Width = rng.MergeArea.Width - 4
        shp.Top = rng.Top + (rng.MergeArea.Height - shp.Height) / 2
        shp.Placement = 3
    Next i
    Aviso "Logo aplicado em todas as abas.", vbInformation, "Logo"
End Sub

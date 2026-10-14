Option Explicit

' ============================================================================
'  ModBatch - historico de batches (tbl_Batches + tbl_Leituras), novo batch,
'  abrir/excluir e importacao da planilha antiga.
'  tbl_Batches (aba Historico): um batch por linha. tbl_Leituras (aba BD_Leituras):
'  uma linha por batch x horario x analise. As duas alimentam o Power BI.
' ============================================================================

' Campos da aba Batch na ordem das primeiras colunas de tbl_Batches
Private Function CamposBatch() As Variant
    Dim c As String, p As Long
    c = "bId|bData|bMin|bIniU|bIniG|bNHor|bLabA|bLabB|bRef|bTit|bSub"
    For p = 1 To NP
        c = c & "|bTol_" & PChave(p)
    Next p
    c = c & "|bIncFrase|bCom1|bCom2|bCom3|bCom4|bCom5|bCom6|bPfData|bPfFe|bPfSiO2|bPfP|bPfSE"
    CamposBatch = Split(c, "|")
End Function

' ---------------------------------------------------------------- tabelas
Private Function PrimeiraLinhaH() As Long
    PrimeiraLinhaH = H_ROW_HDR + 1
End Function

' Ultima linha com batch na tabela tbl_Batches (PrimeiraLinhaH - 1 se vazia)
Private Function UltimaLinhaH() As Long
    Dim r As Long
    r = PrimeiraLinhaH()
    Do While Not Vazio(shHist.Cells(r, H_COL1).Value)
        r = r + 1
    Loop
    UltimaLinhaH = r - 1
End Function

Private Function UltimaLinhaL() As Long
    Dim r As Long
    r = 2
    Do While Not Vazio(shLeituras.Cells(r, 1).Value)
        r = r + 1
    Loop
    UltimaLinhaL = r - 1
End Function

' Linha do batch id em tbl_Batches (0 = nao existe)
Public Function LinhaHistorico(ByVal id As Variant) As Long
    Dim r As Long, ult As Long, v As Variant
    If Not ENumero(id) Then Exit Function
    ult = UltimaLinhaH()
    For r = PrimeiraLinhaH() To ult
        v = shHist.Cells(r, H_COL1).Value
        If ENumero(v) Then
            If CDbl(v) = CDbl(id) Then
                LinhaHistorico = r
                Exit Function
            End If
        End If
    Next r
End Function

' Ajusta o tamanho da tabela do Excel (no LibreOffice nao ha ListObject: ignora)
Private Sub AjustarTabela(ByVal ws As Worksheet, ByVal nome As String, ByVal rHdr As Long, ByVal c1 As Long, _
                          ByVal nCol As Long, ByVal rUlt As Long)
    Dim lo As Object
    On Error Resume Next
    Set lo = ws.ListObjects(nome)
    If lo Is Nothing Then Exit Sub
    If rUlt <= rHdr Then rUlt = rHdr + 1
    lo.Resize ws.Range(ws.Cells(rHdr, c1), ws.Cells(rUlt, c1 + nCol - 1))
    Err.Clear
End Sub

Private Sub ExcluirLeituras(ByVal id As Variant)
    Dim r As Long, v As Variant
    For r = UltimaLinhaL() To 2 Step -1
        v = shLeituras.Cells(r, 1).Value
        If ENumero(v) Then
            If CDbl(v) = CDbl(id) Then shLeituras.Rows(r).Delete
        End If
    Next r
End Sub

' ---------------------------------------------------------------- salvar
Public Sub SalvarHistorico()
    Dim id As Variant, r As Long, novo As Boolean
    Recalcular
    id = Nm("bId").Value
    If Not ENumero(id) Then
        Aviso "Informe o número do batch (aba Batch).", vbExclamation, "Salvar no histórico"
        Exit Sub
    End If
    r = LinhaHistorico(id)
    If r > 0 And Not gSilencioso Then
        If Aviso("O batch " & id & " já está no histórico. Substituir pelos dados atuais?", _
                 vbQuestion + vbYesNo, "Salvar no histórico") <> vbYes Then Exit Sub
    End If
    novo = (r = 0)
    GravarHistorico id
    gAlterado = False
    Aviso "Batch " & id & IIf(novo, " salvo", " atualizado") & " no histórico (" & ContarLeituras() & _
          " leituras).", vbInformation, "Salvar no histórico"
End Sub

Private Function ContarLeituras() As Long
    Dim p As Long, lab As Long
    For p = 1 To NP
        For lab = 0 To 1
            ContarLeituras = ContarLeituras + Application.WorksheetFunction.Count( _
                shBatch.Range(shBatch.Cells(B_ROW_G0, ColValor(p, lab)), shBatch.Cells(B_ROW_G0 + NMAX - 1, ColValor(p, lab))))
        Next lab
    Next p
End Function

Public Sub GravarHistorico(ByVal id As Variant)
    Dim r As Long, campos As Variant, j As Long, p As Long, c As Long, v As Variant, nok As Long
    Dim ws As Worksheet, rL As Long, i As Long, lab As Long, a As Variant, b As Variant, tU As Variant, tG As Variant
    Dim tol As Variant, dif As Variant, refA As Boolean, erro As String

    ModoCarga True
    On Error GoTo Fim
    Set ws = shHist
    r = LinhaHistorico(id)
    If r = 0 Then r = UltimaLinhaH() + 1
    campos = CamposBatch()
    For j = LBound(campos) To UBound(campos)
        v = Nm(CStr(campos(j))).Value
        If IsError(v) Then v = Empty
        c = H_COL1 + j - LBound(campos)
        If Vazio(v) Then
            ws.Cells(r, c).ClearContents
        ElseIf VarType(v) = vbString Then
            ws.Cells(r, c).Value = "'" & v
        Else
            ws.Cells(r, c).Value = v
        End If
    Next j
    c = H_COL1 + H_COL_STAT1 - 1
    For p = 1 To NP
        For j = 0 To 3
            v = shGraf.Cells(G_ROW_ST + p - 1, Choose(j + 1, G_COL_MA, G_COL_MB, G_COL_D, G_COL_ST)).Value
            If IsError(v) Then v = Empty
            If Vazio(v) Then
                ws.Cells(r, c).ClearContents
            Else
                ws.Cells(r, c).Value = v
            End If
            If j = 3 And CStr(v) = "NOK" Then nok = nok + 1
            c = c + 1
        Next j
    Next p
    ws.Cells(r, c).Value = nok
    ws.Cells(r, c + 1).Value = ContarLeituras()
    ws.Cells(r, c + 2).Value = Now
    ws.Cells(r, c + 3).Value = "'" & Usuario()
    AjustarTabela ws, "tbl_Batches", H_ROW_HDR, H_COL1, H_NCOL, UltimaLinhaH()

    ' leituras: apaga as antigas do batch e grava as atuais
    ExcluirLeituras id
    rL = UltimaLinhaL() + 1
    refA = (CfgTxt("bRef") <> CfgTxt("bLabB"))
    For i = 1 To NMAX
        tU = shBatch.Cells(B_ROW_G0 + i - 1, B_COL_TU).Value
        tG = shBatch.Cells(B_ROW_G0 + i - 1, B_COL_TG).Value
        For p = 1 To NP
            a = shBatch.Cells(B_ROW_G0 + i - 1, ColValor(p, 0)).Value
            b = shBatch.Cells(B_ROW_G0 + i - 1, ColValor(p, 1)).Value
            If ENumero(a) Or ENumero(b) Then
                tol = Nm("bTol_" & PChave(p)).Value
                shLeituras.Cells(rL, 1).Value = id
                shLeituras.Cells(rL, 2).Value = i
                If EData(tU) Then shLeituras.Cells(rL, 3).Value = CDate(tU)
                If EData(tG) Then shLeituras.Cells(rL, 4).Value = CDate(tG)
                shLeituras.Cells(rL, 5).Value = "'" & PRotulo(p)
                If ENumero(a) Then shLeituras.Cells(rL, 6).Value = CDbl(a)
                If ENumero(b) Then shLeituras.Cells(rL, 7).Value = CDbl(b)
                If ENumero(a) And ENumero(b) Then
                    If refA Then dif = CDbl(b) - CDbl(a) Else dif = CDbl(a) - CDbl(b)
                    shLeituras.Cells(rL, 8).Value = dif
                    If ENumero(tol) Then
                        shLeituras.Cells(rL, 10).Value = IIf(Abs(Round(dif, PDec(p))) <= CDbl(tol) + 0.000000001, "OK", "NOK")
                    End If
                End If
                If ENumero(tol) Then shLeituras.Cells(rL, 9).Value = CDbl(tol)
                rL = rL + 1
            End If
        Next p
    Next i
    AjustarTabela shLeituras, "tbl_Leituras", 1, 1, L_NCOL, UltimaLinhaL()
    Err.Clear
Fim:
    erro = Err.Description
    ModoCarga False
    If erro <> "" Then Aviso "Erro ao gravar o histórico: " & erro, vbExclamation, "Salvar no histórico"
End Sub

Private Function Usuario() As String
    On Error Resume Next
    Usuario = Application.UserName
    If Usuario = "" Then Usuario = Environ$("USERNAME")
End Function

' ---------------------------------------------------------------- abrir
Public Sub AbrirBatch(ByVal id As Variant)
    Dim r As Long, campos As Variant, j As Long, v As Variant, campo As String, rL As Long, i As Long
    Dim p As Long, lab As Long, tU As Variant, c As Range, prev As Variant, erro As String

    r = LinhaHistorico(id)
    If r = 0 Then
        Aviso "Batch " & id & " não encontrado no histórico.", vbExclamation, "Abrir batch"
        Exit Sub
    End If
    ModoCarga True
    On Error GoTo Fim
    Desproteger shBatch
    campos = CamposBatch()
    For j = LBound(campos) To UBound(campos)
        campo = CStr(campos(j))
        v = shHist.Cells(r, H_COL1 + j - LBound(campos)).Value
        If Vazio(v) Then
            Nm(campo).ClearContents
        ElseIf VarType(v) = vbString Then
            Nm(campo).Value = "'" & v
        Else
            Nm(campo).Value = v
        End If
    Next j
    ' leituras
    shBatch.Range(shBatch.Cells(B_ROW_G0, B_COL_V1), shBatch.Cells(B_ROW_G0 + NMAX - 1, B_COL_V1 + 2 * NP - 1)).ClearContents
    RestaurarHorarios
    Recalcular
    For rL = 2 To UltimaLinhaL()
        v = shLeituras.Cells(rL, 1).Value
        If ENumero(v) Then
            If CDbl(v) = CDbl(id) Then
                i = CLng(shLeituras.Cells(rL, 2).Value)
                p = PIndice(CStr(shLeituras.Cells(rL, 5).Value))
                If i >= 1 And i <= NMAX And p > 0 Then
                    tU = shLeituras.Cells(rL, 3).Value
                    Set c = shBatch.Cells(B_ROW_G0 + i - 1, B_COL_TU)
                    If EData(tU) Then
                        prev = c.Value
                        If Not EData(prev) Then
                            c.Value = CDate(tU)
                        ElseIf Abs(CDbl(prev) - CDbl(tU)) > 1 / 1440 Then
                            c.Value = CDate(tU)
                        End If
                    End If
                    For lab = 0 To 1
                        v = shLeituras.Cells(rL, 6 + lab).Value
                        If ENumero(v) Then shBatch.Cells(B_ROW_G0 + i - 1, ColValor(p, lab)).Value = CDbl(v)
                    Next lab
                End If
            End If
        End If
    Next rL
    Nm("pSel").Value = CDbl(id)
    Err.Clear
Fim:
    erro = Err.Description
    Proteger shBatch
    ModoCarga False
    If erro <> "" Then
        Aviso "Erro ao abrir o batch: " & erro, vbExclamation, "Abrir batch"
        Exit Sub
    End If
    gAlterado = False
    Recalcular
    AjustarGraficos
End Sub

' Pergunta se salva o batch atual (alterado) antes de trocar. False = cancelar.
Private Function ConfirmarTroca() As Boolean
    Dim resp As Long
    ConfirmarTroca = True
    If Not gAlterado Or gSilencioso Then Exit Function
    resp = Aviso("O batch " & CfgTxt("bId") & " foi alterado e não foi salvo." & vbCrLf & vbCrLf & _
                 "Salvar no histórico antes de continuar?", vbQuestion + vbYesNoCancel, TITULO_APP)
    If resp = vbCancel Then
        ConfirmarTroca = False
    ElseIf resp = vbYes Then
        If ENumero(Nm("bId").Value) Then GravarHistorico Nm("bId").Value
        gAlterado = False
    End If
End Function

' Painel: lista "Batch em exibicao"
Public Sub SelecionarBatch()
    Dim id As Variant
    id = Nm("pSel").Value
    If Not ENumero(id) Then Exit Sub
    If CStr(id) = CfgTxt("bId") And Not gAlterado Then
        AjustarGraficos
        Exit Sub
    End If
    If LinhaHistorico(id) = 0 Then
        Aviso "Batch " & id & " não está no histórico.", vbExclamation, "Batch em exibição"
        ModoCarga True
        Nm("pSel").Value = Nm("bId").Value
        ModoCarga False
        Exit Sub
    End If
    If Not ConfirmarTroca() Then
        ModoCarga True
        Nm("pSel").Value = Nm("bId").Value
        ModoCarga False
        Exit Sub
    End If
    AbrirBatch id
End Sub

Private Function BatchDaSelecao() As Variant
    Dim r As Long
    On Error Resume Next
    r = ActiveCell.Row
    If ActiveSheet.Name <> shHist.Name Or r < PrimeiraLinhaH() Or r > UltimaLinhaH() Then
        BatchDaSelecao = Empty
    Else
        BatchDaSelecao = shHist.Cells(r, H_COL1).Value
    End If
End Function

Public Sub AbrirBatchSelecionado()
    Dim id As Variant
    id = BatchDaSelecao()
    If Not ENumero(id) Then
        Aviso "Clique numa célula da linha do batch (tabela abaixo) e use o botão de novo.", vbInformation, "Abrir batch"
        Exit Sub
    End If
    If Not ConfirmarTroca() Then Exit Sub
    AbrirBatch id
    IrPainel
End Sub

Public Sub ExcluirBatchSelecionado()
    Dim id As Variant, r As Long
    id = BatchDaSelecao()
    If Not ENumero(id) Then
        Aviso "Clique numa célula da linha do batch (tabela abaixo) e use o botão de novo.", vbInformation, "Excluir batch"
        Exit Sub
    End If
    If Aviso("Excluir o batch " & id & " do histórico (e as leituras dele)?", vbQuestion + vbYesNo, _
             "Excluir batch") <> vbYes Then Exit Sub
    ExcluirDoHistorico id
End Sub

Public Sub ExcluirDoHistorico(ByVal id As Variant)
    Dim r As Long
    ModoCarga True
    r = LinhaHistorico(id)
    If r > 0 Then shHist.Rows(r).Delete
    AjustarTabela shHist, "tbl_Batches", H_ROW_HDR, H_COL1, H_NCOL, UltimaLinhaH()
    ExcluirLeituras id
    AjustarTabela shLeituras, "tbl_Leituras", 1, 1, L_NCOL, UltimaLinhaL()
    ModoCarga False
End Sub

' ---------------------------------------------------------------- novo / limpar
Public Sub NovoBatch()
    Dim novo As Long, r As Long, v As Variant, p As Long, i As Long
    If Not ConfirmarTroca() Then Exit Sub
    novo = 0
    For r = PrimeiraLinhaH() To UltimaLinhaH()
        v = shHist.Cells(r, H_COL1).Value
        If ENumero(v) Then
            If CLng(v) > novo Then novo = CLng(v)
        End If
    Next r
    If ENumero(Nm("bId").Value) Then
        If CLng(Nm("bId").Value) > novo Then novo = CLng(Nm("bId").Value)
    End If
    novo = novo + 1
    ModoCarga True
    Desproteger shBatch
    Nm("bId").Value = novo
    Nm("bData").Value = Date
    Nm("bMin").Formula = "=padMin"
    Nm("bLabA").Formula = "=padLabA"
    Nm("bLabB").Formula = "=padLabB"
    Nm("bRef").Formula = "=padRef"
    Nm("bTit").Formula = "=padTit"
    Nm("bSub").Formula = "=padSub"
    Nm("bIniU").ClearContents
    Nm("bIniG").ClearContents
    Nm("bNHor").Formula = "=padNHor"
    For p = 1 To NP
        Nm("bTol_" & PChave(p)).Formula = "=padTol_" & PChave(p)
    Next p
    Nm("bIncFrase").Value = "Sim"
    Nm("bCom1").Formula = "=padCom"
    For i = 2 To 6
        Nm("bCom" & i).ClearContents
    Next i
    Nm("bPfData").ClearContents
    Nm("bPfFe").ClearContents
    Nm("bPfSiO2").ClearContents
    Nm("bPfP").ClearContents
    Nm("bPfSE").ClearContents
    shBatch.Range(shBatch.Cells(B_ROW_G0, B_COL_V1), shBatch.Cells(B_ROW_G0 + NMAX - 1, B_COL_V1 + 2 * NP - 1)).ClearContents
    RestaurarHorarios
    Nm("pSel").Value = novo
    Proteger shBatch
    ModoCarga False
    gAlterado = False
    AjustarGraficos
    On Error Resume Next
    shBatch.Activate
    Nm("bIniU").Select
    On Error GoTo 0
    Aviso "Novo batch " & novo & " em branco." & vbCrLf & vbCrLf & _
          "Informe o início no processo em Ubu (e em Germano) e use Buscar no MES.", vbInformation, "Novo batch"
End Sub

Public Sub LimparLeituras()
    If Aviso("Apagar todas as leituras do batch " & CfgTxt("bId") & " (só nesta aba; o histórico não muda)?", _
             vbQuestion + vbYesNo, "Limpar leituras") <> vbYes Then Exit Sub
    ModoCarga True
    Desproteger shBatch
    shBatch.Range(shBatch.Cells(B_ROW_G0, B_COL_V1), shBatch.Cells(B_ROW_G0 + NMAX - 1, B_COL_V1 + 2 * NP - 1)).ClearContents
    RestaurarHorarios
    Proteger shBatch
    ModoCarga False
    gAlterado = True
    AjustarGraficos
End Sub

' ---------------------------------------------------------------- importar da planilha antiga
Public Sub LimparImportacao()
    Nm("impArea").ClearContents
End Sub

' Parametro pelo rotulo da planilha antiga ("S.E Germano", "< 325 # Alimentação 04", "P Germano"...)
Private Function ParamDoRotulo(ByVal rot As String) As Long
    Dim s As String
    s = LCase$(rot)
    s = Replace(s, "germano", "")
    s = Replace(s, "alimentação", "")
    s = Replace(s, "alimentacao", "")
    s = Replace(s, "ubu", "")
    s = Replace(s, "04", "")
    s = Replace(s, "03", "")
    s = Trim$(s)
    If InStr(s, "325") > 0 Then
        ParamDoRotulo = PIndice("325")
    ElseIf InStr(s, "sio2") > 0 Or InStr(s, "sílica") > 0 Or InStr(s, "silica") > 0 Then
        ParamDoRotulo = PIndice("SiO2")
    ElseIf InStr(s, "ppc") > 0 Then
        ParamDoRotulo = PIndice("PPC")
    ElseIf Left$(s, 3) = "s.e" Or Left$(s, 2) = "se" Or InStr(s, "superf") > 0 Then
        ParamDoRotulo = PIndice("SE")
    ElseIf s = "p" Or Left$(s, 2) = "p " Or InStr(s, "fósforo") > 0 Or InStr(s, "fosforo") > 0 Then
        ParamDoRotulo = PIndice("P")
    End If
End Function

Public Sub ImportarColado()
    Dim area As Range, r As Long, c As Long, nC As Long, rot As String, low As String, id As Variant
    Dim datas() As Variant, chaves() As Double, colChave() As Long, nCh As Long, k As Long, lab As Long, p As Long
    Dim vals() As Variant, v As Variant, t As Double, achou As Boolean, i As Long, qtd As Long, novo As Boolean

    Set area = Nm("impArea")
    nC = area.Columns.Count
    ReDim datas(1 To nC)
    ReDim colChave(1 To nC)
    ReDim chaves(1 To NMAX)
    ReDim vals(1 To NMAX, 1 To NP, 0 To 1)
    For r = 1 To area.Rows.Count
        rot = Trim$(CStr(area.Cells(r, 1).Text))
        low = LCase$(rot)
        If low = "" Then
        ElseIf Left$(low, 5) = "batch" Then
            For c = 2 To nC
                If ENumero(area.Cells(r, c).Value) Then
                    id = CLng(area.Cells(r, c).Value)
                    Exit For
                End If
            Next c
        ElseIf Left$(low, 4) = "data" Then
            For c = 2 To nC
                datas(c) = area.Cells(r, c).Value
            Next c
        ElseIf InStr(low, "inicio") > 0 Or InStr(low, "início") > 0 Or InStr(low, "horário") > 0 Or _
               InStr(low, "horario") > 0 Then
            For c = 2 To nC
                colChave(c) = 0
                v = area.Cells(r, c).Value
                If ENumero(v) And EData(datas(c)) Then
                    t = Int(CDbl(datas(c))) + (CDbl(v) - Int(CDbl(v)))
                    achou = False
                    For k = 1 To nCh
                        If Abs(chaves(k) - t) < 1 / 1440 Then
                            colChave(c) = k
                            achou = True
                            Exit For
                        End If
                    Next k
                    If Not achou And nCh < NMAX Then
                        nCh = nCh + 1
                        chaves(nCh) = t
                        colChave(c) = nCh
                    End If
                End If
            Next c
        ElseIf InStr(low, "diferen") > 0 Or InStr(low, "retorno") > 0 Or InStr(low, "legenda") > 0 Then
        Else
            lab = -1
            If InStr(low, "germano") > 0 Then
                lab = 0
            ElseIf InStr(low, "aliment") > 0 Or InStr(low, "ubu") > 0 Then
                lab = 1
            End If
            p = ParamDoRotulo(rot)
            If lab >= 0 And p > 0 Then
                For c = 2 To nC
                    If colChave(c) > 0 Then
                        v = area.Cells(r, c).Value
                        If ENumero(v) Then
                            vals(colChave(c), p, lab) = CDbl(v)
                            qtd = qtd + 1
                        End If
                    End If
                Next c
            End If
        End If
    Next r
    If nCh = 0 Or qtd = 0 Then
        Aviso "Nada reconhecido. Cole a partir da célula B7 as linhas Batch, Data, Inicio aproximado no processo " & _
              "e as linhas das análises (Germano / Alimentação 04).", vbExclamation, "Importar"
        Exit Sub
    End If
    If Not ENumero(id) Then id = Nm("bId").Value
    novo = (CStr(id) <> CfgTxt("bId"))
    If novo Then
        If Not ConfirmarTroca() Then Exit Sub
    End If
    ' ordena os horarios (as colunas da planilha ja vem em ordem, mas por garantia)
    OrdenarHorarios chaves, vals, nCh
    ModoCarga True
    Desproteger shBatch
    Nm("bId").Value = id
    Nm("bIniU").Value = CDate(chaves(1))
    Nm("bIniG").ClearContents
    Nm("bNHor").Value = nCh
    If Not EData(Nm("bData").Value) Or novo Then Nm("bData").Value = CDate(Int(chaves(1)))
    shBatch.Range(shBatch.Cells(B_ROW_G0, B_COL_V1), shBatch.Cells(B_ROW_G0 + NMAX - 1, B_COL_V1 + 2 * NP - 1)).ClearContents
    RestaurarHorarios
    Recalcular
    For i = 1 To nCh
        v = shBatch.Cells(B_ROW_G0 + i - 1, B_COL_TU).Value
        If Not EData(v) Then
            shBatch.Cells(B_ROW_G0 + i - 1, B_COL_TU).Value = CDate(chaves(i))
        ElseIf Abs(CDbl(v) - chaves(i)) > 1 / 1440 Then
            shBatch.Cells(B_ROW_G0 + i - 1, B_COL_TU).Value = CDate(chaves(i))
        End If
        For p = 1 To NP
            For lab = 0 To 1
                If ENumero(vals(i, p, lab)) Then shBatch.Cells(B_ROW_G0 + i - 1, ColValor(p, lab)).Value = vals(i, p, lab)
            Next lab
        Next p
    Next i
    Nm("pSel").Value = id
    Proteger shBatch
    ModoCarga False
    gAlterado = True
    AjustarGraficos
    IrBatch
    Aviso "Batch " & id & " importado: " & nCh & " horários, " & qtd & " leituras." & vbCrLf & vbCrLf & _
          "Confira a aba Batch e use Salvar no histórico.", vbInformation, "Importar"
End Sub

Private Sub OrdenarHorarios(ByRef chaves() As Double, ByRef vals() As Variant, ByVal n As Long)
    Dim i As Long, j As Long, p As Long, lab As Long, t As Double, tv As Variant
    For i = 1 To n - 1
        For j = i + 1 To n
            If chaves(j) < chaves(i) Then
                t = chaves(i)
                chaves(i) = chaves(j)
                chaves(j) = t
                For p = 1 To NP
                    For lab = 0 To 1
                        tv = vals(i, p, lab)
                        vals(i, p, lab) = vals(j, p, lab)
                        vals(j, p, lab) = tv
                    Next lab
                Next p
            End If
        Next j
    Next i
End Sub

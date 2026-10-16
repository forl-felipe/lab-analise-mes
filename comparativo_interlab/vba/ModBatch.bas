Option Explicit

' ============================================================================
'  ModBatch - historico de batches (tbl_Batches + tbl_Leituras), novo batch,
'  abrir/excluir e importacao da planilha antiga.
'  tbl_Batches (aba Historico): um batch por linha. tbl_Leituras (aba BD_Leituras):
'  uma linha por batch x horario x analise. As duas alimentam o Power BI.
' ============================================================================

' Campos da aba Batch na ordem das primeiras colunas de tbl_Batches
Private Function CamposBatch() As Variant
    CamposBatch = Split(CAMPOS_BATCH(), "|")
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
    Dim id As Variant, r As Long, novo As Boolean, n As Long
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
    Desproteger shBatch
    If Not EData(Nm("bData").Value) Then
        Nm("bData").Value = Date
    ElseIf Nm("bData").HasFormula Then
        Nm("bData").Value = CDate(Nm("bData").Value)
    End If
    GravarHistorico id
    gAlterado = False
    n = ContarLeituras()
    ' a aba Batch fica pronta para o proximo batch, com as tolerancias e padroes da aba Configuracoes
    LimparBatch
    Aviso "Batch " & id & IIf(novo, " salvo", " atualizado") & " no histórico (" & n & " leituras)." & vbCrLf & vbCrLf & _
          "A aba Batch já está pronta para o próximo batch. Para ver o " & id & " de novo: Painel > Batch em exibição.", _
          vbInformation, "Salvar no histórico"
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
    v = Nm("bIniP").Value
    If EData(v) Then ws.Cells(r, c).Value = CDate(v) Else ws.Cells(r, c).ClearContents
    v = Nm("bDef").Value
    If ENumero(v) Then ws.Cells(r, c + 1).Value = CDbl(v) Else ws.Cells(r, c + 1).ClearContents
    ws.Cells(r, c + 2).Value = nok
    ws.Cells(r, c + 3).Value = ContarLeituras()
    ws.Cells(r, c + 4).Value = Now
    ws.Cells(r, c + 5).Value = "'" & Usuario()
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

' Frase dos parametros fora e acompanhamento: formula automatica (todos = tambem os ja editados)
Public Sub RestaurarComentariosAuto(ByVal todos As Boolean)
    If todos Or Vazio(Nm("bFrase").Value) Then Nm("bFrase").Formula = "=gFrase"
    If todos Or Vazio(Nm("bAcomp").Value) Then Nm("bAcomp").Formula = "=gAcomp"
End Sub

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
            Nm(campo).MergeArea.ClearContents
        ElseIf VarType(v) = vbString Then
            Nm(campo).Value = "'" & v
        Else
            Nm(campo).Value = v
        End If
    Next j
    RestaurarComentariosAuto False
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
    If Not ConfirmarTroca() Then Exit Sub
    LimparBatch
    On Error Resume Next
    shBatch.Activate
    Nm("bId").Select
End Sub

' Deixa a aba Batch em branco para o proximo batch (numero em branco; tolerancias, laboratorios,
' titulos e comentarios voltam aos padroes da aba Configuracoes)
Public Sub LimparBatch()
    Dim p As Long, i As Long
    ModoCarga True
    Desproteger shBatch
    Nm("bId").MergeArea.ClearContents
    Nm("bData").Value = Date
    Nm("bMin").Formula = "=padMin"
    Nm("bLabA").Formula = "=padLabA"
    Nm("bLabB").Formula = "=padLabB"
    Nm("bRef").Formula = "=padRef"
    Nm("bTit").Formula = "=padTit"
    Nm("bSub").Formula = "=padSub"
    Nm("bIniU").MergeArea.ClearContents
    Nm("bIniG").MergeArea.ClearContents
    Nm("bFimG").MergeArea.ClearContents
    Nm("bFimU").MergeArea.ClearContents
    Nm("bNivel").MergeArea.ClearContents
    Nm("bIniPMan").MergeArea.ClearContents
    Nm("bIncAcomp").Value = "Sim"
    RestaurarComentariosAuto True
    Nm("bFasesMsg").MergeArea.ClearContents
    Nm("bNHor").Formula = "=IF(ISNUMBER(gNGer),gNGer,IF(ISNUMBER(gNHorFim),gNHorFim,padNHor))"
    For p = 1 To NP
        Nm("bTol_" & PChave(p)).Formula = "=padTol_" & PChave(p)
    Next p
    Nm("bIncFrase").Value = "Sim"
    Nm("bCom1").Formula = "=padCom"
    For i = 2 To 6
        Nm("bCom" & i).MergeArea.ClearContents
    Next i
    shBatch.Range(shBatch.Cells(B_ROW_G0, B_COL_V1), shBatch.Cells(B_ROW_G0 + NMAX - 1, B_COL_V1 + 2 * NP - 1)).ClearContents
    RestaurarHorarios
    Nm("pSel").MergeArea.ClearContents
    Proteger shBatch
    ModoCarga False
    gAlterado = False
    AjustarGraficos
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






' ---------------------------------------------------------------- colar as Fases do Evento do MES
' Le as linhas coladas em bFases (Fase, Descricao, Inicio, Termino - com ou sem cabecalho) e preenche
' inicio/termino de Germano (fase cfgFaseG, ex.: EB06) e de Ubu (fase que comeca com cfgFaseU, ex.: UBU3).
Public Sub LerFases()
    Dim area As Range, r As Long, c As Long, v As Variant, cod As String, d1 As Variant, d2 As Variant
    Dim faseG As String, faseU As String, gIni As Variant, gFim As Variant, uIni As Variant, uFim As Variant
    Dim gCod As String, uCod As String, primIni As Variant, primFim As Variant, primCod As String, msg As String
    Dim t As String, dd As Variant

    Set area = Nm("bFases")
    faseG = UCase$(CfgTxt("cfgFaseG"))
    faseU = UCase$(CfgTxt("cfgFaseU"))
    If faseU = "" Then faseU = "UBU"
    For r = 1 To area.Rows.Count
        cod = ""
        d1 = Empty
        d2 = Empty
        For c = 1 To area.Columns.Count
            v = area.Cells(r, c).Value
            dd = Empty
            If EData(v) Then
                dd = CDate(v)
            ElseIf VarType(v) = vbString Then
                t = Trim$(v)
                If t <> "" And InStr(t, "/") > 0 And IsDate(t) Then
                    dd = CDate(t)
                ElseIf cod = "" And t <> "" And InStr(t, " ") = 0 And Len(t) <= 8 Then
                    cod = UCase$(t)
                End If
            End If
            If Not IsEmpty(dd) Then
                If IsEmpty(d1) Then
                    d1 = dd
                ElseIf IsEmpty(d2) Then
                    d2 = dd
                End If
            End If
        Next c
        If cod <> "" And Not IsEmpty(d1) Then
            If Left$(cod, Len(faseU)) = faseU Then
                uCod = cod: uIni = d1: uFim = d2
            ElseIf cod = faseG Then
                gCod = cod: gIni = d1: gFim = d2
            ElseIf primCod = "" Then
                primCod = cod: primIni = d1: primFim = d2
            ElseIf CDbl(d1) < CDbl(primIni) Then
                primCod = cod: primIni = d1: primFim = d2
            End If
        End If
    Next r
    ' sem a fase de Germano configurada: usa a primeira fase (mais cedo) que nao e de Ubu
    If gCod = "" And primCod <> "" Then
        gCod = primCod: gIni = primIni: gFim = primFim
    End If
    If gCod = "" And uCod = "" Then
        msg = "Nada reconhecido. Cole as linhas da grade Fases do Evento (Fase, Descrição, Início, Término)."
    Else
        ModoCarga True
        Desproteger shBatch
        If uCod <> "" Then
            Nm("bIniU").Value = uIni
            If Not IsEmpty(uFim) Then Nm("bFimU").Value = uFim
            If IsNumeric(Mid$(uCod, Len(faseU) + 1)) And Len(uCod) > Len(faseU) Then
                Nm("bMin").Value = "'" & Format$(CLng(Mid$(uCod, Len(faseU) + 1)), "00")
            End If
        End If
        If gCod <> "" Then
            Nm("bIniG").Value = gIni
            If Not IsEmpty(gFim) Then Nm("bFimG").Value = gFim
        End If
        area.ClearContents
        msg = "Fases lidas:"
        If gCod <> "" Then msg = msg & "  Germano " & gCod & " " & Format$(gIni, "dd\/mm hh:mm") & _
                                  IIf(IsEmpty(gFim), "", " a " & Format$(gFim, "dd\/mm hh:mm"))
        If uCod <> "" Then msg = msg & "   ·   Ubu " & uCod & " " & Format$(uIni, "dd\/mm hh:mm") & _
                                  IIf(IsEmpty(uFim), "", " a " & Format$(uFim, "dd\/mm hh:mm"))
        Nm("bFasesMsg").Value = "'" & msg
        Proteger shBatch
        ModoCarga False
        gAlterado = True
        Exit Sub
    End If
    ModoCarga True
    Desproteger shBatch
    Nm("bFasesMsg").Value = "'" & msg
    Proteger shBatch
    ModoCarga False
End Sub

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
Public Function PrimeiraLinhaH() As Long
    PrimeiraLinhaH = H_ROW_HDR + 1
End Function

' Ultima linha com batch na tabela tbl_Batches (PrimeiraLinhaH - 1 se vazia)
Public Function UltimaLinhaH() As Long
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
    If ModoAcomp() Then
        Aviso "O tipo da busca está em Acompanhamento (batch em processamento): os dados ficam nas abas " & _
              "Acomp. Químico e Acomp. Físico, atualizadas a cada Buscar no MES." & vbCrLf & vbCrLf & _
              "O histórico é para o comparativo do batch já processado: troque o tipo da busca para Comparativo, " & _
              "busque no MES e salve.", vbInformation, "Salvar no histórico"
        Exit Sub
    End If
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
    RestaurarHorariosG
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
                    ' horario de Germano gravado (resultados em sequencia)
                    tU = shLeituras.Cells(rL, 4).Value
                    Set c = shBatch.Cells(B_ROW_G0 + i - 1, B_COL_TG)
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
    AtualizarUltimos
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
    AtualizarUltimos
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
    Nm("bExtH").MergeArea.ClearContents
    Nm("bIncAcomp").Value = "Sim"
    RestaurarComentariosAuto True
    Nm("bFasesMsg").MergeArea.ClearContents
    ' pellet feed ultra fino: composto do dia anterior a data do report (buscado no MES)
    Nm("bPfData").Formula = "=IF(ISNUMBER(bData),bData-cfgPfDias,"""")"
    For i = 1 To NPF
        Nm("bPf" & Split(PF_CHAVES, "|")(i - 1)).ClearContents
    Next i
    Nm("bPfMsg").MergeArea.ClearContents
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
    RestaurarHorariosG
    Nm("pSel").MergeArea.ClearContents
    Proteger shBatch
    ModoCarga False
    gAlterado = False
    AjustarGraficos
    AtualizarUltimos
End Sub

' Planilha de acompanhamento: limpa a aba de preenchimento para buscar o proximo batch (os blocos da aba de
' acompanhamento ficam)
' Nivel dos tanques OU inicio no processo manual: preencher um apaga o outro (evento da aba de preenchimento)
Public Sub NivelOuManual(ByVal veioNivel As Boolean)
    Dim apagar As Range
    If veioNivel Then
        If Vazio(Nm("bNivel").Value) Then Exit Sub
        Set apagar = Nm("bIniPMan").MergeArea
    Else
        If Vazio(Nm("bIniPMan").Value) Then Exit Sub
        Set apagar = Nm("bNivel").MergeArea
    End If
    If Vazio(apagar.Cells(1, 1).Value) Then Exit Sub
    ModoCarga True
    apagar.ClearContents
    ModoCarga False
End Sub

Public Sub ProximoBatchAcomp()
    If Aviso("Limpar a aba " & shBatch.Name & " para o próximo batch?" & vbCrLf & vbCrLf & _
             "O que foi lançado na aba " & shAcomp.Name & " fica como está.", vbQuestion + vbYesNo, "Próximo batch", 6) <> vbYes Then Exit Sub
    LimparBatch
    On Error Resume Next
    Nm("bId").Select
End Sub

Public Sub LimparLeituras()
    If Aviso("Apagar todas as leituras do batch " & CfgTxt("bId") & " (só nesta aba; o histórico não muda)?", _
             vbQuestion + vbYesNo, "Limpar leituras") <> vbYes Then Exit Sub
    ModoCarga True
    Desproteger shBatch
    shBatch.Range(shBatch.Cells(B_ROW_G0, B_COL_V1), shBatch.Cells(B_ROW_G0 + NMAX - 1, B_COL_V1 + 2 * NP - 1)).ClearContents
    RestaurarHorarios
    RestaurarHorariosG
    Proteger shBatch
    ModoCarga False
    gAlterado = True
    AjustarGraficos
End Sub






' ---------------------------------------------------------------- conferencia colada (Ctrl+V) na etapa 3
' A linha de valores da CONFERENCIA tem o mesmo desenho aqui e na planilha dos tecnicos. Copiada la e colada aqui,
' chegam valores e formulas da outra planilha; vale o texto copiado (o que o tecnico via na tela): ele preenche o
' batch e as formulas daqui (inicio no processo, defasagem, data) sao refeitas.
Public Function LinhaConf() As Range
    Dim r As Long
    r = Nm("bIniG").Row
    Set LinhaConf = shBatch.Range(shBatch.Cells(r, 2), shBatch.Cells(r, 14))
End Function

Public Sub ConferenciaAlterada(ByVal alvo As Range)
    Dim v As Variant
    If alvo.Columns.Count >= 5 Then
        ConferenciaColada
    ElseIf Not Intersect(alvo, Nm("bIniP").MergeArea) Is Nothing Then
        ' digitou o inicio no processo: vira inicio manual (e a formula volta)
        v = Nm("bIniP").Value
        RestaurarLinhaConf False
        If EData(v) Then
            ModoCarga True
            Nm("bIniPMan").Value = CDate(v)
            Nm("bNivel").MergeArea.ClearContents
            ModoCarga False
            Recalcular
        End If
    ElseIf Not Intersect(alvo, Nm("bDef")) Is Nothing Then
        RestaurarLinhaConf False
    End If
End Sub

Public Sub ConferenciaColada()
    Dim msg As String, t As String, c As Long, lin As Range
    msg = LerConferencia(TextoAreaTransferencia())
    If msg = "" Then
        ' sem a area de transferencia: o que ficou nas celulas
        Set lin = LinhaConf()
        For c = 1 To lin.Columns.Count
            t = t & IIf(c > 1, vbTab, "") & lin.Cells(1, c).Text
        Next c
        msg = LerConferencia(t)
    End If
    RestaurarLinhaConf True
    If msg = "" Then
        Aviso "Não reconheci a conferência colada. Na planilha dos técnicos, copie a linha de valores da " & _
              "CONFERÊNCIA (de Germano · início a Mineroduto) e cole na 1ª célula da linha da conferência.", _
              vbExclamation, "Conferência"
    Else
        Aviso msg, vbInformation, "Conferência"
    End If
End Sub

' Formulas e formatos da linha da conferencia depois de um Ctrl+V vindo de outra planilha
Private Sub RestaurarLinhaConf(ByVal comData As Boolean)
    ModoCarga True
    Desproteger shBatch
    On Error Resume Next
    With Nm("bIniP")
        .Formula = "=IF(ISNUMBER(bIniPMan),bIniPMan,IF(ISNUMBER(bIniU),bIniU+N(gAtrasoNivel)/24,""""))"
        .MergeArea.NumberFormat = "dd/mm/yy hh:mm"
    End With
    With Nm("bDef")
        .Formula = "=IF(AND(ISNUMBER(gJanU),ISNUMBER(gJanG)),ROUND((gJanU-gJanG)*24,2),"""")"
        .NumberFormat = "0.0"" h"""
    End With
    If comData Then
        If Not EData(Nm("bData").Value) Or Nm("bData").HasFormula Then Nm("bData").Formula = "=TODAY()"
        Nm("bData").NumberFormat = "dd/mm/yy"
    End If
    LinhaConf().Locked = False
    On Error GoTo 0
    Proteger shBatch
    ModoCarga False
    Recalcular
End Sub

' ---------------------------------------------------------------- colar a CONFERENCIA da planilha dos tecnicos
' Na planilha de acompanhamento (aba Preenchimento) o tecnico ja tem a linha CONFERENCIA: Germano inicio/termino,
' Ubu chegada/termino, inicio no processo, bi-horarios, estendido, data e mineroduto. Copiada (Ctrl+C), o botao
' Colar conferencia le o texto da area de transferencia e preenche o batch daqui.
Public Sub ColarConferencia()
    Dim t As String, msg As String
    t = TextoAreaTransferencia()
    If Trim$(t) = "" Then
        Aviso "Nada copiado. Na planilha dos técnicos (aba Preenchimento), selecione a linha da CONFERÊNCIA " & _
              "(de Germano · início a Mineroduto), Ctrl+C, e clique de novo em Colar conferência.", vbExclamation, _
              "Colar conferência"
        Exit Sub
    End If
    msg = LerConferencia(t)
    If msg = "" Then
        Aviso "Não reconheci a conferência no que foi copiado. Copie a linha de valores da CONFERÊNCIA da planilha " & _
              "dos técnicos (de Germano · início a Mineroduto).", vbExclamation, "Colar conferência"
    Else
        Aviso msg, vbInformation, "Colar conferência"
    End If
End Sub

' Texto da area de transferencia (Ctrl+C no Excel, no aplicativo ou no navegador)
Private Function TextoAreaTransferencia() As String
    Dim d As Object, v As Variant
    On Error Resume Next
    Set d = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")      ' MSForms.DataObject
    If Not d Is Nothing Then
        d.GetFromClipboard
        TextoAreaTransferencia = d.GetText(1)
    End If
    If Len(TextoAreaTransferencia) = 0 Or InStr(TextoAreaTransferencia, "??") > 0 Then
        Err.Clear
        v = CreateObject("htmlfile").parentWindow.clipboardData.GetData("text")
        If Not IsNull(v) And Not IsEmpty(v) Then TextoAreaTransferencia = CStr(v)
    End If
End Function

' Le a linha copiada (colunas separadas por tabulacao) e preenche o batch. Devolve a mensagem; "" = nao reconheceu.
' Posicoes (como na aba Preenchimento dos tecnicos, celulas mescladas vem vazias): 0 Germano inicio, 2 Germano termino,
' 3 Ubu chegada, 5 Ubu termino, 7 inicio no processo, 9 bi-horarios, 10 estendido, 11 data, 12 mineroduto.
Public Function LerConferencia(ByVal texto As String) As String
    Dim linhas As Variant, tk As Variant, i As Long, j As Long, p0 As Long, d As Date, nDatas As Long
    Dim gIni As Variant, gFim As Variant, uIni As Variant, uFim As Variant, ip As Variant, nBi As Variant
    Dim ext As Double, mn As String, t As String, atr As Double, nivel As String, msg As String
    Dim k As Long

    texto = Replace(texto, vbCrLf, vbLf)
    texto = Replace(texto, vbCr, vbLf)
    linhas = Split(texto, vbLf)
    For i = LBound(linhas) To UBound(linhas)
        tk = Split(CStr(linhas(i)), vbTab)
        nDatas = 0
        p0 = -1
        For j = LBound(tk) To UBound(tk)
            If DataBR(CStr(tk(j)), d, True) Then
                nDatas = nDatas + 1
                If p0 < 0 Then p0 = j
            End If
        Next j
        If nDatas >= 2 Then Exit For
    Next i
    If nDatas < 2 Then Exit Function

    gIni = Empty: gFim = Empty: uIni = Empty: uFim = Empty: ip = Empty: nBi = Empty
    If UBound(tk) >= p0 + 7 And Tok(tk, p0 + 1) = "" And Tok(tk, p0 + 4) = "" Then
        ' linha copiada com as celulas mescladas (posicoes fixas: a 2a coluna de cada mescla vem vazia)
        If DataBR(Tok(tk, p0), d, True) Then gIni = d
        If DataBR(Tok(tk, p0 + 2), d, True) Then gFim = d
        If DataBR(Tok(tk, p0 + 3), d, True) Then uIni = d
        If DataBR(Tok(tk, p0 + 5), d, True) Then uFim = d
        If DataBR(Tok(tk, p0 + 7), d, True) Then ip = d
        t = Tok(tk, p0 + 9)
        If IsNumeric(t) And t <> "" Then nBi = CLng(t)
        t = Tok(tk, p0 + 10)
        If Left$(t, 1) = "+" Then ext = Val(Mid$(t, 2))
        mn = Tok(tk, p0 + 12)
    Else
        ' sem as celulas vazias: pela ordem (4 datas = sem o termino em Ubu)
        k = 0
        For j = LBound(tk) To UBound(tk)
            t = Trim$(CStr(tk(j)))
            If DataBR(t, d, True) Then
                Select Case k
                    Case 0: gIni = d
                    Case 1: gFim = d
                    Case 2: uIni = d
                    Case 3: uFim = d
                    Case 4: ip = d
                End Select
                k = k + 1
            ElseIf Left$(t, 1) = "+" Then
                ext = Val(Mid$(t, 2))
            ElseIf IsNumeric(t) And t <> "" And InStr(t, "/") = 0 And k >= 3 Then
                If IsEmpty(nBi) Then nBi = CLng(t) Else mn = t
            End If
        Next j
        If k = 4 Then ip = uFim: uFim = Empty
    End If
    If IsEmpty(gIni) And IsEmpty(uIni) Then Exit Function

    ModoCarga True
    Desproteger shBatch
    Nm("bIniG").MergeArea.ClearContents
    Nm("bFimG").MergeArea.ClearContents
    Nm("bIniU").MergeArea.ClearContents
    Nm("bFimU").MergeArea.ClearContents
    If Not IsEmpty(gIni) Then Nm("bIniG").Value = gIni
    If Not IsEmpty(gFim) Then Nm("bFimG").Value = gFim
    If Not IsEmpty(uIni) Then Nm("bIniU").Value = uIni
    If Not IsEmpty(uFim) Then Nm("bFimU").Value = uFim
    ' inicio no processo: se bate com chegada + atraso de um nivel dos tanques, fica o nivel; senao, inicio manual
    Nm("bNivel").MergeArea.ClearContents
    Nm("bIniPMan").MergeArea.ClearContents
    If Not IsEmpty(ip) Then
        If Not IsEmpty(uIni) Then
            atr = Round((CDbl(ip) - CDbl(uIni)) * 24, 2)
            nivel = NivelDoAtraso(atr)
        End If
        If nivel <> "" Then
            Nm("bNivel").Value = nivel
        Else
            Nm("bIniPMan").Value = ip
        End If
    End If
    If Not IsEmpty(nBi) Then
        If nBi >= 1 And nBi <= NMAX Then Nm("bNHor").Value = nBi
    End If
    Nm("bExtH").MergeArea.ClearContents
    If ext > 0 Then Nm("bExtH").Value = ext
    If mn <> "" Then Nm("bMin").Value = "'" & Format$(Val(mn), "00")
    msg = "Conferência colada:" & IIf(IsEmpty(gIni), "", "  Germano " & Format$(gIni, "dd\/mm hh:mm") & _
          IIf(IsEmpty(gFim), "", " a " & Format$(gFim, "dd\/mm hh:mm"))) & _
          IIf(IsEmpty(uIni), "", "   ·   Ubu " & Format$(uIni, "dd\/mm hh:mm") & _
          IIf(IsEmpty(uFim), "", " a " & Format$(uFim, "dd\/mm hh:mm"))) & _
          IIf(IsEmpty(ip), "", "   ·   início " & Format$(ip, "dd\/mm hh:mm") & _
          IIf(nivel <> "", " (nível " & LCase$(nivel) & ")", " (manual)")) & _
          IIf(IsEmpty(nBi), "", "   ·   " & nBi & " bi-horários") & IIf(ext > 0, " (estendido +" & ext & " h)", "")
    Nm("bFasesMsg").Value = "'" & msg
    Proteger shBatch
    ModoCarga False
    gAlterado = True
    Recalcular
    On Error Resume Next
    AtualizarUltimos
    LerConferencia = msg & vbCrLf & vbCrLf & "Agora é só Buscar no MES." & IIf(ext > 0, " Os " & Int(ext / HORAS_SLOT + 0.5) & _
        " bi-horários estendidos recebem a média de Germano, como na planilha dos técnicos.", "")
End Function

Private Function Tok(tk As Variant, ByVal i As Long) As String
    If i >= LBound(tk) And i <= UBound(tk) Then Tok = Trim$(CStr(tk(i)))
End Function

' Nivel dos tanques cujo atraso (Configuracoes) e h horas; "" se nenhum
Private Function NivelDoAtraso(ByVal h As Double) As String
    Dim rN As Range, rH As Range, i As Long
    On Error Resume Next
    Set rN = Nm("cfgAtrasoN")
    Set rH = Nm("cfgAtrasoH")
    For i = 1 To rN.Cells.Count
        If ENumero(rH.Cells(i).Value) And CStr(rN.Cells(i).Value) <> "" Then
            If Abs(CDbl(rH.Cells(i).Value) - h) < 0.02 Then
                NivelDoAtraso = CStr(rN.Cells(i).Value)
                Exit Function
            End If
        End If
    Next i
End Function

' "dd/mm/aa hh:mm", "dd/mm/aaaa hh:mm" ou (comHora = False) so a data; nao depende do idioma do Windows
Private Function DataBR(ByVal t As String, ByRef d As Date, ByVal comHora As Boolean) As Boolean
    Dim pd As Variant, ph As Variant, partes As Variant, a As Long
    t = Trim$(t)
    partes = Split(t, " ")
    If UBound(partes) < IIf(comHora, 1, 0) Then Exit Function
    pd = Split(partes(0), "/")
    If UBound(pd) <> 2 Then Exit Function
    If Not (IsNumeric(pd(0)) And IsNumeric(pd(1)) And IsNumeric(pd(2))) Then Exit Function
    a = CLng(pd(2))
    If a < 100 Then a = a + 2000
    If CLng(pd(1)) < 1 Or CLng(pd(1)) > 12 Or CLng(pd(0)) < 1 Or CLng(pd(0)) > 31 Then Exit Function
    d = DateSerial(a, CLng(pd(1)), CLng(pd(0)))
    If comHora Then
        ph = Split(partes(UBound(partes)), ":")
        If UBound(ph) < 1 Then Exit Function
        If Not (IsNumeric(ph(0)) And IsNumeric(ph(1))) Then Exit Function
        d = d + TimeSerial(CLng(ph(0)), CLng(ph(1)), 0)
    End If
    DataBR = True
End Function

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

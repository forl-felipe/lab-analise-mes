Option Explicit

' por que o ultimo bloco colado nao foi lido (mensagem)
Private mMotivoBloco As String

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
          "A aba " & shBatch.Name & " já está pronta para o próximo batch. Para ver o " & id & " de novo: Painel > Batch em exibição.", _
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
    If VARIANTE = "COMP" Then Nm("bConfCola").Cells(1, 1).Select Else Nm("bId").Select
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
    ' pellet feed ultra fino: data livre; batch novo comeca com o dia anterior (Configuracoes: dias antes do report)
    Nm("bPfData").Value = CDate(Int(CDbl(Date)) - IIf(ENumero(Nm("cfgPfDias").Value), Nm("cfgPfDias").Value, 1))
    For i = 1 To NPF
        Nm("bPf" & Split(PF_CHAVES, "|")(i - 1)).ClearContents
    Next i
    Nm("bPfMsg").MergeArea.ClearContents
    If VARIANTE = "ACOMP" Then
        Nm("bNHor").Formula = "=" & Nm("bNHorAuto").Address(True, True)       ' sem nomes (copiavel)
    Else
        Nm("bNHor").Formula = "=IF(ISNUMBER(gNGer),gNGer,IF(ISNUMBER(gNHorFim),gNHorFim,padNHor))"
    End If
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






' ---------------------------------------------------------------- conferencia colada (Ctrl+V) na linha COLE AQUI
' A linha do batch e copiada na planilha dos tecnicos (tabela CONFERENCIA DOS BATCHES, com o numero do batch, ou a
' linha 3 da conferencia) e colada na linha amarela COLE AQUI, que tem as mesmas mesclas e fica sempre vazia: o VBA
' le o que foi colado, preenche a conferencia do batch (linha de cima) e esvazia a linha de novo.
Public Function LinhaConf() As Range
    Set LinhaConf = Nm("bConfCola")
End Function

Public Sub ConferenciaColada()
    Static ocupado As Boolean
    Dim msg As String, t As String, c As Long, lin As Range
    If ocupado Then Exit Sub                     ' nunca roda por cima de si mesma
    ocupado = True
    On Error GoTo Falha
    Set lin = LinhaConf()
    ' o que ficou nas celulas (o colar do Excel); a propria faixa com o texto dela nao conta
    For c = 1 To lin.Columns.Count
        If Trim$(lin.Cells(1, c).Text) <> "" Then t = t & IIf(t <> "", vbTab, "") & lin.Cells(1, c).Text
    Next c
    If InStr(t, TEXTO_COLA_RESTO) > 0 Then t = ""
    If t <> "" Then
        ' vale o texto copiado (sem formulas da outra planilha); sem ele, o que ficou nas celulas
        msg = LerColado(TextoAreaTransferencia())
        If msg = "" Then msg = LerColado(t)
    End If
    EsvaziarLinhaCola
    ocupado = False
    If t = "" Then Exit Sub
    AvisoConferencia msg
    Exit Sub
Falha:
    ocupado = False
    RestaurarExcel
End Sub

' Ctrl+V deste arquivo (Application.OnKey, ligado enquanto ele esta ativo): nas linhas COLE AQUI (conferencia) e do
' PFUF o Excel nao cola nada - a planilha le so o texto copiado. Assim formulas e nomes da outra planilha nunca
' chegam aqui (sem o aviso "o nome ja existe") e os 4 resultados do PFUF ja ficam empilhados. Fora delas: colar normal.
Public Sub ColarInteligente()
    Dim t As String, msg As String
    On Error GoTo Normal
    If VARIANTE <> "ACOMP" And ActiveSheet Is shBatch Then
        If Not Intersect(Selection, Nm("bConfCola")) Is Nothing Then
            t = TextoAreaTransferencia()
            If Trim$(t) <> "" Then
                msg = LerColado(t)
                EsvaziarLinhaCola
                AvisoConferencia msg
                Exit Sub
            End If
        ElseIf Not Intersect(Selection, Nm("bPfCola")) Is Nothing Then
            t = TextoAreaTransferencia()
            If Trim$(t) <> "" Then
                If ColarPFUFTexto(t) Then Exit Sub
            End If
        End If
    End If
Normal:
    On Error Resume Next
    ActiveSheet.Paste
End Sub

' Liga/desliga o Ctrl+V deste arquivo (ThisWorkbook: ao abrir/ativar liga; ao sair/fechar devolve o normal)
Public Sub LigarColar(ByVal ligar As Boolean)
    On Error Resume Next
    If ligar Then
        Application.OnKey "^v", "'" & ThisWorkbook.Name & "'!ColarInteligente"
    Else
        Application.OnKey "^v"
    End If
End Sub

' O que foi copiado: o bloco inteiro do batch (aba Batch dos tecnicos) ou so a linha da conferencia
Public Function LerColado(ByVal t As String) As String
    If InStr(1, t, "HOR", vbTextCompare) > 0 And InStr(1, t, "Germano", vbTextCompare) > 0 And _
       InStr(1, t, "DATA", vbTextCompare) > 0 Then
        LerColado = ColarBloco(t)
    Else
        LerColado = LerConferencia(t)
    End If
End Function

Private Sub AvisoConferencia(ByVal msg As String)
    If msg = "" Then
        Aviso "Não reconheci o que foi colado. Na planilha dos técnicos, aba Batch, selecione o bloco do batch " & _
              "inteiro (da faixa BATCH até a última linha da grade), Ctrl+C, e aqui Ctrl+V na linha amarela COLE AQUI." & _
              IIf(mMotivoBloco <> "", vbCrLf & vbCrLf & "(" & mMotivoBloco & ")", ""), vbExclamation, "Colar batch"
    Else
        Aviso msg, vbInformation, "Colar batch"
    End If
End Sub

' Deixa a linha COLE AQUI vazia e com o formato dela (o Ctrl+V traz o formato da outra planilha)
Private Sub EsvaziarLinhaCola()
    Dim lin As Range
    On Error Resume Next
    ModoCarga True
    Desproteger shBatch
    Set lin = LinhaConf()
    lin.ClearContents
    If lin.Columns.Count > 1 And Not lin.Cells(1, 1).MergeCells Then lin.Merge
    lin.Cells(1, 1).Value = ChrW(9654) & "   " & TEXTO_COLA_RESTO
    lin.NumberFormat = "General"
    lin.Interior.Color = RGB(255, 242, 204)
    lin.Font.Color = RGB(0, 38, 67)
    lin.Font.Bold = True
    lin.Font.Size = 12
    lin.HorizontalAlignment = -4108
    lin.Locked = False
    Proteger shBatch
    ModoCarga False
End Sub

' ---------------------------------------------------------------- conferencia direto da planilha dos tecnicos
' Com a planilha de acompanhamento (tecnicos) aberta no mesmo Excel, a conferencia do batch vem do registro dela
' (aba _Acomp: chegada/termino em Ubu, nivel, inicio no processo, Germano, bi-horarios) e a extensao do bloco do
' batch (ou da coluna Estendido do registro). Nao precisa copiar nada.
Public Function TrazerConferencia(ByVal id As Variant, Optional ByVal avisar As Boolean = False) As String
    Dim wb As Workbook, ws As Worksheet, r As Long, v As Variant, ext As Double, nivel As String, msg As String
    If Not ENumero(id) Then Exit Function
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            Set ws = Nothing
            On Error Resume Next
            Set ws = wb.Worksheets("_Acomp")
            On Error GoTo 0
            If Not ws Is Nothing Then
                If CStr(ws.Cells(1, 2).Value) = "Chegada Ubu (UBU3)" Then
                    r = 2
                    Do While Not Vazio(ws.Cells(r, 1).Value) And r < 5000
                        v = ws.Cells(r, 1).Value
                        If ENumero(v) Then
                            If CDbl(v) = CDbl(id) Then Exit Do
                        End If
                        r = r + 1
                    Loop
                    If Not Vazio(ws.Cells(r, 1).Value) Then
                        ext = ExtensaoNosTecnicos(wb, ws, r, CDbl(id))
                        nivel = Trim$(CStr(ws.Cells(r, 4).Value))
                        msg = AplicarConferencia(CLng(id), DataOuVazio(ws.Cells(r, 6).Value), DataOuVazio(ws.Cells(r, 7).Value), _
                                                 DataOuVazio(ws.Cells(r, 2).Value), DataOuVazio(ws.Cells(r, 3).Value), _
                                                 DataOuVazio(ws.Cells(r, 5).Value), NumeroOuVazio(ws.Cells(r, 8).Value), _
                                                 ext, "", nivel, "da planilha dos técnicos (" & wb.Name & ")")
                        TrazerConferencia = msg
                        If avisar Then Aviso msg, vbInformation, "Conferência"
                        Exit Function
                    End If
                End If
            End If
        End If
    Next wb
End Function

' Horas estendidas do batch na planilha dos tecnicos: coluna Estendido (h) do registro ou, nas versoes sem ela, o
' marcador do bloco do batch na aba de acompanhamento (bi-horarios estendidos x 2 h)
Private Function ExtensaoNosTecnicos(ByVal wb As Workbook, ByVal reg As Worksheet, ByVal r As Long, ByVal id As Double) As Double
    Dim ws As Worksheet, i As Long, v As Variant
    On Error Resume Next
    If CStr(reg.Cells(1, 10).Value) <> "" Then
        If ENumero(reg.Cells(r, 10).Value) Then ExtensaoNosTecnicos = CDbl(reg.Cells(r, 10).Value)
        Exit Function
    End If
    For Each ws In wb.Worksheets
        For i = 1 To 400
            v = ws.Cells(i, 40).Value
            If ENumero(v) Then
                If CDbl(v) = id And CStr(ws.Cells(i, 42).Value) Like "ACOMP*" Then
                    If ENumero(ws.Cells(i, 44).Value) Then ExtensaoNosTecnicos = CDbl(ws.Cells(i, 44).Value) * HORAS_SLOT
                    Exit Function
                End If
            End If
        Next i
    Next ws
End Function

Private Function DataOuVazio(ByVal v As Variant) As Variant
    If EData(v) Then DataOuVazio = CDate(v) Else DataOuVazio = Empty
End Function

Private Function NumeroOuVazio(ByVal v As Variant) As Variant
    If ENumero(v) Then NumeroOuVazio = CLng(v) Else NumeroOuVazio = Empty
End Function

' ---------------------------------------------------------------- pellet feed ultra fino (lancamento manual)
' Os 4 resultados do dia copiados do MES na ordem ORDEM_PF (Fe, S.E., SiO2, P) e colados em Fe. Pelo Ctrl+V da
' planilha (ColarInteligente) o texto copiado e repartido nas 4 linhas; colado pelo Excel, se tudo cair numa celula
' so (o MES manda o texto em linhas), ela e repartida aqui. Em cada linha vale o 1o numero; o resto e limpo.
Public Sub PFUFColado()
    Dim area As Range, k As Long, c As Long, t As String, v As String
    Set area = Nm("bPfCola")
    ' tudo numa celula so, com quebras de linha: reparte
    For k = 1 To NPF
        For c = 1 To area.Columns.Count
            v = CStr(area.Cells(k, c).Value)
            If InStr(v, vbLf) > 0 Or InStr(v, vbCr) > 0 Then
                t = v
                Exit For
            End If
        Next c
        If t <> "" Then Exit For
    Next k
    If t <> "" Then
        ColarPFUFTexto t
        Exit Sub
    End If
    ArrumarPFUF
End Sub

' Texto copiado do MES (uma linha por resultado, ou separado por tabulacao): os numeros, na ordem, vao para as linhas
Public Function ColarPFUFTexto(ByVal texto As String) As Boolean
    Dim partes As Variant, i As Long, k As Long, x As Double, area As Range
    texto = Replace(Replace(Replace(texto, vbCrLf, vbLf), vbCr, vbLf), vbTab, vbLf)
    partes = Split(texto, vbLf)
    Set area = Nm("bPfCola")
    ModoCarga True
    Desproteger shBatch
    area.ClearContents
    For i = LBound(partes) To UBound(partes)
        If k < NPF Then
            If NumeroColado(partes(i), x) Then
                k = k + 1
                area.Cells(k, 1).Value = x
            End If
        End If
    Next i
    Proteger shBatch
    ModoCarga False
    ArrumarPFUF
    ColarPFUFTexto = (k > 0)
End Function

' Em cada linha, o 1o numero vai para a coluna do resultado; formatos e mensagem
Private Sub ArrumarPFUF()
    Dim area As Range, k As Long, c As Long, v As Variant, x As Double, achou As Boolean, txt As String, chave As String
    Set area = Nm("bPfCola")
    ModoCarga True
    Desproteger shBatch
    For k = 1 To NPF
        chave = Split(ORDEM_PF, "|")(k - 1)
        achou = False
        For c = 1 To area.Columns.Count
            v = area.Cells(k, c).Value
            If Not achou And NumeroColado(v, x) Then
                area.Cells(k, 1).Value = x
                achou = True
            End If
            If c > 1 Then area.Cells(k, c).ClearContents
        Next c
        If Not achou Then area.Cells(k, 1).ClearContents
        area.Cells(k, 1).NumberFormat = IIf(DecimaisPF(chave) = 0, "0", "0." & String$(DecimaisPF(chave), "0"))
        area.Cells(k, 1).Interior.Color = RGB(255, 242, 204)
        area.Cells(k, 1).HorizontalAlignment = -4108
        If achou Then txt = txt & IIf(txt <> "", "  ·  ", "") & RotuloPF(chave) & " " & area.Cells(k, 1).Text
    Next k
    area.Locked = False
    Nm("bPfMsg").Value = "'" & IIf(txt = "", "", "Colado: " & txt)
    Proteger shBatch
    ModoCarga False
    gAlterado = True
End Sub

Private Function NumeroColado(ByVal v As Variant, ByRef x As Double) As Boolean
    Dim t As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        x = CDbl(v)
        NumeroColado = True
        Exit Function
    End If
    t = Trim$(Replace(CStr(v), Chr(160), ""))
    If t = "" Then Exit Function
    If InStr(t, ",") > 0 And InStr(t, ".") > 0 Then t = Replace(t, ".", "")      ' 2.236,5
    t = Replace(t, ",", ".")
    If Not IsNumeric(Replace(t, ".", "0")) Then Exit Function
    x = Val(t)
    NumeroColado = True
End Function

Private Function RotuloPF(ByVal chave As String) As String
    Select Case chave
        Case "SE": RotuloPF = "S.E."
        Case Else: RotuloPF = chave
    End Select
End Function

Private Function DecimaisPF(ByVal chave As String) As Long
    Dim ch As Variant, i As Long
    ch = Split(PF_CHAVES, "|")
    For i = LBound(ch) To UBound(ch)
        If ch(i) = chave Then DecimaisPF = CLng(Split(PF_DECIMAIS, "|")(i))
    Next i
End Function

' ---------------------------------------------------------------- bloco do batch colado (planilha dos tecnicos)
' Na planilha dos tecnicos, aba Batch, o bloco do batch (faixa BATCH n, resumo, observacoes e a grade com DATA,
' HORARIO e, por analise, Germano / Alimentacao 04 / Diferenca) e copiado inteiro (Ctrl+C) e colado aqui (Ctrl+V na
' linha COLE AQUI): o texto copiado e lido - batch, chegada em Ubu, nivel, inicio no processo, Germano (EB06),
' extensao, horarios e os resultados de Germano e de Ubu - e vira o batch desta planilha (sem consulta ao MES).
Public Function ColarBloco(ByVal texto As String) As String
    Dim lin As Collection, cel As Variant, r As Long, j As Long, c0 As Long, idB As Long, t As String, obs As String
    Dim datas As Variant, horas As Variant, n As Long, s As Long, p As Long, subRot As String, rotAtual As String
    Dim ger() As Variant, ubu() As Variant, tU() As Variant, x As Double, d As Variant, k As Long, msg As String
    Dim uIni As Variant, ip As Variant, gIni As Variant, gFim As Variant, nivel As String, ext As Double
    Dim rD As Long, rH As Long

    mMotivoBloco = ""
    Set lin = LinhasDoTexto(texto)
    ' numero do batch (faixa "BATCH 326") e observacoes (texto com a chegada em Ubu)
    For r = 1 To lin.Count
        cel = CelulasDaLinha(lin(r))
        For j = LBound(cel) To UBound(cel)
            t = Trim$(CStr(cel(j)))
            If idB = 0 And UCase$(Left$(t, 6)) = "BATCH " Then
                t = Trim$(Mid$(t, 7))
                k = 1
                Do While k <= Len(t)
                    If Mid$(t, k, 1) < "0" Or Mid$(t, k, 1) > "9" Then Exit Do
                    k = k + 1
                Loop
                If k > 1 Then idB = CLng(Left$(t, k - 1))
            ElseIf obs = "" And (InStr(1, t, "Chegada em Ubu", vbTextCompare) > 0 Or InStr(1, t, "Germano (EB06)", vbTextCompare) > 0) Then
                obs = t
            End If
        Next j
        If rD = 0 And UBound(cel) >= LBound(cel) Then
            For j = LBound(cel) To UBound(cel)
                If UCase$(Trim$(CStr(cel(j)))) = "DATA" Then
                    rD = r
                    c0 = j
                    Exit For
                End If
            Next j
        End If
    Next r
    If rD = 0 Or idB = 0 Then
        mMotivoBloco = "sem DATA (" & rD & ") ou sem BATCH (" & idB & ")"
        Exit Function
    End If
    ' horario: a linha seguinte a DATA
    rH = rD + 1
    If rH > lin.Count Then Exit Function
    datas = CelulasDaLinha(lin(rD))
    horas = CelulasDaLinha(lin(rH))
    If InStr(1, CStr(CelulaDe(horas, c0)), "HOR", vbTextCompare) = 0 Then
        mMotivoBloco = "linha HORARIO nao veio depois de DATA"
        Exit Function
    End If
    ReDim tU(1 To NMAX)
    For s = 1 To NMAX
        d = DataDoBloco(CStr(CelulaDe(datas, c0 + 1 + s)), CStr(CelulaDe(horas, c0 + 1 + s)))
        If IsEmpty(d) Then Exit For
        tU(s) = d
        n = s
    Next s
    If n = 0 Then
        mMotivoBloco = "nenhum horario lido (" & CelulaDe(datas, c0 + 2) & " / " & CelulaDe(horas, c0 + 2) & ")"
        Exit Function
    End If
    ' resultados: por analise, a linha Germano e a linha Alimentacao 04 (a de baixo)
    ReDim ger(1 To NP, 1 To NMAX)
    ReDim ubu(1 To NP, 1 To NMAX)
    For r = rH + 1 To lin.Count
        cel = CelulasDaLinha(lin(r))
        t = Trim$(CStr(CelulaDe(cel, c0)))
        If t <> "" Then rotAtual = t
        p = PIndice(Replace(rotAtual, "SiO2", "SiO" & ChrW(8322)))
        If p = 0 Then p = PIndice(rotAtual)
        subRot = Trim$(CStr(CelulaDe(cel, c0 + 1)))
        If p > 0 And subRot <> "" Then
            For s = 1 To n
                If NumeroColado(CelulaDe(cel, c0 + 1 + s), x) Then
                    If InStr(1, subRot, "Germano", vbTextCompare) > 0 Then
                        ger(p, s) = x
                    ElseIf InStr(1, subRot, "Diferen", vbTextCompare) = 0 Then
                        ubu(p, s) = x                                   ' Alimentacao 04 (ou o nome de Ubu)
                    End If
                End If
            Next s
        End If
    Next r
    ' observacoes: chegada em Ubu, nivel, inicio no processo, Germano (EB06), batch estendido
    uIni = DataDepois(obs, "Chegada em Ubu (UBU3):")
    ip = DataDepois(obs, "Início no processo:")
    gIni = DataDepois(obs, "Germano (EB06):")
    gFim = DataDepois(obs, " a ", InStr(1, obs, "Germano (EB06):", vbTextCompare))
    nivel = TextoDepois(obs, "Nível dos tanques:")
    k = InStr(1, obs, "estendido +", vbTextCompare)
    If k > 0 Then ext = Val(Mid$(obs, k + Len("estendido +")))

    ' batch novo: aba limpa (PFUF e comentarios em branco); o mesmo batch colado de novo: so os dados do bloco
    If Not ENumero(Nm("bId").Value) Then
        LimparBatch
    ElseIf CLng(Nm("bId").Value) <> idB Then
        LimparBatch
    End If
    ModoCarga True
    Desproteger shBatch
    Nm("bId").Value = idB
    Nm("bIniU").MergeArea.ClearContents
    Nm("bIniG").MergeArea.ClearContents
    Nm("bFimG").MergeArea.ClearContents
    Nm("bFimU").MergeArea.ClearContents
    Nm("bNivel").MergeArea.ClearContents
    Nm("bIniPMan").MergeArea.ClearContents
    Nm("bExtH").MergeArea.ClearContents
    If Not IsEmpty(uIni) Then Nm("bIniU").Value = uIni
    If Not IsEmpty(gIni) Then Nm("bIniG").Value = gIni
    If Not IsEmpty(gFim) Then Nm("bFimG").Value = gFim
    If nivel <> "" Then
        Nm("bNivel").Value = nivel
    ElseIf Not IsEmpty(ip) Then
        Nm("bIniPMan").Value = ip
    End If
    If ext > 0 Then Nm("bExtH").Value = ext
    Nm("bNHor").Value = n
    ' grade: horario de Ubu e resultados de cada analise (Germano e Ubu)
    shBatch.Range(shBatch.Cells(B_ROW_G0, B_COL_V1), shBatch.Cells(B_ROW_G0 + NMAX - 1, B_COL_V1 + 2 * NP - 1)).ClearContents
    For s = 1 To NMAX
        If s <= n Then
            shBatch.Cells(B_ROW_G0 + s - 1, B_COL_TU).Value = CDate(tU(s))
        Else
            shBatch.Cells(B_ROW_G0 + s - 1, B_COL_TU).ClearContents
        End If
        For p = 1 To NP
            If s <= n Then
                If Not IsEmpty(ger(p, s)) Then shBatch.Cells(B_ROW_G0 + s - 1, ColValor(p, 0)).Value = ger(p, s)
                If Not IsEmpty(ubu(p, s)) Then shBatch.Cells(B_ROW_G0 + s - 1, ColValor(p, 1)).Value = ubu(p, s)
            End If
        Next p
    Next s
    msg = "Batch " & idB & " colado da planilha dos técnicos: " & n & " bi-horários (" & _
          Format$(CDate(tU(1)), "dd\/mm hh:mm") & " a " & Format$(CDate(tU(n)), "dd\/mm hh:mm") & ")" & _
          IIf(IsEmpty(uIni), "", "   ·   chegada em Ubu " & Format$(uIni, "dd\/mm hh:mm")) & _
          IIf(IsEmpty(ip), "", "   ·   início " & Format$(ip, "dd\/mm hh:mm")) & IIf(ext > 0, "   ·   estendido +" & ext & " h", "")
    Nm("bFasesMsg").Value = "'" & msg
    Proteger shBatch
    ModoCarga False
    gAlterado = True
    Recalcular
    On Error Resume Next
    AjustarGraficos
    ColarBloco = msg & vbCrLf & vbCrLf & "Confira o Painel. Depois: pellet feed ultra fino (etapa 3), comentários e " & _
                 "Salvar no histórico."
End Function

Public Function MotivoBloco() As String
    MotivoBloco = mMotivoBloco
End Function

' Celula j da linha (vazio fora do limite)
Private Function CelulaDe(cel As Variant, ByVal j As Long) As Variant
    If j >= LBound(cel) And j <= UBound(cel) Then CelulaDe = cel(j) Else CelulaDe = ""
End Function

' Texto da area de transferencia em linhas (texto de cada linha; celulas separadas por tabulacao). Celula entre aspas
' pode ter quebra de linha: dentro das aspas a quebra vira Chr(1) e volta em CelulasDaLinha.
Private Function LinhasDoTexto(ByVal texto As String) As Collection
    Dim res As New Collection, i As Long, ch As String, aspas As Boolean, s2 As String, partes As Variant
    For i = 1 To Len(texto)
        ch = Mid$(texto, i, 1)
        If ch = """" Then aspas = Not aspas
        If aspas And (ch = vbLf Or ch = vbCr) Then
            If ch = vbLf Then s2 = s2 & Chr(1)
        ElseIf ch <> vbCr Then
            s2 = s2 & ch
        End If
    Next i
    partes = Split(s2, vbLf)
    For i = LBound(partes) To UBound(partes)
        res.Add CStr(partes(i))
    Next i
    Set LinhasDoTexto = res
End Function

' Celulas de uma linha (aspas de fora tiradas, "" vira ", quebras de linha de volta)
Private Function CelulasDaLinha(ByVal linha As String) As Variant
    Dim c As Variant, i As Long, t As String
    c = Split(linha, vbTab)
    For i = LBound(c) To UBound(c)
        t = CStr(c(i))
        If Len(t) >= 2 And Left$(t, 1) = """" And Right$(t, 1) = """" Then t = Replace(Mid$(t, 2, Len(t) - 2), """""", """")
        c(i) = Replace(t, Chr(1), vbLf)
    Next i
    CelulasDaLinha = c
End Function

' "23/11" + "9:30" -> data com o ano (o ano atual; se cair muito no futuro, o anterior)
Private Function DataDoBloco(ByVal dm As String, ByVal hm As String) As Variant
    Dim pd As Variant, ph As Variant, a As Long, d As Date
    DataDoBloco = Empty
    dm = Trim$(dm)
    hm = Trim$(hm)
    pd = Split(dm, "/")
    ph = Split(hm, ":")
    If UBound(pd) < 1 Or UBound(ph) < 1 Then Exit Function
    If Not (IsNumeric(pd(0)) And IsNumeric(pd(1)) And IsNumeric(ph(0)) And IsNumeric(ph(1))) Then Exit Function
    If CLng(pd(1)) < 1 Or CLng(pd(1)) > 12 Or CLng(pd(0)) < 1 Or CLng(pd(0)) > 31 Then Exit Function
    If UBound(pd) >= 2 Then
        a = CLng(Val(pd(2)))
        If a < 100 Then a = a + 2000
    Else
        a = Year(Date)
        If DateSerial(a, CLng(pd(1)), CLng(pd(0))) > Date + 60 Then a = a - 1
    End If
    d = DateSerial(a, CLng(pd(1)), CLng(pd(0))) + TimeSerial(CLng(ph(0)), CLng(Val(ph(1))), 0)
    DataDoBloco = d
End Function

' Primeira "dd/mm hh:mm" depois do rotulo no texto das observacoes
Private Function DataDepois(ByVal texto As String, ByVal rotulo As String, Optional ByVal desde As Long = 1) As Variant
    Dim k As Long, resto As Variant
    DataDepois = Empty
    If desde < 1 Then Exit Function
    k = InStr(desde, texto, rotulo, vbTextCompare)
    If k = 0 Then Exit Function
    resto = Split(Trim$(Replace(Replace(Mid$(texto, k + Len(rotulo)), vbLf, " "), vbCr, " ")), " ")
    If UBound(resto) < 1 Then Exit Function
    DataDepois = DataDoBloco(CStr(resto(0)), CStr(resto(1)))
End Function

' Texto depois do rotulo, ate o fim da linha (ou ate o proximo rotulo conhecido)
Private Function TextoDepois(ByVal texto As String, ByVal rotulo As String) As String
    Dim k As Long, t As String, fim As Variant, i As Long, q As Long
    k = InStr(1, texto, rotulo, vbTextCompare)
    If k = 0 Then Exit Function
    t = Mid$(texto, k + Len(rotulo))
    q = InStr(t, vbLf)
    If q > 0 Then t = Left$(t, q - 1)
    fim = Array("Início no processo", "Germano (EB06)", "Batch estendido", "Chegada em Ubu")
    For i = LBound(fim) To UBound(fim)
        q = InStr(1, t, CStr(fim(i)), vbTextCompare)
        If q > 0 Then t = Left$(t, q - 1)
    Next i
    TextoDepois = Trim$(Replace(t, vbCr, ""))
End Function

' ---------------------------------------------------------------- colar a CONFERENCIA da planilha dos tecnicos
' Na planilha de acompanhamento (aba Preenchimento) o tecnico ja tem a linha CONFERENCIA: Germano inicio/termino,
' Ubu chegada/termino, inicio no processo, bi-horarios, estendido, data e mineroduto. Copiada (Ctrl+C), o botao
' Colar conferencia le o texto da area de transferencia e preenche o batch daqui.
Public Sub ColarConferencia()
    Dim t As String
    t = TextoAreaTransferencia()
    If Trim$(t) = "" Then
        Aviso "Nada copiado. Na planilha dos técnicos, aba Batch, selecione o bloco do batch inteiro (da faixa BATCH " & _
              "até a última linha da grade), Ctrl+C, e clique de novo em Colar batch.", vbExclamation, "Colar batch"
        Exit Sub
    End If
    AvisoConferencia LerColado(t)
End Sub

' Texto da area de transferencia (Ctrl+C no Excel, no aplicativo ou no navegador)
' (so MSForms.DataObject: o outro caminho, pelo Internet Explorer, pode abrir um aviso escondido e travar o Excel)
Private Function TextoAreaTransferencia() As String
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")      ' MSForms.DataObject
    If Not d Is Nothing Then
        d.GetFromClipboard
        TextoAreaTransferencia = d.GetText(1)
    End If
    If InStr(TextoAreaTransferencia, "??") > 0 Then TextoAreaTransferencia = ""
End Function

' Le a linha copiada (colunas separadas por tabulacao) e preenche o batch. Devolve a mensagem; "" = nao reconheceu.
' Posicoes (como na aba Preenchimento dos tecnicos, celulas mescladas vem vazias): 0 Germano inicio, 2 Germano termino,
' 3 Ubu chegada, 5 Ubu termino, 7 inicio no processo, 9 bi-horarios, 10 estendido, 11 data, 12 mineroduto.
Public Function LerConferencia(ByVal texto As String) As String
    Dim linhas As Variant, tk As Variant, i As Long, j As Long, p0 As Long, d As Date, nDatas As Long
    Dim gIni As Variant, gFim As Variant, uIni As Variant, uFim As Variant, ip As Variant, nBi As Variant
    Dim ext As Double, mn As String, t As String, atr As Double, nivel As String, msg As String
    Dim k As Long, idB As Long

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
    ' numero do batch antes da 1a data (linha da tabela de conferencias dos tecnicos)
    For j = p0 - 1 To LBound(tk) Step -1
        t = Trim$(Replace(Tok(tk, j), "Batch", "", , , vbTextCompare))
        If t <> "" Then
            If IsNumeric(t) And InStr(t, "/") = 0 And InStr(t, ":") = 0 Then idB = CLng(t)
            Exit For
        End If
    Next j
    ' mineroduto: so um codigo de 1 ou 2 algarismos (03)
    If Not (Len(mn) >= 1 And Len(mn) <= 2 And IsNumeric(mn)) Then mn = ""

    ' sem o numero do batch (linha 3 dos tecnicos): pergunta, para nao gravar no batch anterior
    If idB = 0 Then
        idB = PerguntarBatch()
        If idB = 0 Then
            LerConferencia = "Conferência não colada: falta o número do batch."
            Exit Function
        End If
    End If
    LerConferencia = AplicarConferencia(idB, gIni, gFim, uIni, uFim, ip, nBi, ext, mn, "", "colada")
End Function

' Preenche o batch com a conferencia (colada ou lida da planilha dos tecnicos) e devolve a mensagem
Public Function AplicarConferencia(ByVal idB As Long, ByVal gIni As Variant, ByVal gFim As Variant, ByVal uIni As Variant, _
                                   ByVal uFim As Variant, ByVal ip As Variant, ByVal nBi As Variant, ByVal ext As Double, _
                                   ByVal mn As String, ByVal nivelDado As String, ByVal origem As String) As String
    Dim atr As Double, nivel As String, msg As String
    ModoCarga True
    Desproteger shBatch
    If idB > 0 Then Nm("bId").Value = idB
    Nm("bIniG").MergeArea.ClearContents
    Nm("bFimG").MergeArea.ClearContents
    Nm("bIniU").MergeArea.ClearContents
    Nm("bFimU").MergeArea.ClearContents
    If Not IsEmpty(gIni) Then Nm("bIniG").Value = gIni
    If Not IsEmpty(gFim) Then Nm("bFimG").Value = gFim
    If Not IsEmpty(uIni) Then Nm("bIniU").Value = uIni
    If Not IsEmpty(uFim) Then Nm("bFimU").Value = uFim
    ' inicio no processo: o nivel que veio, ou o que bate com chegada + atraso; senao, inicio manual
    Nm("bNivel").MergeArea.ClearContents
    Nm("bIniPMan").MergeArea.ClearContents
    nivel = nivelDado
    If Not IsEmpty(ip) And nivel = "" Then
        If Not IsEmpty(uIni) Then
            atr = Round((CDbl(ip) - CDbl(uIni)) * 24, 2)
            nivel = NivelDoAtraso(atr)
        End If
    End If
    If nivel <> "" Then
        Nm("bNivel").Value = nivel
    ElseIf Not IsEmpty(ip) Then
        Nm("bIniPMan").Value = ip
    End If
    If Not IsEmpty(nBi) Then
        If nBi >= 1 And nBi <= NMAX Then Nm("bNHor").Value = nBi
    End If
    Nm("bExtH").MergeArea.ClearContents
    If ext > 0 Then Nm("bExtH").Value = ext
    If mn <> "" Then Nm("bMin").Value = "'" & Format$(Val(mn), "00")
    msg = IIf(idB > 0, "Batch " & idB & " - c", "C") & "onferência " & origem & ":" & IIf(IsEmpty(gIni), "", "  Germano " & Format$(gIni, "dd\/mm hh:mm") & _
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
    AplicarConferencia = msg & vbCrLf & vbCrLf & "Agora é só Buscar no MES." & IIf(ext > 0, " Os " & Int(ext / HORAS_SLOT + 0.5) & _
        " bi-horários estendidos recebem a média de Germano, como na planilha dos técnicos.", "")
End Function

' Numero do batch da conferencia colada sem ele (InputBox; no teste automatico, o que ja esta na aba)
Private Function PerguntarBatch() As Long
    Dim r As Variant
    If gSilencioso Then
        If ENumero(Nm("bId").Value) Then PerguntarBatch = CLng(Nm("bId").Value)
        Exit Function
    End If
    r = InputBox("Número do batch desta conferência:", "Conferência", CStr(Nm("bId").Value))
    If IsNumeric(r) And Trim$(CStr(r)) <> "" Then PerguntarBatch = CLng(r)
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

Option Explicit

' Teste automatico (LibreOffice headless): grava o resultado em _Graf!BA1 (texto, linhas separadas por |)
Private mLog As String
Private mFalhas As Long

Private Sub Ok(ByVal cond As Boolean, ByVal desc As String)
    If cond Then
        mLog = mLog & "OK   " & desc & "|"
    Else
        mLog = mLog & "FALHA " & desc & "|"
        mFalhas = mFalhas + 1
    End If
End Sub

Private Function V(ByVal nome As String) As Variant
    V = Nm(nome).Value
End Function

Private Function Celula(ByVal i As Long, ByVal p As Long, ByVal lab As Long) As Variant
    Celula = shBatch.Cells(B_ROW_G0 + i - 1, ColValor(p, lab)).Value
End Function

Public Sub TesteLO()
    Dim n As Long, id As Long, r As Long, nL As Long, i As Long, a As Variant, b As Variant, s As String
    Dim pSE As Long, pSi As Long, salvo1 As Variant, salvo2 As Variant
    On Error GoTo Erro
    gSilencioso = True
    gTesteLO = True
    mLog = ""
    mFalhas = 0
    pSE = PIndice("SE")
    pSi = PIndice("SiO2")
    Recalcular

    ' --- abre em branco; o 284 vem do historico do laboratorio
    Ok Vazio(V("bId")) And Not ENumero(Celula(1, PIndice("SE"), 0)), "abre com batch em branco"
    AbrirBatch 284
    Recalcular
    Ok CStr(V("bId")) = "284", "batch 284 aberto do historico (" & V("bId") & ")"
    Ok Celula(2, pSE, 0) = 1933 And Celula(2, pSE, 1) = 1780, "2a leitura S.E. 1933/1780"
    Ok CStr(shGraf.Cells(G_ROW_ST + pSE - 1, G_COL_ST).Value) = "NOK", "S.E. NOK no farol"
    Ok LinhaHistorico(284) > 0, "historico do laboratorio com o 284"
    Ok CStr(V("bNivel")) = "Baixo", "284 com nivel baixo"
    Ok InStr(CStr(V("bFrase")), "Superfície Específica, Sílica e PPC") > 0, "frase NOK automatica: " & V("bFrase")
    Ok InStr(CStr(V("eResumo")), "3 de 5 fora da tolerância") > 0, "resumo enxuto: " & V("eResumo")
    Ok PRotulo(pSi) = "SiO" & ChrW(8322), "rotulo SiO2 com subscrito"

    ' --- simulacao: novo batch + buscar no MES
    Desproteger shConfig
    Nm("cfgFonte").Value = "SIMULAÇÃO"
    NovoBatch
    Ok Vazio(V("bId")), "novo batch com numero em branco"
    Nm("bId").Value = 291
    id = 291
    Ok Not ENumero(Celula(1, pSE, 0)), "novo batch sem leituras"
    Ok Nm("bAcomp").HasFormula And Nm("bFrase").HasFormula, "novo batch com comentarios automaticos"
    gUltimaMsg = ""
    BuscarMES
    Ok InStr(gUltimaMsg, "não buscadas") > 0 And InStr(gUltimaMsg, "4 de 4") > 0, _
       "Buscar no MES sem chegada em Ubu: busca so o PFUF (" & gUltimaMsg & ")"
    Nm("bIniU").Value = DateSerial(2026, 10, 10) + TimeSerial(7, 30, 0)
    Nm("bFimU").Value = DateSerial(2026, 10, 11) + TimeSerial(8, 10, 0)
    Nm("bIniG").Value = DateSerial(2026, 10, 8) + TimeSerial(5, 30, 0)
    Nm("bFimG").Value = DateSerial(2026, 10, 8) + TimeSerial(20, 0, 0)
    Recalcular
    Ok CLng(V("bNHor")) = 7, "Nº de horarios = bi-horarios de Germano = 7 (" & V("bNHor") & ")"
    Ok CLng(V("gNG")) = 7, "janelas de Germano pelo termino = 7 (" & V("gNG") & ")"
    Ok Abs(CDbl(V("bDef")) - 51) < 0.01, "defasagem 51 h (" & V("bDef") & ")"
    Ok Abs(CDbl(shBatch.Cells(B_ROW_G0, B_COL_TU).Value) - CDbl(DateSerial(2026, 10, 10) + TimeSerial(9, 30, 0))) < 0.0001, _
       "1o horario Ubu 10/10 09:30 (1a janela completa 09:00)"
    Ok Abs(CDbl(shBatch.Cells(B_ROW_G0 + 1, B_COL_TU).Value) - CDbl(DateSerial(2026, 10, 10) + TimeSerial(11, 30, 0))) < 0.0001, _
       "2o horario Ubu 11:30"
    Ok Abs(CDbl(shBatch.Cells(B_ROW_G0, B_COL_TG).Value) - CDbl(DateSerial(2026, 10, 8) + TimeSerial(7, 0, 0))) < 0.0001, _
       "1o horario Germano 08/10 07:00 (janela 06:00-08:00, grade par)"
    Ok shBatch.Cells(B_ROW_G0 + 7, B_COL_TG).Text = "", "8o horario Germano vazio (fim do bombeamento)"
    Ok shBatch.Cells(B_ROW_G0 + 7, B_COL_TU).Text = "", "8o horario Ubu vazio (mesmo Nº de Germano)"
    BuscarMES
    Recalcular
    n = 0
    For i = 1 To 12
        If ENumero(Celula(i, pSE, 0)) Then n = n + 1
    Next i
    Ok n = 7, "simulacao: 7 S.E. Germano (" & n & ")"
    Ok ENumero(Celula(7, pSE, 1)) And Not ENumero(Celula(8, pSE, 1)), "simulacao: S.E. Ubu ate o 7o horario"
    Ok ENumero(Celula(1, pSE, 1)) And Not ENumero(Celula(2, PIndice("325"), 1)), "Ubu: S.E. a cada 2 h, <325# a cada 6 h"
    Ok Not ENumero(Celula(1, pSi, 1)) And ENumero(Celula(2, pSi, 1)) And Not ENumero(Celula(3, pSi, 1)) And _
       ENumero(Celula(4, pSi, 1)), "Ubu SiO2 so na escala de 4 h (11:30, 15:30...)"
    Ok ENumero(Celula(1, pSi, 0)) And ENumero(Celula(2, pSi, 0)) And ENumero(Celula(3, pSi, 0)), "Germano SiO2 de 2 em 2 h"
    Ok NaEscala(pSi, 1, DateSerial(2026, 10, 10) + TimeSerial(23, 30, 0)) And _
       Not NaEscala(pSi, 1, DateSerial(2026, 10, 10) + TimeSerial(21, 30, 0)), "escala 03:30 + 4 h: 23:30 sim, 21:30 nao"
    Ok CLng(shGraf.Cells(G_ROW_ST + pSE - 1, G_COL_NP).Value) = 7, "7 pares S.E. no grafico"
    Ok CStr(shGraf.Cells(2 + 7, (pSE - 1) * G_BLOCO + 1).Value) = "Média", "rotulo Media apos os pares"
    ' pellet feed ultra fino: composto do dia anterior a data do report
    Ok Abs(CDbl(V("bPfData")) - (CDbl(V("bData")) - 1)) < 0.0001, "PFUF: data do composto = report - 1"
    Ok ENumero(V("bPfFe")) And ENumero(V("bPfSiO2")) And ENumero(V("bPfP")) And ENumero(V("bPfSE")), _
       "PFUF: Fe, SiO2, P e S.E. simulados"
    Ok InStr(gUltimaMsg, "PFUF (composto de") > 0, "PFUF na mensagem do Buscar no MES"
    Ok InStr(CStr(V("eResumo")), "PFUF (composto diário") > 0, "PFUF no resumo: " & V("eResumo")
    Ok InStr(CStr(shPainel.Cells(Nm("pSt_SE").Row - 1, 13).Value), "COMPOSTO DIÁRIO — PFUF · ") = 1, _
       "PFUF: titulo do quadro no painel (" & shPainel.Cells(Nm("pSt_SE").Row - 1, 13).Value & ")"
    gUltimaMsg = ""
    BuscarPFUF
    Ok InStr(gUltimaMsg, "4 de 4") > 0, "Buscar so o PFUF: " & gUltimaMsg
    Nm("bPfFe").Value = 66.12
    Nm("bAcomp").Value = "Acompanhamento editado"
    salvo1 = Celula(3, pSi, 0)
    salvo2 = Celula(5, pSE, 1)
    Nm("bTol_SE").Value = 999
    SalvarHistorico
    r = LinhaHistorico(id)
    Ok r > 0, "291 salvo no historico (linha " & r & ")"
    Ok Vazio(V("bId")) And Not ENumero(Celula(1, pSE, 0)), "salvar deixa a aba Batch em branco para o proximo"
    Ok Nm("bTol_SE").HasFormula And CDbl(V("bTol_SE")) = 150, "tolerancia volta ao padrao depois de salvar"
    Ok Not ENumero(V("bPfFe")) And Nm("bPfData").HasFormula, "PFUF em branco no proximo batch"
    Ok CDbl(shHist.Cells(r, H_COL1 + 15).Value) = 999, "tolerancia do batch gravada no historico"
    nL = 0
    For i = 2 To 2000
        If Vazio(shLeituras.Cells(i, 1).Value) Then Exit For
        If CLng(shLeituras.Cells(i, 1).Value) = id Then nL = nL + 1
    Next i
    Ok nL = 35, "leituras do 291 em BD_Leituras (" & nL & ")"
    Ok InStr(gUltimaMsg, "Erro") = 0, "nenhuma mensagem de erro: " & gUltimaMsg
    Ok CStr(shHist.Cells(r, H_COL1 + 2).Value) = "03", "mineroduto gravado como texto 03 (" & shHist.Cells(r, H_COL1 + 2).Value & ")"

    ' --- abrir outro batch e voltar
    AbrirBatch 284
    Ok CStr(V("bId")) = "284" And Celula(2, pSE, 0) = 1933, "abrir 284"
    Ok CStr(V("bMin")) = "03", "mineroduto aberto como 03 (" & V("bMin") & ")"
    Ok CStr(V("pSel")) = "284", "painel em exibicao = 284"
    AbrirBatch id
    Ok Celula(3, pSi, 0) = salvo1 And Celula(5, pSE, 1) = salvo2, "reabrir 291 com os mesmos valores"
    Ok Abs(CDbl(V("bDef")) - 51) < 0.01, "reabrir 291 com a defasagem"
    Ok CStr(V("bAcomp")) = "Acompanhamento editado", "comentario editado reaberto"
    Ok Abs(CDbl(V("bPfFe")) - 66.12) < 0.0001, "PFUF reaberto do historico (" & V("bPfFe") & ")"
    ' selecao pelo painel
    Nm("pSel").Value = 284
    SelecionarBatch
    Ok CStr(V("bId")) = "284", "selecionar 284 pelo painel"
    Ok ENumero(Celula(12, pSE, 1)), "284 com 12 horarios"

    ' --- excluir
    ExcluirDoHistorico id
    Ok LinhaHistorico(id) = 0, "291 excluido"
    nL = 0
    For i = 2 To 2000
        If Vazio(shLeituras.Cells(i, 1).Value) Then Exit For
        If CLng(shLeituras.Cells(i, 1).Value) = id Then nL = nL + 1
    Next i
    Ok nL = 0, "leituras do 291 excluidas"

    ' --- MES: formula e tags recusadas
    shConfig.Cells(C_ROW_TAG0 + 1, C_COL_TAG).Value = "TAG-UBU-SE"
    shConfig.Cells(C_ROW_TAG0 + 2 * (pSi - 1) + 1, C_COL_TAG).Value = "TAG-UBU-SIO2"
    s = ArgumentosLab(1, "1," & pSi, shDadosMES.Range("A70"), shDadosMES.Range("A72"))
    Ok InStr(s, """TAG-UBU-SE,TAG-UBU-SIO2"",""UBU,UBU"",""IP_MESVALOR,IP_MESVALOR"",Dados_MES!$B$6,Dados_MES!$B$7,""2h""") > 0, _
       "argumentos da consulta Ubu: " & Left$(s, 140)
    s = TesteTagsInvalidas("Error: Tag Name TAG-UBU-SIO2 is invalid")
    Ok InStr(s, "TAG-UBU-SIO2") > 0, "tag recusada identificada"

    ' --- batch 285 colado da tela de eventos do MES (Fases do Evento)
    NovoBatch
    Nm("bId").Value = 285
    With Nm("bFases")
        .Cells(1, 1).Value = "Fase"
        .Cells(1, 2).Value = "Descrição"
        .Cells(1, 3).Value = "Início"
        .Cells(1, 4).Value = "Término"
        .Cells(2, 1).Value = "EB06"
        .Cells(2, 2).Value = "Estação de Bombas 06"
        .Cells(2, 3).Value = DateSerial(2026, 10, 11)
        .Cells(2, 4).Value = DateSerial(2026, 10, 11) + TimeSerial(23, 59, 0)
        .Cells(3, 1).Value = "EB07"
        .Cells(3, 2).Value = "Estação de Bombas 07"
        .Cells(3, 3).Value = DateSerial(2026, 10, 12) + TimeSerial(0, 5, 0)
        .Cells(3, 4).Value = DateSerial(2026, 10, 13) + TimeSerial(0, 12, 0)
        .Cells(4, 1).Value = "UBU3"
        .Cells(4, 2).Value = "Terminal Ponta Ubu 03"
        .Cells(4, 3).Value = DateSerial(2026, 10, 13) + TimeSerial(9, 18, 0)
        .Cells(4, 4).Value = DateSerial(2026, 10, 14) + TimeSerial(10, 37, 0)
    End With
    LerFases
    Recalcular
    Ok Abs(CDbl(V("bIniG")) - CDbl(DateSerial(2026, 10, 11))) < 0.0001 And _
       Abs(CDbl(V("bFimU")) - CDbl(DateSerial(2026, 10, 14) + TimeSerial(10, 37, 0))) < 0.0001, _
       "285 colado: Germano EB06 e Ubu UBU3 (" & V("bFasesMsg") & ")"
    Ok Vazio(Nm("bFases").Cells(4, 1).Value), "area de colar limpa depois de ler"
    Ok Abs(CDbl(V("bData")) - CDbl(Date)) < 0.0001, "data do report = hoje"
    Ok CStr(V("bMin")) = "03", "mineroduto pela fase UBU3 = 03"
    Ok Abs(CDbl(V("gJanU")) - CDbl(DateSerial(2026, 10, 13) + TimeSerial(11, 0, 0))) < 0.0001, _
       "285: 1a janela Ubu 13/10 11:00 (sem nivel = sem atraso)"
    Ok Abs(CDbl(V("gJanG")) - CDbl(DateSerial(2026, 10, 11))) < 0.0001, "285: 1a janela Germano 11/10 00:00"
    Ok CLng(V("bNHor")) = 12 And CLng(V("gNG")) = 12, "285: 12 horarios Ubu e 12 Germano (" & V("bNHor") & "/" & V("gNG") & ")"
    Ok Abs(CDbl(V("bDef")) - 59) < 0.01, "285: defasagem 59 h (" & V("bDef") & ")"
    ' nivel dos tanques: atraso da tabela de Configuracoes
    Nm("cfgAtrasoH").Cells(2, 1).Value = 4
    Nm("bNivel").Value = "Médio"
    Recalcular
    Ok Abs(CDbl(V("bIniP")) - CDbl(DateSerial(2026, 10, 13) + TimeSerial(13, 18, 0))) < 0.0001, "nivel medio: processo = chegada + 4 h"
    Ok Abs(CDbl(V("gJanU")) - CDbl(DateSerial(2026, 10, 13) + TimeSerial(15, 0, 0))) < 0.0001, "nivel medio: 1a janela 15:00"
    Ok CLng(V("bNHor")) = 12, "nivel medio: termino tambem anda 4 h (12 horarios)"
    Ok Abs(CDbl(V("bAtraso")) - 4) < 0.01, "atraso exibido 4 h"
    Ok InStr(CStr(V("gAcomp")), "nível dos tanques médio · comparação a partir de 13/10 15:30") > 0, _
       "acompanhamento: " & V("gAcomp")
    Ok InStr(CStr(V("gComentarios")), "Chegada em Ubu 13/10 09:18") > 0, "acompanhamento no comentario do painel"
    Nm("bIniPMan").Value = DateSerial(2026, 10, 13) + TimeSerial(12, 0, 0)
    Recalcular
    Ok Abs(CDbl(V("gJanU")) - CDbl(DateSerial(2026, 10, 13) + TimeSerial(13, 0, 0))) < 0.0001, "ajuste manual: 1a janela 13:00"
    Ok Abs(CDbl(V("bAtraso")) - 2.7) < 0.01, "ajuste manual: atraso 2,7 h (" & V("bAtraso") & ")"
    id = CLng(V("bId"))
    a = V("bIniP")
    SalvarHistorico
    r = LinhaHistorico(id)
    Ok CStr(shHist.Cells(r, H_COL1 + 5).Value) = "Médio", "nivel gravado no historico"
    Ok Vazio(V("bNivel")), "nivel em branco no proximo batch"
    Ok Abs(CDbl(shHist.Cells(r, H_COL1 + H_COL_STAT1 - 1 + 4 * NP).Value) - CDbl(a)) < 0.0001, _
       "inicio no processo gravado no historico"
    AbrirBatch 284
    AbrirBatch id
    Ok CStr(V("bNivel")) = "Médio", "nivel reaberto"

    ' --- acompanhamento (batch em processamento): so Germano; modelo fixo com batch anterior e em processamento
    Dim rb As Long, rb2 As Long, ubu As Double, kSE As Long
    Nm("bModo").Value = "Acompanhamento (em processamento)"
    LimparBatch
    GarantirModeloAcomp
    Ok CStr(shAcomp.Cells(LinhaBlocoN(1), 2).Value) = "Batch" And CStr(shAcomp.Cells(LinhaBlocoN(2), 2).Value) = "Batch", _
       "modelo com os dois blocos (anterior e em processamento)"
    Ok IsEmpty(IdBloco(1)) And IsEmpty(IdBloco(2)), "modelo em branco"
    Nm("bId").Value = 295
    Nm("bIniG").Value = DateSerial(2026, 10, 21) + TimeSerial(18, 10, 0)
    Nm("bFimG").Value = DateSerial(2026, 10, 22) + TimeSerial(20, 0, 0)
    Nm("bIniU").Value = DateSerial(2026, 10, 23) + TimeSerial(9, 18, 0)
    Nm("bNivel").Value = "Médio"
    Recalcular
    ' Ubu de outro batch esquecido na aba Batch nao pode ir para o bloco
    Desproteger shBatch
    shBatch.Cells(B_ROW_G0, ColValor(pSi, 1)).Value = 9.99
    gUltimaMsg = ""
    BuscarMES
    Ok InStr(gUltimaMsg, "Acompanhamento do batch 295") > 0, "acompanhamento: busca so Germano (" & gUltimaMsg & ")"
    Ok ENumero(Celula(1, pSi, 0)) And Not ENumero(Celula(1, pSi, 1)), "acompanhamento: Germano sim, Ubu nao"
    rb = LinhaBloco(shAcomp, 295)
    Ok rb = LinhaBlocoN(2) And IsEmpty(IdBloco(1)), "295 no bloco em processamento; anterior vazio (" & rb & ")"
    Ok CStr(shAcomp.Cells(rb + 3, 2).Value) = PRotulo(pSi) And CStr(shAcomp.Cells(rb + 3, 3).Value) = "Germano", _
       "rotulo SiO2 / Germano"
    Ok CStr(shAcomp.Cells(rb + 4, 3).Value) = ChrW(9998) & " Alimentação 04", "rotulo Alimentacao 04 (para lancar)"
    Ok ENumero(shAcomp.Cells(rb + 3, 4).Value), "SiO2 Germano no bloco"
    Ok Not ENumero(shAcomp.Cells(rb + 4, 4).Value), "Ubu de outro batch (aba Batch) nao entra no bloco"
    kSE = rb + 3 + 3 * 3
    Ok CStr(shAcomp.Cells(kSE, 2).Value) = PRotulo(PIndice("SE")) And ENumero(shAcomp.Cells(kSE, 4).Value), _
       "S.E. no mesmo bloco (fisico)"
    Ok Abs(CDbl(shAcomp.Cells(rb + 2, 4).Value) - CDbl(shBatch.Cells(B_ROW_G0, B_COL_TU).Value)) < 0.0001, _
       "inicio aproximado no processo = horario de Ubu"
    n = CLng(V("bNHor"))
    Ok CStr(shAcomp.Cells(rb, 4 + n).Value) = "Média" And CStr(shAcomp.Cells(rb, 6 + n).Value) = "Status", "coluna Media depois dos " & n & " horarios"
    ' Ubu lancado a mao: diferenca e media; refazer a busca nao apaga
    shAcomp.Unprotect
    shAcomp.Cells(rb + 4, 4).Value = CDbl(shAcomp.Cells(rb + 3, 4).Value) + 0.1
    shAcomp.Calculate
    Ok Abs(CDbl(shAcomp.Cells(rb + 5, 4).Value) - 0.1) < 0.0001, "diferenca Ubu - Germano"
    ' Status compara as medias: Ubu em todos os horarios = Germano + 0,1 -> OK; + 0,5 -> NOK
    For i = 1 To n
        If ENumero(shAcomp.Cells(rb + 3, 3 + i).Value) Then shAcomp.Cells(rb + 4, 3 + i).Value = CDbl(shAcomp.Cells(rb + 3, 3 + i).Value) + 0.1
    Next i
    shAcomp.Calculate
    Ok CStr(shAcomp.Cells(rb + 3, 6 + n).Value) = "OK", "status da SiO2: OK (" & shAcomp.Cells(rb + 3, 6 + n).Value & ")"
    For i = 1 To n
        If ENumero(shAcomp.Cells(rb + 3, 3 + i).Value) Then shAcomp.Cells(rb + 4, 3 + i).Value = CDbl(shAcomp.Cells(rb + 3, 3 + i).Value) + 0.5
    Next i
    shAcomp.Calculate
    Ok CStr(shAcomp.Cells(rb + 3, 6 + n).Value) = "NOK", "status da SiO2: NOK com diferenca 0,5"
    shAcomp.Range(shAcomp.Cells(rb + 4, 5), shAcomp.Cells(rb + 4, 3 + n)).ClearContents
    shAcomp.Cells(rb + 4, 4).Value = CDbl(shAcomp.Cells(rb + 3, 4).Value) + 0.1
    shAcomp.Calculate
    Ok ENumero(shAcomp.Cells(rb + 4, 4 + n).Value), "media de Ubu"
    ubu = CDbl(shAcomp.Cells(rb + 4, 4).Value)
    BuscarMES
    rb = LinhaBloco(shAcomp, 295)
    Ok rb = LinhaBlocoN(2) And Abs(CDbl(shAcomp.Cells(rb + 4, 4).Value) - ubu) < 0.0001, "Ubu lancado fica depois de buscar de novo"
    gUltimaMsg = ""
    SalvarHistorico
    Ok LinhaHistorico(295) = 0 And InStr(gUltimaMsg, "Acompanhamento") > 0, "acompanhamento nao vai para o historico"
    ' proximo batch: 295 vira anterior (com o Ubu lancado), 296 entra em processamento sem Ubu nenhum
    Nm("bId").Value = 296
    AtualizarUltimos
    Ok CStr(shBatch.Cells(B_ROW_ULT, U_COL_BATCH).Value) = "Batch 295", "ultimo batch = 295 (" & _
       shBatch.Cells(B_ROW_ULT, U_COL_BATCH).Value & ")"
    Ok CStr(shBatch.Cells(B_ROW_ULT, U_COL_NIVEL).Value) = "Médio", "ultimo batch: nivel medio"
    Ok EData(shBatch.Cells(B_ROW_ULT, U_COL_CHEG).Value), "ultimo batch: chegada em Ubu"
    Ok CStr(shBatch.Cells(B_ROW_ULT, U_COL_ORIGEM).Value) = "Acompanhamento", "ultimo batch: origem acompanhamento"
    Ok InStr(CStr(shBatch.Cells(B_ROW_ULT + 1, U_COL_BATCH).Value), "Batch 2") = 1, "2o ultimo batch do historico"
    BuscarMES
    rb = LinhaBloco(shAcomp, 296)
    rb2 = LinhaBloco(shAcomp, 295)
    Ok rb = LinhaBlocoN(2) And rb2 = LinhaBlocoN(1), "295 anterior em cima, 296 em processamento embaixo (" & rb2 & ", " & rb & ")"
    Ok Abs(CDbl(shAcomp.Cells(rb2 + 4, 4).Value) - ubu) < 0.0001, "Ubu do 295 vai junto para o bloco anterior"
    Ok Not ENumero(shAcomp.Cells(rb + 4, 4).Value), "296 sem Ubu (nao repete o do 295)"
    Ok ENumero(shAcomp.Cells(rb2 + 3, 4).Value) And ENumero(shAcomp.Cells(rb + 3, 4).Value), "Germano nos dois blocos"
    ' 297: 296 vira anterior e o 295 sai da aba
    Nm("bId").Value = 297
    BuscarMES
    Ok IdBloco(1) = 296 And IdBloco(2) = 297 And LinhaBloco(shAcomp, 295) = 0, "297: anterior 296, 295 sai da aba"
    ' buscar de novo o anterior so atualiza o bloco dele
    Nm("bId").Value = 296
    BuscarMES
    Ok IdBloco(1) = 296 And IdBloco(2) = 297, "buscar o 296 de novo atualiza so o bloco anterior"
    ' depois, comparativo do 295: chegada e nivel vem do acompanhamento
    Nm("bModo").Value = "Comparativo (batch processado)"
    LimparBatch
    Nm("bId").Value = 295
    PreencherDoAcomp
    Ok EData(V("bIniU")) And CStr(V("bNivel")) = "Médio" And EData(V("bIniG")), "comparativo do 295 com os dados do acompanhamento"
    LimparBatch
    AbrirBatch id

    ' --- MES sem o suplemento Aspen (LibreOffice): #NOME? avisado uma vez, sem esperar
    Nm("cfgFonte").Value = "MES"
    gUltimaMsg = ""
    BuscarMES
    Ok InStr(gUltimaMsg, "não reconheceu a função do MES") > 0 And _
       (Len(gUltimaMsg) - Len(Replace(gUltimaMsg, "não reconheceu a função", ""))) = Len("não reconheceu a função"), _
       "sem Aspen: aviso #NOME? uma vez"
    Ok Not shDadosMES.Range("A70").HasFormula And Not shDadosMES.Range("A10").HasFormula And _
       Not shDadosMES.Range("A130").HasFormula, "consulta desligada depois da falha"
    Nm("cfgFonte").Value = "SIMULAÇÃO"

    Ok Limpo("M4250010000-0004-ABLQ" & vbTab) = "M4250010000-0004-ABLQ", "tag com tabulacao no fim e limpa"

    ' --- Germano em sequencia (como na tela do MES): horario sem resultado nao deixa coluna vazia;
    '     23:00 e 23:59 ficam separados (antes a janela de 2 h tirava a media dos dois)
    Ok TesteSequenciaG() = "11|178|207|202|15:00", "Germano em sequencia (" & TesteSequenciaG() & ")"
    Ok SequenciaG(), "Germano em sequencia ligado por padrao"

    ' --- variantes do codigo da tag
    Ok VarianteTag("M4650060010-0008-HHLQU") = "M4650060010-8-HHLQU", "variante -0008- -> -8-"
    Ok VarianteTag("M250010000-17-ABLF") = "M250010000-0017-ABLF", "variante -17- -> -0017-"
    Ok VarianteTag("306TP001-FX001-R") = "", "tag sem codigo numerico: sem variante"

    ' --- mapa recusado pelo MES: tenta o proximo da lista e grava o que funcionou
    Ok TesteMapas("Erro:(M0650024000-0003-DDLQU) Map Name IP_MESVALOR is invalid(M0650024000-0004-DDLQU) Map Name " & _
       "IP_MESVALOR is invalid") = "2|IP_DiscreteMap|IP_AnalogMap", "mapa recusado: 2 tags, troca para o proximo mapa"

    ' --- leitura do PFUF: janela das 23:30; S.E. gravado em outro horario do mesmo dia; Fe fora da faixa
    Ok TesteLeituraPF() = "-|1250|31|2350000|nota SE|nota Fe", "PFUF: leitura da saida do MES (" & TesteLeituraPF() & ")"

    ' --- respostas do Aspen em texto
    Dim tj As Date, xx As Double
    Ok ComoData("12/10/2026 15:00:00", tj) And Abs(CDbl(tj) - CDbl(DateSerial(2026, 10, 12) + TimeSerial(15, 0, 0))) < 0.0001, _
       "data do Aspen em texto"
    Ok ComoData(DateSerial(2026, 10, 12), tj), "data do Aspen como data"
    Ok ComoData("10/12/2026 3:00:00 PM", tj, DateSerial(2026, 10, 12)) And _
       Abs(CDbl(tj) - CDbl(DateSerial(2026, 10, 12) + TimeSerial(15, 0, 0))) < 0.0001, "data do Aspen em texto americano"
    Ok Not ComoData("Success", tj), "texto Success nao e data"
    Ok ComoNumero("2052.0", xx) And Abs(xx - 2052) < 0.0001, "numero em texto 2052.0 (" & xx & ")"
    Ok ComoNumero("0.058", xx) And Abs(xx - 0.058) < 0.000001, "numero em texto 0.058"
    Ok ComoNumero(1.95, xx) And Abs(xx - 1.95) < 0.000001, "numero 1,95"
    Ok Not ComoNumero("Tag Name X is invalid", xx), "texto de erro nao e numero"

    GoTo Fim
Erro:
    mLog = mLog & "ERRO " & Err.Number & " " & Err.Description & "|"
    mFalhas = mFalhas + 1
Fim:
    shGraf.Range("BA1").Value = "'" & IIf(mFalhas = 0, "TUDO OK", mFalhas & " FALHA(S)") & "|" & mLog
End Sub



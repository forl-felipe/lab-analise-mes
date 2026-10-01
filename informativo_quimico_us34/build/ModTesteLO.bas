Option Explicit

' Modulo usado somente na validacao automatica (LibreOffice). Nao vai para a versao final.

Private Function LinhaChave(ByVal ws As Worksheet, ByVal chave As String) As Long
    Dim r As Long
    For r = 1 To ws.UsedRange.Row + ws.UsedRange.Rows.Count
        If CStr(ws.Cells(r, 1).Value) = chave Then
            LinhaChave = r
            Exit Function
        End If
    Next r
End Function

' Linha do Resumo cujo texto (coluna C) e igual a t
Private Function LinhaTexto(ByVal ws As Worksheet, ByVal t As String) As Long
    Dim r As Long
    For r = 1 To 120
        If ws.Cells(r, 3).Text = t Then
            LinhaTexto = r
            Exit Function
        End If
    Next r
End Function

Public Sub TesteLO()
    Dim r As String, passo As Long, lin As Long, ws As Worksheet, nv As Long, jj As Long
    On Error GoTo Erro
    gSilencioso = True
    passo = 1
    ' estado apos a 1a consulta real: formula do Aspen desligada (texto), prefixo guardado
    Nm("mesPrefixo").Value = "'" & Left$(shDadosMES.Range("A7").Formula, InStr(shDadosMES.Range("A7").Formula, "("))
    shDadosMES.Range("A7").Value = "'Consulta concluída"
    shDadosMES.Range("A22").Value = "'Consulta concluída"
    Nm("cfgFonte").Value = "SIMULAÇÃO"
    Nm("pData").Value = DateSerial(2026, 9, 28)
    Nm("dLetra").Value = "A"
    Nm("dTecnico").Value = "Fulano"
    Nm("dReal").Cells(1, 1).Value = "Controle da Produção: PF 03 04X04h"
    Nm("dReal").Cells(2, 1).Value = " - Descarte das amostras bi-horárias;"
    Nm("dSol").Cells(1, 1).Value = "sol. 1432026 PF"
    Nm("dComUS3").Value = "Usina 03 estável." & vbLf & "Sílica da mistura acima do alvo às 11:30."
    Nm("dRecebe").Value = "B"
    Nm("dH2Coque").Value = "0,45"
    Nm("dProdUS3").Value = "PDR/STD"
    Nm("dProdUS4").Value = "PBF/MB45"
    Nm("nLetra").Value = "C"
    Nm("nTecnico").Value = "Beltrano"
    Nm("nReal").Cells(1, 1).Value = "Tarefa da noite"
    passo = 2
    Set ws = shResumoDia
    AtualizarMESDia
    passo = 21
    lin = LinhaChave(ws, "P01U3")
    r = "dSiO2PF_slot1=" & ws.Cells(lin, 5).Value & "|dSiO2PF_media=" & ws.Cells(lin, 11).Value & _
        "|PF_umaLinha=" & (LinhaChave(ws, "P01U4") = 0) & "|PFrotulo=" & ws.Cells(lin, 4).Value & _
        "|dHora1=" & Nm("dHoras").Cells(1, 1).Value & "|dAtual=" & Left$(Nm("dAtualizado").Value, 40) & _
        "|FePelotaUS3=" & ws.Cells(LinhaChave(ws, "P06U3"), 11).Value & _
        "|SiO2PFvisivel=" & (Not ws.Rows(lin).Hidden) & _
        "|embarqueNoResumo=" & (LinhaChave(ws, "P13U3") > 0) & "|semPPCpelota=" & (LinhaChave(ws, "P12U3") > 0 And CStr(ws.Cells(LinhaChave(ws, "P12U3"), 2).Value) = "Mn") & _
        "|B2US4=" & ws.Cells(LinhaChave(ws, "P11U4"), 11).Value
    ' farol: limites do produto (colunas auxiliares Q=min, R=max)
    passo = 22
    ws.Calculate
    r = r & "|limSiO2US3=" & ws.Cells(LinhaChave(ws, "P07U3"), 17).Text & "/" & ws.Cells(LinhaChave(ws, "P07U3"), 18).Text & _
        "|limSiO2US4=" & ws.Cells(LinhaChave(ws, "P07U4"), 18).Text & _
        "|limFeUS3min=" & ws.Cells(LinhaChave(ws, "P06U3"), 17).Text & _
        "|limCaOUS3min=" & ws.Cells(LinhaChave(ws, "P09U3"), 17).Text & _
        "|limCaOUS4=" & ws.Cells(LinhaChave(ws, "P09U4"), 17).Text & "/" & ws.Cells(LinhaChave(ws, "P09U4"), 18).Text & _
        "|limB2US4max=" & ws.Cells(LinhaChave(ws, "P11U4"), 18).Text & _
        "|limSiO2PF_US3=" & ws.Cells(LinhaChave(ws, "P01U3"), 18).Text & "|limSiO2PF_US4=n/a" & _
        "|limP_PF_US3=" & ws.Cells(LinhaChave(ws, "P02U3"), 18).Text & "|limPPCPF=" & ws.Cells(LinhaChave(ws, "P05U3"), 18).Text & _
        "|produtos=" & Application.WorksheetFunction.CountA(shLimites.Range("B9:B33")) & "|gProdUS3=" & Nm("gProdUS3").Text
    passo = 3
    ' migracao por formula
    lin = LinhaTexto(ws, "Controle da Produção: PF 03 04X04h")
    r = r & "|realLinha1=" & (lin > 0) & "|realVisivel=" & (Not ws.Rows(lin).Hidden) & _
        "|real2=" & ws.Cells(lin + 1, 3).Value & "|real3oculta=" & ws.Rows(lin + 2).Hidden & _
        "|equipVazio=" & (LinhaTexto(ws, "Sem registro") > 0) & _
        "|data=" & ws.Cells(6, 3).Text & "|letra=" & ws.Cells(6, 10).Text & "|recebe=" & ws.Cells(6, 13).Text & _
        "|tecnico=" & ws.Cells(7, 3).Text & "|prodUS3=" & ws.Cells(7, 10).Text
    For jj = 1 To 120
        If CStr(ws.Cells(jj, 2).Value) = "Comentário US3" Then r = r & "|comUS3=" & Replace(ws.Cells(jj, 3).Text, vbLf, "/") & _
            "|altComUS3=" & ws.Rows(jj).RowHeight
        If CStr(ws.Cells(jj, 2).Value) = "Comentário US4" Then r = r & "|comUS4=" & ws.Cells(jj, 3).Text
        If Left$(CStr(ws.Cells(jj, 2).Value), 11) = "Observações" Then r = r & "|obs=" & ws.Cells(jj, 3).Text
        If CStr(ws.Cells(jj, 2).Value) = "Programa em uso" Then r = r & "|prog=" & Left$(ws.Cells(jj, 3).Text, 12)
        If CStr(ws.Cells(jj, 2).Value) = "Sistema de ar utilizado" Then r = r & "|ar=" & ws.Cells(jj, 3).Text
    Next jj
    passo = 4
    AtualizarMESNoite
    Set ws = shResumoNoite
    lin = LinhaChave(ws, "P01U3")
    r = r & "|nHora1=" & Nm("nHoras").Cells(1, 1).Value & "|nSiO2PF_media=" & ws.Cells(lin, 11).Value & _
        "|nTecnico=" & ws.Cells(7, 3).Text & "|nReal=" & (LinhaTexto(ws, "Tarefa da noite") > 0) & _
        "|dMediaIntacta=" & shResumoDia.Cells(LinhaChave(shResumoDia, "P01U3"), 11).Value
    passo = 5
    Nm("gIni").Value = DateSerial(2026, 9, 28) + TimeSerial(7, 0, 0)
    Nm("gFim").Value = DateSerial(2026, 9, 29) + TimeSerial(7, 0, 0)
    AtualizarResultados
    lin = LinhaChave(shResultados, "P07U3")
    For jj = 0 To NSLOT_MAX - 1
        If ENumero(shResultados.Cells(lin, COL_SLOT1 + jj).Value) Then nv = nv + 1
    Next jj
    r = r & "|g24h_slots=" & nv & "|gH12=" & Nm("gHoras").Cells(1, 12).Value & "|gCol12oculta=" & _
        shResultados.Columns(COL_SLOT1 + 11).Hidden
    Nm("gFim").Value = DateSerial(2026, 9, 28) + TimeSerial(17, 0, 0)
    AtualizarResultados
    r = r & "|g10h_col6visivel=" & (Not shResultados.Columns(COL_SLOT1 + 4).Hidden) & _
        "|g10h_col6oculta=" & shResultados.Columns(COL_SLOT1 + 5).Hidden
    passo = 55
    ' periodo em andamento (comecou ha 3 h, termina daqui a 3 h): janelas futuras em branco
    Dim ag As Date, cheias As Long, vazia3 As Boolean
    ag = CDate(Round(CDbl(Now) * 1440, 0) / 1440)
    Nm("gIni").Value = CDate(CDbl(ag) - 3 / 24)
    Nm("gFim").Value = CDate(CDbl(ag) + 3 / 24)
    AtualizarResultados
    AtualizarResultados
    lin = LinhaChave(shResultados, "P07U3")
    For jj = 0 To 2
        If ENumero(shResultados.Cells(lin, COL_SLOT1 + jj).Value) Then cheias = cheias + 1
    Next jj
    vazia3 = Vazio(shResultados.Cells(lin, COL_SLOT1 + 2).Value)
    r = r & "|andamento_cheias=" & cheias & "|andamento_janela3vazia=" & vazia3 & _
        "|andamento_atual=" & Left$(Nm("gAtualizado").Value, 60)
    passo = 58
    ' Embarque (testes): consulta propria, uma linha por analise
    Nm("eIni").Value = DateSerial(2026, 9, 28) + TimeSerial(7, 0, 0)
    Nm("eFim").Value = DateSerial(2026, 9, 28) + TimeSerial(19, 0, 0)
    AtualizarEmbarque
    nv = 0
    lin = LinhaChave(shEmbarque, "P13U3")
    For jj = 0 To 5
        If ENumero(shEmbarque.Cells(lin, COL_SLOT1 + jj).Value) Then nv = nv + 1
    Next jj
    r = r & "|emb_FeSlots=" & nv & "|emb_FeMedia=" & shEmbarque.Cells(lin, 17).Text & _
        "|emb_semUS4=" & (LinhaChave(shEmbarque, "P13U4") = 0) & "|emb_tag=" & shConfig.Cells(CFG_ROW1 + 12, CFG_COL_TAG3).Value & _
        "|emb_semFiltragem=" & (LinhaChave(shEmbarque, "P01U3") = 0) & _
        "|emb_ultima=" & (LinhaChave(shEmbarque, "P" & Format$(NPARAM, "00") & "U3") > 0) & _
        "|emb_col7oculta=" & shEmbarque.Columns(COL_SLOT1 + 6).Hidden & _
        "|emb_atual=" & Left$(Nm("eAtualizado").Value, 45)
    passo = 59
    ' tags recusadas pelo MES: o texto de erro do Aspen e lido e as tags saem da consulta
    Nm("eFase").Value = "TN"
    Application.Calculate
    r = r & "|tagTN=" & shConfig.Cells(CFG_ROW1 + 12, CFG_COL_TAG3).Value
    shDadosMES.Range("A22").Value = "'Erro:(M620010000-0002-TNLQU) Tag Name M620010000-0002-TNLQU is invalid(M620010000-0002-TNLQU) Tag Name M620010000-0004-TNLQU is invalid"
    r = r & "|invalidas=" & Replace(TesteTagsInvalidas("E"), vbCrLf, "/")
    shDadosMES.Range("A22").Value = "'Consulta concluída"
    Nm("eFase").Value = "HH"
    passo = 6
    LimparSaidasMES
    r = r & "|limpo=" & Vazio(shDadosMES.Range("A9").Formula)
    RestaurarSaidasMES
    r = r & "|restaurado=" & (InStr(1, shDadosMES.Range("A9").Formula, "ShowCalculationValues", 1) > 0)
    shConfig.Range("Z1").Value = r
    Exit Sub
Erro:
    r = r & "|ERRO passo " & passo & ": " & Err.Number & " " & Err.Description
    shConfig.Range("Z1").Value = r
End Sub

' Preenche um turno de exemplo (para as imagens de previa)
Public Sub DemoLO()
    gSilencioso = True
    Nm("cfgFonte").Value = "SIMULAÇÃO"
    Nm("pData").Value = DateSerial(2026, 9, 1)
    Nm("dLetra").Value = "C"
    Nm("dTecnico").Value = "Técnico da letra C"
    Nm("dProdUS3").Value = "PDR/MX"
    Nm("dProdUS4").Value = "PBF/MB45"
    Nm("dReal").Cells(1, 1).Value = "***Controle da Produção: PF 03 04X04h e LM 03 02X02h (Mix Coque/Moinha de Carvão, Calcário e Aglomerante) - PDR/STD2"
    Nm("dReal").Cells(2, 1).Value = " - Acompanhamento na Planilha do Batch - %SiO2, %P e PPC - Batch 242 - Mineroduto 03"
    Nm("dReal").Cells(3, 1).Value = " - Descarte das amostras bi-horárias;"
    Nm("dReal").Cells(4, 1).Value = " - Verificação da calibração Raio X, Leco SC832 e Leco CS230;"
    Nm("dReal").Cells(5, 1).Value = " - Análise de coque e calcário consumido - Composto diário - ref.: 01/09/26;"
    Nm("dReal").Cells(6, 1).Value = " - Acompanhamento do embarque no Navio SAKIZAYA MIRACLE - PBF MB45 para AM Gent;"
    Nm("dEquip").Cells(1, 1).Value = " - Linha de oxigênio de alta com vazamento - fechar a rede após testes de PCS"
    Nm("dEquip").Cells(2, 1).Value = " - Tubulação de exaustão da capela 04 com furo. Manutenção ciente."
    Nm("dAReal").Cells(1, 1).Value = "***Controle da Produção: PF 03 04X04h e LM 03 02X02h - PDR/STD2"
    Nm("dAReal").Cells(2, 1).Value = " - Acompanhamento na Planilha do Batch - %SiO2, %P e PPC - Batch 240/241 - Mineroduto 03;"
    Nm("dAReal").Cells(3, 1).Value = " - PQ e PF 03/04 - Composto diário - ref.: 30/08/26;"
    Nm("dAReal").Cells(4, 1).Value = " - Acompanhamento do embarque no Navio Ultra Cougar p/ Nucor - PDR/STD;"
    Nm("dArStatus").Value = "Desligado"
    Nm("dCompressor").Value = "Regular"
    Nm("dNitrogenio").Value = "Bom"
    Nm("dRecebe").Value = "A"
    Nm("dCadinhos").Value = "15 cadinhos: 9 p/ minérios, 3 p/ insumos e 3 retirados de uso"
    Nm("dComUS3").Value = "Produção estável. SiO2 da mistura acima do limite do PDR/MX a partir das 09:30; operação ciente."
    Nm("dComUS4").Value = "Sem desvios no turno."
    Nm("dH2Carvao").Value = "4,12"
    Nm("dObs").Value = "Linha de oxigênio fechada após os testes de PCS."
    AtualizarMESDia
    Nm("gIni").Value = DateSerial(2026, 9, 1) + TimeSerial(7, 0, 0)
    Nm("gFim").Value = DateSerial(2026, 9, 2) + TimeSerial(7, 0, 0)
    AtualizarResultados
    Nm("eIni").Value = DateSerial(2026, 9, 1) + TimeSerial(7, 0, 0)
    Nm("eFim").Value = DateSerial(2026, 9, 1) + TimeSerial(19, 0, 0)
    AtualizarEmbarque
End Sub

' Grava os argumentos que o VBA monta para cada consulta (comparados com o Python no teste)
Public Sub TesteFormulas()
    Dim b As Long
    For b = 1 To NumBlocos()
        shConfig.Cells(b, 27).Value = "'" & ArgumentosBloco(b)
    Next b
End Sub

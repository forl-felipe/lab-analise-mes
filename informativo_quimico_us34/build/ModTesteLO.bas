Option Explicit

' Modulo usado somente na validacao automatica (LibreOffice). Nao vai para a versao final.

Private Function LinhaRel(ByVal ws As Worksheet, ByVal chave As String) As Long
    Dim r As Long
    For r = 1 To REL_FIM
        If CStr(ws.Cells(r, 1).Value) = chave Then
            LinhaRel = r
            Exit Function
        End If
    Next r
End Function

Public Sub TesteLO()
    Dim r As String, passo As Long, lin As Long
    On Error GoTo Erro
    gSilencioso = True
    passo = 1
    Nm("cfgFonte").Value = "SIMULAÇÃO"
    Nm("pData").Value = DateSerial(2026, 9, 28)
    Nm("pTurno").Value = TurnoDia()
    Nm("pTurma").Value = "A"
    Nm("pResp").Value = "Teste Automatico"
    passo = 2
    AtualizarMES
    passo = 3
    lin = LinhaRel(shRelDia, "P02U3")
    r = "info=" & Nm("iTurno").Value & "|dGerado=" & (Not Vazio(Nm("rlDGerado").Value)) & _
        "|dSiO2PF=" & shRelDia.Cells(lin, 11).Value & "|dHora1=" & Nm("rlDHoras").Cells(1, 1).Value & _
        "|FeTPFoculta=" & shRelDia.Rows(LinhaRel(shRelDia, "P01U3")).Hidden & _
        "|SiO2PFvisivel=" & (Not shRelDia.Rows(lin).Hidden) & _
        "|infFeToculta=" & shInformativo.Rows(LinhaInformativo(1, 1)).Hidden & _
        "|B2US4=" & shRelDia.Cells(LinhaRel(shRelDia, "P16U4"), 11).Value
    passo = 4
    Nm("ocReal").Cells(1, 1).Value = "Controle da Produção: PF 03 04X04h e LM 03 02X02h"
    Nm("ocReal").Cells(2, 1).Value = " - Descarte das amostras bi-horárias;"
    Nm("ocAReal").Cells(1, 1).Value = "Acompanhamento do embarque"
    Nm("ocAReal").Cells(2, 1).Value = "Análise de coque consumido"
    Nm("ocSol").Cells(1, 1).Value = "sol. 1432026 PF"
    Nm("ocAus").Value = "Sim"
    Nm("ocAusQuem").Value = "Fulano"
    Nm("ocRecebe").Value = "B"
    passo = 5
    FecharTurno
    r = r & "|dFinal=" & (Not Vazio(Nm("rlDFinal").Value)) & "|dReal=" & Replace(Nm("rlDReal").Value, vbLf, "/") & _
        "|dPessoal=" & Nm("rlDPessoal").Value & "|dSol=" & Nm("rlDSol").Value & _
        "|ocRealDepois=" & Nm("ocReal").Cells(1, 1).Value & ";" & Nm("ocReal").Cells(2, 1).Value & _
        "|ocARealDepois=" & Nm("ocAReal").Cells(1, 1).Value & "|ocAusDepois=" & Nm("ocAus").Value & _
        "|ocProg=" & Left$(Nm("ocProg").Value, 10) & "|altReal=" & Nm("rlDReal").RowHeight
    passo = 6
    SelecionarNoite
    Nm("pTurma").Value = "C"
    Nm("pResp").Value = "Beltrano"
    Nm("ocSol").Cells(1, 1).Value = "Nada pendente"
    FecharTurno
    r = r & "|nFinal=" & (Not Vazio(Nm("rlNFinal").Value)) & "|nInfo=" & Nm("rlNInfo").Value & _
        "|nSol=" & Nm("rlNSol").Value & "|dSolDepois=" & Nm("rlDSol").Value & _
        "|nHora1=" & Nm("rlNHoras").Cells(1, 1).Value & _
        "|dSiO2PFdepois=" & shRelDia.Cells(lin, 11).Value & "|nSiO2PF=" & shRelNoite.Cells(lin, 11).Value
    passo = 7
    PrepararRelatorioParaCopia "rlD"
    r = r & "|dSolAposCopia=" & Nm("rlDSol").Value
    passo = 8
    Nm("iSelIni").Value = DateSerial(2026, 9, 28) + TimeSerial(7, 0, 0)
    Nm("iSelFim").Value = DateSerial(2026, 9, 29) + TimeSerial(7, 0, 0)
    AtualizarMESPeriodo
    Dim nv As Long, jj As Long
    For jj = 0 To NSLOT_INF - 1
        If ENumero(shInformativo.Cells(LinhaInformativo(12, 1), INF_COL_H1 + jj).Value) Then nv = nv + 1
    Next jj
    r = r & "|per=" & Nm("iTurno").Value & "|perSlots=" & nv
    passo = 9
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
    Nm("pTurno").Value = TurnoDia()
    Nm("pTurma").Value = "C"
    Nm("pResp").Value = "Técnico da letra C"
    shConfig.Cells(CFG_ROW1 + 11, CFG_COL_LSE).Value = 1.85
    AtualizarMES
    Nm("ocReal").Cells(1, 1).Value = "***Controle da Produção: PF 03 04X04h e LM 03 02X02h (Mix Coque/Moinha de Carvão, Calcário e Aglomerante) - PDR/STD2"
    Nm("ocReal").Cells(2, 1).Value = " - Acompanhamento na Planilha do Batch - %SiO2, %P e PPC - Batch 242 - Mineroduto 03"
    Nm("ocReal").Cells(3, 1).Value = " - Descarte das amostras bi-horárias;"
    Nm("ocReal").Cells(4, 1).Value = " - Verificação da calibração Raio X, Leco SC832 e Leco CS230;"
    Nm("ocReal").Cells(5, 1).Value = " - Análise de coque e calcário consumido - Composto diário - ref.: 01/09/26;"
    Nm("ocReal").Cells(6, 1).Value = " - Acompanhamento do embarque no Navio SAKIZAYA MIRACLE - PBF MB45 para AM Gent;"
    Nm("ocEquip").Cells(1, 1).Value = " - Linha de oxigênio de alta com vazamento - fechar a rede após testes de PCS"
    Nm("ocEquip").Cells(2, 1).Value = " - Tubulação de exaustão da capela 04 com furo. Manutenção ciente."
    Nm("ocAReal").Cells(1, 1).Value = "***Controle da Produção: PF 03 04X04h e LM 03 02X02h - PDR/STD2"
    Nm("ocAReal").Cells(2, 1).Value = " - Acompanhamento na Planilha do Batch - %SiO2, %P e PPC - Batch 240/241 - Mineroduto 03;"
    Nm("ocAReal").Cells(3, 1).Value = " - PQ e PF 03/04 - Composto diário - ref.: 30/08/26;"
    Nm("ocAReal").Cells(4, 1).Value = " - Acompanhamento do embarque no Navio Ultra Cougar p/ Nucor - PDR/STD;"
    Nm("ocArStatus").Value = "Desligado"
    Nm("ocCompressor").Value = "Regular"
    Nm("ocNitrogenio").Value = "Bom"
    Nm("ocRecebe").Value = "A"
    Nm("ocCadinhos").Value = "15 cadinhos: 9 p/ minérios, 3 p/ insumos e 3 retirados de uso"
    GerarRelatorio "rlD"
End Sub

' Grava os argumentos que o VBA monta para cada consulta (comparados com o Python no teste)
Public Sub TesteFormulas()
    Dim b As Long
    For b = 1 To NumBlocos()
        shConfig.Cells(b, 27).Value = "'" & ArgumentosBloco(b)
    Next b
End Sub

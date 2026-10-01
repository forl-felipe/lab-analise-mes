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

' Linha do Resumo cujo texto (coluna B ou C) e igual a t
Private Function LinhaTexto(ByVal ws As Worksheet, ByVal t As String) As Long
    Dim r As Long
    For r = 1 To 150
        If ws.Cells(r, 2).Text = t Or ws.Cells(r, 3).Text = t Then
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
    ' estado apos a 1a consulta real: formulas do Aspen desligadas (texto), prefixo guardado
    Nm("mesPrefixo").Value = "'" & Left$(shDadosMES.Range("A7").Formula, InStr(shDadosMES.Range("A7").Formula, "("))
    shDadosMES.Range("A7").Value = "'Consulta concluída"
    shDadosMES.Range("A17").Value = "'Consulta concluída"
    shDadosMES.Range("A27").Value = "'Consulta concluída"
    Nm("cfgFonte").Value = "SIMULAÇÃO"
    Nm("pData").Value = DateSerial(2026, 9, 28)
    Nm("dLetra").Value = "A"
    Nm("dTecnico").Value = "Fulano"
    Nm("dProdUS3").Value = "PDR/STD"
    Nm("dProdUS4").Value = "PBF/MB45"
    Nm("dComUS3").Value = "Campanha PDR/STD - qualidade física não atendeu aos limites de processo durante todo o turno."
    Nm("dDragas").Cells(1, 1).Value = "Draga 01: parada disponível;"
    Nm("dBatch").Value = 267
    Nm("dBatchSiO2").Value = 1.83
    Nm("dBatchPPC").Value = 3.64
    Nm("dFiltragem").Cells(1, 1).Value = "Elutriando durante o turno - períodos intercalados, devido aos níveis de tanques da polpa 01, a elutriação está gerando mais polpa do que a capacidade dos filtros prensa tem de processar;"
    Nm("dCarreg").Cells(1, 1).Value = "Carregamento para navio IOANNIS, cliente GREEN STEEL"
    Nm("nLetra").Value = "C"
    Nm("nTecnico").Value = "Beltrano"
    Nm("nDragas").Cells(1, 1).Value = "Dragas da noite"
    AjustarAlturas shPreenchimento, Nm("dFiltragem").Cells(1, 1)
    r = "altPreLonga=" & shPreenchimento.Rows(Nm("dFiltragem").Cells(1, 1).Row).RowHeight & "|"
    passo = 2
    Set ws = shResumoDia
    AtualizarMESDia
    lin = LinhaChave(ws, "P08U3")
    r = r & "SiO2PF_media=" & ws.Cells(lin, 11).Value & "|PF_umaLinha=" & (LinhaChave(ws, "P08U4") = 0) & _
        "|PFrotulo=" & ws.Cells(lin, 4).Value & "|dHora1=" & Nm("dHoras").Cells(1, 1).Value & _
        "|prodSoma=" & (Abs(ws.Cells(LinhaChave(ws, "P21U3"), 11).Value - _
            Application.WorksheetFunction.Sum(ws.Range(ws.Cells(LinhaChave(ws, "P21U3"), 5), ws.Cells(LinhaChave(ws, "P21U3"), 10)))) < 0.001) & _
        "|ritmoUltimo=" & ws.Cells(LinhaChave(ws, "P23U3"), 11).Value
    passo = 22
    ws.Calculate
    r = r & "|limTambUS3=" & ws.Cells(LinhaChave(ws, "P17U3"), 17).Text & _
        "|limTambUS4=" & ws.Cells(LinhaChave(ws, "P17U4"), 17).Text & _
        "|limFinosUS4=" & ws.Cells(LinhaChave(ws, "P20U4"), 18).Text & _
        "|limRG=" & ws.Cells(LinhaChave(ws, "P16U3"), 17).Text & "/" & ws.Cells(LinhaChave(ws, "P16U3"), 18).Text & _
        "|limSiO2MistUS3=" & ws.Cells(LinhaChave(ws, "P10U3"), 18).Text & _
        "|limSiO2PF=" & ws.Cells(LinhaChave(ws, "P08U3"), 18).Text & "|limPPCPF=" & ws.Cells(LinhaChave(ws, "P07U3"), 18).Text & _
        "|limB2US4=" & ws.Cells(LinhaChave(ws, "P12U4"), 18).Text
    passo = 3
    r = r & "|batch=" & ws.Cells(LinhaTexto(ws, "Mineroduto 03") + 1, 2).Text & _
        "|dragasVisivel=" & (Not ws.Rows(LinhaTexto(ws, "Dragas")).Hidden) & _
        "|mercIntOculto=" & ws.Rows(LinhaTexto(ws, "Mercado interno")).Hidden & _
        "|equipOculto=" & ws.Rows(LinhaTexto(ws, "Equipamentos em operação")).Hidden & _
        "|obsOculto=" & ws.Rows(LinhaTexto(ws, "Observações gerais")).Hidden & _
        "|altFiltragem=" & ws.Rows(LinhaTexto(ws, "Filtragem") + 1).RowHeight & "/" & Len(ws.Cells(LinhaTexto(ws, "Filtragem") + 1, 2).Text) & "/" & ws.Cells(LinhaTexto(ws, "Filtragem") + 1, 2).MergeArea.Width & "/" & ws.Cells(LinhaTexto(ws, "Filtragem") + 1, 2).WrapText & _
        "|letra=" & ws.Cells(6, 11).Text & "|tecnico=" & ws.Cells(7, 3).Text
    For jj = 1 To 150
        If CStr(ws.Cells(jj, 2).Value) = "Comentário US4" Then r = r & "|comUS4oculto=" & ws.Rows(jj).Hidden
    Next jj
    passo = 4
    AtualizarMESNoite
    Set ws = shResumoNoite
    r = r & "|nHora1=" & Nm("nHoras").Cells(1, 1).Value & "|nTecnico=" & ws.Cells(7, 3).Text & _
        "|nDragas=" & (LinhaTexto(ws, Chr(149) & "   Dragas da noite") > 0) & _
        "|nMinerodutoOculto=" & ws.Rows(LinhaTexto(ws, "Mineroduto 03")).Hidden
    passo = 5
    Nm("gIni").Value = DateSerial(2026, 9, 28) + TimeSerial(7, 0, 0)
    Nm("gFim").Value = DateSerial(2026, 9, 29) + TimeSerial(7, 0, 0)
    AtualizarResultados
    lin = LinhaChave(shResultados, "P17U3")
    For jj = 0 To NSLOT_MAX - 1
        If ENumero(shResultados.Cells(lin, COL_SLOT1 + jj).Value) Then nv = nv + 1
    Next jj
    r = r & "|g24h_slots=" & nv & "|gH1=" & Nm("gHoras").Cells(1, 1).Value & "|gH12=" & Nm("gHoras").Cells(1, 12).Value
    Nm("gIni").Value = DateSerial(2026, 9, 28) + TimeSerial(0, 0, 0)
    Nm("gFim").Value = DateSerial(2026, 9, 28) + TimeSerial(12, 0, 0)
    AtualizarResultados
    r = r & "|gH1_00h=" & Nm("gHoras").Cells(1, 1).Value
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
    Nm("pData").Value = DateSerial(2026, 9, 30)
    Nm("dLetra").Value = "B"
    Nm("dTecnico").Value = "Técnico da letra B"
    Nm("dProdUS3").Value = "PDR/STD"
    Nm("dProdUS4").Value = "PDR/STD"
    Nm("dComUS3").Value = "Campanha PDR/STD - qualidade física não atendeu aos limites de processo durante todo o turno. Perdas nos limites de compressão, impactado diretamente pela elevada alimentação nas usinas e retorno do material das pilhas de emergência."
    Nm("dDragas").Cells(1, 1).Value = "Draga 01: parada disponível;"
    Nm("dDragas").Cells(2, 1).Value = "Draga 02: parada disponível;"
    Nm("dBatch").Value = 267
    Nm("dBatchSiO2").Value = 1.83
    Nm("dBatchP").Value = 0.06
    Nm("dBatchPPC").Value = 3.64
    Nm("dBatchSE").Value = 1983
    Nm("dBatchM325").Value = 85.4
    Nm("dFiltragem").Cells(1, 1).Value = "Elutriando durante o turno - períodos intercalados, devido aos níveis de tanques da polpa 01, a elutriação está gerando mais polpa do que a capacidade dos filtros prensa tem de processar;"
    Nm("dFiltragem").Cells(2, 1).Value = "Dosando surfactante para as usinas;"
    Nm("dFiltragem").Cells(3, 1).Value = "Diluição com água quente no processo."
    Nm("dFiltragem").Cells(4, 1).Value = "Não incorporou material da praça do Roller Press da usina 04;"
    Nm("dFiltragem").Cells(5, 1).Value = "Filtro prensa Matec e demais filtros em operação."
    Nm("dFiltragem").Cells(6, 1).Value = "Não filtrou para o pátio durante o turno;"
    Nm("dSol").Cells(1, 1).Value = "Solicitação 1632026 - Análise do carvão moído - U03 e U04: segue de 12x12h;"
    Nm("dInsumos").Cells(1, 1).Value = "Calcário IMERYS, Calcário Importado IMI, Coproduto CAL BRASIL, Coproduto IMIL"
    Nm("dCarreg").Cells(1, 1).Value = "Carregamento para navio IOANNIS, cliente GREEN STEEL, total de 153.700 TMN de PDR/STD; o carregamento iniciou no dia 25/09/2026 às 16:34h, total de 54.350 TMN embarcado até o momento;"
    Nm("dEquip").Cells(1, 1).Value = "Prensa de rolos US3 e US4, peneiras de pelota e filtros em operação."
    AjustarAlturas shPreenchimento, shPreenchimento.UsedRange
    AtualizarMESDia
    Nm("gIni").Value = DateSerial(2026, 9, 30) + TimeSerial(7, 0, 0)
    Nm("gFim").Value = DateSerial(2026, 10, 1) + TimeSerial(7, 0, 0)
    AtualizarResultados
End Sub

' Grava os argumentos que o VBA monta para cada consulta (comparados com o Python no teste)
Public Sub TesteFormulas()
    Dim b As Long
    For b = 1 To NumBlocos()
        shConfig.Cells(b, 27).Value = "'" & ArgumentosBloco(b)
    Next b
End Sub

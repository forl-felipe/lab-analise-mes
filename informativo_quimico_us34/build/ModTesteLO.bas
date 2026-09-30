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
    For r = 1 To 80
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
    Nm("cfgFonte").Value = "SIMULAÇÃO"
    Nm("pData").Value = DateSerial(2026, 9, 28)
    Nm("dLetra").Value = "A"
    Nm("dTecnico").Value = "Fulano"
    Nm("dReal").Cells(1, 1).Value = "Controle da Produção: PF 03 04X04h"
    Nm("dReal").Cells(2, 1).Value = " - Descarte das amostras bi-horárias;"
    Nm("dSol").Cells(1, 1).Value = "sol. 1432026 PF"
    Nm("dAus").Value = "Sim"
    Nm("dAusQuem").Value = "Ciclano"
    Nm("dRecebe").Value = "B"
    Nm("dH2Coque").Value = "0,45"
    Nm("nLetra").Value = "C"
    Nm("nTecnico").Value = "Beltrano"
    Nm("nReal").Cells(1, 1).Value = "Tarefa da noite"
    passo = 2
    Set ws = shResumoDia
    AtualizarMESDia
    lin = LinhaChave(ws, "P02U3")
    r = "dSiO2PF_slot1=" & ws.Cells(lin, 5).Value & "|dSiO2PF_media=" & ws.Cells(lin, 11).Value & _
        "|dHora1=" & Nm("dHoras").Cells(1, 1).Value & "|dAtual=" & Left$(Nm("dAtualizado").Value, 40) & _
        "|FeTPFoculta=" & ws.Rows(LinhaChave(ws, "P01U3")).Hidden & _
        "|SiO2PFvisivel=" & (Not ws.Rows(lin).Hidden) & _
        "|B2US4=" & ws.Cells(LinhaChave(ws, "P16U4"), 11).Value
    passo = 3
    ' migracao por formula
    lin = LinhaTexto(ws, "Controle da Produção: PF 03 04X04h")
    r = r & "|realLinha1=" & (lin > 0) & "|realVisivel=" & (Not ws.Rows(lin).Hidden) & _
        "|real2=" & ws.Cells(lin + 1, 3).Value & "|real3oculta=" & ws.Rows(lin + 2).Hidden & _
        "|equipVazio=" & (LinhaTexto(ws, "— nada registrado") > 0) & _
        "|info=" & ws.Cells(6, 2).Value
    For jj = 1 To 80
        If Left$(CStr(ws.Cells(jj, 2).Value), 7) = "Pessoal" Then r = r & "|pessoal=" & ws.Cells(jj, 3).Text
        If Left$(CStr(ws.Cells(jj, 2).Value), 11) = "Observações" Then r = r & "|obs=" & ws.Cells(jj, 3).Text
        If Left$(CStr(ws.Cells(jj, 2).Value), 7) = "Control" Then r = r & "|lab=" & Replace(ws.Cells(jj, 3).Text, vbLf, "/")
    Next jj
    passo = 4
    AtualizarMESNoite
    Set ws = shResumoNoite
    lin = LinhaChave(ws, "P02U3")
    r = r & "|nHora1=" & Nm("nHoras").Cells(1, 1).Value & "|nSiO2PF_media=" & ws.Cells(lin, 11).Value & _
        "|nInfo=" & ws.Cells(6, 2).Value & "|nReal=" & (LinhaTexto(ws, "Tarefa da noite") > 0) & _
        "|dMediaIntacta=" & shResumoDia.Cells(LinhaChave(shResumoDia, "P02U3"), 11).Value
    passo = 5
    Nm("gIni").Value = DateSerial(2026, 9, 28) + TimeSerial(7, 0, 0)
    Nm("gFim").Value = DateSerial(2026, 9, 29) + TimeSerial(7, 0, 0)
    AtualizarResultados
    lin = LinhaChave(shResultados, "P12U3")
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
    lin = LinhaChave(shResultados, "P12U3")
    For jj = 0 To 2
        If ENumero(shResultados.Cells(lin, COL_SLOT1 + jj).Value) Then cheias = cheias + 1
    Next jj
    vazia3 = Vazio(shResultados.Cells(lin, COL_SLOT1 + 2).Value)
    r = r & "|andamento_cheias=" & cheias & "|andamento_janela3vazia=" & vazia3 & _
        "|andamento_atual=" & Left$(Nm("gAtualizado").Value, 60)
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
    shConfig.Cells(CFG_ROW1 + 11, CFG_COL_LSE).Value = 1.85
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
    AtualizarMESDia
    Nm("gIni").Value = DateSerial(2026, 9, 1) + TimeSerial(7, 0, 0)
    Nm("gFim").Value = DateSerial(2026, 9, 2) + TimeSerial(7, 0, 0)
    AtualizarResultados
End Sub

' Grava os argumentos que o VBA monta para cada consulta (comparados com o Python no teste)
Public Sub TesteFormulas()
    Dim b As Long
    For b = 1 To NumBlocos()
        shConfig.Cells(b, 27).Value = "'" & ArgumentosBloco(b)
    Next b
End Sub

Option Explicit

' Modulo usado somente na validacao automatica (LibreOffice). Nao vai para a versao final.
Public Sub TesteLO()
    Dim r As String
    Dim c As Range, passo As Long
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
    r = "info=" & Nm("iTurno").Value & "|" & shInformativo.Cells(LinhaInformativo(18, 1), INF_COL_H1).Value
    r = r & "|dGeradoAposMES=" & (Not Vazio(Nm("rsDGerado").Value))
    shPassagem.Range("C10").Value = "Fulano"
    Nm("ptRecebe").Value = "B"
    For Each c In Nm("ptStatus").Cells
        c.Value = "Operando"
    Next c
    Nm("ptStatus").Cells(2, 1).Value = "Não operando"
    Nm("ptPendencias").Value = "Refazer tamboramento US4"
    passo = 4
    FecharTurno
    passo = 5
    r = r & "|pendAposLimpar=" & Nm("ptPendencias").Value & "|dFinal=" & (Not Vazio(Nm("rsDFinal").Value)) & _
        "|dEquipe=" & Nm("rsDEquipe").Value & "|dPend=" & Nm("rsDPend").Value & "|dNaoOper=" & Nm("rsDNaoOper").Value & _
        "|dRes=" & shResumoDia.Cells(11, 4).Value
    passo = 6
    ' segundo turno: Noite do mesmo dia (Finalizar busca o MES sozinho)
    Nm("pTurno").Value = TurnoNoite()
    Nm("pTurma").Value = "C"
    shPassagem.Range("C10").Value = "Beltrano"
    Nm("ptPendencias").Value = "Nada pendente"
    FecharTurno
    r = r & "|noite=" & Nm("iTurno").Value & "|nEquipe=" & Nm("rsNEquipe").Value & "|nPend=" & Nm("rsNPend").Value & _
        "|dEquipeDepois=" & Nm("rsDEquipe").Value & "|dResDepois=" & shResumoDia.Cells(11, 4).Value & _
        "|nRes=" & shResumoNoite.Cells(11, 4).Value
    passo = 7
    ' copia de um resumo finalizado nao pode refazer o outro turno
    PrepararResumoParaCopia "rsD"
    r = r & "|dEquipeAposCopia=" & Nm("rsDEquipe").Value
    Dim d As Date, t As String, hs As Variant, i As Long
    hs = Array(6.5, 7.5, 8.5, 18.5, 19.5, 20.5, 23)
    For i = 0 To UBound(hs)
        TurnoDeReferencia CDbl(DateSerial(2026, 9, 28)) + hs(i) / 24, d, t
        r = r & "|" & hs(i) & "h=" & Format$(d, "dd") & Left$(t, 1)
    Next i
    shConfig.Range("Z1").Value = r
    Exit Sub
Erro:
    r = r & "|ERRO passo " & passo & ": " & Err.Number & " " & Err.Description
    shConfig.Range("Z1").Value = r
End Sub

Public Sub DemoLO()
    gSilencioso = True
    Nm("cfgFonte").Value = "SIMULAÇÃO"
    Nm("pData").Value = DateSerial(2026, 9, 28)
    Nm("pTurno").Value = TurnoDia()
    Nm("pTurma").Value = "A"
    Nm("pResp").Value = "Maria Silva"
    shConfig.Range("K35").Value = 300
    shConfig.Range("L35").Value = 340
    shConfig.Range("K34").Value = 94
    AtualizarMES
    shPassagem.Range("C10").Value = "Maria Silva"
    shPassagem.Range("C11").Value = "Carlos Lima"
    shPassagem.Range("C12").Value = "João Souza"
    Nm("ptRecebe").Value = "B"
    Nm("ptTestes").Value = "Tamboramento e compressão US3/US4 conforme plano."
    Nm("ptPendencias").Value = "Refazer granulometria US4 das 15h."
    Nm("ptStatus").Cells(1, 1).Value = "Operando"
    Nm("ptStatus").Cells(2, 1).Value = "Não operando"
    Nm("ptStatus").Cells(2, 1).Offset(0, 2).Value = "Aguardando manutenção mecânica"
    Nm("ptStatus").Cells(3, 1).Value = "Operando"
    GerarResumo "rsD"
End Sub

' Grava os argumentos que o VBA monta para cada consulta (comparados com o Python no teste)
Public Sub TesteFormulas()
    Dim b As Long
    For b = 1 To NumBlocos()
        shConfig.Cells(b, 27).Value = "'" & ArgumentosBloco(b)
    Next b
End Sub

Public Sub TesteDatas()
    Dim modos As Variant, i As Long, ini As Date
    modos = Array("Data do Excel", "Texto dd/mm/aaaa", "Texto mm/dd/aaaa")
    ini = DateSerial(2026, 9, 28) + TimeSerial(7, 0, 0)
    For i = 0 To 2
        Nm("cfgFormatoData").Value = modos(i)
        GravarPeriodoMES ini, ini + 0.5
        shConfig.Cells(10 + i, 27).Value = "'" & modos(i) & " | " & TypeName(Nm("mesInicio").Value) & " | " & Nm("mesInicio").Text & " | " & Nm("mesFim").Text
    Next i
End Sub

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
    shPassagem.Range("C10").Value = "Fulano"
    Nm("ptRecebe").Value = "B"
    For Each c In Nm("ptStatus").Cells
        c.Value = "OK"
    Next c
    Nm("ptStatus").Cells(2, 1).Value = "NÃO OK"
    Nm("ptPendencias").Value = "Refazer tamboramento US4"
    passo = 4
    FecharTurno
    passo = 5
    r = r & "|apos=" & Nm("pTurno").Value & "|" & Nm("pTurma").Value & "|" & Nm("ptPendRecebidas").Value
    Nm("ptPendRecebidas").Value = ""
    Nm("pResp").Value = "Outro"
    passo = 6
    CarregarPendencias
    r = r & "|pend=" & Nm("ptPendRecebidas").Value
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
    shPassagem.Range("C11").Value = "João Souza"
    Nm("ptRecebe").Value = "B"
    Nm("ptSegFlag").Value = "Não"
    Nm("ptTestes").Value = "Tamboramento e compressão US3/US4 conforme plano."
    Nm("ptPendencias").Value = "Refazer granulometria US4 das 15h."
    Nm("ptStatus").Cells(1, 1).Value = "OK"
    Nm("ptStatus").Cells(2, 1).Value = "NÃO OK"
    shPassagem.Range("H42").Value = "Aguardando manutenção mecânica"
    Nm("ptStatus").Cells(3, 1).Value = "FORA DE OPERAÇÃO"
    shPassagem.Range("C58").Value = "PBF"
End Sub

' Grava os argumentos que o VBA monta para cada consulta (comparados com o Python no teste)
Public Sub TesteFormulas()
    Dim b As Long
    For b = 1 To NumBlocos()
        shConfig.Cells(b, 27).Value = "'" & ArgumentosBloco(b)
    Next b
End Sub

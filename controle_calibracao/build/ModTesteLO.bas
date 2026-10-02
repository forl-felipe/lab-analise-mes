Option Explicit

Public gPasso As String

' Modulo usado somente na validacao automatica (LibreOffice). Nao vai para a versao final.

' Endereco da k-esima entrada do ensaio (ordem do mapaInputs)
Private Function E(ByVal cod As String, ByVal k As Long) As Range
    Dim m As Variant, i As Long, p As Variant, a As String
    m = Nm("mapaInputs").Value
    For i = 1 To UBound(m, 1)
        If CStr(m(i, 1)) = cod Then
            p = Split(CStr(m(i, 3)), ";")
            a = p(k - 1)
            If InStr(a, ":") > 0 Then a = Left$(a, InStr(a, ":") - 1)
            Set E = shLancamento.Range(a)
            Exit Function
        End If
    Next i
End Function

Private Function ContaBD() As Long
    ContaBD = UltimaLinhaBD() - 1
End Function

Private Function TextoLan(ByVal t As String) As Long
    Dim r As Long, c As Long
    For r = 1 To 400
        For c = 2 To 4
            If shLancamento.Cells(r, c).Text = t Then
                TextoLan = r
                Exit Function
            End If
        Next c
    Next r
End Function

Public Sub PreencherExemplo()
    Dim k As Long
    gPasso = "a"
    Nm("pData").Value = DateSerial(2026, 10, 5)
    Nm("pTurno").Value = "19x07"
    Nm("pLetra").Value = "B"
    Nm("pResp").Value = "Felipe"
    gPasso = "b"
    ' Blaine: manual, star, automatico (automatico fora: 2100 > 2074)
    E("BLA", 1).Value = 2011
    E("BLA", 2).Value = 2120
    E("BLA", 3).Value = 2100
    ' Tambores: (min, s) 66TA05 7:30 -> 25,07 rpm; 66TA06 6:00 -> 31,3 rpm (fora); 66TA07 8:00 (200 voltas) -> 25,0
    E("TAM", 1).Value = 7: E("TAM", 2).Value = 30
    E("TAM", 7).Value = 6: E("TAM", 8).Value = 0
    E("TAM", 9).Value = 8: E("TAM", 10).Value = 0
    ' Alpine
    E("ALP", 1).Value = 88
    E("ALP", 2).Value = 88.5
    ' Umidade: estufa pelas pesagens (100 / 600 / 555 -> 9,00 %)
    E("UMI", 2).Value = 100: E("UMI", 3).Value = 600: E("UMI", 4).Value = 555
    E("UMI", 5).Value = 9.03
    E("UMI", 6).Value = 9.1
    ' Compressao FX 16: PS04 365 / DP 40 / 15,2 mm/min; PS10 350
    E("COM", 1).Value = 365: E("COM", 2).Value = 40: E("COM", 3).Value = 15.2
    E("COM", 10).Value = 350
    ' Granulometria (g): LTF e LCE
    Dim ltf As Variant, lce As Variant
    ltf = Array(10, 400, 1300, 1700, 1500, 250, 120, 40, 20, 10)
    lce = Array(12, 380, 1310, 1690, 1520, 240, 125, 45, 18, 9)
    For k = 0 To 9
        E("GRA", 1 + k * 2).Value = ltf(k)
        E("GRA", 2 + k * 2).Value = lce(k)
    Next k
    ' Tamboramento 5 x 15 kg
    E("T515", 1).Value = 4700: E("T515", 2).Value = 100
    E("T515", 3).Value = 14070: E("T515", 4).Value = 330
    gPasso = "c"
    ' Peneiradores: tudo OK, menos 1 NAO OK e 1 SEM TAG
    MarcarPeneiradoresOK
    gPasso = "d"
    E("PEN", 3).Value = "nao ok"
    E("PEN", 6).Value = "SEM TAG"
    ' Fisher
    E("FIS", 1).Value = 1900: E("FIS", 2).Value = 2150
End Sub

Public Sub PreencherObs()
    Nm("obs_BLA").Value = "Blaine automático recalibrado."
    Nm("obs_TAM").Value = "Tambor 66TA06 com correia patinando."
    Nm("obs_UMI").Value = "Analisador 66AN11 em verificação."
    Nm("obs_PEN").Value = "Peneirador com vibração anormal."
End Sub

Public Sub TesteLO()
    Dim r As String, passo As Long, n0 As Long, n1 As Long, nsel As Long, i As Long, lin As Long
    On Error GoTo Erro
    gSilencioso = True
    passo = 1
    n0 = ContaBD()
    PreencherExemplo
    Recalcular
    passo = 2
    lin = TextoLan("66TA05")
    r = "base0=" & n0 & "|rpmTA05=" & Format$(shLancamento.Cells(lin, 9).Value, "0.00") & _
        "|resTA05=" & shLancamento.Cells(lin, 14).Text & _
        "|resTA06=" & shLancamento.Cells(TextoLan("66TA06"), 14).Text & _
        "|estufa=" & Format$(shLancamento.Cells(TextoLan("Pesagens da estufa (g)"), 14).Value, "0.00") & _
        "|resAN11=" & shLancamento.Cells(TextoLan("66AN11"), 12).Text & _
        "|pen3=" & E("PEN", 3).Value & _
        "|agendaTAM=" & shLancamento.Cells(TextoLan("2. Tambor de abrasão · rotação") , 9).Text & _
        "|agendaUMI=" & shLancamento.Cells(TextoLan("4. Umidade · analisadores x estufa"), 11).Text & _
        "|agendaBLAres=" & shLancamento.Cells(TextoLan("1. Calibração Blaine"), 13).Text & _
        "|cntBLA=" & Nm("cnt_BLA").Value & "|ultBLA=" & shLancamento.Cells(TextoLan("1. Calibração Blaine"), 15).Text
    shStaging.Calculate
    For i = 2 To NSTG + 1
        If shStaging.Cells(i, 18).Value = 1 Then nsel = nsel + 1
    Next i
    r = r & "|aRegistrar=" & nsel
    passo = 3
    ' nao conforme sem observacao: registro bloqueado
    RegistrarLancamento
    r = r & "|bloqueado=" & (ContaBD() - n0) & "|msgBloq=" & Left$(Replace(gUltimaMsg, vbCrLf, "/"), 120)
    PreencherObs
    Nm("resp_TAM").Value = "Anderson"
    RegistrarLancamento
    n1 = ContaBD()
    r = r & "|registrados=" & (n1 - n0) & "|id=" & Nm("cfgUltimoID").Value & "|cntBLApos=" & Nm("cnt_BLA").Value & _
        "|pDataMantida=" & Format$(Nm("pData").Value, "dd/mm/yyyy")
    ' conferencia de algumas linhas gravadas
    Dim a As Variant, achou As String
    a = shBD.Range("A" & (n0 + 2) & ":U" & (n1 + 1)).Value
    For i = 1 To UBound(a, 1)
        If a(i, 4) = "66TA06" Then achou = achou & "|TA06=" & a(i, 13) & "/" & Format$(a(i, 7), "0.0")
        If a(i, 5) = "Blaine Automático" Then achou = achou & "|BLAA=" & a(i, 13) & "/obs:" & a(i, 16)
        If a(i, 4) = "66PS04" And Left$(a(i, 5), 4) = "Resi" Then achou = achou & "|PS04=" & a(i, 9) & "/" & a(i, 13) & "/obs:" & a(i, 16)
        If a(i, 5) = "RG (Relação)" Then achou = achou & "|RG=" & Format$(a(i, 7), "0.000") & "/" & Format$(a(i, 8), "0.000") & "/" & a(i, 13)
        If a(i, 5) = "-16,0 +8,0 mm" Then achou = achou & "|F168=" & a(i, 13)
        If a(i, 4) = "PN0482" Then achou = achou & "|PN0482=" & a(i, 7) & "/" & a(i, 8) & "/" & a(i, 13) & "/obs:" & a(i, 16)
        If a(i, 4) = "PN0423" Then achou = achou & "|PN0423=" & a(i, 13) & "/obs:" & a(i, 16)
        If a(i, 5) = "+6,3 mm" Then achou = achou & "|T515=" & Format$(a(i, 9), "0.00") & "/" & a(i, 13)
        If a(i, 4) = "66TA05" Then achou = achou & "|respTA05=" & a(i, 14)
        If a(i, 4) = "66AG09" Then achou = achou & "|respAG09=" & a(i, 14)
        If i = 1 Then achou = achou & "|turno=" & a(i, 18) & "|regEm=" & Format$(a(i, 19), "dd/mm hh:mm") & "|turnoReg=" & a(i, 20) & "|fora=" & a(i, 21)
        If i = 1 Then achou = achou & "|origem=" & a(i, 17) & "|data=" & Format$(a(i, 3), "dd/mm/yyyy") & "|resp=" & a(i, 14) & "|letra=" & a(i, 15)
    Next i
    r = r & achou
    passo = 4
    ' painel
    Nm("painelIni").Value = DateSerial(2026, 10, 1)
    Nm("painelFim").Value = DateSerial(2026, 10, 5)
    AtualizarPainel
    r = r & "|kpiReg=" & Nm("kpiReg").Value & "|kpiConf=" & Format$(Nm("kpiConf").Value, "0.000") & _
        "|kpiNC=" & Nm("kpiNC").Value & "|kpiAder=" & Nm("kpiAder").Text & "|aderSub=" & Nm("kpiAderSub").Value & _
        "|ensBLA=" & shPainel.Cells(PAI_ENS1, 6).Value & "/" & shPainel.Cells(PAI_ENS1, 8).Value & "/" & _
        shPainel.Cells(PAI_ENS1, 10).Value & "/" & shPainel.Cells(PAI_ENS1, 16).Value & _
        "|ensPEN=" & shPainel.Cells(PAI_ENS1 + 7, 10).Value & "/" & shPainel.Cells(PAI_ENS1 + 7, 16).Value & _
        "|eq1=" & shPainel.Cells(PAI_EQ1, 4).Value & ":" & shPainel.Cells(PAI_EQ1, 9).Value & "/" & shPainel.Cells(PAI_EQ1, 16).Value & _
        "|pen1=" & shPainel.Cells(PAI_PEN1, 8).Value & "/" & shPainel.Cells(PAI_PEN1, 10).Value & "/" & shPainel.Cells(PAI_PEN1, 16).Value & _
        "|nc1=" & shPainel.Cells(PAI_NC1, 4).Value & ":" & shPainel.Cells(PAI_NC1, 6).Value & _
        "|gra1=" & shGraficos.Cells(5, 1).Value & "/" & shGraficos.Cells(5, 2).Value & "/" & shGraficos.Cells(5, 5).Value & _
        "|graOcas=" & Application.WorksheetFunction.CountA(shGraficos.Range("A5:A34")) & _
        "|info=" & Left$(Nm("painelInfo").Value, 30) & "|msg=" & gUltimaMsg
    passo = 5
    ' duplicado (mesma data/turno) com resposta "sim" no modo silencioso, depois desfazer
    PreencherExemplo
    PreencherObs
    RegistrarLancamento
    r = r & "|dup=" & (ContaBD() - n1) & "|id2=" & Nm("cfgUltimoID").Value
    DesfazerUltimo
    r = r & "|desfeito=" & (ContaBD() - n1) & "|id3=" & Nm("cfgUltimoID").Value
    passo = 6
    ' limpar
    PreencherExemplo
    LimparTela
    Recalcular
    r = r & "|limpoBLA=" & Nm("cnt_BLA").Value & "|limpoPEN=" & Nm("cnt_PEN").Value & "|tagMantida=" & _
        shLancamento.Cells(E("PEN", 1).Row, 6).Value
    passo = 7
    ' sem dados -> nada registrado
    RegistrarLancamento
    r = r & "|vazio=" & (ContaBD() - n1)
    shConfig.Range("Z1").Value = r
    Exit Sub
Erro:
    r = r & "|ERRO passo " & passo & gPasso & ": " & Err.Number & " " & Err.Description
    shConfig.Range("Z1").Value = r
End Sub

' Importacao de uma planilha antiga (caminho em Config!Z2)
Public Sub TesteImportar()
    Dim n0 As Long, n As Long
    gSilencioso = True
    n0 = ContaBD()
    n = ImportarDe(CStr(shConfig.Range("Z2").Value))
    shConfig.Range("Z3").Value = "importados=" & n & "|base=" & n0 & "->" & ContaBD()
End Sub

' Demonstracao para as previas: alguns lancamentos em dias diferentes
Public Sub DemoLO()
    Dim d As Long, k As Long
    gSilencioso = True
    For d = 1 To 5
        PreencherExemplo
        Nm("pData").Value = DateSerial(2026, 10, d)
        For k = 1 To 3
            E("BLA", k).Value = 2010 + d * 7 + k * 3 + IIf(k = 2, 100, 0)
        Next k
        E("TAM", 1).Value = 7: E("TAM", 2).Value = 25 + d
        E("TAM", 3).Value = 7: E("TAM", 4).Value = 32 - d
        E("TAM", 5).Value = 7: E("TAM", 6).Value = 28
        E("TAM", 7).Value = 7: E("TAM", 8).Value = 31 + d
        E("ALP", 1).Value = 88 + d * 0.1: E("ALP", 2).Value = 88.6 - d * 0.1: E("ALP", 3).Value = 88.3
        E("UMI", 5).Value = 9 + d * 0.01: E("UMI", 6).Value = 9.04 - d * 0.01: E("UMI", 7).Value = 8.97
        E("COM", 4).Value = 352 + d: E("COM", 7).Value = 368 - d: E("COM", 12).Value = 361
        PreencherObs
        RegistrarLancamento
    Next d
    PreencherExemplo
    Nm("pData").Value = DateSerial(2026, 10, 6)
    Nm("pTurno").Value = "07x19"
    Nm("painelIni").Value = DateSerial(2026, 10, 1)
    Nm("painelFim").Value = DateSerial(2026, 10, 6)
    AtualizarPainel
End Sub

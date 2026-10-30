Option Explicit

' Teste automatico (LibreOffice headless): grava o resultado em _Graf (nome gTeste; texto, linhas separadas por |)
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

Private Function H(ByVal r As Long, ByVal c As Long) As Variant
    H = shHist.Cells(r, c).Value
End Function

Private Function Perto(ByVal v As Variant, ByVal alvo As Double, ByVal tol As Double) As Boolean
    If ENumero(v) Then Perto = (Abs(CDbl(v) - alvo) <= tol)
End Function

Private Function GColT(ByVal chave As String) As Long
    Dim l As Variant, i As Long
    l = Split(G_CHAVES, "|")
    For i = LBound(l) To UBound(l)
        If CStr(l(i)) = chave Then GColT = i + 1
    Next i
End Function

Private Sub Data(ByVal d As Date)
    Desproteger shDia
    Nm("dData").Value = d
    Proteger shDia
    CarregarDia
End Sub

Public Sub TesteLO()
    Dim r As Long, r2 As Long, r3 As Long, s As String, tj As Date, xx As Double, k As Long
    On Error GoTo Erro
    gSilencioso = True
    gTesteLO = True
    mLog = ""
    mFalhas = 0
    Recalcular

    ' --- base importada da planilha antiga (exemplo do Word: 18/08/2026)
    r = LinhaHist(DateSerial(2026, 8, 18))
    Ok r > 0, "18/08/2026 no histórico importado (linha " & r & ")"
    Ok H(r, H_COL_EB6) = 2184 And H(r, H_COL_CONCC3) = 23979 And H(r, H_COL_CONCC2) = 19601, _
       "18/08: EB6 2184, conc C3 23979, conc C2 19601"
    Ok Perto(H(r, H_COL_ESPP), 2266.07, 0.01), "18/08: ESP C2+C3 ponderada 2266,07 (" & H(r, H_COL_ESPP) & ")"
    Ok Perto(H(r, H_COL_EB4P), 2155.21, 0.01), "18/08: EB4+ESP C3 ponderada 2155,21 (" & H(r, H_COL_EB4P) & ")"
    Ok Perto(H(r, H_COL_DIFK), -82.07, 0.01) And Perto(H(r, H_COL_DIFJ), 28.79, 0.01), "18/08: diferenças -82,07 e +28,79"
    Ok Perto(H(r, H_COL_PCTLAMA), 0, 0.000001), "18/08: % lama 0"
    r = LinhaHist(DateSerial(2026, 10, 14))
    Ok H(r, H_COL_LAMAC3) = 412 And H(r, H_COL_LAMAC2) = 50 And H(r, H_COL_LAMA) = 462, "14/10: lama 412 + 50 = 462"
    Ok Perto(H(r, H_COL_PCTLAMA), 462 / (16694 + 21216), 0.0000001), "14/10: % lama " & H(r, H_COL_PCTLAMA)
    Ok UltimoDiaHist() = CLng(DateSerial(2026, 10, 15)), "último dia do histórico 15/10/2026"

    ' --- Painel: 15 dias até 15/10 + coluna Média
    Ok DiaDe(Nm("rFim").Value) = CLng(DateSerial(2026, 10, 15)) And DiasPainel() = 15, "Painel abre em 15/10, 15 dias"
    Ok DiaDe(shGraf.Cells(2, 1).Value) = CLng(DateSerial(2026, 10, 1)), "1º dia dos gráficos 01/10"
    Ok CStr(shGraf.Cells(2, 2).Value) = "01/10", "rótulo 01/10 (" & shGraf.Cells(2, 2).Value & ")"
    Ok Nm("gUlt_eb6").Value = 2210 And Nm("gN_eb6").Value = 15, "último EB6 2210 e 15 dias com valor"
    Ok CStr(shGraf.Cells(17, 2).Value) = "Média", "16ª coluna dos gráficos = Média"
    Ok Perto(shGraf.Cells(17, 4).Value, Application.WorksheetFunction.Average(shGraf.Range("D2:D16")), 0.001), _
       "barra Média do EB6 = média do período (" & shGraf.Cells(17, 4).Value & ")"
    Ok InStr(shGraf.Cells(17, 5).Text, "#N") = 1 And Perto(shGraf.Cells(17, 6).Value, Nm("gMed_difK").Value, 0.001), _
       "na Média: linha da diferença vazia e losango = diferença média"
    Ok InStr(shGraf.Cells(18, 4).Text, "#N") = 1, "depois da Média = #N/D (" & shGraf.Cells(18, 4).Text & ")"
    Ok Perto(Nm("gMed_difK").Value, Application.WorksheetFunction.Average(shGraf.Range("E2:E16")), 0.001), _
       "média do período da diferença (" & Nm("gMed_difK").Value & ")"
    Ok shGraf.Cells(2, GColT("tolKP")).Value = 150 And shGraf.Cells(17, GColT("tolJN")).Value = -150, "tolerância ±150 nas duas diferenças"
    Ok CStr(Nm("pSt_difK").Value) = "OK" And CStr(Nm("pSt_difJ").Value) = "OK", "15/10: +114 e +127 dentro da tolerância"
    s = CStr(Nm("gComentarios").Value)
    Ok InStr(s, "Período de 15 dias") > 0 And InStr(s, "15/10") > 0 And InStr(s, "máxima") > 0, "comentários automáticos: " & Left$(s, 90)
    Ok InStr(CStr(Nm("gResumo").Value), "Batch EB6") > 0, "resumo para o e-mail"
    Ok DiaDe(Nm("gPctMaxDia").Value) = CLng(DateSerial(2026, 10, 14)), "% lama máxima em 14/10"
    Ok Perto(Nm("gUlt_lama").Value, 335, 0.001), "caixa Produção e DVR: lama 335 em 15/10"

    ' --- comentarios editaveis na aba Dia e tolerancia
    Ok InStr(CStr(Nm("dCom1").Value), "Último dia") > 0 And Nm("dCom1").HasFormula, "comentário automático na aba Dia"
    Desproteger shDia
    Nm("dCom1").Value = "Texto editado pelo laboratório"
    Nm("dComOn2").Value = "Não"
    Nm("dTop1").Value = "Tópico livre do dia"
    Proteger shDia
    Recalcular
    s = CStr(Nm("gComentarios").Value)
    Ok InStr(s, "Texto editado") > 0 And InStr(s, "Período de") = 0 And InStr(s, "Tópico livre") > 0 And _
       InStr(s, "% lama") > 0, "painel usa o comentário editado, tira o desmarcado e inclui o tópico"
    RestaurarComentarios
    s = CStr(Nm("gComentarios").Value)
    Ok InStr(s, "Último dia") > 0 And InStr(s, "Período de") > 0 And InStr(s, "Tópico livre") = 0, _
       "restaurar volta os comentários automáticos"
    Desproteger shDia
    Nm("tol_difK").Value = 100
    Proteger shDia
    Recalcular
    Ok CStr(Nm("pSt_difK").Value) = "NOK" And shGraf.Cells(2, GColT("tolKP")).Value = 100 And _
       shGraf.Cells(2, GColT("tolJP")).Value = 150, "tolerância 100 só na ESP C2 + C3: +114 vira NOK"
    Ok CStr(Nm("pSt_eb6").Value) = "—" And CStr(Nm("pSt_pctLama").Value) = "—", "item sem tolerância: — no painel"
    Desproteger shDia
    Nm("tol_eb6").Value = 10
    Nm("tol_difJ").ClearContents
    Proteger shDia
    Recalcular
    Ok CStr(Nm("pSt_eb6").Value) = "NOK" And CStr(Nm("pSt_difJ").Value) = "—" And _
       InStr(shGraf.Cells(2, GColT("tolJP")).Text, "#N") = 1, "tolerância no EB6 (|2210 - média| > 10) e sem tolerância na EB4 + ESP C3"
    Desproteger shDia
    Nm("tol_eb6").ClearContents
    Nm("tol_difK").Value = 150
    Nm("tol_difJ").Value = 150
    Proteger shDia
    Recalcular

    ' --- leitura da saida do MES (janela E) e tags
    Ok TesteLeitura() = "1000|23|2011|12", "leitura da janela E: produção média 1000 (23 h), S.E. média 2011 (12 resultados): " & TesteLeitura()
    Ok VarianteTag("M0239010000-0596-HHCC") = "M0239010000-596-HHCC", "variante de tag -0596- -> -596-"
    Ok TesteFormas("M0238010000-0017-HHLQ") = "M0238010000-0017-HHLQ|M0238010000-17-HHLQ|M238010000-0017-HHLQ|M238010000-17-HHLQ", _
       "formas testadas no diagnóstico: " & TesteFormas("M0238010000-0017-HHLQ")
    Ok TesteFormas("M4238010000-1445-DDDVR") = "M4238010000-1445-DDDVR", "tag sem outra forma: " & TesteFormas("M4238010000-1445-DDDVR")
    Ok TesteFormas("M250010000-17-ABLQ") = "M250010000-17-ABLQ|M250010000-0017-ABLQ|M0250010000-17-ABLQ|M0250010000-0017-ABLQ", _
       "formas do EB4: " & TesteFormas("M250010000-17-ABLQ")
    s = TesteTagsInvalidas("Erro: Tag Name M0239010000-0596-HHCC is invalid")
    Ok s = "|M0239010000-596-HHCC", "tag recusada: a consulta passa a usar a variante (" & s & ")"
    Ok TesteMapas("Erro:(M4239000000-0781-DDDVR) Map Name IP_AnalogMap is invalid") = "1|IP_MESVALOR|IP_AnalogMap", _
       "mapa recusado: troca só a tag recusada (" & TesteMapas("Erro:(M4239000000-0781-DDDVR) Map Name IP_AnalogMap is invalid") & ")"
    Ok ComoData("12/10/2026 15:00:00", tj) And Abs(CDbl(tj) - CDbl(DateSerial(2026, 10, 12) + TimeSerial(15, 0, 0))) < 0.0001, _
       "data do Aspen em texto"
    Ok ComoNumero("2052.0", xx) And Abs(xx - 2052) < 0.0001, "número em texto 2052.0"
    Ok Not ComoNumero("Tag Name X is invalid", xx), "texto de erro não é número"

    ' --- simulacao: dia novo, buscar, salvar
    Desproteger shConfig
    Nm("cfgFonte").Value = "SIMULAÇÃO"
    Proteger shConfig
    Data DateSerial(2026, 10, 16)
    Ok Vazio(Nm("d_eb6").Value) And Nm("d_fator").Value = 160, "dia novo 16/10: em branco, fator 160 do dia anterior"
    gUltimaMsg = ""
    BuscarMES
    Ok ENumero(Nm("d_eb6").Value) And ENumero(Nm("d_prodC2").Value) And ENumero(Nm("d_concC2").Value), _
       "Buscar no MES (simulação) preenche EB6, produção e concentrado"
    k = IIndice("tk01")
    Ok ENumero(Nm("d_tk01").Value) And ENumero(Nm("d_tk02").Value) And _
       shDia.Cells(D_ROW_IT0 + k - 1, D_COL_NRES).Value > 0, "TK 01 e TK 02 vêm do MES (tags TBLQ)"
    Ok shDia.Cells(D_ROW_IT0 + IIndice("eb6") - 1, D_COL_NRES).Value = 13, "EB6: 13 resultados no dia"
    Ok LinhaHist(DateSerial(2026, 10, 16)) > 0 And DiaDe(Nm("rFim").Value) = CLng(DateSerial(2026, 10, 16)), _
       "Buscar no MES já grava o dia fechado no histórico e leva o painel para ele"
    Ok InStr(gUltimaMsg, "salvo no histórico") > 0, "mensagem avisa que o dia foi salvo"
    Ok InStr(gUltimaMsg, "12 de 12") > 0, "mensagem da busca: " & Left$(gUltimaMsg, 120)
    Desproteger shDia
    Nm("d_tk01").Value = 2111
    Nm("d_obs").Value = "teste"
    Proteger shDia
    SalvarDia
    r = LinhaHist(DateSerial(2026, 10, 16))
    Ok r = LinhaHist(DateSerial(2026, 10, 15)) + 1, "16/10 salvo logo depois de 15/10 (linha " & r & ")"
    Ok H(r, H_COL_TK01) = 2111 And H(r, H_COL_OBS) = "teste" And H(r, H_COL_FATOR) = 160, _
       "16/10: TK 01, observação e fator gravados"
    Ok Perto(H(r, H_COL_ESPP), (H(r, H_COL_ESPC3) * H(r, H_COL_PRODC3) + H(r, H_COL_ESPC2) * H(r, H_COL_PRODC2)) / _
       (H(r, H_COL_PRODC3) + H(r, H_COL_PRODC2)), 0.001), "16/10: ponderada calculada na linha nova"
    Ok DiaDe(Nm("rFim").Value) = CLng(DateSerial(2026, 10, 16)), "Painel avança para 16/10"
    Ok Nm("gUlt_eb6").Value = H(r, H_COL_EB6), "último dia do Painel = 16/10"

    ' --- fora de ordem: 20/10 antes de 18/10
    Data DateSerial(2026, 10, 20)
    BuscarMES
    SalvarDia
    Data DateSerial(2026, 10, 18)
    BuscarMES
    SalvarDia
    r2 = LinhaHist(DateSerial(2026, 10, 18))
    r3 = LinhaHist(DateSerial(2026, 10, 20))
    Ok r2 > 0 And r3 = r2 + 1, "histórico em ordem: 18/10 antes de 20/10 (" & r2 & ", " & r3 & ")"
    Ok Perto(H(r3, H_COL_DIFK), H(r3, H_COL_EB6) - H(r3, H_COL_ESPP), 0.001), "20/10: diferença da própria linha depois de ordenar"
    ' salvar de novo o mesmo dia nao duplica
    SalvarDia
    Ok LinhaHist(DateSerial(2026, 10, 18)) = r2 And Vazio(H(r3 + 1, H_COL_DATA)), "salvar de novo não duplica o dia"

    ' --- lancamento a mao direto no Historico (dia 17/10): o painel le e a aba Dia abre o dia
    r = LinhaHist(DateSerial(2026, 10, 18)) + 2
    Ok Vazio(H(r, H_COL_DATA)), "linha livre no fim do histórico"
    shHist.Cells(r, H_COL_DATA).Value = DateSerial(2026, 10, 17)
    shHist.Cells(r, H_COL_PRODC3).Value = 25000
    shHist.Cells(r, H_COL_PRODC2).Value = 20000
    shHist.Cells(r, H_COL_ESPC3).Value = 2000
    shHist.Cells(r, H_COL_ESPC2).Value = 1900
    shHist.Cells(r, H_COL_EB6).Value = 2100
    Recalcular
    Ok Perto(H(r, H_COL_DIFK), 2100 - (2000 * 25000 + 1900 * 20000) / 45000, 0.001), "dia lançado à mão: diferença calculada"
    Data DateSerial(2026, 10, 17)
    Ok Nm("d_eb6").Value = 2100 And Nm("d_espC3").Value = 2000, "aba Dia abre o dia lançado à mão no histórico"
    r2 = LinhaHist(DateSerial(2026, 10, 18))
    Data DateSerial(2026, 10, 18)
    Ok Nm("d_tk01").Value = H(r2, H_COL_TK01) And Nm("d_eb6").Value = H(r2, H_COL_EB6), _
       "trocar a data abre o dia do histórico"
    Ok Perto(Nm("_difK").Value, Nm("d_eb6").Value - Nm("_espP").Value, 0.001), "resultado do dia na aba Dia"

    ' --- dia em andamento (hoje): aviso e pergunta antes de salvar
    Data CDate(Int(CDbl(Now)))
    gUltimaMsg = ""
    BuscarMES
    Ok InStr(gUltimaMsg, "ainda não terminou") > 0, "buscar hoje avisa que o dia não terminou"
    Ok LinhaHist(CDate(Int(CDbl(Now)))) = 0, "buscar hoje não grava no histórico"

    ' --- diagnostico sem o Aspen (LibreOffice): avisa e nao trava
    Desproteger shConfig
    Nm("cfgFonte").Value = "MES"
    Proteger shConfig
    Desproteger shDiag
    Nm("xData").Value = DateSerial(2026, 10, 15)
    Proteger shDiag
    gUltimaMsg = ""
    DiagnosticoMES
    Ok InStr(CStr(shDiag.Cells(X_ROW0, X_COL_RES).Value), "suplemento") > 0 And InStr(gUltimaMsg, "#NOME?") > 0, _
       "diagnóstico sem o Aspen: avisa o suplemento (" & shDiag.Cells(X_ROW0, X_COL_RES).Value & ")"
    Ok InStr(CStr(shDiag.Cells(X_ROW0 + IIndice("tk01") - 1, X_COL_RES).Value), "sem tag") > 0 Or _
       InStr(CStr(Nm("xStatus").Value), "15/10/2026") > 0, "diagnóstico grava a última execução"

    GoTo Fim
Erro:
    mLog = mLog & "ERRO " & Err.Number & " " & Err.Description & "|"
    mFalhas = mFalhas + 1
Fim:
    Nm("gTeste").Value = "'" & IIf(mFalhas = 0, "TUDO OK", mFalhas & " FALHA(S)") & "|" & mLog
End Sub

Option Explicit

' ============================================================================
'  ModMES - resultados do dia no MES (Aspen IP.21)
'  Duas consultas GetCalculationValues em janelas de 1 h, uma por janela de horario:
'    E (espessador): D-1 23:00 a D 21:00 - produção C2 e S.E. dos espessadores (2 h antes do batch)
'    D (dia)       : D 00:00 a D 23:59 - batch EB4/EB6, tanques e DVR (valor diario gravado as 23:59)
'  A janela de 1 h comeca 30 min antes do horario: o resultado das 01:00 cai em 00:30-01:30 e o de
'  fechamento do batch (23:59) em 23:30-00:30, cada um na sua linha.
'  Tags, servidores e mapas vao em texto literal (mesmo formato dos relatorios do laboratorio).
'  Tag recusada pelo MES ("Tag Name ... is invalid") sai da consulta, que e refeita com as demais;
'  antes, o mesmo codigo e tentado no outro formato (-0596- <-> -596-). Mapa recusado: tenta os outros.
'  Diagnostico do MES: cada tag sozinha, de D-1 12:00 a D+1 12:00, para ver onde ha valor.
' ============================================================================

' Segundos sem mudanca nos resultados para considerar a consulta concluida
Private Const ESTAVEL_SEG As Double = 2
' Consulta que respondeu "Success" sem nenhuma linha: "sem dados" depois deste tempo
Private Const SEM_DADOS_SEG As Double = 15
Private Const MAX_RODADAS As Long = 10
' Celulas da aba Dados_MES
Private Const ANC_E As String = "A15"
Private Const OUT_E As String = "A17"
Private Const ANC_D As String = "A75"
Private Const OUT_D As String = "A77"
' Linhas lidas da saida de uma consulta (ate 48 janelas de 1 h)
Private Const LIN_SAIDA As Long = 50
Private Const MAPAS_ALT As String = "IP_MESVALOR|IP_AnalogMap|IP_DiscreteMap"
' Mapas testados pelo diagnostico, nesta ordem
Private Const MAPAS_DIAG As String = "IP_MESVALOR|IP_AnalogMap"
' Diagnostico: janelas de 1 h de D-1 12:00 a D+1 12:00
Private Const DIAG_JANELAS As Long = 48

' Funcao do Aspen nao reconhecida (#NOME?): nao repete a consulta nem o aviso
Private mSemAspen As Boolean
' Tags recusadas pelo MES nesta sessao, no formato |TAG1|TAG2|
Private mInvalidas As String
' Mapas recusados pelo MES nesta sessao ("Map Name X is invalid"), no formato |TAG>MAPA|
Private mMapasRuins As String
' Contador da consulta na barra de status: o que esta sendo consultado
Private mEtapa As String
' Ajustes gravados em Configuracoes nesta busca (mapas e tags que o MES aceitou), para a mensagem
Private mAjMapas As String
Private mAjTags As String
' Horario (centro da janela de 1 h) do ultimo valor achado de cada item na ultima busca
Private mHora(1 To 40) As Double
' Avisos da ultima busca (ajustes, tags invalidas, MES sem resposta) - o chamador mostra
Public gNotasMES As String
' Diagnostico: tag/mapa testados no lugar dos de Configuracoes (vazio = usa Configuracoes)
Private mDescobrindo As Boolean
Private mTagOvr(1 To 40) As String
Private mMapOvr(1 To 40) As String

' ---------------------------------------------------------------- busca de um dia
' vals(1..NI): valor agregado (Empty = sem resultado); nres(1..NI): resultados achados; sit(1..NI): situacao
' Devolve False se o MES nao respondeu (ou sem suplemento).
Public Function BuscarDiaMES(ByVal d As Date, ByRef vals As Variant, ByRef nres As Variant, _
                             ByRef sit As Variant) As Boolean
    Dim v() As Variant, n() As Variant, s() As Variant, k As Long, ok As Boolean, andamento As Boolean
    ReDim v(1 To NI)
    ReDim n(1 To NI)
    ReDim s(1 To NI)
    For k = 1 To NI
        n(k) = 0
        mHora(k) = 0
        If ICfg(k, C_COL_TAG) = "" Then s(k) = "sem tag: lançar à mão" Else s(k) = ""
    Next k
    gNotasMES = ""
    mAjMapas = ""
    mAjTags = ""
    d = DateSerial(Year(d), Month(d), Day(d))
    andamento = (CDbl(d) >= Int(CDbl(Now)))
    If UCase$(Left$(CfgTxt("cfgFonte"), 3)) = "SIM" Then
        Simular d, v, n, s
        ok = True
    Else
        mSemAspen = False
        ok = ConsultarJanela("E", d, v, n, s)
        If Not mSemAspen Then ok = ConsultarJanela("D", d, v, n, s) And ok
        Nm("mesUltima").Value = "'" & Format$(Now, "dd\/mm\/yyyy hh:mm:ss") & "  dia " & TxtDia(d)
    End If
    For k = 1 To NI
        If s(k) = "" Then
            If Not IsEmpty(v(k)) Then
                s(k) = "MES · " & n(k) & IIf(n(k) = 1, " valor", " valores")
                If mHora(k) > 0 Then s(k) = s(k) & " · último " & Format$(CDate(mHora(k)), "dd\/mm hh:mm")
            ElseIf andamento And UCase$(ICfg(k, C_COL_JAN)) = "D" Then
                s(k) = "sem resultado: dia ainda em andamento"
            Else
                s(k) = "sem resultado no MES (ver Diagnóstico)"
            End If
        End If
    Next k
    If andamento Then
        gNotasMES = "Atenção: " & TxtDia(d) & " ainda não terminou. O DVR (produção C3, lama e concentrados) e o " & _
                    "fechamento do batch só entram no MES depois das 23:59 - para o report, busque o dia anterior." & _
                    vbCrLf & gNotasMES
    End If
    If mAjMapas <> "" Then gNotasMES = gNotasMES & "Mapa trocado em Configurações (o MES recusou o anterior): " & _
                                       Mid$(mAjMapas, 3) & "." & vbCrLf
    If mAjTags <> "" Then gNotasMES = gNotasMES & "Tag ajustada em Configurações (o MES aceitou em outro formato): " & _
                                      Mid$(mAjTags, 3) & "." & vbCrLf
    vals = v
    nres = n
    sit = s
    Application.StatusBar = False
    BuscarDiaMES = ok
End Function

' Horario inicial da 1a janela de 1 h e numero de janelas de uma janela de horario
Public Sub PeriodoJanela(ByVal j As String, ByVal d As Date, ByRef qIni As Date, ByRef nJan As Long)
    Dim h1 As Double
    h1 = CfgNum("cfgJan" & j & "H1", IIf(j = "E", -1, 1))
    nJan = CLng(CfgNum("cfgJan" & j & "N", IIf(j = "E", 23, 24)))
    qIni = CDate(CDbl(d) + (h1 - 0.5) / 24)
End Sub

' Itens com tag da janela j (lista "1,3,5")
Private Function ItensJanela(ByVal j As String) As String
    Dim k As Long, t As String, r As String
    For k = 1 To NI
        If UCase$(ICfg(k, C_COL_JAN)) = j Then
            t = TagAtual(k)
            If t <> "" And Not TagInvalida(t) Then
                If r <> "" Then r = r & ","
                r = r & k
            End If
        End If
    Next k
    ItensJanela = r
End Function

' So os itens de ps que continuam com tag valida (depois de uma tag recusada)
Private Function FiltrarValidos(ByVal ps As String) As String
    Dim lst As Variant, i As Long, k As Long, r As String
    If ps = "" Then Exit Function
    lst = Split(ps, ",")
    For i = LBound(lst) To UBound(lst)
        k = CLng(lst(i))
        If TagAtual(k) <> "" And Not TagInvalida(TagAtual(k)) And MapaAtual(k) <> "" Then
            If r <> "" Then r = r & ","
            r = r & k
        End If
    Next i
    FiltrarValidos = r
End Function

Private Function ConsultarJanela(ByVal j As String, ByVal d As Date, v() As Variant, n() As Variant, _
                                 s() As Variant) As Boolean
    Dim ps As String, nJan As Long, nTot As Long, qIni As Date, estado As Long, resp As String, bruto As Variant
    Dim i As Long
    ps = ItensJanela(j)
    If ps = "" Then
        ConsultarJanela = True           ' nada a consultar nesta janela
        Exit Function
    End If
    PeriodoJanela j, d, qIni, nTot
    ' nunca pede ao MES janelas que ainda nao comecaram
    nJan = nTot
    If CDbl(qIni) + nJan / 24 > CDbl(Now) Then nJan = Int((CDbl(Now) - CDbl(qIni)) * 24) + 1
    If nJan > nTot Then nJan = nTot
    If nJan < 1 Then
        ConsultarJanela = True           ' dia ainda nao comecou
        Exit Function
    End If
    mEtapa = IIf(j = "E", "espessador/produção", "batch e DVR") & " " & TxtDia(d)
    estado = Consultar(ps, qIni, nJan, j, bruto, resp)
    If estado = 1 Then
        LerSaida bruto, qIni, nJan, ps, v, n
    ElseIf estado = 0 Or estado = 4 Or estado = 5 Then
        gNotasMES = gNotasMES & IIf(estado = 0, "O MES não respondeu", "O MES devolveu erro") & " (" & mEtapa & _
                    "): " & Left$(resp, 160) & vbCrLf
    End If
    ' itens que ficaram de fora por tag recusada
    For i = 1 To NI
        If UCase$(ICfg(i, C_COL_JAN)) = j And ICfg(i, C_COL_TAG) <> "" Then
            If TagInvalida(TagAtual(i)) Then s(i) = "tag recusada pelo MES (ver Diagnóstico)"
        End If
    Next i
    ConsultarJanela = (estado = 1 Or estado = 2 Or estado = 9)
End Function

' Uma consulta GetCalculationValues (itens de ps, janelas de 1 h a partir de qIni) com as tentativas de tag e de
' mapa. ps volta so com os itens consultados (na ordem das colunas de bruto). bruto = matriz de saida lida.
' Devolve 1 = resultados; 2 = "Sucesso" sem linhas; 0 = sem resposta; 4 = mapa recusado; 5 = erro do Aspen;
' 3 = todas as tags recusadas; 9 = nada a consultar; -1 = sem o suplemento do Aspen.
Private Function Consultar(ByRef ps As String, ByVal qIni As Date, ByVal nJan As Long, ByVal j As String, _
                           ByRef bruto As Variant, ByRef resp As String) As Long
    Dim anc As Range, sai As Range, rodada As Long, limite As Double, fim As Date, estado As Long, e2 As Long
    Dim restaurado As Boolean, invalidas As String
    If mSemAspen Then
        Consultar = -1
        Exit Function
    End If
    If j = "E" Then
        Set anc = shDadosMES.Range(ANC_E)
        Set sai = shDadosMES.Range(OUT_E)
    Else
        Set anc = shDadosMES.Range(ANC_D)
        Set sai = shDadosMES.Range(OUT_D)
    End If
    fim = CDate(CDbl(qIni) + nJan / 24)
    limite = CfgNum("cfgTimeout", 120)
    GravarPeriodo j, qIni, fim
    estado = 9
    For rodada = 0 To MAX_RODADAS
        ps = FiltrarValidos(ps)
        If ps = "" Then
            If estado <> 3 Then estado = 9
            Exit For
        End If
        LimparSaida sai
        If Not EscreverConsulta(anc, ArgumentosJanela(j, ps, anc, sai)) Then
            Nm("mesResp" & j).Value = "'" & Format$(Now, "dd\/mm hh:mm:ss") & "  #NOME? (suplemento Aspen)"
            anc.Value = "'Suplemento Aspen não reconhecido (#NOME?)"
            LimparSaida sai
            mSemAspen = True
            Aviso "O Excel não reconheceu a função do MES (GetCalculationValues) - aparece #NOME?." & vbCrLf & _
                  vbCrLf & "Verificar se o suplemento Aspen Process Data está ativo neste Excel: Arquivo > " & _
                  "Opções > Suplementos > Gerenciar: Suplementos COM > Ir... (o mesmo usado nos relatórios do " & _
                  "laboratório). Depois feche e abra a planilha.", vbExclamation, "MES"
            Consultar = -1
            Exit Function
        End If
        estado = Esperar(anc, sai, qIni, limite, NumCols(ps))
        If estado = 3 Then
            If ColetarInvalidas(anc, sai) = 0 Then Exit For
        ElseIf estado = 4 Then
            If ColetarMapas(anc, sai) = 0 Then Exit For
        ElseIf estado = 5 Then
            Exit For
        ElseIf (estado = 0 Or estado = 2) And Not restaurado Then
            ' respondeu sem resultados na celula de saida: monta a matriz ShowCalculationValues
            ' (forma da planilha de referencia do MES) e espera mais uma vez
            restaurado = True
            RestaurarSaida anc, sai, nJan, NumCols(ps)
            anc.Calculate
            sai.Resize(nJan, NumCols(ps)).Calculate
            e2 = Esperar(anc, sai, qIni, limite, NumCols(ps))
            If e2 = 1 Or estado = 0 Then estado = e2
            Exit For
        Else
            Exit For
        End If
    Next rodada
    If estado = 1 And Not mDescobrindo Then
        CorrigirTags
        CorrigirMapas
    End If
    invalidas = ListaInvalidas()
    If invalidas <> "" Then Nm("mesInvalidas").Value = "'" & Replace(invalidas, vbCrLf, " ")
    resp = Trim$(anc.Text)
    If resp = "" Or Left$(resp, 1) = "=" Then resp = "(sem texto)"
    Nm("mesResp" & j).Value = "'" & Format$(Now, "dd\/mm hh:mm:ss") & "  " & Left$(resp & "  |  1a célula: " & _
                              Trim$(sai.Cells(1, 1).Text) & "  " & Trim$(sai.Cells(1, 2).Text), 250)
    If estado = 1 And ps <> "" Then bruto = sai.Resize(LIN_SAIDA, NumCols(ps)).Value
    DesligarConsulta anc, sai
    Consultar = estado
End Function

' Le a matriz de saida (bruto): 1a coluna = inicio da janela de 1 h; demais = itens na ordem de ps.
' Agrega por item (Soma / Media / Ultimo) os valores validos das janelas 1..nJan.
Public Sub LerSaida(ByVal bruto As Variant, ByVal qIni As Date, ByVal nJan As Long, ByVal ps As String, _
                    v() As Variant, n() As Variant)
    Dim lst As Variant, i As Long, k As Long, r As Long, sj As Long, tj As Date, x As Double, cel As Variant
    Dim soma() As Double, cont() As Long, ult() As Double, sUlt() As Long, agr As String
    If ps = "" Then Exit Sub
    lst = Split(ps, ",")
    ReDim soma(LBound(lst) To UBound(lst))
    ReDim cont(LBound(lst) To UBound(lst))
    ReDim ult(LBound(lst) To UBound(lst))
    ReDim sUlt(LBound(lst) To UBound(lst))
    For r = 1 To UBound(bruto, 1)
        If Not ComoData(bruto(r, 1), tj, qIni) Then Exit For
        sj = CLng((CDbl(tj) - CDbl(qIni)) * 24) + 1
        If sj >= 1 And sj <= nJan Then
            For i = LBound(lst) To UBound(lst)
                If 2 + i - LBound(lst) <= UBound(bruto, 2) Then
                    k = CLng(lst(i))
                    cel = bruto(r, 2 + i - LBound(lst))
                    If ComoNumero(cel, x) Then
                        If ValorValido(k, x) Then
                            soma(i) = soma(i) + x
                            cont(i) = cont(i) + 1
                            If sj >= sUlt(i) Then
                                sUlt(i) = sj
                                ult(i) = x
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    Next r
    For i = LBound(lst) To UBound(lst)
        k = CLng(lst(i))
        n(k) = cont(i)
        If cont(i) > 0 Then
            mHora(k) = CDbl(qIni) + (sUlt(i) - 0.5) / 24
            agr = LCase$(Left$(ICfg(k, C_COL_AGR), 1))
            If agr = "s" Then
                v(k) = soma(i)
            ElseIf agr = "m" Then
                v(k) = soma(i) / cont(i)
            Else
                v(k) = ult(i)
            End If
        Else
            v(k) = Empty
        End If
    Next i
End Sub

Private Function NumCols(ByVal ps As String) As Long
    NumCols = UBound(Split(ps, ",")) + 2
End Function

' ---------------------------------------------------------------- diagnostico do MES (descobre tag e mapa)
' O MES pode aceitar uma tag (resposta "Sucesso") e nao trazer nenhum valor quando o formato do codigo ou o mapa
' nao sao os do historico. O diagnostico consulta de D-1 12:00 a D+1 12:00 e, para cada item ainda sem valor,
' tenta em rodadas as formas da tag (-0596- / -596-, M0... / M...) com IP_MESVALOR e IP_AnalogMap. Todas as tags
' pendentes vao juntas em cada rodada (uma consulta por rodada). A combinacao que traz valor fica gravada em
' Configuracoes; o resultado vai para a aba Diagnostico.
Public Sub DiagnosticoMES()
    Dim d As Date, k As Long, r As Long, ps As String, qIni As Date, nJan As Long, estado As Long, resp As String
    Dim bruto As Variant, mapas As Variant, f As Long, im As Long, i As Long, pend As Long, t As String
    Dim formas As Variant, achou() As Boolean, tent() As String, nOk As Long, nTag As Long, t0 As Single
    Dim rodada As Long, lst As Variant, col As Long, ajustes As String
    If Not EData(Nm("xData").Value) Then
        Aviso "Informe o dia do teste (célula amarela).", vbExclamation, "Diagnóstico do MES"
        Exit Sub
    End If
    d = CDate(DiaDe(Nm("xData").Value))
    If UCase$(Left$(CfgTxt("cfgFonte"), 3)) = "SIM" Then
        Aviso "Fonte = SIMULAÇÃO em Configurações: o diagnóstico precisa do MES.", vbExclamation, "Diagnóstico do MES"
        Exit Sub
    End If
    t0 = Timer
    mSemAspen = False
    mInvalidas = ""
    mMapasRuins = ""
    Ampulheta True
    Desproteger shDiag
    shDiag.Range(shDiag.Cells(X_ROW0, 3), shDiag.Cells(X_ROW0 + NI - 1, 12)).ClearContents
    Nm("xStatus").Value = "'Rodando..."
    qIni = CDate(CDbl(d) - 0.5 - 0.5 / 24)
    nJan = DIAG_JANELAS
    If CDbl(qIni) + nJan / 24 > CDbl(Now) Then nJan = Int((CDbl(Now) - CDbl(qIni)) * 24) + 1
    ReDim achou(1 To NI)
    ReDim tent(1 To NI)
    mapas = Split(MAPAS_DIAG, "|")
    For k = 1 To NI
        If ICfg(k, C_COL_TAG) = "" Then
            shDiag.Cells(X_ROW0 + k - 1, 6).Value = "sem tag (lançamento à mão)"
            achou(k) = True
        Else
            nTag = nTag + 1
        End If
    Next k
    mDescobrindo = True
    ' rodadas: forma 1 (a de Configuracoes) com cada mapa, depois as outras formas
    For f = 0 To 3
        For im = LBound(mapas) To UBound(mapas)
            ps = ""
            For k = 1 To NI
                mTagOvr(k) = ""
                mMapOvr(k) = ""
                If Not achou(k) Then
                    formas = FormasTag(ICfg(k, C_COL_TAG))
                    If f <= UBound(formas) Then
                        mTagOvr(k) = CStr(formas(f))
                        mMapOvr(k) = CStr(mapas(im))
                        If Not TagInvalida(CStr(formas(f))) Then
                            If ps <> "" Then ps = ps & ","
                            ps = ps & k
                            tent(k) = tent(k) & IIf(tent(k) = "", "", "; ") & formas(f) & " / " & mapas(im)
                        End If
                    End If
                End If
            Next k
            If ps <> "" Then
                rodada = rodada + 1
                mEtapa = "diagnóstico, rodada " & rodada & " (" & UBound(Split(ps, ",")) + 1 & " tags)"
                estado = Consultar(ps, qIni, nJan, "D", bruto, resp)
                If estado = -1 Then GoTo SemAspen
                If estado = 1 And ps <> "" Then
                    lst = Split(ps, ",")
                    For i = LBound(lst) To UBound(lst)
                        k = CLng(lst(i))
                        col = 2 + i - LBound(lst)
                        If ResultadoDiag(k, d, bruto, col, qIni, ps, resp) > 0 Then
                            achou(k) = True
                            nOk = nOk + 1
                            ajustes = ajustes & GravarAchado(k, mTagOvr(k), mMapOvr(k))
                        End If
                    Next i
                End If
            End If
        Next im
    Next f
    ' o que nao achou em nenhuma forma
    For k = 1 To NI
        If Not achou(k) Then
            r = X_ROW0 + k - 1
            shDiag.Cells(r, 3).Value = "'" & ICfg(k, C_COL_TAG)
            shDiag.Cells(r, 4).Value = ICfg(k, C_COL_MAPA)
            shDiag.Cells(r, 5).Value = JanelaTxt(k)
            shDiag.Cells(r, 6).Value = "sem valor em nenhuma forma testada"
            shDiag.Cells(r, 7).Value = 0
            shDiag.Cells(r, 11).Value = "'Testadas: " & Left$(tent(k), 240)
        End If
    Next k
    GoTo Fim
SemAspen:
    shDiag.Cells(X_ROW0, 6).Value = "sem o suplemento do Aspen (#NOME?)"
Fim:
    mDescobrindo = False
    For k = 1 To NI
        mTagOvr(k) = ""
        mMapOvr(k) = ""
    Next k
    Nm("xStatus").Value = "'" & Format$(Now, "dd\/mm\/yyyy hh:mm") & "  dia " & TxtDia(d) & ": " & nOk & " de " & _
                          nTag & " tags com valor  (" & rodada & " consultas, " & Format$(Timer - t0, "0") & " s)"
    Proteger shDiag
    Application.StatusBar = False
    Ampulheta False
    shDiag.Activate
    Aviso "Diagnóstico de " & TxtDia(d) & ": " & nOk & " de " & nTag & " tags com valor no MES." & _
          IIf(ajustes <> "", vbCrLf & vbCrLf & "Gravado em Configurações (forma que trouxe valor):" & vbCrLf & _
              Left$(ajustes, 700), "") & vbCrLf & vbCrLf & "Detalhes (horários e valores) na aba Diagnóstico." & _
          IIf(nOk = nTag, vbCrLf & "Agora é só voltar ao Dia e clicar em Buscar no MES.", ""), _
          IIf(nOk = nTag, vbInformation, vbExclamation), "Diagnóstico do MES"
End Sub

' Formas da tag testadas pelo diagnostico: como esta, -0596- <-> -596-, M0... <-> M... e as duas trocas
Public Function FormasTag(ByVal t As String) As Variant
    Dim c(0 To 3) As String, r As String, i As Long
    c(0) = t
    c(1) = VarianteTag(t)
    c(2) = VarianteM(t)
    If c(1) <> "" Then c(3) = VarianteM(c(1))
    For i = 0 To 3
        If c(i) <> "" Then
            If InStr(1, "|" & r & "|", "|" & c(i) & "|", vbTextCompare) = 0 Then r = r & IIf(r = "", "", "|") & c(i)
        End If
    Next i
    FormasTag = Split(r, "|")
End Function

' M0238010000-... <-> M238010000-... (o zero depois do M aparece nas duas formas no MES)
Public Function VarianteM(ByVal t As String) As String
    Dim pt As Variant, s As String
    pt = Split(t, "-")
    If UBound(pt) < 2 Then Exit Function
    s = pt(0)
    If Len(s) < 3 Or UCase$(Left$(s, 1)) <> "M" Then Exit Function
    If Not IsNumeric(Mid$(s, 2)) Then Exit Function
    If Len(s) = 11 And Mid$(s, 2, 1) = "0" Then
        pt(0) = "M" & Mid$(s, 3)
    ElseIf Len(s) = 10 Then
        pt(0) = "M0" & Mid$(s, 2)
    Else
        Exit Function
    End If
    VarianteM = Join(pt, "-")
End Function

Public Function TesteFormas(ByVal t As String) As String
    TesteFormas = Join(FormasTag(t), "|")
End Function

Private Function JanelaTxt(ByVal k As Long) As String
    JanelaTxt = IIf(UCase$(ICfg(k, C_COL_JAN)) = "E", "E: D-1 23:00 a D 21:00", "D: D 00:00 a 23:59") & " · " & _
                ICfg(k, C_COL_AGR)
End Function

' Preenche a linha do item k na aba Diagnostico com a coluna col da saida; devolve quantos valores achou
Private Function ResultadoDiag(ByVal k As Long, ByVal d As Date, ByVal bruto As Variant, ByVal col As Long, _
                               ByVal qIni As Date, ByVal ps As String, ByVal resp As String) As Long
    Dim i As Long, tj As Date, x As Double, nV As Long, t1 As Double, t2 As Double, lista As String, r As Long
    Dim v() As Variant, n() As Variant, qN As Date, nN As Long, res As String
    For i = 1 To UBound(bruto, 1)
        If Not ComoData(bruto(i, 1), tj, qIni) Then Exit For
        If col <= UBound(bruto, 2) Then
            If ComoNumero(bruto(i, col), x) Then
                nV = nV + 1
                If t1 = 0 Then t1 = CDbl(tj) + 0.5 / 24
                t2 = CDbl(tj) + 0.5 / 24
                If nV <= 14 Then lista = lista & IIf(lista = "", "", "  ·  ") & _
                    Format$(CDate(CDbl(tj) + 0.5 / 24), "dd\/mm hh:mm") & " = " & Format$(x, "0.###")
            End If
        End If
    Next i
    ResultadoDiag = nV
    If nV = 0 Then Exit Function
    If nV > 14 Then lista = lista & "  ·  ... (" & nV & " no total)"
    ' o que a janela normal do item daria (mesma agregacao da busca)
    ReDim v(1 To NI)
    ReDim n(1 To NI)
    PeriodoJanela UCase$(ICfg(k, C_COL_JAN)), d, qN, nN
    LerSaida bruto, qN, nN, ps, v, n
    r = X_ROW0 + k - 1
    res = "OK · " & nV & " valor(es)"
    If IsEmpty(v(k)) Then res = res & " · nenhum na janela do item"
    shDiag.Cells(r, 3).Value = "'" & mTagOvr(k)
    shDiag.Cells(r, 4).Value = mMapOvr(k)
    shDiag.Cells(r, 5).Value = JanelaTxt(k)
    shDiag.Cells(r, 6).Value = res
    shDiag.Cells(r, 7).Value = nV
    shDiag.Cells(r, 8).Value = "'" & Format$(CDate(t1), "dd\/mm hh:mm")
    shDiag.Cells(r, 9).Value = "'" & Format$(CDate(t2), "dd\/mm hh:mm")
    If Not IsEmpty(v(k)) Then shDiag.Cells(r, 10).Value = Round(CDbl(v(k)), 3)
    shDiag.Cells(r, 11).Value = "'" & lista
    shDiag.Cells(r, 12).Value = "'" & Left$(resp, 250)
End Function

' Grava em Configuracoes a tag/mapa que trouxe valor (se diferente); devolve a linha da mensagem
Private Function GravarAchado(ByVal k As Long, ByVal t As String, ByVal m As String) As String
    If UCase$(t) = UCase$(ICfg(k, C_COL_TAG)) And UCase$(m) = UCase$(ICfg(k, C_COL_MAPA)) Then Exit Function
    GravarAchado = "  - " & IRotulo(k) & ": " & t & " / " & m & vbCrLf
    On Error Resume Next
    Desproteger shConfig
    shConfig.Cells(C_ROW_IT0 + k - 1, C_COL_TAG).Value = "'" & t
    shConfig.Cells(C_ROW_IT0 + k - 1, C_COL_MAPA).Value = m
    Proteger shConfig
End Function

' ---------------------------------------------------------------- tags
Private Function TagInvalida(ByVal t As String) As Boolean
    If t = "" Then Exit Function
    TagInvalida = (InStr(1, mInvalidas, "|" & UCase$(t) & "|") > 0)
End Function

' Mesma tag com o codigo do meio em outro formato: -0596- <-> -596- (o MES aceita so um deles)
Public Function VarianteTag(ByVal t As String) As String
    Dim pt As Variant
    pt = Split(t, "-")
    If UBound(pt) < 2 Then Exit Function
    If Not IsNumeric(pt(1)) Or Len(pt(1)) = 0 Then Exit Function
    If Len(pt(1)) >= 4 And Left$(pt(1), 1) = "0" Then
        pt(1) = CStr(CLng(pt(1)))
    ElseIf Len(pt(1)) < 4 Then
        pt(1) = Format$(CLng(pt(1)), "0000")
    Else
        Exit Function
    End If
    VarianteTag = Join(pt, "-")
End Function

' Tag usada na consulta: a da aba Configuracoes ou, se o MES a recusou, a variante do codigo
Private Function TagAtual(ByVal k As Long) As String
    Dim t As String, alt As String
    If mTagOvr(k) <> "" Then
        TagAtual = mTagOvr(k)
        Exit Function
    End If
    t = ICfg(k, C_COL_TAG)
    TagAtual = t
    If t = "" Then Exit Function
    If TagInvalida(t) Then
        alt = VarianteTag(t)
        If alt <> "" Then
            If Not TagInvalida(alt) Then TagAtual = alt
        End If
    End If
End Function

Private Function ValorValido(ByVal k As Long, ByVal x As Double) As Boolean
    Dim vmin As String, vmax As String
    vmin = ICfg(k, C_COL_VMIN)
    vmax = ICfg(k, C_COL_VMAX)
    If IsNumeric(vmin) And vmin <> "" Then
        If x < CDbl(vmin) Then Exit Function
    End If
    If IsNumeric(vmax) And vmax <> "" Then
        If x > CDbl(vmax) Then Exit Function
    End If
    ValorValido = True
End Function

Private Function ColetarInvalidas(ByVal anc As Range, ByVal sai As Range) As Long
    Dim t As String, i As Long, j As Long, tag As String
    t = TextoResposta(anc, sai)
    i = InStr(1, t, "Tag Name ", vbTextCompare)
    Do While i > 0
        j = InStr(i + 9, t, " is invalid", vbTextCompare)
        If j = 0 Then Exit Do
        tag = UCase$(Trim$(Mid$(t, i + 9, j - i - 9)))
        If tag <> "" And InStr(tag, " ") = 0 And Not TagInvalida(tag) Then
            If mInvalidas = "" Then mInvalidas = "|"
            mInvalidas = mInvalidas & tag & "|"
            ColetarInvalidas = ColetarInvalidas + 1
        End If
        i = InStr(j, t, "Tag Name ", vbTextCompare)
    Loop
End Function

' Grava na aba Configuracoes as variantes de tag que funcionaram; devolve a lista
Private Function CorrigirTags() As String
    Dim k As Long, t As String, r As String
    For k = 1 To NI
        t = TagAtual(k)
        If t <> "" And t <> ICfg(k, C_COL_TAG) And Not TagInvalida(t) Then
            r = r & "  - " & IRotulo(k) & ": " & ICfg(k, C_COL_TAG) & "  ->  " & t & vbCrLf
            mAjTags = mAjTags & "; " & IRotulo(k) & " -> " & t
            On Error Resume Next
            Desproteger shConfig
            shConfig.Cells(C_ROW_IT0 + k - 1, C_COL_TAG).Value = t
            Proteger shConfig
            On Error GoTo 0
        End If
    Next k
    CorrigirTags = r
End Function

Public Function ListaInvalidas() As String
    Dim k As Long, t As String, r As String
    For k = 1 To NI
        t = TagAtual(k)
        If TagInvalida(t) Then
            r = r & "  - " & IRotulo(k) & ": " & ICfg(k, C_COL_TAG)
            If t <> ICfg(k, C_COL_TAG) Then r = r & "  (e " & t & ")"
            r = r & vbCrLf
        End If
    Next k
    ListaInvalidas = r
End Function

' Item com tag recusada pelo MES nesta sessao (para a coluna Situacao)
Public Function ItemRecusado(ByVal k As Long) As Boolean
    ItemRecusado = TagInvalida(TagAtual(k))
End Function

' ---------------------------------------------------------------- mapas
Private Function MapaRecusado(ByVal tag As String, ByVal mapa As String) As Boolean
    MapaRecusado = (InStr(1, mMapasRuins, "|" & UCase$(tag) & ">" & UCase$(mapa) & "|") > 0)
End Function

' Mapa usado na consulta: o da aba Configuracoes ou, se o MES o recusou, o 1o da lista MAPAS_ALT ainda nao
' recusado ("" = nenhum)
Private Function MapaAtual(ByVal k As Long) As String
    Dim m As String, t As String, alt As Variant, i As Long
    If mMapOvr(k) <> "" Then
        If Not MapaRecusado(TagAtual(k), mMapOvr(k)) Then MapaAtual = mMapOvr(k)
        Exit Function
    End If
    m = ICfg(k, C_COL_MAPA)
    If m = "" Then m = "IP_MESVALOR"
    t = TagAtual(k)
    If Not MapaRecusado(t, m) Then
        MapaAtual = m
        Exit Function
    End If
    alt = Split(MAPAS_ALT, "|")
    For i = LBound(alt) To UBound(alt)
        If Not MapaRecusado(t, CStr(alt(i))) Then
            MapaAtual = CStr(alt(i))
            Exit Function
        End If
    Next i
End Function

Private Function CorrigirMapas() As String
    Dim k As Long, m As String, r As String
    For k = 1 To NI
        If TagAtual(k) <> "" And Not TagInvalida(TagAtual(k)) Then
            m = MapaAtual(k)
            If m <> "" And UCase$(m) <> UCase$(ICfg(k, C_COL_MAPA)) And ICfg(k, C_COL_MAPA) <> "" Then
                r = r & "  - " & IRotulo(k) & ": mapa " & ICfg(k, C_COL_MAPA) & "  ->  " & m & vbCrLf
                mAjMapas = mAjMapas & "; " & IRotulo(k) & " -> " & m
                On Error Resume Next
                Desproteger shConfig
                shConfig.Cells(C_ROW_IT0 + k - 1, C_COL_MAPA).Value = m
                Proteger shConfig
                On Error GoTo 0
            End If
        End If
    Next k
    CorrigirMapas = r
End Function

' Marca os pares tag/mapa recusados (gravar = True); devolve quantos sao novos
Private Function MapasRecusados(ByVal t As String, ByVal gravar As Boolean) As Long
    Dim i As Long, j As Long, a As Long, tag As String, mapa As String, chave As String
    i = InStr(1, t, "Map Name ", vbTextCompare)
    Do While i > 0
        j = InStr(i + 9, t, " is invalid", vbTextCompare)
        If j = 0 Then Exit Do
        mapa = UCase$(Trim$(Mid$(t, i + 9, j - i - 9)))
        a = InStrRev(t, "(", i)
        tag = ""
        If a > 0 Then
            If InStr(a, t, ")") < i Then tag = UCase$(Trim$(Mid$(t, a + 1, InStr(a, t, ")") - a - 1)))
        End If
        If tag <> "" And mapa <> "" And InStr(tag, " ") = 0 Then
            chave = "|" & tag & ">" & mapa & "|"
            If InStr(1, mMapasRuins, chave) = 0 Then
                MapasRecusados = MapasRecusados + 1
                If gravar Then
                    If mMapasRuins = "" Then mMapasRuins = "|"
                    mMapasRuins = mMapasRuins & tag & ">" & mapa & "|"
                End If
            End If
        End If
        i = InStr(j, t, "Map Name ", vbTextCompare)
    Loop
End Function

' Depois de "Map Name ... is invalid": guarda os recusados; a tag sem nenhum mapa aceito sai da consulta
Private Function ColetarMapas(ByVal anc As Range, ByVal sai As Range) As Long
    Dim k As Long, t As String
    ColetarMapas = MapasRecusados(TextoResposta(anc, sai), True)
    If mDescobrindo Then Exit Function          ' no diagnostico a tag continua valida para o outro mapa
    For k = 1 To NI
        t = TagAtual(k)
        If t <> "" And Not TagInvalida(t) And MapaAtual(k) = "" Then
            If mInvalidas = "" Then mInvalidas = "|"
            mInvalidas = mInvalidas & UCase$(t) & "|"
        End If
    Next k
End Function

' Usado so na validacao automatica: "n|mapa do 1o item|mapa do 2o"
Public Function TesteMapas(ByVal texto As String) As String
    Dim k1 As Long, k2 As Long, n As Long
    mMapasRuins = ""
    n = MapasRecusados(texto, True)
    k1 = IIndice("lamaC3")
    k2 = IIndice("lamaC2")
    TesteMapas = n & "|" & MapaAtual(k1) & "|" & MapaAtual(k2)
    mMapasRuins = ""
End Function

' Usado so na validacao automatica: tags recusadas numa resposta do Aspen
Public Function TesteTagsInvalidas(ByVal texto As String) As String
    Dim anc As Range
    Set anc = shDadosMES.Range(ANC_D)
    mInvalidas = ""
    anc.Value = "'" & texto
    ColetarInvalidas anc, shDadosMES.Range(OUT_D)
    TesteTagsInvalidas = ListaInvalidas() & "|" & TagAtual(IIndice("prodC2"))
    mInvalidas = ""
    anc.ClearContents
End Function

' ---------------------------------------------------------------- formula do Aspen
' Texto literal para formula, em pedacos de 250 caracteres unidos com & (limite do Excel: 255)
Private Function Literal(ByVal texto As String) As String
    Dim r As String, i As Long
    For i = 1 To Len(texto) Step 250
        If r <> "" Then r = r & "&"
        r = r & """" & Mid$(texto, i, 250) & """"
    Next i
    If r = "" Then r = """"""
    Literal = r
End Function

Public Function ArgumentosJanela(ByVal j As String, ByVal ps As String, ByVal anc As Range, _
                                 ByVal sai As Range) As String
    Dim lst As Variant, i As Long, k As Long, tags As String, servs As String, mapas As String
    Dim s As String, ci As String, cf As String, calc As String
    lst = Split(ps, ",")
    For i = LBound(lst) To UBound(lst)
        k = CLng(lst(i))
        s = ICfg(k, C_COL_SERV)
        If s = "" Then s = "UBU"
        If tags <> "" Then
            tags = tags & ","
            servs = servs & ","
            mapas = mapas & ","
        End If
        tags = tags & TagAtual(k)
        servs = servs & s
        mapas = mapas & MapaAtual(k)
    Next i
    ci = "Dados_MES!" & Nm("mesIni" & j).Address
    cf = "Dados_MES!" & Nm("mesFim" & j).Address
    calc = CfgTxt("cfgCalc")
    If calc = "" Then calc = "1"
    ArgumentosJanela = """time,attribute""," & Literal(tags) & "," & Literal(servs) & "," & Literal(mapas) & _
        "," & ci & "," & cf & ",""1h"",0,"""",0,""" & calc & """,0,1560,0,0,1,1," & _
        "ADDRESS(ROW(Dados_MES!" & anc.Address(False, False) & "),COLUMN(Dados_MES!" & anc.Address(False, False) & _
        "),1,,""Dados_MES""),""Dados_MES!" & sai.Address(False, False) & """,1)"
End Function

' Escreve a consulta na ancora. O nome da funcao do Aspen aparece de formas diferentes conforme a instalacao
' (com ou sem "_xll."); testa as formas e guarda a que o Excel reconhece (sem #NOME?) em mesPrefixo.
Private Function EscreverConsulta(ByVal anc As Range, ByVal args As String) As Boolean
    Dim cand As Variant, i As Long, pre As String, lista As String
    lista = CfgTxt("mesPrefixo") & "|=AspenTech.PME.ProcessData.Functions.GetCalculationValues(" & _
            "|=_xll.AspenTech.PME.ProcessData.Functions.GetCalculationValues(" & "|=GetCalculationValues("
    cand = Split(lista, "|")
    For i = LBound(cand) To UBound(cand)
        pre = CStr(cand(i))
        If pre <> "" Then
            On Error Resume Next
            Err.Clear
            anc.Formula = pre & args
            If Err.Number = 0 Then
                anc.Calculate
                If Not CelulaErroNome(anc) Then
                    On Error GoTo 0
                    If CfgTxt("mesPrefixo") <> pre Then GravarCfg "mesPrefixo", "'" & pre
                    EscreverConsulta = True
                    Exit Function
                End If
            End If
            On Error GoTo 0
        End If
    Next i
End Function

Private Sub GravarCfg(ByVal nome As String, ByVal v As Variant)
    On Error Resume Next
    Desproteger shConfig
    Nm(nome).Value = v
    Proteger shConfig
End Sub

Private Function EhErroNome(ByVal v As Variant) As Boolean
    On Error Resume Next
    If IsError(v) Then EhErroNome = (v = CVErr(2029))
End Function

' Celula com #NOME? / #NAME? (funcao nao reconhecida)
Private Function CelulaErroNome(ByVal c As Range) As Boolean
    Dim t As String
    On Error Resume Next
    If EhErroNome(c.Value) Then
        CelulaErroNome = True
        Exit Function
    End If
    t = UCase$(Trim$(c.Text))
    CelulaErroNome = (t = "#NAME?" Or t = "#NOME?" Or t = "#¿NOMBRE?")
End Function

Private Function PrefixoAspen(ByVal anc As Range) As String
    Dim atual As String
    atual = anc.Formula
    If InStr(atual, "(") > 0 And InStr(1, atual, "GetCalculationValues", vbTextCompare) > 0 Then
        PrefixoAspen = Left$(atual, InStr(atual, "("))
    ElseIf Len(CfgTxt("mesPrefixo")) > 0 Then
        PrefixoAspen = CfgTxt("mesPrefixo")
    Else
        PrefixoAspen = "=AspenTech.PME.ProcessData.Functions.GetCalculationValues("
    End If
End Function

' Depois de ler: a formula vira texto (o Aspen nao refaz a consulta a cada recalculo)
Private Sub DesligarConsulta(ByVal anc As Range, ByVal sai As Range)
    On Error Resume Next
    If anc.HasFormula Then
        anc.Offset(0, 7).Value = "'" & anc.Formula
        anc.Value = "'Consulta encerrada em " & Format$(Now, "dd\/mm\/yyyy hh:mm:ss") & " (fórmula em H)"
    End If
    LimparSaida sai
End Sub

Private Sub LimparSaida(ByVal sai As Range)
    Dim area As Range, c As Range
    On Error Resume Next
    Set area = sai.Resize(LIN_SAIDA, NI + 1)
    For Each c In area.Cells
        If c.HasArray Then c.CurrentArray.ClearContents
    Next c
    area.ClearContents
End Sub

Private Sub RestaurarSaida(ByVal anc As Range, ByVal sai As Range, ByVal nLin As Long, ByVal nCol As Long)
    Dim prefixo As String, ender As String
    On Error Resume Next
    prefixo = PrefixoAspen(anc)
    prefixo = Left$(prefixo, InStr(1, prefixo, "GetCalculationValues", vbTextCompare) - 1)
    ender = "Dados_MES!" & anc.Address(False, False)
    sai.Resize(nLin, nCol).FormulaArray = prefixo & "ShowCalculationValues(ADDRESS(ROW(" & ender & "),COLUMN(" & _
        ender & "),1,,""Dados_MES"")," & ender & ", 0)"
End Sub

' Inicio/fim do periodo nas celulas usadas pela formula
Private Sub GravarPeriodo(ByVal j As String, ByVal ini As Date, ByVal fim As Date)
    Dim ci As Range, cf As Range, modo As String
    Set ci = Nm("mesIni" & j)
    Set cf = Nm("mesFim" & j)
    modo = UCase$(CfgTxt("cfgFormatoData"))
    If InStr(modo, "MM/DD") > 0 Then
        ci.NumberFormat = "@"
        ci.Value = Format$(ini, "mm\/dd\/yyyy hh:mm:ss")
        cf.NumberFormat = "@"
        cf.Value = Format$(fim, "mm\/dd\/yyyy hh:mm:ss")
    ElseIf InStr(modo, "TEXTO") > 0 Then
        ci.NumberFormat = "@"
        ci.Value = Format$(ini, "dd\/mm\/yyyy hh:mm:ss")
        cf.NumberFormat = "@"
        cf.Value = Format$(fim, "dd\/mm\/yyyy hh:mm:ss")
    Else
        ci.NumberFormat = "dd/mm/yyyy hh:mm:ss"
        ci.Value = ini
        cf.NumberFormat = "dd/mm/yyyy hh:mm:ss"
        cf.Value = fim
    End If
End Sub

' ---------------------------------------------------------------- espera da resposta
' 1 = resultados prontos; 0 = sem resposta no tempo maximo; 2 = "Sucesso" sem linhas;
' 3 = tag recusada; 4 = mapa recusado; 5 = erro do Aspen
Private Function Esperar(ByVal anc As Range, ByVal sai As Range, ByVal ini As Date, ByVal limite As Double, _
                         ByVal nCol As Long) As Long
    Dim t0 As Single, tEst As Single, assin As String, ant As String, v As Variant, pronto As Boolean, tj As Date
    Dim seg As Long, tol As Double
    seg = -1
    tol = 1 / 24 + 1 / 1440
    t0 = Timer
    tEst = Timer
    Do
        DoEvents
        If Timer < t0 Then t0 = t0 - 86400
        If Timer < tEst Then tEst = tEst - 86400
        If Int(Timer - t0) <> seg Then
            seg = Int(Timer - t0)
            Application.StatusBar = "Consultando MES - " & mEtapa & "...  " & seg & " s  (máx. " & limite & " s)"
        End If
        If RespostaTagInvalida(anc, sai) Then
            Esperar = 3
            Exit Function
        End If
        If MapasRecusados(TextoResposta(anc, sai), False) > 0 Then
            Esperar = 4
            Exit Function
        End If
        v = sai.Cells(1, 1).Value
        pronto = False
        If ComoData(v, tj, ini) Then pronto = (Abs(CDbl(tj) - CDbl(ini)) <= tol)
        assin = Assinatura(anc, sai, nCol)
        If assin <> ant Then
            ant = assin
            tEst = Timer
        End If
        If pronto And Timer - tEst >= ESTAVEL_SEG Then
            Esperar = 1
            Exit Function
        End If
        If Not pronto And Timer - tEst >= ESTAVEL_SEG Then
            If RespondeuErro(anc) Then
                Esperar = 5
                Exit Function
            End If
        End If
        If Not pronto And Timer - t0 >= SEM_DADOS_SEG Then
            If RespondeuSucesso(anc) And Vazio(sai.Cells(1, 1).Value) Then
                Esperar = 2
                Exit Function
            End If
        End If
        If Timer - t0 > limite Then
            If pronto Then Esperar = 1 Else Esperar = 0
            Exit Function
        End If
    Loop
End Function

' "Success" / "Sucesso" (o suplemento responde no idioma do Office)
Private Function RespondeuSucesso(ByVal anc As Range) As Boolean
    Dim t As String
    t = LCase$(Trim$(anc.Text))
    RespondeuSucesso = (Left$(t, 7) = "success" Or Left$(t, 7) = "sucesso" Or Left$(t, 5) = "éxito")
End Function

Private Function RespondeuErro(ByVal anc As Range) As Boolean
    Dim t As String
    t = LCase$(Trim$(anc.Text))
    RespondeuErro = (Left$(t, 4) = "erro" Or Left$(t, 5) = "error" Or InStr(t, " is invalid") > 0)
End Function

Private Function Assinatura(ByVal anc As Range, ByVal sai As Range, ByVal nCol As Long) As String
    Dim arr As Variant, i As Long, j As Long, n As Long, soma As Double
    arr = sai.Resize(LIN_SAIDA, nCol).Value
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            If Not Vazio(arr(i, j)) Then
                n = n + 1
                If ENumero(arr(i, j)) Then soma = soma + CDbl(arr(i, j)) Else soma = soma + Len(CStr(arr(i, j)))
            End If
        Next j
    Next i
    Assinatura = anc.Text & "|" & n & "|" & soma
End Function

Private Function TextoResposta(ByVal anc As Range, ByVal sai As Range) As String
    On Error Resume Next
    TextoResposta = anc.Text & " " & sai.Cells(1, 1).Text
End Function

Private Function RespostaTagInvalida(ByVal anc As Range, ByVal sai As Range) As Boolean
    Dim t As String, i As Long, j As Long, tag As String
    t = TextoResposta(anc, sai)
    i = InStr(1, t, "Tag Name ", vbTextCompare)
    Do While i > 0
        j = InStr(i + 9, t, " is invalid", vbTextCompare)
        If j = 0 Then Exit Do
        tag = UCase$(Trim$(Mid$(t, i + 9, j - i - 9)))
        If tag <> "" And Not TagInvalida(tag) Then
            RespostaTagInvalida = True
            Exit Function
        End If
        i = InStr(j, t, "Tag Name ", vbTextCompare)
    Loop
End Function

' ---------------------------------------------------------------- simulacao (treinar sem MES)
Private Sub Simular(ByVal d As Date, v() As Variant, n() As Variant, s() As Variant)
    Dim k As Long, x As Double, ch As String
    For k = 1 To NI
        If ICfg(k, C_COL_TAG) <> "" Then
            x = Sin(CDbl(d) * 0.7 + k)
            ch = IChave(k)
            Select Case ch
                Case "prodC3", "prodC2", "concC3", "concC2"
                    v(k) = Round(24000 + 3500 * x, 0)
                    n(k) = IIf(UCase$(ICfg(k, C_COL_JAN)) = "E", 23, 1)
                Case "lamaC3"
                    v(k) = IIf(x > 0.3, Round(400 * x, 0), 0)
                    n(k) = 1
                Case "lamaC2"
                    v(k) = IIf(x > 0.6, Round(80 * x, 0), 0)
                    n(k) = 1
                Case Else
                    v(k) = Round(2050 + 140 * x + 9 * k, 0)
                    n(k) = IIf(UCase$(ICfg(k, C_COL_JAN)) = "E", 12, 13)
            End Select
        End If
    Next k
End Sub

' Usado so na validacao automatica: saida simulada da janela E (17/08 22:30, 23 janelas) com a producao C2
' (1000 t em cada hora) e a S.E. do espessador C2 nos horarios impares (2000 + k); 99999 fora da faixa e
' uma linha alem das 23 janelas -> "soma|n|media|n"
Public Function TesteLeitura() As String
    Dim sai As Range, q As Date, k As Long, kP As Long, kE As Long, v() As Variant, n() As Variant, ps As String
    Set sai = shDadosMES.Range(OUT_E)
    LimparSaida sai
    kP = IIndice("prodC2")
    kE = IIndice("espC2")
    ps = kP & "," & kE
    q = DateSerial(2026, 8, 17) + TimeSerial(22, 30, 0)
    For k = 0 To 23
        sai.Cells(k + 1, 1).Value = Format$(CDate(CDbl(q) + k / 24), "dd\/mm\/yyyy hh:mm:ss")
        sai.Cells(k + 1, 2).Value = IIf(k = 23, 5000, "1000.0")
        If k Mod 2 = 0 Then sai.Cells(k + 1, 3).Value = 2000 + k
        If k = 1 Then sai.Cells(k + 1, 3).Value = 99999
    Next k
    ReDim v(1 To NI)
    ReDim n(1 To NI)
    LerSaida sai.Resize(LIN_SAIDA, 3).Value, q, 23, ps, v, n
    TesteLeitura = v(kP) & "|" & n(kP) & "|" & Round(v(kE), 2) & "|" & n(kE)
    LimparSaida sai
End Function

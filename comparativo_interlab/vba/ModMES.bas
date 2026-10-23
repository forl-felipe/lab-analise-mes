Option Explicit

' ============================================================================
'  ModMES - resultados de Germano e de Ubu no MES (Aspen IP.21)
'  Botao "Buscar no MES" (aba Batch):
'    Ubu     : janelas de 2 h a partir do "Inicio no processo em Ubu"
'    Germano : as mesmas janelas deslocadas pela defasagem (inicio do batch em Germano)
'  Uma consulta GetCalculationValues por laboratorio, com tags/servidores/mapas em
'  texto literal (mesmo formato dos relatorios do laboratorio). Tags recusadas pelo
'  MES ("Tag Name ... is invalid") saem da consulta, que e refeita com as demais.
' ============================================================================

' Segundos sem mudanca nos resultados para considerar a consulta concluida
Private Const ESTAVEL_SEG As Double = 2
' Consulta que respondeu "Success" sem nenhuma linha: "sem dados" depois deste tempo
Private Const SEM_DADOS_SEG As Double = 15
Private Const MAX_RODADAS As Long = 10
' Celulas da aba Dados_MES (consulta de Germano e de Ubu)
Private Const ANC_G As String = "A10"
Private Const OUT_G As String = "A12"
Private Const ANC_U As String = "A70"
Private Const OUT_U As String = "A72"
' Pellet feed ultra fino (composto diario): "laboratorio" 2
Private Const ANC_P As String = "A130"
Private Const OUT_P As String = "A132"
Private Const LAB_PF As Long = 2

Private mOcupado As Boolean
' Funcao do Aspen nao reconhecida nesta busca (#NOME?): nao repete o aviso para o 2o laboratorio
Private mSemAspen As Boolean
' Tags recusadas pelo MES nesta sessao, no formato |TAG1|TAG2|
Private mInvalidas As String
' Mapas recusados pelo MES nesta sessao ("Map Name X is invalid"), no formato |TAG>MAPA|
Private mMapasRuins As String
' Germano em sequencia: resultados achados na ultima consulta (na ordem) e o horario de cada um
Private mContG As Long
Private mTG(1 To 60) As Double
' Linhas lidas da saida de uma consulta (Germano em sequencia usa janelas de 1 h: ate 2 x NMAX + 1)
Private Const LIN_SAIDA As Long = 52
' Contador da consulta na barra de status: o que esta sendo consultado
Private mEtapa As String
' Mapas tentados quando o MES recusa o da aba Configuracoes (o que funcionar fica gravado la)
' Motivo da ultima falha da consulta do PFUF (para a mensagem)
Private mErroPF As String
' Dias em que o MES tem resultado do PFUF na ultima consulta ("19/10 (4), 20/10 (4)")
Private mDiasPF As String
' Por analise do PFUF: onde o valor foi achado / por que ficou de fora (para a mensagem)
Private mNotaPF(1 To 10) As String
' PFUF: janelas de 2 h (igual as leituras do batch, que funcionam) de 12 h antes do dia do composto ate 12 h
' depois; o resultado vale na janela do horario de lancamento (cfgPfHora, 23:30)
Private Const PF_JANELAS As Long = 24
Private Const MAPAS_ALT As String = "IP_AnalogMap|IP_MESVALOR|IP_DiscreteMap"

' ---------------------------------------------------------------- botao
Public Sub BuscarMES()
    Dim n As Long, nG As Long, iniU As Date, iniG As Date
    Dim resG As Variant, resU As Variant, simulado As Boolean, okG As Boolean, okU As Boolean
    Dim qG As Long, qU As Long, msg As String, semTag As String, pf As String, t0 As Single, seg As Long

    If mOcupado Then
        Aviso "Consulta ao MES em andamento. Aguarde.", vbInformation, "MES"
        Exit Sub
    End If
    Recalcular
    If ModoAcomp() Then
        BuscarAcompanhamento
        Exit Sub
    End If
    If Not EData(Nm("gJanU").Value) Then
        ' sem os horarios do batch: busca so o pellet feed ultra fino
        BuscarSoPFUF "Leituras do batch: não buscadas - falta a chegada em Ubu (cole as Fases do Evento ou " & _
                   "informe a chegada/início no processo na aba Batch)."
        Exit Sub
    End If
    n = 0
    If ENumero(Nm("bNHor").Value) Then n = CLng(Nm("bNHor").Value)
    If n < 1 Or n > NMAX Then
        Aviso "Nº de horários deve ficar entre 1 e " & NMAX & ".", vbExclamation, "MES"
        Exit Sub
    End If
    simulado = (Left$(UCase$(CfgTxt("cfgFonte")), 3) = "SIM")
    semTag = TagsFaltando()
    If Not simulado And semTag = "TODAS" Then
        Aviso "Nenhuma tag do MES cadastrada." & vbCrLf & vbCrLf & _
              "Preencha as tags de Germano e de Ubu na aba Configurações (ou use Fonte = SIMULAÇÃO para treinar).", _
              vbExclamation, "MES"
        shConfig.Activate
        Exit Sub
    End If

    ' janelas de 2 h: Ubu a partir da 1a janela completa depois do inicio (UBU3); Germano idem a partir
    ' do inicio do bombeamento (sem inicio de Germano = mesmos horarios de Ubu)
    iniU = CDate(Nm("gJanU").Value)
    nG = n
    If EData(Nm("gJanG").Value) Then
        iniG = CDate(Nm("gJanG").Value)
        If ENumero(Nm("gNG").Value) Then nG = CLng(Nm("gNG").Value)
    Else
        iniG = iniU
    End If
    If nG > n Then nG = n

    mOcupado = True
    mSemAspen = False
    Ampulheta True
    t0 = Timer
    On Error GoTo Falha
    If simulado Then
        If nG > 0 Then resG = DadosSimulados(0, nG, iniG)
        resU = DadosSimulados(1, n, iniU)
        okG = True
        okU = True
    Else
        Application.StatusBar = "Consultando MES - " & CfgTxt("bLabA") & "..."
        If nG > 0 Then okG = ConsultarLab(0, iniG, nG, resG) Else okG = True
        If okG Then AjustarNHorG nG, n
        Application.StatusBar = "Consultando MES - " & CfgTxt("bLabB") & "..."
        okU = ConsultarLab(1, iniU, n, resU)
    End If
    Application.StatusBar = False
    If Not okG And Not okU Then GoTo Fim

    If simulado Then mContG = 0
    qG = GravarResultados(0, resG, nG, okG, iniG)
    qU = GravarResultados(1, resU, n, okU, iniU)
    If okG Then GravarHorariosG nG
    Application.StatusBar = "Consultando MES - pellet feed ultra fino..."
    pf = AtualizarPFUF()
    Application.StatusBar = False
    seg = Int(Timer - t0)
    If seg < 0 Then seg = seg + 86400
    Nm("mesTempo").Value = "'" & Format$(Now, "dd\/mm\/yyyy hh:mm:ss") & IIf(simulado, " (SIMULAÇÃO)", "  -  " & seg & " s")
    AjustarGraficos
    msg = "Resultados gravados no batch " & CfgTxt("bId") & ":" & vbCrLf & vbCrLf & _
          "  " & CfgTxt("bLabA") & ": " & qG & IIf(okG, "", " (consulta falhou)") & _
          "   (" & nG & " janelas a partir de " & Format$(iniG, "dd\/mm hh:mm") & ")" & vbCrLf & _
          "  " & CfgTxt("bLabB") & ": " & qU & IIf(okU, "", " (consulta falhou)") & _
          "   (" & n & " janelas a partir de " & Format$(iniU, "dd\/mm hh:mm") & ")"
    If Not simulado And (CDbl(iniU) + n * HORAS_SLOT / 24 > CDbl(Now) Or CDbl(iniG) + nG * HORAS_SLOT / 24 > CDbl(Now)) Then
        msg = msg & vbCrLf & vbCrLf & "Batch ainda em andamento: horários futuros ficam em branco."
    End If
    If Not simulado And ((okG And qG = 0 And nG > 0) Or (okU And qU = 0)) Then
        msg = msg & vbCrLf & vbCrLf & "Resultado 0 = o MES respondeu, mas não tem resultados dessas tags no período."
    End If
    If semTag <> "" And semTag <> "TODAS" Then
        msg = msg & vbCrLf & vbCrLf & "Sem tag (lançamento manual): " & semTag
    End If
    If pf <> "" Then msg = msg & vbCrLf & vbCrLf & pf
    If simulado Then msg = msg & vbCrLf & vbCrLf & "Fonte: SIMULAÇÃO (dados fictícios)." Else msg = msg & _
        vbCrLf & vbCrLf & "Tempo total da consulta: " & seg & " s."
    Aviso msg, vbInformation, "MES"
Fim:
    Application.StatusBar = False
    Ampulheta False
    mOcupado = False
    Exit Sub
Falha:
    Application.StatusBar = False
    Ampulheta False
    mOcupado = False
    Aviso "Erro ao consultar o MES: " & Err.Description, vbExclamation, "MES"
End Sub

' ---------------------------------------------------------------- batch em processamento (acompanhamento)
' So Germano (Ubu e lancado a mao nas abas Acomp. Quimico / Acomp. Fisico); monta os blocos das duas abas
Private Sub BuscarAcompanhamento()
    Dim n As Long, nG As Long, iniG As Date, resG As Variant, okG As Boolean, qG As Long, simulado As Boolean
    Dim msg As String, t0 As Single, seg As Long

    If Not ENumero(Nm("bId").Value) Then
        Aviso "Informe o número do batch (passo 1).", vbExclamation, "MES - acompanhamento"
        Exit Sub
    End If
    If Not EData(Nm("gJanG").Value) Then
        Aviso "Informe o início do bombeamento em Germano (EB06): cole as Fases do Evento (passo 2) ou preencha " & _
              "Germano · início na aba Batch.", vbExclamation, "MES - acompanhamento"
        Exit Sub
    End If
    n = 0
    If ENumero(Nm("bNHor").Value) Then n = CLng(Nm("bNHor").Value)
    If n < 1 Or n > NMAX Then
        Aviso "Nº de horários deve ficar entre 1 e " & NMAX & ".", vbExclamation, "MES - acompanhamento"
        Exit Sub
    End If
    nG = n
    If ENumero(Nm("gNG").Value) Then nG = CLng(Nm("gNG").Value)
    If nG > n Then nG = n
    iniG = CDate(Nm("gJanG").Value)
    simulado = (Left$(UCase$(CfgTxt("cfgFonte")), 3) = "SIM")

    mOcupado = True
    mSemAspen = False
    Ampulheta True
    t0 = Timer
    On Error GoTo Falha
    If simulado Then
        resG = DadosSimulados(0, nG, iniG)
        okG = True
        mContG = 0
    Else
        okG = ConsultarLab(0, iniG, nG, resG)
        If okG Then AjustarNHorG nG, n
    End If
    Application.StatusBar = False
    If okG Then
        qG = GravarResultados(0, resG, nG, okG, iniG)
        GravarHorariosG nG
    End If
    ' Ubu do batch em processamento e lancado na aba Acompanhamento: o que estiver na aba Batch e de outro
    ' batch e nao pode ir para o bloco
    LimparUbuBatch
    msg = MontarAcompanhamento()
    seg = Int(Timer - t0)
    If seg < 0 Then seg = seg + 86400
    Nm("mesTempo").Value = "'" & Format$(Now, "dd\/mm\/yyyy hh:mm:ss") & IIf(simulado, " (SIMULAÇÃO)", "  -  " & seg & " s")
    AjustarGraficos
    Ampulheta False
    mOcupado = False
    msg = "Acompanhamento do batch " & CfgTxt("bId") & " (só " & CfgTxt("bLabA") & "):" & vbCrLf & vbCrLf & _
          "  " & CfgTxt("bLabA") & ": " & qG & IIf(okG, "", " (consulta falhou)") & "   (" & nG & _
          " janelas a partir de " & Format$(iniG, "dd\/mm hh:mm") & ")" & vbCrLf & vbCrLf & msg & _
          IIf(simulado, vbCrLf & vbCrLf & "Fonte: SIMULAÇÃO (dados fictícios).", vbCrLf & vbCrLf & _
          "Tempo total da consulta: " & seg & " s.")
    Aviso msg, vbInformation, "MES - acompanhamento"
    Exit Sub
Falha:
    Application.StatusBar = False
    Ampulheta False
    mOcupado = False
    Aviso "Erro no acompanhamento: " & Err.Description, vbExclamation, "MES - acompanhamento"
End Sub

Private Sub LimparUbuBatch()
    Dim p As Long
    ModoCarga True
    Desproteger shBatch
    For p = 1 To NP
        shBatch.Range(shBatch.Cells(B_ROW_G0, ColValor(p, 1)), shBatch.Cells(B_ROW_G0 + NMAX - 1, ColValor(p, 1))).ClearContents
    Next p
    Proteger shBatch
    ModoCarga False
End Sub

' ---------------------------------------------------------------- tags (aba Configuracoes)
Public Function TagCfg(ByVal p As Long, ByVal lab As Long, ByVal campo As Long) As String
    Dim v As Variant
    v = shConfig.Cells(LinhaCfg(p, lab), campo).Value
    If IsError(v) Then TagCfg = "" Else TagCfg = Limpo(CStr(v))
End Function

' Linha da tag na aba Configuracoes: lab 0/1 = tabela Germano/Ubu; lab 2 = pellet feed ultra fino
Private Function LinhaCfg(ByVal p As Long, ByVal lab As Long) As Long
    If lab = LAB_PF Then
        LinhaCfg = C_ROW_PF0 + p - 1
    Else
        LinhaCfg = C_ROW_TAG0 + 2 * (p - 1) + lab
    End If
End Function

Private Function NItens(ByVal lab As Long) As Long
    If lab = LAB_PF Then NItens = NPF Else NItens = NP
End Function

Private Function RotuloItem(ByVal p As Long, ByVal lab As Long) As String
    If lab = LAB_PF Then
        RotuloItem = TagCfg(p, lab, C_COL_LAB)
        If RotuloItem = "" Then RotuloItem = Split(PF_CHAVES, "|")(p - 1)
    Else
        RotuloItem = PRotulo(p)
    End If
End Function

' Sem tabulacao, quebra de linha e espaco duro (vem junto quando a tag e copiada de outro programa)
Public Function Limpo(ByVal t As String) As String
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, ChrW(160), "")
    Limpo = Trim$(t)
End Function

Private Function TagInvalida(ByVal t As String) As Boolean
    If t = "" Then Exit Function
    TagInvalida = (InStr(1, mInvalidas, "|" & UCase$(t) & "|") > 0)
End Function

' Analise sem tag: "S.E. Germano, P Ubu..." ; "TODAS" se nenhuma tag
Private Function TagsFaltando() As String
    Dim p As Long, lab As Long, r As String, n As Long
    For p = 1 To NP
        For lab = 0 To 1
            If TagCfg(p, lab, C_COL_TAG) = "" Then
                n = n + 1
                If r <> "" Then r = r & ", "
                r = r & PRotulo(p) & " " & IIf(lab = 0, CfgTxt("bLabA"), CfgTxt("bLabB"))
            End If
        Next lab
    Next p
    If n = 2 * NP Then r = "TODAS"
    TagsFaltando = r
End Function

' Parametros do laboratorio com tag valida (lista "1,3,5")
' Mesma tag com o codigo do meio em outro formato: -0008- <-> -8- (o MES aceita so um deles)
Public Function VarianteTag(ByVal t As String) As String
    Dim pt As Variant, i As Long, r As String
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
Private Function TagAtual(ByVal p As Long, ByVal lab As Long) As String
    Dim t As String, alt As String
    t = TagCfg(p, lab, C_COL_TAG)
    TagAtual = t
    If t = "" Then Exit Function
    If TagInvalida(t) Then
        alt = VarianteTag(t)
        If alt <> "" Then
            If Not TagInvalida(alt) Then TagAtual = alt
        End If
    End If
End Function

Private Function ParamsLab(ByVal lab As Long) As String
    Dim p As Long, t As String, r As String
    For p = 1 To NItens(lab)
        t = TagAtual(p, lab)
        If t <> "" And Not TagInvalida(t) Then
            If r <> "" Then r = r & ","
            r = r & p
        End If
    Next p
    ParamsLab = r
End Function

' ---------------------------------------------------------------- consulta de um laboratorio
' res(1..n, 1..NP): valores das janelas de 2 h (Empty = sem resultado)
Private Function ConsultarLab(ByVal lab As Long, ByVal ini As Date, ByVal n As Long, ByRef res As Variant) As Boolean
    Dim anc As Range, sai As Range, ps As String, rodada As Long, limite As Double, nJan As Long
    Dim fim As Date, t0 As Single, d() As Variant, lst As Variant, i As Long, r As Long, s As Long, v As Variant
    Dim estado As Long, nomeLab As String, invalidas As String, tj As Date, x As Double, resp As String
    Dim restaurado As Boolean, e2 As Long, corrigidas As String, seq As Boolean, passo As Double, qIni As Date
    Dim nTot As Long

    ' Germano em sequencia: janelas de 1 h comecando 30 min depois da grade (cada resultado do MES, 01:00, 03:00...
    ' e o de fechamento 23:59, cai numa janela so); os resultados entram um atras do outro, sem coluna vazia
    seq = (lab = 0 And SequenciaG())
    If seq Then
        ReDim d(1 To NMAX, 1 To NP)
        passo = 1
        qIni = CDate(CDbl(ini) + 0.5 / 24)
        nTot = 2 * n
    Else
        ReDim d(1 To n, 1 To NP)
        passo = HORAS_SLOT
        qIni = ini
        nTot = n
    End If
    If lab = 0 Then mContG = 0
    res = d
    If mSemAspen Then Exit Function
    nomeLab = IIf(lab = 0, CfgTxt("bLabA"), CfgTxt("bLabB"))
    mEtapa = nomeLab
    If lab = 0 Then
        Set anc = shDadosMES.Range(ANC_G)
        Set sai = shDadosMES.Range(OUT_G)
    Else
        Set anc = shDadosMES.Range(ANC_U)
        Set sai = shDadosMES.Range(OUT_U)
    End If
    ps = ParamsLab(lab)
    If ps = "" Then
        ConsultarLab = True          ' nada a consultar (sem tags): lancamento manual
        Exit Function
    End If
    ' nunca pede ao MES janelas que ainda nao comecaram
    nJan = nTot
    If CDbl(qIni) + nJan * passo / 24 > CDbl(Now) Then
        nJan = Int((CDbl(Now) - CDbl(qIni)) * 24 / passo) + 1
        If nJan > nTot Then nJan = nTot
    End If
    If nJan < 1 Then
        ConsultarLab = True          ' batch ainda nao chegou a este laboratorio
        Exit Function
    End If
    fim = CDate(CDbl(qIni) + nJan * passo / 24)
    limite = 120
    If ENumero(Nm("cfgTimeout").Value) Then limite = CDbl(Nm("cfgTimeout").Value)

    GravarPeriodo lab, qIni, fim
    For rodada = 0 To MAX_RODADAS
        ps = ParamsLab(lab)
        If ps = "" Then Exit For
        LimparSaida sai
        If Not EscreverConsulta(anc, ArgumentosLab(lab, ps, anc, sai)) Then
            Nm(IIf(lab = 0, "mesRespG", "mesRespU")).Value = "'" & Format$(Now, "dd\/mm hh:mm:ss") & "  #NOME? (suplemento Aspen)"
            anc.Value = "'Suplemento Aspen não reconhecido (#NOME?)"
            LimparSaida sai
            mSemAspen = True
            Aviso "O Excel não reconheceu a função do MES (GetCalculationValues) - aparece #NOME?." & vbCrLf & vbCrLf & _
                  "Verificar se o suplemento Aspen Process Data está ativo neste Excel: Arquivo > Opções > " & _
                  "Suplementos > Gerenciar: Suplementos do Excel / COM > Ir... (o mesmo usado no relatório do " & _
                  "laboratório físico). Depois feche e abra a planilha.", vbExclamation, "MES"
            Exit Function
        End If
        estado = Esperar(anc, sai, qIni, limite)
        If estado = 3 Then
            ' tag recusada: guarda e consulta de novo sem ela (ou com a variante do codigo)
            If ColetarInvalidas(anc, sai) = 0 Then Exit For
        ElseIf estado = 4 Then
            ' mapa recusado: consulta de novo com o proximo mapa da lista MAPAS_ALT
            If ColetarMapas(anc, sai, lab) = 0 Then Exit For
        ElseIf estado = 5 Then
            Exit For
        ElseIf (estado = 0 Or estado = 2) And Not restaurado Then
            ' respondeu sem resultados na celula de saida: monta a matriz ShowCalculationValues
            ' (forma da planilha de referencia do MES) e espera mais uma vez
            restaurado = True
            RestaurarSaida anc, sai, nJan, UBound(Split(ps, ",")) + 2
            anc.Calculate
            sai.Resize(nJan, UBound(Split(ps, ",")) + 2).Calculate
            e2 = Esperar(anc, sai, qIni, limite)
            If e2 = 1 Or estado = 0 Then estado = e2
            Exit For
        Else
            Exit For
        End If
    Next rodada
    ' variantes de codigo que o MES aceitou: passam para a aba Configuracoes
    If estado = 1 Then corrigidas = CorrigirTags(lab) & CorrigirMapas(lab)

    If corrigidas <> "" Then
        Aviso "Tags de " & nomeLab & " corrigidas na aba Configurações (o MES aceitou o código em outro formato):" & _
              vbCrLf & vbCrLf & corrigidas, vbInformation, "MES - tags"
    End If
    invalidas = ListaInvalidas(lab)
    If invalidas <> "" Then
        Nm("mesInvalidas").Value = "'" & Replace(invalidas, vbCrLf, " ")
        Aviso "O MES não reconheceu estas tags de " & nomeLab & " (ficam em branco):" & vbCrLf & vbCrLf & _
              invalidas & vbCrLf & "Corrigir na aba Configurações.", vbExclamation, "MES - tags inválidas"
    End If
    resp = Left$(Trim$(anc.Text) & "  |  1a célula: " & Trim$(sai.Cells(1, 1).Text) & "  " & _
                 Trim$(sai.Cells(1, 2).Text), 250)
    Nm(IIf(lab = 0, "mesRespG", "mesRespU")).Value = "'" & Format$(Now, "dd\/mm hh:mm:ss") & "  " & resp
    If estado = 0 Or estado = 4 Or estado = 5 Then
        ' a formula vira texto (com a formula guardada para conferencia): uma consulta ativa refaz a busca
        ' no MES a cada recalculo e deixa a planilha lenta
        DesligarConsulta anc, sai
        If Aviso(IIf(estado = 0, "O MES não respondeu à consulta de ", "O MES devolveu erro na consulta de ") & nomeLab & " (" & Format$(ini, "dd\/mm\/yyyy hh:mm") & ")." & _
                 vbCrLf & vbCrLf & "Resposta do Aspen: " & resp & vbCrLf & vbCrLf & _
                 "Verificar: suplemento Aspen ativo, servidor, tags e formato da data " & _
                 "(aba Configurações)." & vbCrLf & vbCrLf & "Abrir a aba Dados_MES?", vbExclamation + vbYesNo, _
                 "MES sem resposta", 7) = vbYes Then MostrarDadosMES
        Exit Function
    End If

    ' le a matriz: 1a coluna = inicio da janela; demais = analises na ordem de ps
    lst = Split(ParamsLab(lab), ",")
    If estado = 1 And seq Then
        mContG = LerSequenciaG(sai, qIni, ParamsLab(lab), d)
    ElseIf estado = 1 Then
        For r = 1 To NMAX + 2
            If Not ComoData(sai.Cells(r, 1).Value, tj, ini) Then Exit For
            s = CLng((CDbl(tj) - CDbl(ini)) * 24 / HORAS_SLOT) + 1
            If s >= 1 And s <= n Then
                For i = LBound(lst) To UBound(lst)
                    v = sai.Cells(r, 2 + i - LBound(lst)).Value
                    If ComoNumero(v, x) Then
                        If ValorValido(CLng(lst(i)), lab, x) Then d(s, CLng(lst(i))) = x
                    End If
                Next i
            End If
        Next r
    End If
    res = d
    DesligarConsulta anc, sai
    ConsultarLab = True
End Function

Public Function SequenciaG() As Boolean
    SequenciaG = (UCase$(Left$(CfgTxt("cfgSeqG"), 1)) <> "N")
End Function

' Germano em sequencia: cada linha da saida com algum resultado valido vira o proximo bi-horario (1, 2, 3...);
' linhas sem resultado nao ocupam coluna. Guarda o horario de cada resultado (centro da janela de 1 h) em mTG.
' Devolve quantos resultados achou.
Private Function LerSequenciaG(ByVal sai As Range, ByVal qIni As Date, ByVal ps As String, d() As Variant) As Long
    Dim lst As Variant, r As Long, i As Long, k As Long, v As Variant, x As Double, tj As Date, tem As Boolean
    Dim vals() As Variant
    lst = Split(ps, ",")
    ReDim vals(LBound(lst) To UBound(lst))
    For r = 1 To LIN_SAIDA
        If Not ComoData(sai.Cells(r, 1).Value, tj, qIni) Then Exit For
        tem = False
        For i = LBound(lst) To UBound(lst)
            vals(i) = Empty
            v = sai.Cells(r, 2 + i - LBound(lst)).Value
            If ComoNumero(v, x) Then
                If ValorValido(CLng(lst(i)), 0, x) Then
                    vals(i) = x
                    tem = True
                End If
            End If
        Next i
        If tem Then
            k = k + 1
            If k <= NMAX Then
                For i = LBound(lst) To UBound(lst)
                    If Not IsEmpty(vals(i)) Then d(k, CLng(lst(i))) = vals(i)
                Next i
                mTG(k) = CDbl(tj) + 0.5 / 24
            End If
        End If
    Next r
    LerSequenciaG = k
End Function

' Usado so na validacao automatica: saida simulada do MES de Germano (batch 295: sem 11:00 e 13:00; 23:00 e
' 23:59 separados) -> "resultados|SiO2 x 100 do 6o|do 10o|do 11o|hora do 6o"
Public Function TesteSequenciaG() As String
    Dim sai As Range, d() As Variant, q As Date, k As Long, n As Long, j As Long, pSi As Long
    Dim horas As Variant, sio2 As Variant
    Set sai = shDadosMES.Range(OUT_G)
    LimparSaida sai
    pSi = PIndice("SiO2")
    q = DateSerial(2026, 10, 21) + TimeSerial(0, 30, 0)
    horas = Array(1, 3, 5, 7, 9, 15, 17, 19, 21, 23, 23.98)
    sio2 = Array(2.18, 1.78, 1.5, 1.54, 1.68, 1.78, 2.14, 1.95, 2.13, 2.07, 2.02)
    For k = 0 To 23
        sai.Cells(k + 1, 1).Value = CDate(CDbl(q) + k / 24)
    Next k
    For j = 0 To 10
        k = Int((horas(j) - 0.5) + 0.000001)          ' janela de 1 h que contem o horario
        sai.Cells(k + 1, 2).Value = sio2(j)
    Next j
    ReDim d(1 To NMAX, 1 To NP)
    n = LerSequenciaG(sai, q, CStr(pSi), d)
    TesteSequenciaG = n & "|" & Format$(CDbl(d(6, pSi)) * 100, "0") & "|" & Format$(CDbl(d(10, pSi)) * 100, "0") & "|" & _
                      Format$(CDbl(d(11, pSi)) * 100, "0") & "|" & Format$(CDate(mTG(6)), "hh:mm")
    LimparSaida sai
End Function

' Depois da consulta de Germano em sequencia: horario de cada resultado na coluna de Germano da aba Batch
' (as colunas alem dos resultados ficam sem horario)
Private Sub GravarHorariosG(ByVal n As Long)
    Dim k As Long, c As Range
    If mContG = 0 Then
        RestaurarHorariosG                ' sem resultado em sequencia: horarios da grade
        Exit Sub
    End If
    ModoCarga True
    Desproteger shBatch
    For k = 1 To NMAX
        Set c = shBatch.Cells(B_ROW_G0 + k - 1, B_COL_TG)
        If k <= mContG And k <= n Then
            c.Value = CDate(mTG(k))
        Else
            c.ClearContents
        End If
    Next k
    Proteger shBatch
    ModoCarga False
End Sub

' Horarios de Germano voltam a ser calculados pela grade (mesma formula de build_interlab.f_hora_g)
Public Sub RestaurarHorariosG()
    Dim i As Long, cu As String
    For i = 1 To NMAX
        cu = shBatch.Cells(B_ROW_G0 + i - 1, B_COL_TU).Address(False, False)
        shBatch.Cells(B_ROW_G0 + i - 1, B_COL_TG).Formula = "=IF(ISNUMBER(" & cu & "),IF(ISNUMBER(gJanG),IF(" & i & _
            "<=gNG,gJanG+" & (i - 1) & "/12+cfgMinG/1440,"""")," & cu & "),"""")"
    Next i
End Sub

' Mais resultados de Germano que bi-horarios: o Nº de horarios do batch passa a ser o Nº de resultados
Private Sub AjustarNHorG(ByRef nG As Long, ByRef n As Long)
    If mContG <= nG Then Exit Sub
    nG = mContG
    If nG > NMAX Then nG = NMAX
    If nG > n Then
        n = nG
        Desproteger shBatch
        Nm("bNHor").Value = n
        Proteger shBatch
    End If
End Sub

' 0 = sem resposta, 1 = resultados prontos, 2 = respondeu sem dados, 3 = tag recusada, 4 = mapa recusado,
' 5 = outro erro do Aspen
Private Function Esperar(ByVal anc As Range, ByVal sai As Range, ByVal ini As Date, ByVal limite As Double, _
                         Optional ByVal tol As Double = -1) As Long
    Dim t0 As Single, tEst As Single, assin As String, ant As String, v As Variant, pronto As Boolean, tj As Date
    Dim seg As Long
    seg = -1
    If tol < 0 Then tol = HORAS_SLOT / 24 + 1 / 1440
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
        If RespostaMapaInvalido(anc, sai) Then
            Esperar = 4
            Exit Function
        End If
        v = sai.Cells(1, 1).Value
        pronto = False
        If ComoData(v, tj, ini) Then pronto = (Abs(CDbl(tj) - CDbl(ini)) <= tol)
        assin = Assinatura(anc, sai)
        If assin <> ant Then
            ant = assin
            tEst = Timer
        End If
        If pronto And Timer - tEst >= ESTAVEL_SEG Then
            Esperar = 1
            Exit Function
        End If
        ' o Aspen respondeu com erro ("Erro: ...", "Error: ..."): nao adianta esperar o tempo maximo
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

Private Function Assinatura(ByVal anc As Range, ByVal sai As Range) As String
    Dim arr As Variant, i As Long, j As Long, n As Long, soma As Double
    arr = sai.Resize(LIN_SAIDA, NP + 1).Value
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

Private Function RespondeuErro(ByVal anc As Range) As Boolean
    Dim t As String
    t = LCase$(Trim$(anc.Text))
    RespondeuErro = (Left$(t, 4) = "erro" Or Left$(t, 5) = "error" Or InStr(t, " is invalid") > 0)
End Function

' "(TAG) Map Name IP_MESVALOR is invalid": mapa ainda nao marcado como recusado para a tag?
Private Function RespostaMapaInvalido(ByVal anc As Range, ByVal sai As Range) As Boolean
    RespostaMapaInvalido = (MapasRecusados(TextoResposta(anc, sai), False) > 0)
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
Private Function ColetarMapas(ByVal anc As Range, ByVal sai As Range, ByVal lab As Long) As Long
    Dim p As Long, t As String
    ColetarMapas = MapasRecusados(TextoResposta(anc, sai), True)
    For p = 1 To NItens(lab)
        t = TagAtual(p, lab)
        If t <> "" And Not TagInvalida(t) And MapaAtual(p, lab) = "" Then
            If mInvalidas = "" Then mInvalidas = "|"
            mInvalidas = mInvalidas & UCase$(t) & "|"
        End If
    Next p
End Function

Private Function MapaRecusado(ByVal tag As String, ByVal mapa As String) As Boolean
    MapaRecusado = (InStr(1, mMapasRuins, "|" & UCase$(tag) & ">" & UCase$(mapa) & "|") > 0)
End Function

' Mapa usado na consulta: o da aba Configuracoes ou, se o MES o recusou, o 1o da lista MAPAS_ALT ainda nao
' recusado ("" = nenhum)
Private Function MapaAtual(ByVal p As Long, ByVal lab As Long) As String
    Dim m As String, t As String, alt As Variant, i As Long
    m = TagCfg(p, lab, C_COL_MAPA)
    If m = "" Then m = "IP_MESVALOR"
    t = TagAtual(p, lab)
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

' Grava na aba Configuracoes os mapas alternativos que o MES aceitou; devolve a lista
Private Function CorrigirMapas(ByVal lab As Long) As String
    Dim p As Long, m As String, r As String
    For p = 1 To NItens(lab)
        If TagAtual(p, lab) <> "" And Not TagInvalida(TagAtual(p, lab)) Then
            m = MapaAtual(p, lab)
            If m <> "" And UCase$(m) <> UCase$(TagCfg(p, lab, C_COL_MAPA)) And TagCfg(p, lab, C_COL_MAPA) <> "" Then
                r = r & "  - " & RotuloItem(p, lab) & ": mapa " & TagCfg(p, lab, C_COL_MAPA) & "  ->  " & m & vbCrLf
                On Error Resume Next
                Desproteger shConfig
                shConfig.Cells(LinhaCfg(p, lab), C_COL_MAPA).Value = m
                Proteger shConfig
                On Error GoTo 0
            End If
        End If
    Next p
    CorrigirMapas = r
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

' Usado so na validacao automatica
Public Function TesteTagsInvalidas(ByVal texto As String) As String
    Dim anc As Range
    Set anc = shDadosMES.Range(ANC_U)
    mInvalidas = ""
    anc.Value = "'" & texto
    ColetarInvalidas anc, shDadosMES.Range(OUT_U)
    TesteTagsInvalidas = ListaInvalidas(1)
    mInvalidas = ""
End Function

' Grava na aba Configuracoes as variantes de tag que funcionaram; devolve a lista
Private Function CorrigirTags(ByVal lab As Long) As String
    Dim p As Long, t As String, r As String
    For p = 1 To NItens(lab)
        t = TagAtual(p, lab)
        If t <> "" And t <> TagCfg(p, lab, C_COL_TAG) And Not TagInvalida(t) Then
            r = r & "  - " & RotuloItem(p, lab) & ": " & TagCfg(p, lab, C_COL_TAG) & "  ->  " & t & vbCrLf
            On Error Resume Next
            Desproteger shConfig
            shConfig.Cells(LinhaCfg(p, lab), C_COL_TAG).Value = t
            Proteger shConfig
            On Error GoTo 0
        End If
    Next p
    CorrigirTags = r
End Function

Private Function ListaInvalidas(ByVal lab As Long) As String
    Dim p As Long, t As String, r As String
    For p = 1 To NItens(lab)
        t = TagAtual(p, lab)
        If TagInvalida(t) Then
            r = r & "  - " & RotuloItem(p, lab) & ": " & TagCfg(p, lab, C_COL_TAG)
            If t <> TagCfg(p, lab, C_COL_TAG) Then r = r & "  (e " & t & ")"
            r = r & vbCrLf
        End If
    Next p
    ListaInvalidas = r
End Function

Private Function ValorValido(ByVal p As Long, ByVal lab As Long, ByVal v As Variant) As Boolean
    Dim x As Double, vmin As String, vmax As String
    If Not ENumero(v) Then Exit Function
    x = CDbl(v)
    vmin = TagCfg(p, lab, C_COL_VMIN)
    vmax = TagCfg(p, lab, C_COL_VMAX)
    If IsNumeric(vmin) And vmin <> "" Then
        If x < CDbl(vmin) Then Exit Function
    End If
    If IsNumeric(vmax) And vmax <> "" Then
        If x > CDbl(vmax) Then Exit Function
    End If
    ValorValido = True
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

Public Function ArgumentosLab(ByVal lab As Long, ByVal ps As String, ByVal anc As Range, ByVal sai As Range) As String
    Dim lst As Variant, i As Long, p As Long, tags As String, servs As String, mapas As String
    Dim s As String, m As String, ci As String, cf As String, calc As String, intervalo As String
    lst = Split(ps, ",")
    For i = LBound(lst) To UBound(lst)
        p = CLng(lst(i))
        s = TagCfg(p, lab, C_COL_SERV)
        If s = "" Then s = "UBU"
        m = MapaAtual(p, lab)
        If tags <> "" Then
            tags = tags & ","
            servs = servs & ","
            mapas = mapas & ","
        End If
        tags = tags & TagAtual(p, lab)
        servs = servs & s
        mapas = mapas & m
    Next i
    intervalo = "2h"
    If lab = 0 And SequenciaG() Then intervalo = "1h"
    If lab = 0 Then
        ci = "Dados_MES!$B$3"
        cf = "Dados_MES!$B$4"
    ElseIf lab = 1 Then
        ci = "Dados_MES!$B$6"
        cf = "Dados_MES!$B$7"
    Else
        ci = "Dados_MES!$B$9"
        cf = "Dados_MES!$B$10"
    End If
    calc = CfgTxt("cfgCalc")
    If calc = "" Then calc = "1"
    ArgumentosLab = """time,attribute""," & Literal(tags) & "," & Literal(servs) & "," & Literal(mapas) & _
        "," & ci & "," & cf & ",""" & intervalo & """,0,"""",0,""" & calc & """,0,1560,0,0,1,1," & _
        "ADDRESS(ROW(Dados_MES!" & anc.Address(False, False) & "),COLUMN(Dados_MES!" & anc.Address(False, False) & _
        "),1,,""Dados_MES""),""Dados_MES!" & sai.Address(False, False) & """,1)"
End Function

' Inicio da formula do Aspen ("=...GetCalculationValues("), como o Excel a mostra
' Escreve a consulta na ancora. O nome da funcao do Aspen aparece de formas diferentes conforme a instalacao
' (com ou sem "_xll."); testa as formas e guarda a que o Excel reconhece (sem #NOME?) em mesPrefixo.
Private Function EscreverConsulta(ByVal anc As Range, ByVal args As String) As Boolean
    Dim cand As Variant, i As Long, v As Variant, pre As String, lista As String
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
                    If CfgTxt("mesPrefixo") <> pre Then Nm("mesPrefixo").Value = "'" & pre
                    EscreverConsulta = True
                    Exit Function
                End If
            End If
            On Error GoTo 0
        End If
    Next i
End Function

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
        If Not CelulaErroNome(anc) And CfgTxt("mesPrefixo") <> PrefixoAspen Then Nm("mesPrefixo").Value = "'" & PrefixoAspen
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
        Call PrefixoAspen(anc)
        anc.Offset(0, 7).Value = "'" & anc.Formula
        anc.Value = "'Consulta encerrada em " & Format$(Now, "dd\/mm\/yyyy hh:mm:ss") & " (fórmula em H)"
    End If
    LimparSaida sai
End Sub

Private Sub LimparSaida(ByVal sai As Range)
    Dim area As Range, c As Range
    On Error Resume Next
    Set area = sai.Resize(LIN_SAIDA, NP + 1)
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

' Inicio/fim do periodo nas celulas usadas pela formula (Dados_MES!B3:B4 Germano, B6:B7 Ubu, B9:B10 PFUF)
Private Sub GravarPeriodo(ByVal lab As Long, ByVal ini As Date, ByVal fim As Date)
    Dim ci As Range, cf As Range, modo As String
    If lab = 0 Then
        Set ci = Nm("mesIniG")
        Set cf = Nm("mesFimG")
    ElseIf lab = 1 Then
        Set ci = Nm("mesIniU")
        Set cf = Nm("mesFimU")
    Else
        Set ci = Nm("mesIniP")
        Set cf = Nm("mesFimP")
    End If
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

' ---------------------------------------------------------------- gravacao na aba Batch
' Grava os valores do laboratorio nas colunas das analises COM tag (as sem tag ficam como estao).
' Devolve quantos valores foram gravados.
Private Function GravarResultados(ByVal lab As Long, ByVal res As Variant, ByVal n As Long, ByVal ok As Boolean, _
                                  ByVal ini As Date) As Long
    Dim p As Long, s As Long, c As Range, qtd As Long, temTag As Boolean, simulado As Boolean, t As Date, minLab As Double
    If Not ok Then Exit Function
    simulado = (Left$(UCase$(CfgTxt("cfgFonte")), 3) = "SIM")
    ' horario mostrado na linha (Ubu: inicio da janela + 30 min; Germano: + cfgMinG)
    minLab = MinutosRotulo()
    If lab = 0 And ENumero(Nm("cfgMinG").Value) Then minLab = CDbl(Nm("cfgMinG").Value)
    ModoCarga True
    Desproteger shBatch
    For p = 1 To NP
        temTag = (TagAtual(p, lab) <> "" And Not TagInvalida(TagAtual(p, lab)))
        If temTag Or simulado Then
            For s = 1 To NMAX
                Set c = shBatch.Cells(B_ROW_G0 + s - 1, ColValor(p, lab))
                t = CDate(CDbl(ini) + (s - 1) * HORAS_SLOT / 24 + minLab / 1440)
                If s <= n Then
                    If ENumero(res(s, p)) And NaEscala(p, lab, t) Then
                        c.Value = CDbl(res(s, p))
                        qtd = qtd + 1
                    Else
                        c.ClearContents
                    End If
                Else
                    c.ClearContents
                End If
            Next s
        End If
    Next p
    RestaurarHorarios
    Proteger shBatch
    ModoCarga False
    gAlterado = True
    GravarResultados = qtd
End Function

' Horario t esta na escala da analise (aba Configuracoes: frequencia e 1o horario)? A analise de 4 h aparece
' no MES em duas janelas de 2 h (valor repetido); so o horario da escala vale.
Public Function NaEscala(ByVal p As Long, ByVal lab As Long, ByVal t As Date) As Boolean
    Dim fq As Variant, h1 As Variant, dif As Double, r As Double
    NaEscala = True
    fq = shConfig.Cells(C_ROW_TAG0 + 2 * (p - 1) + lab, C_COL_FREQ).Value
    h1 = shConfig.Cells(C_ROW_TAG0 + 2 * (p - 1) + lab, C_COL_HORA1).Value
    If Not ENumero(fq) Or Vazio(h1) Then Exit Function
    If Not (VarType(h1) = vbDate Or IsNumeric(h1)) Then Exit Function
    If CDbl(fq) <= HORAS_SLOT + 0.001 Then Exit Function
    dif = (CDbl(t) - CDbl(h1)) * 24
    r = dif - Int(dif / CDbl(fq)) * CDbl(fq)
    NaEscala = (r < 0.1 Or r > CDbl(fq) - 0.1)
End Function

' Horarios de Ubu voltam a ser calculados pelo inicio e pelo Nº de horarios
Public Sub RestaurarHorarios()
    Dim i As Long
    For i = 1 To NMAX
        shBatch.Cells(B_ROW_G0 + i - 1, B_COL_TU).Formula = _
            "=IF(AND(ISNUMBER(gJanU)," & i & "<=bNHor),gJanU+" & (i - 1) & "/12+cfgMinAntes/1440,"""")"
    Next i
End Sub

' ---------------------------------------------------------------- pellet feed ultra fino (composto diario)
' So o composto diario, sem as leituras do batch (Buscar no MES sem a chegada em Ubu)
Public Sub BuscarPFUF()
    BuscarSoPFUF ""
End Sub

Private Sub BuscarSoPFUF(ByVal antes As String)
    Dim msg As String
    If mOcupado Then
        Aviso "Consulta ao MES em andamento. Aguarde.", vbInformation, "MES"
        Exit Sub
    End If
    mOcupado = True
    mSemAspen = False
    Ampulheta True
    On Error GoTo Falha
    Application.StatusBar = "Consultando MES - pellet feed ultra fino..."
    msg = AtualizarPFUF()
    Application.StatusBar = False
    Ampulheta False
    mOcupado = False
    If antes <> "" Then msg = antes & vbCrLf & vbCrLf & msg
    If msg <> "" Then Aviso msg, vbInformation, "MES"
    Exit Sub
Falha:
    Application.StatusBar = False
    Ampulheta False
    mOcupado = False
    Aviso "Erro ao consultar o PFUF no MES: " & Err.Description, vbExclamation, "MES - PFUF"
End Sub

' Composto diario do dia bPfData (padrao: data do report - 1): grava Fe, SiO2, P e S.E. nos campos bPf*
' (so os que o MES devolveu; os demais ficam como estao). Devolve uma linha para a mensagem.
Public Function AtualizarPFUF() As String
    Dim dia As Double, h As Double, ini As Date, fim As Date, res As Variant, k As Long, q As Long
    Dim simulado As Boolean, ok As Boolean, falta As String, tip As String, txtDia As String, t0 As Single
    Dim seg As Long, qIni As Date, nd As Long

    Recalcular
    If Not EData(Nm("bPfData").Value) Then
        AtualizarPFUF = "PFUF: sem data do composto (aba Batch)."
        Exit Function
    End If
    dia = Int(CDbl(CDate(Nm("bPfData").Value)))
    h = 0
    h = 23.5 / 24
    If ENumero(Nm("cfgPfHora").Value) Then h = CDbl(Nm("cfgPfHora").Value) - Int(CDbl(Nm("cfgPfHora").Value))
    ini = CDate(dia)              ' inicio do dia do composto
    fim = CDate(dia + h)          ' horario em que o composto e lancado no MES
    txtDia = Format$(CDate(dia), "dd\/mm\/yyyy")
    simulado = (Left$(UCase$(CfgTxt("cfgFonte")), 3) = "SIM")
    ReDim res(1 To NPF)
    If simulado Then
        Randomize
        For k = 1 To NPF
            tip = TagCfg(k, LAB_PF, C_COL_TIP)
            If IsNumeric(tip) And tip <> "" Then res(k) = Round(CDbl(tip) * (1 + (Rnd() - 0.5) * 0.04), PfDec(k) + 1)
        Next k
        ok = True
    Else
        If ParamsLab(LAB_PF) = "" And mInvalidas = "" Then
            AtualizarPFUF = "PFUF: sem tags na aba Configurações (lançamento manual)."
            Exit Function
        End If
        If CDbl(fim) > CDbl(Now) Then
            AtualizarPFUF = "PFUF: o composto de " & txtDia & " é lançado às " & Format$(fim, "hh:mm") & _
                            " (ainda não chegou essa hora)."
            Exit Function
        End If
        mErroPF = ""
        mDiasPF = ""
        ' consulta alguns dias em volta: o resultado do dia fica com o composto; os outros so mostram
        ' em que dias o MES tem resultado (ajuda a achar a data/hora em que o composto e gravado)
        qIni = CDate(CDbl(ini) - 0.5)
        nd = PF_JANELAS
        If CDbl(qIni) + nd * HORAS_SLOT / 24 > CDbl(Now) Then nd = Int((CDbl(Now) - CDbl(qIni)) * 24 / HORAS_SLOT) + 1
        If nd > PF_JANELAS Then nd = PF_JANELAS
        t0 = Timer
        ok = ConsultarPF(qIni, CDate(CDbl(qIni) + nd * HORAS_SLOT / 24), fim, res)
        seg = Int(Timer - t0)
        If seg < 0 Then seg = seg + 86400
    End If
    If Not ok Then
        AtualizarPFUF = "PFUF (composto de " & txtDia & "): consulta falhou" & IIf(mErroPF <> "", " - " & mErroPF, "") & _
                        " (" & seg & " s). Resposta completa: aba Dados_MES, Resposta PFUF."
        Exit Function
    End If

    ModoCarga True
    Desproteger shBatch
    For k = 1 To NPF
        If ENumero(res(k)) Then
            Nm("bPf" & PfChave(k)).Value = CDbl(res(k))
            q = q + 1
        Else
            If falta <> "" Then falta = falta & ", "
            falta = falta & RotuloItem(k, LAB_PF)
        End If
    Next k
    Nm("bPfMsg").Value = "'" & IIf(simulado, "SIMULAÇÃO", "MES") & " " & Format$(Now, "dd\/mm hh:mm") & _
                         " · " & q & " de " & NPF & IIf(simulado, "", " · " & seg & " s")
    Proteger shBatch
    ModoCarga False
    gAlterado = True
    AtualizarPFUF = "PFUF (composto de " & txtDia & "): " & q & " de " & NPF & " resultados" & _
                    IIf(falta <> "", " - sem resultado: " & falta, "") & IIf(simulado, "", " (" & seg & " s)") & "."
    If Not simulado Then
        For k = 1 To NPF
            If mNotaPF(k) <> "" Then AtualizarPFUF = AtualizarPFUF & vbCrLf & "  " & RotuloItem(k, LAB_PF) & ": " & mNotaPF(k)
        Next k
    End If
    If Not simulado And q < NPF Then
        If mDiasPF <> "" Then
            AtualizarPFUF = AtualizarPFUF & vbCrLf & "Resultados do PFUF no MES (janelas de 2 h de " & _
                Format$(qIni, "dd\/mm hh:mm") & " a " & Format$(CDate(CDbl(qIni) + nd * HORAS_SLOT / 24), "dd\/mm hh:mm") & _
                "): " & mDiasPF & vbCrLf & "O composto de " & Format$(CDate(dia), "dd\/mm") & " é procurado às " & _
                Format$(fim, "hh:mm") & ". Se ele aparece em outro horário, ajuste ""Horário de lançamento"" ou " & _
                """Dias antes da data do report"" (aba Configurações)."
        Else
            AtualizarPFUF = AtualizarPFUF & vbCrLf & "O MES não devolveu nenhum resultado dessas tags de " & _
                Format$(qIni, "dd\/mm hh:mm") & " a " & Format$(CDate(CDbl(qIni) + nd * HORAS_SLOT / 24), "dd\/mm hh:mm") & _
                ": conferir tag, servidor e mapa (aba Configurações) e a Resposta PFUF na aba Dados_MES."
        End If
    End If
End Function

' Uma consulta em janelas de 2 h de ini a fim (em volta do dia do composto).
' res(1..NPF): resultado da janela que contem alvo (horario de lancamento); mDiasPF: janelas com resultado
Private Function ConsultarPF(ByVal ini As Date, ByVal fim As Date, ByVal alvo As Date, ByRef res As Variant) As Boolean
    Dim anc As Range, sai As Range, ps As String, rodada As Long, limite As Double, estado As Long, e2 As Long
    Dim restaurado As Boolean, lst As Variant, r As Long, i As Long, k As Long, v As Variant, x As Double
    Dim tj As Date, resp As String, invalidas As String, corrigidas As String, d As Variant, n As Long

    ReDim d(1 To NPF)
    res = d
    If mSemAspen Then
        mErroPF = "suplemento Aspen não reconhecido (#NOME?)"
        Exit Function
    End If
    For k = 1 To NPF
        mNotaPF(k) = ""
    Next k
    mEtapa = "pellet feed ultra fino"
    Set anc = shDadosMES.Range(ANC_P)
    Set sai = shDadosMES.Range(OUT_P)
    limite = 120
    If ENumero(Nm("cfgTimeout").Value) Then limite = CDbl(Nm("cfgTimeout").Value)
    GravarPeriodo LAB_PF, ini, fim
    For rodada = 0 To MAX_RODADAS
        ps = ParamsLab(LAB_PF)
        If ps = "" Then Exit For
        LimparSaida sai
        If Not EscreverConsulta(anc, ArgumentosLab(LAB_PF, ps, anc, sai)) Then
            Nm("mesRespP").Value = "'" & Format$(Now, "dd\/mm hh:mm:ss") & "  #NOME? (suplemento Aspen)"
            mErroPF = "suplemento Aspen não reconhecido (#NOME?)"
            anc.Value = "'Suplemento Aspen não reconhecido (#NOME?)"
            LimparSaida sai
            mSemAspen = True
            Exit Function
        End If
        ' o resultado pode vir com a hora do inicio ou do fim do dia: aceita qualquer hora ate o fim
        estado = Esperar(anc, sai, ini, limite, -1)
        If estado = 3 Then
            If ColetarInvalidas(anc, sai) = 0 Then Exit For
        ElseIf estado = 4 Then
            If ColetarMapas(anc, sai, LAB_PF) = 0 Then Exit For
        ElseIf estado = 5 Then
            Exit For
        ElseIf (estado = 0 Or estado = 2) And Not restaurado Then
            restaurado = True
            RestaurarSaida anc, sai, PF_JANELAS + 1, UBound(Split(ps, ",")) + 2
            anc.Calculate
            sai.Resize(PF_JANELAS + 1, UBound(Split(ps, ",")) + 2).Calculate
            e2 = Esperar(anc, sai, ini, limite, -1)
            If e2 = 1 Or estado = 0 Then estado = e2
            Exit For
        Else
            Exit For
        End If
    Next rodada
    If estado = 1 Then corrigidas = CorrigirTags(LAB_PF) & CorrigirMapas(LAB_PF)
    If corrigidas <> "" Then
        Aviso "Tags do pellet feed ultra fino corrigidas na aba Configurações (o MES aceitou o código em outro " & _
              "formato):" & vbCrLf & vbCrLf & corrigidas, vbInformation, "MES - tags"
    End If
    invalidas = ListaInvalidas(LAB_PF)
    If invalidas <> "" Then
        Nm("mesInvalidas").Value = "'" & Replace(invalidas, vbCrLf, " ")
        Aviso "O MES não aceitou estas tags do pellet feed ultra fino (tag ou mapa; ficam como estão):" & vbCrLf & vbCrLf & _
              invalidas & vbCrLf & "Corrigir na aba Configurações.", vbExclamation, "MES - tags inválidas"
    End If
    resp = Left$(Trim$(anc.Text) & "  |  1a célula: " & Trim$(sai.Cells(1, 1).Text) & "  " & _
                 Trim$(sai.Cells(1, 2).Text), 250)
    Nm("mesRespP").Value = "'" & Format$(Now, "dd\/mm hh:mm:ss") & "  " & resp & "  |  período " & _
        Format$(ini, "dd\/mm hh:mm") & " a " & Format$(fim, "dd\/mm hh:mm")
    If estado = 0 Or estado = 4 Or estado = 5 Then
        mErroPF = IIf(estado = 0, "o MES não respondeu", "erro do MES: " & Left$(Trim$(anc.Text), 160))
        DesligarConsulta anc, sai
        Exit Function
    End If
    If estado = 1 Then d = LerSaidaPF(sai, ini, alvo, ParamsLab(LAB_PF))
    res = d
    If mDiasPF <> "" Then Nm("mesRespP").Value = CStr(Nm("mesRespP").Value) & "  |  dias com resultado: " & mDiasPF
    DesligarConsulta anc, sai
    ConsultarPF = True
End Function

' Le a saida da consulta do PFUF (ps = analises na ordem das colunas): valor da janela de alvo; sem ele, o
' ultimo do mesmo dia do composto. Preenche mDiasPF e mNotaPF.
Private Function LerSaidaPF(ByVal sai As Range, ByVal ini As Date, ByVal alvo As Date, ByVal ps As String) As Variant
    Dim lst As Variant, r As Long, i As Long, k As Long, v As Variant, x As Double, tj As Date, n As Long, d() As Variant
    Dim noDia() As Variant, tDia() As Double, foraFaixa() As String, txtJan As String, dia As Double
    lst = Split(ps, ",")
    ReDim d(1 To NPF)
    ReDim noDia(1 To NPF)
    ReDim tDia(1 To NPF)
    ReDim foraFaixa(1 To NPF)
    dia = Int(CDbl(alvo))
    For k = 1 To NPF
        mNotaPF(k) = ""
    Next k
    For r = 1 To NMAX + 2
        If Not ComoData(sai.Cells(r, 1).Value, tj, ini) Then Exit For
        txtJan = Format$(tj, "dd\/mm hh:mm") & "-" & Format$(CDate(CDbl(tj) + HORAS_SLOT / 24), "hh:mm")
        n = 0
        For i = LBound(lst) To UBound(lst)
            k = CLng(lst(i))
            v = sai.Cells(r, 2 + i - LBound(lst)).Value
            If ComoNumero(v, x) Then
                If ValorValido(k, LAB_PF, x) Then
                    n = n + 1
                    If CDbl(tj) <= CDbl(alvo) + 1 / 1440 And CDbl(alvo) < CDbl(tj) + HORAS_SLOT / 24 - 1 / 1440 Then d(k) = x
                    ' outro horario do mesmo dia do composto: fica de reserva (vale o ultimo do dia)
                    If CDbl(tj) >= dia - 1 / 1440 And CDbl(tj) < dia + 1 - 1 / 1440 Then
                        noDia(k) = x
                        tDia(k) = CDbl(tj)
                    End If
                Else
                    If foraFaixa(k) = "" Then foraFaixa(k) = CStr(x) & " em " & txtJan
                End If
            End If
        Next i
        If n > 0 Then
            If mDiasPF <> "" Then mDiasPF = mDiasPF & ", "
            mDiasPF = mDiasPF & txtJan & " (" & n & ")"
        End If
    Next r
    For k = 1 To NPF
        If Not ENumero(d(k)) And ENumero(noDia(k)) Then
            d(k) = noDia(k)
            mNotaPF(k) = "achado às " & Format$(CDate(tDia(k)), "hh:mm") & "-" & _
                         Format$(CDate(tDia(k) + HORAS_SLOT / 24), "hh:mm") & " (fora das " & Format$(alvo, "hh:mm") & ")"
        ElseIf Not ENumero(d(k)) And foraFaixa(k) <> "" Then
            mNotaPF(k) = "valor " & foraFaixa(k) & " fora da faixa válida (" & TagCfg(k, LAB_PF, C_COL_VMIN) & _
                         " a " & TagCfg(k, LAB_PF, C_COL_VMAX) & ", aba Configurações)"
        End If
    Next k
    LerSaidaPF = d
End Function

' Usado so na validacao automatica: saida simulada do MES (S.E. as 08:00 do dia, Fe fora da faixa)
Public Function TesteLeituraPF() As String
    Dim sai As Range, d As Variant, ini As Date, k As Long, r As String
    Set sai = shDadosMES.Range(OUT_P)
    LimparSaida sai
    ini = DateSerial(2026, 10, 20) + TimeSerial(12, 0, 0)
    For k = 0 To 23
        sai.Cells(k + 1, 1).Value = CDate(CDbl(ini) + k / 12)
    Next k
    ' janela 21/10 22:00 (linha 18): Fe 999 (fora), SiO2 1.25, P 0.031; S.E. so em 21/10 08:00 (linha 11)
    sai.Cells(18, 2).Value = 999
    sai.Cells(18, 3).Value = 1.25
    sai.Cells(18, 4).Value = 0.031
    sai.Cells(11, 5).Value = 2350
    sai.Cells(23, 5).Value = 2400            ' 22/10 08:00: outro dia, nao vale
    mDiasPF = ""
    d = LerSaidaPF(sai, ini, DateSerial(2026, 10, 21) + TimeSerial(23, 30, 0), "1,2,3,4")
    For k = 1 To NPF
        r = r & IIf(ENumero(d(k)), Format$(CDbl(d(k)) * 1000, "0"), "-") & "|"
    Next k
    TesteLeituraPF = r & IIf(InStr(mNotaPF(4), "08:00") > 0, "nota SE", "") & "|" & _
                     IIf(InStr(mNotaPF(1), "fora da faixa") > 0, "nota Fe", "")
    LimparSaida sai
End Function

Private Function PfChave(ByVal k As Long) As String
    PfChave = Split(PF_CHAVES, "|")(k - 1)
End Function

Private Function PfDec(ByVal k As Long) As Long
    PfDec = CLng(Split(PF_DECIMAIS, "|")(k - 1))
End Function

' ---------------------------------------------------------------- simulacao
' Em torno do "valor tipico" (aba Configuracoes). Ubu: S.E. a cada 2 h, quimica a cada 4 h,
' <325# a cada 6 h (como nas planilhas de comparacao).
Private Function DadosSimulados(ByVal lab As Long, ByVal n As Long, ByVal ini As Date) As Variant
    Dim d() As Variant, s As Long, p As Long, tip As String, passo As Long
    ReDim d(1 To n, 1 To NP)
    Randomize
    For s = 1 To n
        For p = 1 To NP
            tip = TagCfg(p, lab, C_COL_TIP)
            passo = 1
            If lab = 1 Then
                If PChave(p) = "325" Then passo = 3
            End If
            If IsNumeric(tip) And tip <> "" And ((s - 1) Mod passo = 0) Then
                d(s, p) = Round(CDbl(tip) * (1 + (Rnd() - 0.5) * 0.06), PDec(p) + 1)
            End If
        Next p
    Next s
    DadosSimulados = d
End Function

' Usado so na validacao automatica: "novos|mapa do Fe|mapa do SiO2" depois da resposta de erro
' (Fe recusado tambem em IP_AnalogMap -> IP_DiscreteMap; SiO2 continua com IP_AnalogMap)
Public Function TesteMapas(ByVal texto As String) As String
    Dim n As Long
    mMapasRuins = "|M0650024000-0003-DDLQU>IP_ANALOGMAP|"
    mInvalidas = ""
    n = MapasRecusados(texto, True)
    TesteMapas = n & "|" & MapaAtual(1, LAB_PF) & "|" & MapaAtual(2, LAB_PF)
    mMapasRuins = ""
End Function

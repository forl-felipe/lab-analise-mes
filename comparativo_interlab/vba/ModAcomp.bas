Option Explicit

' ============================================================================
'  ModAcomp - batch em processamento (tipo da busca = Acompanhamento)
'    Buscar no MES traz so Germano e preenche a aba Acompanhamento: modelo fixo com o batch anterior (em
'    cima) e o batch em processamento (embaixo), quimico e fisico juntos. Ubu e lancado a mao a cada 2 h
'    (planilha no SharePoint); ao refazer a busca os valores de Ubu ficam.
'  Tambem: tabela "Ultimos batches em Ubu" da aba Batch (historico + acompanhamento).
' ============================================================================

' Coluna auxiliar (oculta) com o teste de tolerancia da media: a formatacao condicional so aponta para ela
' (formula de formatacao condicional e escrita no idioma do Excel; referencia simples funciona em todos)
' Bloco (padrao do Painel): faixa do batch, informacoes, RESUMO (tipo farol) + OBSERVACOES e a grade LEITURAS POR
' BI-HORARIO (uma linha por horario; para cada analise: Germano | Alimentacao 04 | Diferenca; observacao do horario).
' Colunas (1-based): B = No, C = horario (Ubu), D... = 3 por analise, depois a observacao do horario.
Private Const A_COL_N As Long = 2
Private Const A_COL_HORA As Long = 3
Private Const A_COL_G1 As Long = 4
' Colunas ocultas: numero do batch, Nº de horarios e versao do desenho (na 1a linha do bloco) e os testes de tolerancia
' (a formatacao condicional so aponta para eles: formula de formatacao condicional e escrita no idioma do Excel)
Private Const A_COL_AUX As Long = 40
Private Const VERSAO_BLOCO As String = "ACOMP3"
' Paleta Samarco (kit visual): RGB() da cor
Private Const COR_AZUL As Long = 5911296          ' #00335A Azul Samarco
Private Const COR_AZUL_MEDIO As Long = 7356416    ' #004070 Azul medio
Private Const COR_AZUL_TIT As Long = 7092992      ' #003B6C Azul de titulo
Private Const COR_AZUL_PROF As Long = 4400640     ' #002643 Azul profundo
Private Const COR_AMARELO As Long = 49407         ' #FFC000 Amarelo
Private Const COR_AZUL_CLARO As Long = 15910489   ' #59C6F2 Azul claro
Private Const COR_LARANJA As Long = 2191603       ' #F37021 Laranja
Private Const COR_CINZA_LOGO As Long = 12893877   ' #B5BEC4 Cinza da marca
Private Const COR_CINZA_CLARO As Long = 15262684  ' #DCE3E8
Private Const COR_FUNDO As Long = 16250355        ' #F3F5F7
Private Const COR_ENTRADA As Long = 14809343      ' #FFF8E1 (celulas para lancar)
Private Const COR_BRANCO As Long = 16777215
Private Const COR_DIGITAR As Long = 13431551       ' #FFF2CC amarelo claro: celulas para lancar
Private Const COR_DOURADO As Long = 37055          ' #BF9000 borda das celulas para lancar
Private Const COR_OK As Long = 5274634             ' #0A7C50
Private Const COR_OK_BG As Long = 15529187         ' #E3F4EC
Private Const COR_NOK As Long = 2631878            ' #C62828
Private Const COR_NOK_BG As Long = 15132411        ' #FBE6E6
Private Const COR_TH As Long = 15723494            ' #E6EBEF cabecalho de tabela (Painel)
Private Const COR_LINHA As Long = 14143945         ' #C9D1D7
Private Const COR_LINHA2 As Long = 15000541        ' #DDE3E4
Private Const COR_CINZA_TXT As Long = 7826267      ' #5B6B77

' planilha online aberta por esta macro (fecha depois)
Private mAbertaAqui As Workbook

Public Function ModoAcomp() As Boolean
    ModoAcomp = (InStr(1, CfgTxt("bModo"), "ACOMP", vbTextCompare) > 0)
End Function

' ---------------------------------------------------------------- blocos da aba Acompanhamento
' Modelo fixo com dois blocos: 1 = batch ANTERIOR (em cima), 2 = batch EM PROCESSAMENTO (embaixo).
' Buscar no MES so preenche: o batch da aba Batch atualiza o bloco dele; um batch novo (numero maior que o
' em processamento) passa o em processamento para anterior (com tudo o que foi lancado) e entra embaixo.
' Ubu e observacoes vem SO do que foi lancado no proprio bloco (nunca da aba Batch).
' Planilha online (SharePoint, .xlsx sem macros): mesma aba, para o fisico e o quimico lancarem juntos no navegador.
' Com ela configurada, a busca atualiza a online (que e a que vale para Ubu) e esta aba vira copia dela.

Private Function ChavesAcomp() As String
    ChavesAcomp = ACOMP_Q & "|" & ACOMP_F
End Function

Private Function NAnalises() As Long
    NAnalises = UBound(Split(ChavesAcomp(), "|")) + 1
End Function

' Linhas do bloco: faixa, informacoes, folga, titulos, cabecalho do resumo, 1 por analise, legenda, folga,
' titulo da grade, 2 de cabecalho e NMAX horarios
Public Function AlturaBloco() As Long
    AlturaBloco = 10 + NAnalises() + NMAX
End Function

' 1a linha do bloco i (1 = anterior, 2 = em processamento)
Public Function LinhaBlocoN(ByVal i As Long) As Long
    LinhaBlocoN = A_ROW0 + 1 + (i - 1) * (AlturaBloco() + 2)
End Function

Private Function LinhaResumo(ByVal r As Long, ByVal k As Long) As Long
    LinhaResumo = r + 4 + k
End Function

' Linha do horario s na grade do bloco que comeca em r
Public Function LinhaGrade(ByVal r As Long, ByVal s As Long) As Long
    LinhaGrade = r + 9 + NAnalises() + s
End Function

' Coluna de Germano da analise k (Alimentacao 04 = +1, Diferenca = +2)
Public Function ColAnalise(ByVal k As Long) As Long
    ColAnalise = A_COL_G1 + 3 * (k - 1)
End Function

Public Function ColObsHora() As Long
    ColObsHora = A_COL_G1 + 3 * NAnalises()
End Function

' Observacao geral do bloco (canto direito do resumo)
Public Function CelObs(ByVal ws As Worksheet, ByVal r As Long) As Range
    Set CelObs = ws.Cells(r + 4, A_COL_G1 + 10)
End Function

' Numero do batch do bloco i da aba Acompanhamento (Empty = bloco vazio)
Public Function IdBloco(ByVal i As Long) As Variant
    IdBloco = IdBlocoEm(shAcomp, i)
End Function

Private Function IdBlocoEm(ByVal ws As Worksheet, ByVal i As Long) As Variant
    Dim v As Variant
    IdBlocoEm = Empty
    v = ws.Cells(LinhaBlocoN(i), A_COL_AUX).Value
    If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        IdBlocoEm = CDbl(v)
    End If
End Function

' 1a linha do bloco do batch; 0 = nao esta na aba
Public Function LinhaBloco(ByVal ws As Worksheet, ByVal id As Variant) As Long
    Dim i As Long, v As Variant
    LinhaBloco = 0
    If Not ENumero(id) Then Exit Function
    For i = 1 To 2
        v = IdBlocoEm(ws, i)
        If ENumero(v) Then
            If CDbl(v) = CDbl(id) Then
                LinhaBloco = LinhaBlocoN(i)
                Exit Function
            End If
        End If
    Next i
End Function

' Le o bloco i de ws (numero, Nº de horarios, horarios, Germano, Ubu, observacoes e a linha de informacoes)
Private Sub LerBloco(ByVal ws As Worksheet, ByVal i As Long, ByRef id As Variant, ByRef n As Long, tU() As Variant, _
                     ger() As Variant, ubu() As Variant, ByRef obs As Variant, obsH() As Variant, ByRef info As Variant)
    Dim r As Long, s As Long, k As Long, rr As Long, v As Variant
    r = LinhaBlocoN(i)
    id = IdBlocoEm(ws, i)
    n = 0
    v = ws.Cells(r, A_COL_AUX + 1).Value
    If ENumero(v) Then n = CLng(v)
    If n > NMAX Then n = NMAX
    ReDim tU(1 To NMAX)
    ReDim ger(1 To NAnalises(), 1 To NMAX)
    ReDim ubu(1 To NAnalises(), 1 To NMAX)
    ReDim obsH(1 To NMAX)
    obs = Empty
    info = Empty
    If CStr(ws.Cells(r, A_COL_AUX + 2).Value) <> VERSAO_BLOCO Then Exit Sub
    For s = 1 To n
        rr = LinhaGrade(r, s)
        tU(s) = ws.Cells(rr, A_COL_HORA).Value
        For k = 1 To NAnalises()
            ger(k, s) = ws.Cells(rr, ColAnalise(k)).Value
            ubu(k, s) = ws.Cells(rr, ColAnalise(k) + 1).Value
        Next k
        obsH(s) = ws.Cells(rr, ColObsHora()).Value
    Next s
    obs = CelObs(ws, r).Value
    info = ws.Cells(r + 1, A_COL_N).Value
End Sub

' Copia os dois blocos de uma aba para outra (online -> esta, ou esta -> online)
Private Sub EspelharBlocos(ByVal origem As Worksheet, ByVal destino As Worksheet)
    Dim i As Long, id As Variant, n As Long, tU() As Variant, ger() As Variant, ubu() As Variant, obs As Variant
    Dim obsH() As Variant, info As Variant
    For i = 1 To 2
        LerBloco origem, i, id, n, tU, ger, ubu, obs, obsH, info
        DesenharBloco destino, i, id, n, tU, ger, ubu, obs, obsH, info
    Next i
End Sub

' Linha de informacoes do batch (aba Batch)
Private Function InfoBatch(ByVal n As Long) As String
    Dim t As String
    If EData(Nm("bIniU").Value) Then t = "Chegada em Ubu (UBU3): " & Format$(CDate(Nm("bIniU").Value), "dd\/mm hh:mm")
    If CfgTxt("bNivel") <> "" Then t = t & IIf(t <> "", "   " & ChrW(183) & "   ", "") & "Nível dos tanques: " & CfgTxt("bNivel")
    If EData(Nm("bIniP").Value) Then t = t & IIf(t <> "", "   " & ChrW(183) & "   ", "") & "Início no processo: " & _
        Format$(CDate(Nm("bIniP").Value), "dd\/mm hh:mm")
    If EData(Nm("bIniG").Value) Then t = t & IIf(t <> "", "   " & ChrW(183) & "   ", "") & "Germano (EB06): " & _
        Format$(CDate(Nm("bIniG").Value), "dd\/mm hh:mm") & IIf(EData(Nm("bFimG").Value), " a " & _
        Format$(CDate(Nm("bFimG").Value), "dd\/mm hh:mm"), "")
    t = t & IIf(t <> "", "   " & ChrW(183) & "   ", "") & n & " bi-horários"
    InfoBatch = t
End Function

' Monta/atualiza o bloco do batch da aba Batch; devolve uma linha para a mensagem
Public Function MontarAcompanhamento() As String
    Dim id As Variant, n As Long, s As Long, k As Long, p As Long, ps As Variant, msg As String
    Dim tU() As Variant, tG() As Variant, ger() As Variant, v As Variant, vazios As String, temG As Boolean
    Dim wbO As Workbook, wsO As Worksheet, msgO As String, r As String, info As String

    id = Nm("bId").Value
    If Not ENumero(id) Then Exit Function
    id = CDbl(id)
    GarantirModeloAcomp
    n = NMAX
    If ENumero(Nm("bNHor").Value) Then n = CLng(Nm("bNHor").Value)
    If n < 1 Then n = 1
    If n > NMAX Then n = NMAX
    ps = Split(ChavesAcomp(), "|")
    ReDim tU(1 To NMAX)
    ReDim tG(1 To NMAX)
    ReDim ger(1 To NAnalises(), 1 To NMAX)
    For s = 1 To n
        tU(s) = shBatch.Cells(B_ROW_G0 + s - 1, B_COL_TU).Value
        tG(s) = shBatch.Cells(B_ROW_G0 + s - 1, B_COL_TG).Value
        temG = False
        For k = 1 To NAnalises()
            p = PIndice(CStr(ps(k - 1)))
            v = shBatch.Cells(B_ROW_G0 + s - 1, ColValor(p, 0)).Value
            If ENumero(v) Then
                ger(k, s) = CDbl(v)
                temG = True
            End If
        Next k
        ' janelas de Germano ja passadas sem nenhum resultado: avisa (conferir no MES)
        If Not temG And EData(tG(s)) Then
            If CDbl(tG(s)) + 1 / 24 < CDbl(Now) Then
                If vazios <> "" Then vazios = vazios & ", "
                vazios = vazios & Format$(CDate(tG(s)), "dd\/mm hh:mm")
            End If
        End If
    Next s
    info = InfoBatch(n)

    ' planilha online: e la que o fisico e o quimico lancam Ubu; atualiza la e copia para esta aba
    If CfgTxt("cfgArqOnline") <> "" Then
        Set wbO = AbrirOnline(msgO)
        If Not wbO Is Nothing Then
            Set wsO = AbaOnline(wbO)
            r = AtualizarBlocos(wsO, id, n, tU, ger, info)
            If r <> "CANCELADO" Then
                wbO.Save
                EspelharBlocos wsO, shAcomp
                msgO = "Planilha online atualizada (" & wbO.Name & ")."
            End If
            FecharOnline wbO, False
        Else
            r = AtualizarBlocos(shAcomp, id, n, tU, ger, info)
        End If
    Else
        r = AtualizarBlocos(shAcomp, id, n, tU, ger, info)
    End If
    If r = "CANCELADO" Then
        MontarAcompanhamento = "Aba Acompanhamento não alterada."
        Exit Function
    End If
    RegistrarAcomp id
    AtualizarUltimos
    msg = r
    If msgO <> "" Then msg = msg & vbCrLf & msgO
    If Not EData(tU(1)) Then msg = msg & vbCrLf & _
        "Sem os horários de Ubu: informe a chegada em Ubu (UBU3) e o nível dos tanques (ou o início no processo) " & _
        "e busque de novo."
    If vazios <> "" Then msg = msg & vbCrLf & "Germano sem resultado no MES nas janelas de " & vazios & _
        " (horário de Germano) - conferir no MES."
    MontarAcompanhamento = msg
End Function

' Coloca o batch nos blocos de ws (anterior / em processamento) com Germano e horarios novos; Ubu e observacoes
' ficam os que ja estavam no bloco do batch. Devolve a mensagem ("CANCELADO" se o usuario desistiu).
Private Function AtualizarBlocos(ByVal ws As Worksheet, ByVal id As Double, ByVal n As Long, tU() As Variant, _
                                 ger() As Variant, ByVal info As String) As String
    Dim alvo As Long, s As Long, k As Long, msg As String, id1 As Variant, id2 As Variant
    Dim ubu() As Variant, obs As Variant, obsH() As Variant
    Dim vId As Variant, nV As Long, vU() As Variant, vGer() As Variant, vUbu() As Variant, vObs As Variant
    Dim vObsH() As Variant, vInfo As Variant

    GarantirModeloEm ws
    ReDim ubu(1 To NAnalises(), 1 To NMAX)
    ReDim obsH(1 To NMAX)
    id1 = IdBlocoEm(ws, 1)
    id2 = IdBlocoEm(ws, 2)
    If ENumero(id2) And CDbl(id2) = id Then
        alvo = 2
    ElseIf ENumero(id1) And CDbl(id1) = id Then
        alvo = 1
    ElseIf Not ENumero(id2) Then
        alvo = 2
    ElseIf id > CDbl(id2) Then
        ' batch novo: o em processamento passa a anterior, com tudo o que foi lancado
        LerBloco ws, 2, vId, nV, vU, vGer, vUbu, vObs, vObsH, vInfo
        DesenharBloco ws, 1, vId, nV, vU, vGer, vUbu, vObs, vObsH, vInfo
        DesenharBloco ws, 2, Empty, 0, vU, vGer, vUbu, Empty, vObsH, Empty
        msg = "Batch " & id2 & " passou para ""batch anterior"". "
        alvo = 2
    Else
        If Aviso("O batch " & id & " é anterior ao em processamento (" & id2 & ")." & vbCrLf & vbCrLf & _
                 "Colocar o " & id & " como batch anterior" & IIf(ENumero(id1), ", no lugar do " & id1, "") & "?", _
                 vbQuestion + vbYesNo, "Acompanhamento", 6) <> vbYes Then
            AtualizarBlocos = "CANCELADO"
            Exit Function
        End If
        alvo = 1
    End If
    ' Ubu e observacoes: so o que ja estava lancado no bloco deste batch (por horario de Ubu; sem horario, posicao)
    If LinhaBloco(ws, id) > 0 Then
        LerBloco ws, alvo, vId, nV, vU, vGer, vUbu, vObs, vObsH, vInfo
        For s = 1 To n
            For k = 1 To NAnalises()
                ubu(k, s) = ValorVelho(vU, vUbu, nV, k, s, tU(s))
            Next k
            obsH(s) = ObsVelha(vU, vObsH, nV, s, tU(s))
        Next s
        obs = vObs
    End If
    DesenharBloco ws, alvo, id, n, tU, ger, ubu, obs, obsH, info
    AtualizarBlocos = msg & "Aba Acompanhamento: batch " & id & IIf(alvo = 2, " em processamento", " (batch anterior)") & _
                      ", " & n & " horários."
End Function

' Desenha o bloco i de ws (id vazio = modelo em branco) com os dados passados
Private Sub DesenharBloco(ByVal ws As Worksheet, ByVal i As Long, ByVal id As Variant, ByVal n As Long, tU() As Variant, _
                          ger() As Variant, ubu() As Variant, ByVal obs As Variant, obsH() As Variant, ByVal info As Variant)
    Dim ps As Variant, na As Long, h As Long, r As Long, k As Long, s As Long, p As Long, c As Long, rr As Long
    Dim dec As Long, fmt As String, rg As Range, nomeUbu As String, chave As String, area As Range, aux As Range
    Dim cO As Long, rG1 As Long, rGn As Long, colG As String, colU As String, tol As Range, dif As Range
    Dim st As Range, nVis As Long, vazio As Boolean

    ps = Split(ChavesAcomp(), "|")
    na = NAnalises()
    h = AlturaBloco()
    r = LinhaBlocoN(i)
    cO = ColObsHora()
    vazio = Not ENumero(id)
    nomeUbu = CfgTxt("padNomeUbuAcomp")
    If nomeUbu = "" Then nomeUbu = "Ubu"
    nVis = n
    If nVis < 1 Then nVis = 12                       ' modelo em branco
    Desproteger ws
    Set area = ws.Range(ws.Cells(r - 1, 1), ws.Cells(r + h, A_COL_AUX + na + 2))
    area.UnMerge
    area.Clear
    SemFormatoCondicional area
    area.EntireRow.Hidden = False
    ws.Range(ws.Cells(r, 1), ws.Cells(r + h - 1, 1)).EntireRow.RowHeight = 17
    ' marcadores ocultos
    If Not vazio Then ws.Cells(r, A_COL_AUX).Value = CDbl(id)
    ws.Cells(r, A_COL_AUX + 1).Value = IIf(vazio, 0, n)
    ws.Cells(r, A_COL_AUX + 2).Value = VERSAO_BLOCO

    ' ---- faixa do batch (como o cabecalho do Painel) e informacoes
    ws.Rows(r).RowHeight = 30
    Set rg = ws.Range(ws.Cells(r, A_COL_N), ws.Cells(r, A_COL_G1 + 7))
    Faixa rg, COR_AZUL, COR_BRANCO, 14
    rg.Cells(1, 1).Value = IIf(vazio, "BATCH " & ChrW(8212), "BATCH " & id) & "   |   " & _
                           IIf(i = 1, "BATCH ANTERIOR", "EM PROCESSAMENTO")
    rg.IndentLevel = 1
    Set rg = ws.Range(ws.Cells(r, A_COL_G1 + 8), ws.Cells(r, cO))
    Faixa rg, COR_AZUL, COR_AMARELO, 10
    rg.HorizontalAlignment = -4152
    rg.IndentLevel = 1
    ws.Rows(r + 1).RowHeight = 20
    Set rg = ws.Range(ws.Cells(r + 1, A_COL_N), ws.Cells(r + 1, cO))
    Faixa rg, COR_AZUL_MEDIO, COR_BRANCO, 9
    If Not vazio And Not Vazio(info) Then rg.Cells(1, 1).Value = CStr(info)
    rg.IndentLevel = 1
    LinhaAmarela rg
    ws.Rows(r + 2).RowHeight = 8

    ' ---- RESUMO (tipo farol do Painel) e OBSERVACOES
    Titulo ws.Range(ws.Cells(r + 3, A_COL_N), ws.Cells(r + 3, A_COL_G1 + 8)), "RESUMO DO BATCH   (média de todas as leituras)"
    Titulo ws.Range(ws.Cells(r + 3, A_COL_G1 + 10), ws.Cells(r + 3, cO)), "OBSERVAÇÕES"
    ws.Rows(r + 4).RowHeight = 28
    Cabec ws.Range(ws.Cells(r + 4, A_COL_N), ws.Cells(r + 4, A_COL_HORA)), "ANÁLISE"
    Cabec ws.Range(ws.Cells(r + 4, A_COL_G1), ws.Cells(r + 4, A_COL_G1 + 1)), "GERMANO" & vbLf & "(MÉDIA)"
    Cabec ws.Range(ws.Cells(r + 4, A_COL_G1 + 2), ws.Cells(r + 4, A_COL_G1 + 3)), UCase$(nomeUbu) & vbLf & "(MÉDIA)"
    Cabec ws.Range(ws.Cells(r + 4, A_COL_G1 + 4), ws.Cells(r + 4, A_COL_G1 + 5)), "DIFERENÇA" & vbLf & "(UBU " & ChrW(8722) & " GERMANO)"
    Cabec ws.Range(ws.Cells(r + 4, A_COL_G1 + 6), ws.Cells(r + 4, A_COL_G1 + 6)), "TOLER." & vbLf & "(±)"
    Cabec ws.Range(ws.Cells(r + 4, A_COL_G1 + 7), ws.Cells(r + 4, A_COL_G1 + 7)), "STATUS"
    Cabec ws.Range(ws.Cells(r + 4, A_COL_G1 + 8), ws.Cells(r + 4, A_COL_G1 + 8)), "LEITURAS" & vbLf & "G / U"
    rG1 = LinhaGrade(r, 1)
    rGn = LinhaGrade(r, nVis)
    For k = 1 To na
        chave = CStr(ps(k - 1))
        p = PIndice(chave)
        dec = PDec(p)
        If dec = 0 Then fmt = "0" Else fmt = "0." & String$(dec, "0")
        rr = LinhaResumo(r, k)
        ws.Rows(rr).RowHeight = 21
        colG = ws.Range(ws.Cells(rG1, ColAnalise(k)), ws.Cells(rGn, ColAnalise(k))).Address(False, False)
        colU = ws.Range(ws.Cells(rG1, ColAnalise(k) + 1), ws.Cells(rGn, ColAnalise(k) + 1)).Address(False, False)
        Set rg = ws.Range(ws.Cells(rr, A_COL_N), ws.Cells(rr, A_COL_HORA))
        CelulaTabela rg, True
        rg.Merge
        rg.Cells(1, 1).Value = PRotulo(p)
        rg.HorizontalAlignment = -4131
        rg.IndentLevel = 1
        Set rg = ws.Range(ws.Cells(rr, A_COL_G1), ws.Cells(rr, A_COL_G1 + 1))
        CelulaTabela rg, False
        rg.Merge
        rg.Cells(1, 1).Formula = "=IF(COUNT(" & colG & ")=0,""—"",AVERAGE(" & colG & "))"
        rg.NumberFormat = fmt
        Set rg = ws.Range(ws.Cells(rr, A_COL_G1 + 2), ws.Cells(rr, A_COL_G1 + 3))
        CelulaTabela rg, False
        rg.Merge
        rg.Cells(1, 1).Formula = "=IF(COUNT(" & colU & ")=0,""—"",AVERAGE(" & colU & "))"
        rg.NumberFormat = fmt
        Set tol = ws.Cells(rr, A_COL_G1 + 6)
        CelulaTabela tol, False
        ' tolerancia (Configuracoes); na planilha online vai o numero (ela nao depende desta)
        If ws.Parent Is ThisWorkbook Then
            tol.Formula = "=padTol_" & chave
        Else
            tol.Value = CDbl(Nm("padTol_" & chave).Value)
        End If
        tol.NumberFormat = fmt
        Set dif = ws.Range(ws.Cells(rr, A_COL_G1 + 4), ws.Cells(rr, A_COL_G1 + 5))
        CelulaTabela dif, True
        dif.Merge
        dif.Cells(1, 1).Formula = "=IF(AND(ISNUMBER(" & ws.Cells(rr, A_COL_G1 + 2).Address(False, False) & "),ISNUMBER(" & _
            ws.Cells(rr, A_COL_G1).Address(False, False) & "))," & ws.Cells(rr, A_COL_G1 + 2).Address(False, False) & "-" & _
            ws.Cells(rr, A_COL_G1).Address(False, False) & ",""—"")"
        dif.NumberFormat = "+" & fmt & ";" & ChrW(8722) & fmt & ";" & fmt
        Set aux = ws.Cells(rr, A_COL_AUX)
        aux.Formula = ForaTol(dif.Cells(1, 1), dec, tol)
        MarcarForaTol dif, aux
        Set st = ws.Cells(rr, A_COL_G1 + 7)
        CelulaTabela st, True
        st.Formula = "=IF(ISNUMBER(" & dif.Cells(1, 1).Address(False, False) & "),IF(" & aux.Address(False, False) & _
                     ",""NOK"",""OK""),""—"")"
        MarcarStatus st
        Set rg = ws.Cells(rr, A_COL_G1 + 8)
        CelulaTabela rg, False
        rg.Formula = "=COUNT(" & colG & ")&"" / ""&COUNT(" & colU & ")"
        rg.Font.Color = COR_CINZA_TXT
    Next k
    ' legenda (como no Painel)
    rr = LinhaResumo(r, na) + 1
    ws.Rows(rr).RowHeight = 15
    Set rg = ws.Range(ws.Cells(rr, A_COL_N), ws.Cells(rr, A_COL_G1 + 3))
    Legenda rg, "OK " & ChrW(8212) & " DENTRO DA TOLERÂNCIA", -4131
    Set rg = ws.Range(ws.Cells(rr, A_COL_G1 + 4), ws.Cells(rr, A_COL_G1 + 8))
    Legenda rg, "NOK " & ChrW(8212) & " FORA DA TOLERÂNCIA", -4152
    ' resumo no canto da faixa: quantas analises fora
    Set st = ws.Range(ws.Cells(LinhaResumo(r, 1), A_COL_G1 + 7), ws.Cells(LinhaResumo(r, na), A_COL_G1 + 7))
    ws.Cells(r, A_COL_G1 + 8).Formula = "=IF(COUNTIF(" & st.Address(False, False) & ",""NOK"")>0,COUNTIF(" & _
        st.Address(False, False) & ",""NOK"")&IF(COUNTIF(" & st.Address(False, False) & ",""NOK"")=1,"" análise fora"","" análises fora"")&"" da tolerância"",IF(COUNTIF(" & _
        st.Address(False, False) & ",""OK"")>0,""" & ChrW(10004) & " Todas dentro da tolerância"",""""))"
    ' observacoes gerais (para comentar o que ficou fora)
    Set rg = ws.Range(ws.Cells(r + 4, A_COL_G1 + 10), ws.Cells(rr, cO))
    CelulaDigitar rg
    rg.Merge
    rg.HorizontalAlignment = -4131
    rg.VerticalAlignment = -4160
    rg.WrapText = True
    rg.Font.Bold = False
    rg.IndentLevel = 1
    If Not vazio And Not Vazio(obs) Then rg.Cells(1, 1).Value = obs
    ws.Rows(rr + 1).RowHeight = 10

    ' ---- LEITURAS POR BI-HORARIO: uma linha por horario
    rr = LinhaGrade(r, 0) - 2
    Titulo ws.Range(ws.Cells(rr, A_COL_N), ws.Cells(rr, cO)), "LEITURAS POR BI-HORÁRIO   " & ChrW(8212) & "   " & _
        ChrW(9998) & " lançar " & nomeUbu & " nas células amarelas"
    ws.Rows(rr + 1).RowHeight = 18
    ws.Rows(rr + 2).RowHeight = 26
    Set rg = ws.Range(ws.Cells(rr + 1, A_COL_N), ws.Cells(rr + 2, A_COL_N))
    Cabec rg, "Nº"
    Set rg = ws.Range(ws.Cells(rr + 1, A_COL_HORA), ws.Cells(rr + 2, A_COL_HORA))
    Cabec rg, "HORÁRIO" & vbLf & "(UBU)"
    For k = 1 To na
        p = PIndice(CStr(ps(k - 1)))
        Set rg = ws.Range(ws.Cells(rr + 1, ColAnalise(k)), ws.Cells(rr + 1, ColAnalise(k) + 2))
        Faixa rg, COR_AZUL, COR_BRANCO, 10
        rg.Cells(1, 1).Value = PRotulo(p)
        rg.HorizontalAlignment = -4108
        Cabec ws.Cells(rr + 2, ColAnalise(k)), "Germano"
        Set rg = ws.Cells(rr + 2, ColAnalise(k) + 1)
        Cabec rg, ChrW(9998) & " " & nomeUbu
        rg.Interior.Color = COR_AMARELO
        rg.Font.Color = COR_AZUL_PROF
        Cabec ws.Cells(rr + 2, ColAnalise(k) + 2), "Dif."
    Next k
    Set rg = ws.Range(ws.Cells(rr + 1, cO), ws.Cells(rr + 2, cO))
    Cabec rg, "OBSERVAÇÃO DO HORÁRIO"
    For s = 1 To NMAX
        rr = LinhaGrade(r, s)
        If s > nVis Then
            ws.Rows(rr).Hidden = True
        Else
            ws.Rows(rr).RowHeight = 19
            Set rg = ws.Cells(rr, A_COL_N)
            CelulaTabela rg, False
            rg.Value = s
            rg.Font.Color = COR_CINZA_TXT
            Set rg = ws.Cells(rr, A_COL_HORA)
            CelulaTabela rg, True
            rg.Font.Color = COR_AZUL_TIT
            rg.NumberFormat = "dd/mm hh:mm"
            If Not vazio Then
                If EData(tU(s)) Then rg.Value = CDate(tU(s))
            End If
            For k = 1 To na
                p = PIndice(CStr(ps(k - 1)))
                dec = PDec(p)
                If dec = 0 Then fmt = "0" Else fmt = "0." & String$(dec, "0")
                c = ColAnalise(k)
                Set rg = ws.Cells(rr, c)
                Celulas rg, COR_FUNDO, COR_AZUL_PROF, True
                rg.NumberFormat = fmt
                CelulaDigitar ws.Cells(rr, c + 1)
                ws.Cells(rr, c + 1).NumberFormat = fmt
                Set dif = ws.Cells(rr, c + 2)
                Celulas dif, COR_BRANCO, COR_AZUL_TIT, False
                dif.NumberFormat = "+" & fmt & ";" & ChrW(8722) & fmt & ";" & fmt
                If Not vazio Then
                    If ENumero(ger(k, s)) Then rg.Value = CDbl(ger(k, s))
                    If ENumero(ubu(k, s)) Then ws.Cells(rr, c + 1).Value = CDbl(ubu(k, s))
                End If
                dif.Formula = "=IF(AND(ISNUMBER(" & ws.Cells(rr, c + 1).Address(False, False) & "),ISNUMBER(" & _
                    rg.Address(False, False) & "))," & ws.Cells(rr, c + 1).Address(False, False) & "-" & _
                    rg.Address(False, False) & ","""")"
                Set aux = ws.Cells(rr, A_COL_AUX + k)
                aux.Formula = ForaTol(dif, dec, ws.Cells(LinhaResumo(r, k), A_COL_G1 + 6))
                MarcarForaTol dif, aux
                ' separacao entre as analises
                With ws.Cells(rr, c).Borders(7)
                    .LineStyle = 1
                    .Weight = -4138
                    .Color = COR_AZUL
                End With
            Next k
            Set rg = ws.Cells(rr, cO)
            CelulaDigitar rg
            rg.Font.Bold = False
            rg.HorizontalAlignment = -4131
            rg.IndentLevel = 1
            If Not vazio And Not Vazio(obsH(s)) Then rg.Value = obsH(s)
        End If
    Next s
    ' linha de fechamento
    With ws.Range(ws.Cells(LinhaGrade(r, nVis), A_COL_N), ws.Cells(LinhaGrade(r, nVis), cO)).Borders(9)
        .LineStyle = 1
        .Weight = -4138
        .Color = COR_AZUL
    End With
    ws.Range(ws.Columns(A_COL_AUX), ws.Columns(A_COL_AUX + na + 2)).Hidden = True
    Proteger ws
End Sub

Private Sub Faixa(ByVal rg As Range, ByVal fundo As Long, ByVal cor As Long, ByVal tam As Double)
    On Error Resume Next
    rg.Merge
    rg.Interior.Color = fundo
    rg.Font.Name = "Segoe UI"
    rg.Font.Bold = True
    rg.Font.Size = tam
    rg.Font.Color = cor
    rg.VerticalAlignment = -4108
End Sub

' Titulo de quadro (como "COMENTARIOS" no Painel)
Private Sub Titulo(ByVal rg As Range, ByVal texto As String)
    On Error Resume Next
    rg.Merge
    rg.Cells(1, 1).Value = texto
    rg.Interior.Color = COR_FUNDO
    rg.Font.Name = "Segoe UI"
    rg.Font.Bold = True
    rg.Font.Size = 8
    rg.Font.Color = COR_AZUL_PROF
    rg.IndentLevel = 1
    rg.VerticalAlignment = -4108
    rg.Borders.LineStyle = 1
    rg.Borders.Color = COR_LINHA
    rg.EntireRow.RowHeight = 20
End Sub

' Cabecalho de tabela (cinza, 7,5, com quebra)
Private Sub Cabec(ByVal rg As Range, ByVal texto As String)
    On Error Resume Next
    If rg.Cells.Count > 1 Then rg.Merge
    rg.Cells(1, 1).Value = texto
    rg.Interior.Color = COR_TH
    rg.Font.Name = "Segoe UI"
    rg.Font.Bold = True
    rg.Font.Size = 7.5
    rg.Font.Color = COR_AZUL_TIT
    rg.HorizontalAlignment = -4108
    rg.VerticalAlignment = -4108
    rg.WrapText = True
    rg.Borders.LineStyle = 1
    rg.Borders.Color = COR_LINHA
End Sub

Private Sub CelulaTabela(ByVal rg As Range, ByVal negrito As Boolean)
    On Error Resume Next
    rg.Interior.Color = COR_BRANCO
    rg.Font.Name = "Segoe UI"
    rg.Font.Size = 10
    rg.Font.Bold = negrito
    rg.Font.Color = COR_AZUL_PROF
    rg.HorizontalAlignment = -4108
    rg.VerticalAlignment = -4108
    rg.Borders.LineStyle = 1
    rg.Borders.Color = COR_LINHA2
End Sub

Private Sub Legenda(ByVal rg As Range, ByVal texto As String, ByVal alinh As Long)
    On Error Resume Next
    rg.Merge
    rg.Cells(1, 1).Value = texto
    rg.Font.Name = "Segoe UI"
    rg.Font.Size = 7
    rg.Font.Color = COR_CINZA_TXT
    rg.HorizontalAlignment = alinh
    rg.IndentLevel = 1
End Sub

' Celula para lancar: amarela, borda dourada, desbloqueada
Private Sub CelulaDigitar(ByVal rg As Range)
    Celulas rg, COR_DIGITAR, COR_AZUL_PROF, True
    On Error Resume Next
    With rg.Borders
        .LineStyle = 1
        .Color = COR_DOURADO
    End With
    rg.Locked = False
End Sub

Private Function ForaTol(ByVal c As Range, ByVal dec As Long, ByVal tol As Range) As String
    ForaTol = "=AND(ISNUMBER(" & c.Address(False, False) & "),ISNUMBER(" & tol.Address(True, True) & "),ABS(ROUND(" & _
              c.Address(False, False) & "," & dec & "))>" & tol.Address(True, True) & "+1E-9)"
End Function

' Status OK (verde) / NOK (vermelho): condicao por valor da celula (independe do idioma do Excel)
Private Sub MarcarStatus(ByVal rg As Range)
    Dim fc As Object
    On Error Resume Next
    If gTesteLO Then Exit Sub
    Set fc = rg.FormatConditions.Add(Type:=1, Operator:=3, Formula1:="=""NOK""")
    fc.Font.Color = COR_NOK
    fc.Interior.Color = COR_NOK_BG
    Set fc = rg.FormatConditions.Add(Type:=1, Operator:=3, Formula1:="=""OK""")
    fc.Font.Color = COR_OK
    fc.Interior.Color = COR_OK_BG
End Sub

' Modelo em branco nos dois blocos (aba nova, desenho antigo ou depois de limpar)
Public Sub GarantirModeloAcomp()
    GarantirModeloEm shAcomp
End Sub

Private Sub GarantirModeloEm(ByVal ws As Worksheet)
    Dim i As Long, vU() As Variant, vGer() As Variant, vUbu() As Variant, vObsH() As Variant
    ReDim vU(1 To NMAX)
    ReDim vGer(1 To NAnalises(), 1 To NMAX)
    ReDim vUbu(1 To NAnalises(), 1 To NMAX)
    ReDim vObsH(1 To NMAX)
    For i = 1 To 2
        If CStr(ws.Cells(LinhaBlocoN(i), A_COL_AUX + 2).Value) <> VERSAO_BLOCO Then
            DesenharBloco ws, i, Empty, 0, vU, vGer, vUbu, Empty, vObsH, Empty
        End If
    Next i
End Sub

Private Function ValorVelho(velhoT() As Variant, velhoU() As Variant, ByVal nVelho As Long, ByVal k As Long, _
                            ByVal s As Long, ByVal t As Variant) As Variant
    Dim j As Long
    ValorVelho = Empty
    If nVelho = 0 Then Exit Function
    If EData(t) Then
        For j = 1 To nVelho
            If EData(velhoT(j)) Then
                If Abs(CDbl(velhoT(j)) - CDbl(t)) < 1 / 1440 Then
                    ValorVelho = velhoU(k, j)
                    Exit Function
                End If
            End If
        Next j
        ' horarios mudaram (chegada corrigida): fica a posicao so se o bloco antigo nao tinha horarios
        For j = 1 To nVelho
            If EData(velhoT(j)) Then Exit Function
        Next j
    End If
    If s <= nVelho Then ValorVelho = velhoU(k, s)
End Function

' Observacao do horario: mesma regra do Ubu (pelo horario; sem horario, pela posicao)
Private Function ObsVelha(velhoT() As Variant, velhaObs() As Variant, ByVal nVelho As Long, ByVal s As Long, _
                          ByVal t As Variant) As Variant
    Dim j As Long
    ObsVelha = Empty
    If nVelho = 0 Then Exit Function
    If EData(t) Then
        For j = 1 To nVelho
            If EData(velhoT(j)) Then
                If Abs(CDbl(velhoT(j)) - CDbl(t)) < 1 / 1440 Then
                    ObsVelha = velhaObs(j)
                    Exit Function
                End If
            End If
        Next j
        For j = 1 To nVelho
            If EData(velhoT(j)) Then Exit Function
        Next j
    End If
    If s <= nVelho Then ObsVelha = velhaObs(s)
End Function

Private Sub Celulas(ByVal rg As Range, ByVal fundo As Long, ByVal cor As Long, ByVal negrito As Boolean)
    On Error Resume Next
    rg.Interior.Color = fundo
    rg.Font.Color = cor
    rg.Font.Bold = negrito
    rg.Font.Name = "Segoe UI"
    rg.Font.Size = 10
    rg.HorizontalAlignment = -4108
    rg.VerticalAlignment = -4108
    rg.Borders.LineStyle = 1
    rg.Borders.Color = COR_LINHA2
End Sub

Private Sub LinhaAmarela(ByVal rg As Range)
    On Error Resume Next
    With rg.Borders(9)
        .LineStyle = 1
        .Weight = -4138
        .Color = COR_AMARELO
    End With
End Sub

Private Sub SemFormatoCondicional(ByVal rg As Range)
    On Error Resume Next
    If gTesteLO Then Exit Sub
    rg.FormatConditions.Delete
End Sub

Private Sub MarcarForaTol(ByVal c As Range, ByVal aux As Range)
    Dim fc As Object
    On Error Resume Next
    If gTesteLO Then Exit Sub
    c.FormatConditions.Delete
    Set fc = c.FormatConditions.Add(Type:=2, Formula1:="=" & aux.Address(True, True))
    fc.Font.Color = COR_NOK
    fc.Font.Bold = True
    fc.Interior.Color = COR_NOK_BG
End Sub

' ---------------------------------------------------------------- planilha online (SharePoint)
' Arquivo .xlsx sem macros com a mesma aba Acompanhamento: o fisico e o quimico lancam juntos no navegador
' (coautoria do SharePoint). Endereco em Configuracoes (cfgArqOnline): link do arquivo no SharePoint ou caminho
' na pasta sincronizada do OneDrive.

Private Function AbrirOnline(ByRef msg As String, Optional ByVal soLeitura As Boolean = False) As Workbook
    Dim arq As String, wb As Workbook, nome As String
    arq = CfgTxt("cfgArqOnline")
    If arq = "" Then Exit Function
    nome = Mid$(arq, InStrRev(Replace(arq, "/", "\"), "\") + 1)
    If InStr(nome, "?") > 0 Then nome = Left$(nome, InStr(nome, "?") - 1)
    nome = Replace(nome, "%20", " ")                ' link do SharePoint: espacos vem como %20
    Set mAbertaAqui = Nothing
    On Error Resume Next
    Set wb = Workbooks(nome)                       ' ja aberta neste Excel: usa e nao fecha
    Err.Clear
    If wb Is Nothing Then
        Application.StatusBar = "Abrindo a planilha online..."
        Set wb = Workbooks.Open(arq, 0, soLeitura)      ' sem argumento nomeado (LibreOffice)
        Application.StatusBar = False
        If Not wb Is Nothing Then Set mAbertaAqui = wb
    End If
    If wb Is Nothing Or Err.Number <> 0 Then
        msg = "Não consegui abrir a planilha online (" & arq & "): " & Err.Description & _
              " - atualizado só nesta planilha."
        Set wb = Nothing
        Set mAbertaAqui = Nothing
    End If
    Set AbrirOnline = wb
End Function

Private Sub FecharOnline(ByVal wb As Workbook, ByVal salvar As Boolean)
    On Error Resume Next
    If salvar Then wb.Save
    If wb Is mAbertaAqui Then wb.Close SaveChanges:=False     ' so fecha se foi aberta pela macro
    Set mAbertaAqui = Nothing
    ThisWorkbook.Activate
End Sub

' Aba de acompanhamento da planilha online (cria se nao existir)
Private Function AbaOnline(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("Acompanhamento")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets(1)
        ws.Name = "Acompanhamento"
    End If
    PrepararAbaOnline ws
    Set AbaOnline = ws
End Function

' Larguras, titulo e congelamento iguais aos da aba Acompanhamento
Private Sub PrepararAbaOnline(ByVal ws As Worksheet)
    Dim c As Long
    On Error Resume Next
    Desproteger ws
    ws.Parent.Windows(1).DisplayGridlines = False
    ' mesmas larguras da aba Acompanhamento (build_interlab.abas_acomp)
    ws.Columns(1).ColumnWidth = 1.5
    ws.Columns(A_COL_N).ColumnWidth = 5
    ws.Columns(A_COL_HORA).ColumnWidth = 12
    For c = A_COL_G1 To ColObsHora() - 1
        ws.Columns(c).ColumnWidth = 8.3
    Next c
    ws.Columns(ColObsHora()).ColumnWidth = 34
    With ws.Range(ws.Cells(2, 2), ws.Cells(2, ColObsHora()))
        .UnMerge
        .ClearContents
        .Merge
        .Cells(1, 1).Value = "ACOMPANHAMENTO DO BATCH EM PROCESSAMENTO"
        .Interior.Color = COR_AZUL
        .Font.Color = COR_BRANCO
        .Font.Bold = True
        .Font.Size = 14
        .Font.Name = "Segoe UI"
        .RowHeight = 30
        .IndentLevel = 1
    End With
    LinhaAmarela ws.Range(ws.Cells(2, 2), ws.Cells(2, ColObsHora()))
    With ws.Range(ws.Cells(3, 2), ws.Cells(3, ColObsHora()))
        .UnMerge
        .ClearContents
        .Merge
        .Cells(1, 1).Value = "Germano vem do MES. " & CfgTxt("padNomeUbuAcomp") & ": lançar a cada 2 h nas células amarelas (" & _
                 ChrW(9998) & "). Em cima o batch anterior, embaixo o em processamento."
        .Interior.Color = COR_AZUL_MEDIO
        .Font.Color = COR_BRANCO
        .Font.Bold = True
        .Font.Size = 9.5
        .Font.Name = "Segoe UI"
        .IndentLevel = 1
    End With
End Sub

' Cria (ou refaz) a planilha online a partir desta aba. Se ela ja existe, traz antes o que foi lancado nela.
Public Sub GerarPlanilhaOnline()
    Dim arq As Variant, wb As Workbook, ws As Worksheet, msg As String, nova As Boolean
    arq = CfgTxt("cfgArqOnline")
    If arq <> "" Then
        If Aviso("Atualizar a planilha online com os blocos desta aba?" & vbCrLf & vbCrLf & CStr(arq) & vbCrLf & vbCrLf & _
                 "(o que já foi lançado nela é mantido)", vbQuestion + vbYesNo, "Planilha online", 6) <> vbYes Then Exit Sub
        Set wb = AbrirOnline(msg)
        If wb Is Nothing Then
            Aviso msg & vbCrLf & vbCrLf & "Confira o endereço na aba Configurações.", vbExclamation, "Planilha online"
            Exit Sub
        End If
        Set ws = AbaOnline(wb)
        ' o que foi lancado online vale: traz para ca antes
        GarantirModeloEm ws
        MesclarLancamentos ws, shAcomp
        EspelharBlocos shAcomp, ws
        Proteger ws
        FecharOnline wb, True
        Aviso "Planilha online atualizada: " & CStr(arq), vbInformation, "Planilha online"
        Exit Sub
    End If
    arq = Application.GetSaveAsFilename(InitialFileName:="Acompanhamento Germano x Ubu.xlsx", _
                                        FileFilter:="Pasta de Trabalho do Excel (*.xlsx),*.xlsx", _
                                        Title:="Salvar a planilha online (pasta do SharePoint sincronizada)")
    If VarType(arq) = vbBoolean Then Exit Sub
    If Not CriarPlanilhaOnline(CStr(arq)) Then
        Aviso "Não foi possível criar a planilha online em " & CStr(arq) & ".", vbExclamation, "Planilha online"
        Exit Sub
    End If
    GuardarEnderecoOnline CStr(arq)
    Aviso "Planilha online criada: " & CStr(arq) & vbCrLf & vbCrLf & _
          "Se a pasta não for a do SharePoint sincronizada, envie o arquivo para o SharePoint e cole o link dele em " & _
          "Configurações > Planilha de acompanhamento online. A partir daí o Buscar no MES (acompanhamento) " & _
          "atualiza a online e o físico e o químico lançam lá, juntos, pelo navegador.", vbInformation, "Planilha online"
End Sub

' Cria o arquivo .xlsx com a aba de acompanhamento (copia dos blocos desta aba)
Public Function CriarPlanilhaOnline(ByVal arq As String) As Boolean
    Dim wb As Workbook, ws As Worksheet
    On Error GoTo Falha
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    Set ws = AbaOnline(wb)
    EspelharBlocos shAcomp, ws
    ws.Activate
    ws.Cells(A_ROW0, A_COL_G1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    Proteger ws
    wb.SaveAs Filename:=arq, FileFormat:=51
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    CriarPlanilhaOnline = True
    Exit Function
Falha:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

Private Sub GuardarEnderecoOnline(ByVal arq As String)
    On Error Resume Next
    Desproteger shConfig
    Nm("cfgArqOnline").Value = "'" & arq
    Proteger shConfig
End Sub

' Ubu e observacao lancados na online passam para esta aba (bloco do mesmo batch)
Private Sub MesclarLancamentos(ByVal origem As Worksheet, ByVal destino As Worksheet)
    Dim i As Long, j As Long, id As Variant, n As Long, tU() As Variant, ger() As Variant, ubu() As Variant
    Dim obs As Variant, dId As Variant, dn As Long, dU() As Variant, dGer() As Variant, dUbu() As Variant
    Dim dObs As Variant, s As Long, k As Long, v As Variant, obsH() As Variant, info As Variant
    Dim dObsH() As Variant, dInfo As Variant
    For i = 1 To 2
        LerBloco origem, i, id, n, tU, ger, ubu, obs, obsH, info
        If ENumero(id) Then
            For j = 1 To 2
                LerBloco destino, j, dId, dn, dU, dGer, dUbu, dObs, dObsH, dInfo
                If ENumero(dId) Then
                    If CDbl(dId) = CDbl(id) Then
                        For s = 1 To dn
                            For k = 1 To NAnalises()
                                v = ValorVelho(tU, ubu, n, k, s, dU(s))
                                If ENumero(v) Then dUbu(k, s) = v
                            Next k
                            v = ObsVelha(tU, obsH, n, s, dU(s))
                            If Not Vazio(v) Then dObsH(s) = v
                        Next s
                        If Not Vazio(obs) Then dObs = obs
                        DesenharBloco destino, j, dId, dn, dU, dGer, dUbu, dObs, dObsH, dInfo
                    End If
                End If
            Next j
        End If
    Next i
End Sub

' Botao: traz para esta aba o que foi lancado na planilha online
Public Sub TrazerDoOnline()
    Dim wb As Workbook, msg As String
    If CfgTxt("cfgArqOnline") = "" Then
        Aviso "Nenhuma planilha online configurada (botão Gerar planilha online).", vbInformation, "Planilha online"
        Exit Sub
    End If
    Set wb = AbrirOnline(msg, True)
    If wb Is Nothing Then
        Aviso msg, vbExclamation, "Planilha online"
        Exit Sub
    End If
    EspelharBlocos AbaOnline(wb), shAcomp
    FecharOnline wb, False
    Aviso "Lançamentos da planilha online trazidos para a aba Acompanhamento.", vbInformation, "Planilha online"
End Sub

' ---------------------------------------------------------------- copiar (e-mail do turno)
' Copia como imagem o bloco onde esta a celula selecionada (fora dos blocos: o batch em processamento)
Public Sub CopiarAcompanhamento()
    Dim i As Long, r As Long, n As Long, area As Range, id As Variant, wb As Workbook, msg As String
    i = 2
    On Error Resume Next
    If ActiveSheet Is shAcomp Then
        If ActiveCell.Row < LinhaBlocoN(2) Then i = 1
    End If
    On Error GoTo 0
    ' com planilha online: copia com o que foi lancado la
    If CfgTxt("cfgArqOnline") <> "" Then
        Set wb = AbrirOnline(msg, True)
        If Not wb Is Nothing Then
            EspelharBlocos AbaOnline(wb), shAcomp
            FecharOnline wb, False
        End If
        shAcomp.Activate
    End If
    r = LinhaBlocoN(i)
    id = IdBloco(i)
    If Not ENumero(id) Then
        Aviso "O bloco " & IIf(i = 1, "do batch anterior", "do batch em processamento") & " está vazio " & _
              "(aba Batch: tipo da busca = Acompanhamento e Buscar no MES).", vbInformation, "Copiar acompanhamento"
        Exit Sub
    End If
    n = 12
    If ENumero(shAcomp.Cells(r, A_COL_AUX + 1).Value) Then n = CLng(shAcomp.Cells(r, A_COL_AUX + 1).Value)
    If n < 1 Then n = 1
    Set area = shAcomp.Range(shAcomp.Cells(r, A_COL_N), shAcomp.Cells(LinhaGrade(r, n), ColObsHora()))
    If CopiarAreaImagem(shAcomp, area) Then
        Aviso "Acompanhamento do batch " & id & " copiado. Cole no e-mail com Ctrl+V.", vbInformation, _
              "Copiar acompanhamento"
    Else
        Aviso "Não foi possível copiar a imagem. Use Windows + Shift + S.", vbExclamation, "Copiar acompanhamento"
    End If
End Sub

Public Sub IrAcomp()
    On Error Resume Next
    shAcomp.Activate
End Sub

' ---------------------------------------------------------------- registro (aba oculta _Acomp)
Private Sub RegistrarAcomp(ByVal id As Variant)
    Dim r As Long, v As Variant
    r = 2
    Do While Not Vazio(shAcompReg.Cells(r, 1).Value)
        v = shAcompReg.Cells(r, 1).Value
        If ENumero(v) Then
            If CDbl(v) = CDbl(id) Then Exit Do
        End If
        r = r + 1
    Loop
    shAcompReg.Cells(r, 1).Value = CDbl(id)
    CopiarData Nm("bIniU").Value, shAcompReg.Cells(r, 2)
    CopiarData Nm("bFimU").Value, shAcompReg.Cells(r, 3)
    shAcompReg.Cells(r, 4).Value = "'" & CfgTxt("bNivel")
    CopiarData Nm("bIniP").Value, shAcompReg.Cells(r, 5)
    CopiarData Nm("bIniG").Value, shAcompReg.Cells(r, 6)
    CopiarData Nm("bFimG").Value, shAcompReg.Cells(r, 7)
    If ENumero(Nm("bNHor").Value) Then shAcompReg.Cells(r, 8).Value = CLng(Nm("bNHor").Value)
    shAcompReg.Cells(r, 9).Value = Now
End Sub

Private Sub CopiarData(ByVal v As Variant, ByVal c As Range)
    If EData(v) Then
        c.Value = CDate(v)
        c.NumberFormat = "dd/mm/yyyy hh:mm"
    Else
        c.ClearContents
    End If
End Sub

Public Function LinhaRegistro(ByVal id As Variant) As Long
    Dim r As Long, v As Variant
    If Not ENumero(id) Then Exit Function
    r = 2
    Do While Not Vazio(shAcompReg.Cells(r, 1).Value)
        v = shAcompReg.Cells(r, 1).Value
        If ENumero(v) Then
            If CDbl(v) = CDbl(id) Then
                LinhaRegistro = r
                Exit Function
            End If
        End If
        r = r + 1
    Loop
End Function

' Batch digitado na aba Batch que ja esta em acompanhamento: completa os campos em branco (chegada, termino,
' nivel, Germano) com o que foi usado no acompanhamento
Public Sub PreencherDoAcomp()
    Dim r As Long
    r = LinhaRegistro(Nm("bId").Value)
    If r = 0 Then Exit Sub
    ModoCarga True
    Desproteger shBatch
    Completar "bIniU", shAcompReg.Cells(r, 2).Value
    Completar "bFimU", shAcompReg.Cells(r, 3).Value
    Completar "bNivel", shAcompReg.Cells(r, 4).Value
    Completar "bIniG", shAcompReg.Cells(r, 6).Value
    Completar "bFimG", shAcompReg.Cells(r, 7).Value
    Proteger shBatch
    ModoCarga False
End Sub

Private Sub Completar(ByVal nome As String, ByVal v As Variant)
    If Vazio(v) Then Exit Sub
    If Not Vazio(Nm(nome).Value) Then Exit Sub
    If VarType(v) = vbString Then Nm(nome).Value = "'" & v Else Nm(nome).Value = v
End Sub

' ---------------------------------------------------------------- ultimos batches em Ubu (aba Batch)
' Os N_ULT batches anteriores ao da aba Batch (sem numero: os mais recentes), do historico e do acompanhamento
Public Sub AtualizarUltimos()
    Dim ids() As Double, orig() As String, lin() As Long, n As Long, r As Long, v As Variant, atual As Variant
    Dim i As Long, j As Long, k As Long, td As Double, ts As String, tl As Long, c As Range, cheg As Variant
    Dim fim As Variant, nivel As Variant, inip As Variant, rh As Long
    On Error GoTo Fim
    atual = Nm("bId").Value
    ReDim ids(1 To 1)
    ReDim orig(1 To 1)
    ReDim lin(1 To 1)
    ' historico
    For r = PrimeiraLinhaH() To UltimaLinhaH()
        v = shHist.Cells(r, H_COL1).Value
        If ENumero(v) Then AdicionarUlt ids, orig, lin, n, CDbl(v), "H", r, atual
    Next r
    ' acompanhamento (so os que nao estao no historico)
    r = 2
    Do While Not Vazio(shAcompReg.Cells(r, 1).Value)
        v = shAcompReg.Cells(r, 1).Value
        If ENumero(v) Then AdicionarUlt ids, orig, lin, n, CDbl(v), "A", r, atual
        r = r + 1
    Loop
    ' ordem decrescente de numero
    For i = 1 To n - 1
        For j = i + 1 To n
            If ids(j) > ids(i) Then
                td = ids(i): ids(i) = ids(j): ids(j) = td
                ts = orig(i): orig(i) = orig(j): orig(j) = ts
                tl = lin(i): lin(i) = lin(j): lin(j) = tl
            End If
        Next j
    Next i
    ModoCarga True
    Desproteger shBatch
    For k = 1 To N_ULT
        rh = B_ROW_ULT + k - 1
        For Each c In shBatch.Range(shBatch.Cells(rh, U_COL_BATCH), shBatch.Cells(rh, U_COL_ORIGEM + 1)).Cells
            If c.MergeCells Then c.MergeArea.ClearContents Else c.ClearContents
        Next c
        If k <= n Then
            If orig(k) = "H" Then
                cheg = shHist.Cells(lin(k), H_COL1 + H_COL_INIU - 1).Value
                fim = shHist.Cells(lin(k), H_COL1 + H_COL_FIMU - 1).Value
                nivel = shHist.Cells(lin(k), H_COL1 + H_COL_NIVEL - 1).Value
                inip = shHist.Cells(lin(k), H_COL1 + H_COL_INIP - 1).Value
            Else
                cheg = shAcompReg.Cells(lin(k), 2).Value
                fim = shAcompReg.Cells(lin(k), 3).Value
                nivel = shAcompReg.Cells(lin(k), 4).Value
                inip = shAcompReg.Cells(lin(k), 5).Value
            End If
            shBatch.Cells(rh, U_COL_BATCH).Value = "Batch " & ids(k)
            If Not Vazio(nivel) Then shBatch.Cells(rh, U_COL_NIVEL).Value = "'" & CStr(nivel)
            If EData(cheg) Then shBatch.Cells(rh, U_COL_CHEG).Value = CDate(cheg)
            If EData(fim) Then shBatch.Cells(rh, U_COL_FIM).Value = CDate(fim)
            If EData(inip) Then shBatch.Cells(rh, U_COL_INIP).Value = CDate(inip)
            If EData(cheg) And EData(fim) Then shBatch.Cells(rh, U_COL_DUR).Value = Round((CDbl(fim) - CDbl(cheg)) * 24, 1)
            shBatch.Cells(rh, U_COL_ORIGEM).Value = IIf(orig(k) = "H", "Histórico", "Acompanhamento")
        End If
    Next k
    Proteger shBatch
    ModoCarga False
    Exit Sub
Fim:
    ModoCarga False
End Sub

Private Sub AdicionarUlt(ids() As Double, orig() As String, lin() As Long, n As Long, ByVal id As Double, _
                        ByVal o As String, ByVal r As Long, ByVal atual As Variant)
    Dim i As Long
    If ENumero(atual) Then
        If id >= CDbl(atual) Then Exit Sub
    End If
    For i = 1 To n
        If ids(i) = id Then Exit Sub        ' ja veio do historico
    Next i
    n = n + 1
    ReDim Preserve ids(1 To n)
    ReDim Preserve orig(1 To n)
    ReDim Preserve lin(1 To n)
    ids(n) = id
    orig(n) = o
    lin(n) = r
End Sub

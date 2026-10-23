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
Private Const A_COL_AUX As Long = 40
' colunas do bloco: B = analise (ou "Batch"/"Data"...), C = Germano / Alimentacao 04 / Diferenca, D... = bi-horarios,
' depois Media, Tol. e Status
Private Const A_COL_ROT As Long = 2
Private Const A_COL_SUB As Long = 3
Private Const A_COL_V1 As Long = 4
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

' planilha online aberta por esta macro (fecha depois)
Private mAbertaAqui As Workbook

Public Function ModoAcomp() As Boolean
    ModoAcomp = (InStr(1, CfgTxt("bModo"), "ACOMP", vbTextCompare) > 0)
End Function

' ---------------------------------------------------------------- blocos da aba Acompanhamento
' Modelo fixo com dois blocos: 1 = batch ANTERIOR (em cima), 2 = batch EM PROCESSAMENTO (embaixo).
' Buscar no MES so preenche: o batch da aba Batch atualiza o bloco dele; um batch novo (numero maior que o
' em processamento) passa o em processamento para anterior (com tudo o que foi lancado) e entra embaixo.
' Ubu e observacao vem SO do que foi lancado no proprio bloco (nunca da aba Batch).
' Planilha online (SharePoint, .xlsx sem macros): mesma aba, para o fisico e o quimico lancarem juntos no navegador.
' Com ela configurada, a busca atualiza a online (que e a que vale para Ubu) e esta aba vira copia dela.

Private Function ChavesAcomp() As String
    ChavesAcomp = ACOMP_Q & "|" & ACOMP_F
End Function

Private Function NAnalises() As Long
    NAnalises = UBound(Split(ChavesAcomp(), "|")) + 1
End Function

' Linhas: Batch, Data, Inicio (Ubu); por analise Germano, Alimentacao 04, Diferenca e uma linha de folga;
' no fim, Observacao
Public Function AlturaBloco() As Long
    AlturaBloco = 3 + 4 * NAnalises()
End Function

' 1a linha do bloco i (1 = anterior, 2 = em processamento); a linha de cima e o titulo do bloco
Public Function LinhaBlocoN(ByVal i As Long) As Long
    LinhaBlocoN = A_ROW0 + 1 + (i - 1) * (AlturaBloco() + 2)
End Function

' 1a linha (Germano) da analise k do bloco que comeca em r
Private Function LinhaAnalise(ByVal r As Long, ByVal k As Long) As Long
    LinhaAnalise = r + 3 + 4 * (k - 1)
End Function

Private Function LinhaObs(ByVal r As Long) As Long
    LinhaObs = r + AlturaBloco() - 1
End Function

' Numero do batch do bloco i da aba Acompanhamento (Empty = bloco vazio)
Public Function IdBloco(ByVal i As Long) As Variant
    IdBloco = IdBlocoEm(shAcomp, i)
End Function

Private Function IdBlocoEm(ByVal ws As Worksheet, ByVal i As Long) As Variant
    Dim v As Variant
    IdBlocoEm = Empty
    v = ws.Cells(LinhaBlocoN(i), A_COL_V1).Value
    ' so celula numerica ("—" do modelo em branco nao e numero)
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

' Coluna "Media" do bloco que comeca em r
Private Function ColunaMedia(ByVal ws As Worksheet, ByVal r As Long) As Long
    Dim c As Long
    ColunaMedia = 0
    For c = A_COL_V1 To A_COL_V1 + NMAX
        If CStr(ws.Cells(r, c).Value) = "Média" Then
            ColunaMedia = c
            Exit Function
        End If
    Next c
End Function

' Le o bloco i de ws: numero, horarios, horarios de Ubu, Germano e Ubu por analise e a observacao
Private Sub LerBloco(ByVal ws As Worksheet, ByVal i As Long, ByRef id As Variant, ByRef n As Long, tU() As Variant, _
                     ger() As Variant, ubu() As Variant, ByRef obs As Variant)
    Dim r As Long, s As Long, k As Long, c As Long
    r = LinhaBlocoN(i)
    id = IdBlocoEm(ws, i)
    n = ColunaMedia(ws, r) - A_COL_V1
    If n < 0 Then n = 0
    ReDim tU(1 To NMAX)
    ReDim ger(1 To NAnalises(), 1 To NMAX)
    ReDim ubu(1 To NAnalises(), 1 To NMAX)
    For s = 1 To n
        c = A_COL_V1 + s - 1
        tU(s) = ws.Cells(r + 2, c).Value
        For k = 1 To NAnalises()
            ger(k, s) = ws.Cells(LinhaAnalise(r, k), c).Value
            ubu(k, s) = ws.Cells(LinhaAnalise(r, k) + 1, c).Value
        Next k
    Next s
    obs = Empty
    If CStr(ws.Cells(LinhaObs(r), A_COL_ROT).Value) = "Observação" Then obs = ws.Cells(LinhaObs(r), A_COL_V1).Value
End Sub

' Copia os dois blocos de uma aba para outra (online -> esta, ou esta -> online)
Private Sub EspelharBlocos(ByVal origem As Worksheet, ByVal destino As Worksheet)
    Dim i As Long, id As Variant, n As Long, tU() As Variant, ger() As Variant, ubu() As Variant, obs As Variant
    For i = 1 To 2
        LerBloco origem, i, id, n, tU, ger, ubu, obs
        DesenharBloco destino, i, id, n, tU, ger, ubu, obs
    Next i
End Sub

' Monta/atualiza o bloco do batch da aba Batch; devolve uma linha para a mensagem
Public Function MontarAcompanhamento() As String
    Dim id As Variant, n As Long, s As Long, k As Long, p As Long, ps As Variant, msg As String
    Dim tU() As Variant, tG() As Variant, ger() As Variant, v As Variant, vazios As String, temG As Boolean
    Dim wbO As Workbook, wsO As Worksheet, msgO As String, r As String

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

    ' planilha online: e la que o fisico e o quimico lancam Ubu; atualiza la e copia para esta aba
    If CfgTxt("cfgArqOnline") <> "" Then
        Set wbO = AbrirOnline(msgO)
        If Not wbO Is Nothing Then
            Set wsO = AbaOnline(wbO)
            r = AtualizarBlocos(wsO, id, n, tU, ger)
            If r <> "CANCELADO" Then
                wbO.Save
                EspelharBlocos wsO, shAcomp
                msgO = "Planilha online atualizada (" & wbO.Name & ")."
            End If
            FecharOnline wbO, False
        Else
            r = AtualizarBlocos(shAcomp, id, n, tU, ger)
        End If
    Else
        r = AtualizarBlocos(shAcomp, id, n, tU, ger)
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

' Coloca o batch nos blocos de ws (anterior / em processamento) com Germano e horarios novos; Ubu e observacao
' ficam os que ja estavam no bloco do batch. Devolve a mensagem ("CANCELADO" se o usuario desistiu).
Private Function AtualizarBlocos(ByVal ws As Worksheet, ByVal id As Double, ByVal n As Long, tU() As Variant, _
                                 ger() As Variant) As String
    Dim alvo As Long, s As Long, k As Long, msg As String, id1 As Variant, id2 As Variant
    Dim ubu() As Variant, obs As Variant
    Dim vId As Variant, nV As Long, vU() As Variant, vGer() As Variant, vUbu() As Variant, vObs As Variant

    GarantirModeloEm ws
    ReDim ubu(1 To NAnalises(), 1 To NMAX)
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
        LerBloco ws, 2, vId, nV, vU, vGer, vUbu, vObs
        DesenharBloco ws, 1, vId, nV, vU, vGer, vUbu, vObs
        DesenharBloco ws, 2, Empty, 0, vU, vGer, vUbu, Empty
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
    ' Ubu e observacao: so o que ja estava lancado no bloco deste batch (por horario de Ubu; sem horario, posicao)
    If LinhaBloco(ws, id) > 0 Then
        LerBloco ws, alvo, vId, nV, vU, vGer, vUbu, vObs
        For s = 1 To n
            For k = 1 To NAnalises()
                ubu(k, s) = ValorVelho(vU, vUbu, nV, k, s, tU(s))
            Next k
        Next s
        obs = vObs
    End If
    DesenharBloco ws, alvo, id, n, tU, ger, ubu, obs
    AtualizarBlocos = msg & "Aba Acompanhamento: batch " & id & IIf(alvo = 2, " em processamento", " (batch anterior)") & _
                      ", " & n & " horários."
End Function

' Desenha o bloco i de ws (id vazio = modelo em branco) com os dados passados
Private Sub DesenharBloco(ByVal ws As Worksheet, ByVal i As Long, ByVal id As Variant, ByVal n As Long, tU() As Variant, _
                          ger() As Variant, ubu() As Variant, ByVal obs As Variant)
    Dim ps As Variant, h As Long, r As Long, k As Long, s As Long, p As Long, c As Long
    Dim cM As Long, cT As Long, cS As Long, rr As Long, dec As Long, fmt As String, rg As Range, nomeUbu As String
    Dim chave As String, cG As String, cU As String, area As Range, aux As Range, ro As Long

    ps = Split(ChavesAcomp(), "|")
    h = AlturaBloco()
    r = LinhaBlocoN(i)
    nomeUbu = CfgTxt("padNomeUbuAcomp")
    If nomeUbu = "" Then nomeUbu = "Ubu"
    If n < 1 Then n = 12                       ' modelo em branco
    Desproteger ws
    ' limpa a area do bloco (com o titulo) sem mexer nas linhas da aba
    Set area = ws.Range(ws.Cells(r - 1, 1), ws.Cells(r + h, A_COL_AUX + NMAX))
    area.UnMerge
    area.Clear
    SemFormatoCondicional area
    cM = A_COL_V1 + n
    cT = cM + 1
    cS = cM + 2
    ws.Range(ws.Cells(r, 1), ws.Cells(r + h - 1, 1)).EntireRow.RowHeight = 18
    ws.Rows(r).RowHeight = 26
    ws.Rows(r - 1).RowHeight = 22

    ' titulo do bloco
    With ws.Cells(r - 1, A_COL_ROT)
        .Value = IIf(i = 1, "BATCH ANTERIOR", "BATCH EM PROCESSAMENTO")
        .Font.Bold = True
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Color = COR_AZUL_TIT
        .VerticalAlignment = -4107
    End With
    With ws.Range(ws.Cells(r - 1, A_COL_ROT), ws.Cells(r - 1, cS)).Borders(9)
        .LineStyle = 1
        .Weight = -4138
        .Color = COR_AMARELO
    End With

    ' cabecalho: Batch / Data / Inicio aproximado no processo (horario oficial = Ubu)
    Juntar ws.Range(ws.Cells(r, A_COL_ROT), ws.Cells(r, A_COL_SUB)), "Batch", COR_AZUL, COR_BRANCO, True
    Set rg = ws.Range(ws.Cells(r, A_COL_V1), ws.Cells(r, cM - 1))
    Celulas rg, COR_AZUL, COR_BRANCO, True
    rg.Merge
    If ENumero(id) Then
        rg.Cells(1, 1).Value = CDbl(id)
    Else
        rg.Cells(1, 1).Value = "—"
    End If
    rg.Cells(1, 1).NumberFormat = """Batch ""0"
    rg.Font.Size = 14
    rg.HorizontalAlignment = -4108
    Rotulo ws.Cells(r, cM), "Média", COR_AZUL_MEDIO, COR_BRANCO, True
    Rotulo ws.Cells(r, cT), "Tol. (±)", COR_AZUL_MEDIO, COR_BRANCO, True
    Rotulo ws.Cells(r, cS), "Status", COR_AZUL_MEDIO, COR_BRANCO, True
    LinhaAmarela ws.Range(ws.Cells(r, A_COL_ROT), ws.Cells(r, cS))
    Juntar ws.Range(ws.Cells(r + 1, A_COL_ROT), ws.Cells(r + 1, A_COL_SUB)), "Data", COR_CINZA_CLARO, COR_AZUL_TIT, True
    Juntar ws.Range(ws.Cells(r + 2, A_COL_ROT), ws.Cells(r + 2, A_COL_SUB)), "Início aproximado no processo", _
           COR_CINZA_CLARO, COR_AZUL_TIT, True
    ws.Cells(r + 2, A_COL_ROT).Font.Size = 8.5
    For s = 1 To n
        c = A_COL_V1 + s - 1
        Celulas ws.Cells(r + 1, c), COR_BRANCO, COR_AZUL_TIT, False
        Celulas ws.Cells(r + 2, c), COR_BRANCO, COR_LARANJA, True
        If ENumero(id) Then
            If EData(tU(s)) Then
                ws.Cells(r + 1, c).Value = CDate(Int(CDbl(tU(s))))
                ws.Cells(r + 2, c).Value = CDate(tU(s))
            End If
        End If
        ws.Cells(r + 1, c).NumberFormat = "dd/mm"
        ws.Cells(r + 2, c).NumberFormat = "hh:mm"
    Next s
    Set rg = ws.Range(ws.Cells(r + 1, cM), ws.Cells(r + 2, cS))
    Celulas rg, COR_BRANCO, COR_AZUL_TIT, False
    rg.Merge

    ' analises (quimicas e fisicas juntas): nome | Germano (MES) / Alimentacao 04 (a mao) / Diferenca (Ubu - Germano)
    For k = 1 To NAnalises()
        chave = CStr(ps(k - 1))
        p = PIndice(chave)
        dec = PDec(p)
        If dec = 0 Then fmt = "0" Else fmt = "0." & String$(dec, "0")
        rr = LinhaAnalise(r, k)
        Set rg = ws.Range(ws.Cells(rr, A_COL_ROT), ws.Cells(rr + 2, A_COL_ROT))
        Celulas rg, COR_AZUL, COR_BRANCO, True
        rg.Merge
        rg.Cells(1, 1).Value = PRotulo(p)
        rg.Font.Size = 12
        rg.HorizontalAlignment = -4108
        rg.VerticalAlignment = -4108
        Rotulo ws.Cells(rr, A_COL_SUB), "Germano", COR_AZUL_MEDIO, COR_BRANCO, True
        Rotulo ws.Cells(rr + 1, A_COL_SUB), ChrW(9998) & " " & nomeUbu, COR_AMARELO, COR_AZUL_PROF, True
        Rotulo ws.Cells(rr + 2, A_COL_SUB), "Diferença", COR_CINZA_CLARO, COR_AZUL_TIT, False
        Set rg = ws.Range(ws.Cells(rr, cT), ws.Cells(rr + 2, cT))
        Celulas rg, COR_FUNDO, COR_AZUL_TIT, False
        rg.Merge
        ' tolerancia (Configuracoes); na planilha online vai o numero (ela nao depende desta)
        If ws.Parent Is ThisWorkbook Then
            rg.Cells(1, 1).Formula = "=padTol_" & chave
        Else
            rg.Cells(1, 1).Value = CDbl(Nm("padTol_" & chave).Value)
        End If
        rg.NumberFormat = fmt
        rg.HorizontalAlignment = -4108
        rg.VerticalAlignment = -4108
        For s = 1 To n
            c = A_COL_V1 + s - 1
            ' MES e calculo: cinza (nao digita); lancamento: amarelo com borda dourada
            Celulas ws.Cells(rr, c), COR_FUNDO, COR_AZUL_PROF, True
            Celulas ws.Cells(rr + 2, c), COR_FUNDO, COR_AZUL_TIT, False
            CelulaDigitar ws.Cells(rr + 1, c)
            If ENumero(id) Then
                If ENumero(ger(k, s)) Then ws.Cells(rr, c).Value = CDbl(ger(k, s))
                If ENumero(ubu(k, s)) Then ws.Cells(rr + 1, c).Value = CDbl(ubu(k, s))
            End If
            cG = ws.Cells(rr, c).Address(False, False)
            cU = ws.Cells(rr + 1, c).Address(False, False)
            ws.Cells(rr + 2, c).Formula = "=IF(AND(ISNUMBER(" & cU & "),ISNUMBER(" & cG & "))," & cU & "-" & cG & ","""")"
            ws.Range(ws.Cells(rr, c), ws.Cells(rr + 2, c)).NumberFormat = fmt
            ' diferenca do bi-horario fora da tolerancia (celula Tol. do bloco): vermelho
            Set aux = ws.Cells(rr + 2, A_COL_AUX + s)
            aux.Formula = ForaTol(ws.Cells(rr + 2, c), dec, ws.Cells(rr, cT))
            MarcarForaTol ws.Cells(rr + 2, c), aux
        Next s
        ' Media (de todas as leituras), Tol. e Status (media da diferenca x tolerancia)
        Set rg = ws.Range(ws.Cells(rr, A_COL_V1), ws.Cells(rr, cM - 1))
        ws.Cells(rr, cM).Formula = "=IF(COUNT(" & rg.Address(False, False) & ")=0,"""",AVERAGE(" & rg.Address(False, False) & "))"
        Set rg = ws.Range(ws.Cells(rr + 1, A_COL_V1), ws.Cells(rr + 1, cM - 1))
        ws.Cells(rr + 1, cM).Formula = "=IF(COUNT(" & rg.Address(False, False) & ")=0,"""",AVERAGE(" & rg.Address(False, False) & "))"
        cG = ws.Cells(rr, cM).Address(False, False)
        cU = ws.Cells(rr + 1, cM).Address(False, False)
        ws.Cells(rr + 2, cM).Formula = "=IF(AND(ISNUMBER(" & cU & "),ISNUMBER(" & cG & "))," & cU & "-" & cG & ","""")"
        Celulas ws.Range(ws.Cells(rr, cM), ws.Cells(rr + 1, cM)), COR_AZUL, COR_BRANCO, True
        Celulas ws.Cells(rr + 2, cM), COR_CINZA_CLARO, COR_AZUL_PROF, True
        ws.Range(ws.Cells(rr, cM), ws.Cells(rr + 2, cM)).NumberFormat = fmt
        Set aux = ws.Cells(rr + 2, A_COL_AUX)
        aux.Formula = ForaTol(ws.Cells(rr + 2, cM), dec, ws.Cells(rr, cT))
        MarcarForaTol ws.Cells(rr + 2, cM), aux
        Set rg = ws.Range(ws.Cells(rr, cS), ws.Cells(rr + 2, cS))
        Celulas rg, COR_BRANCO, COR_AZUL_TIT, True
        rg.Merge
        rg.Cells(1, 1).Formula = "=IF(" & ws.Cells(rr + 2, cM).Address(False, False) & "="""","""",IF(" & _
            aux.Address(False, False) & ",""NOK"",""OK""))"
        rg.Font.Size = 11
        rg.HorizontalAlignment = -4108
        rg.VerticalAlignment = -4108
        MarcarStatus rg
        ' folga entre as analises
        If k < NAnalises() Then ws.Rows(rr + 3).RowHeight = 6
    Next k
    ws.Range(ws.Cells(r + 1, A_COL_V1), ws.Cells(r + h - 2, cM)).HorizontalAlignment = -4108

    ' observacao (para comentar o que ficou fora da tolerancia)
    ro = LinhaObs(r)
    ws.Rows(ro).RowHeight = 36
    Juntar ws.Range(ws.Cells(ro, A_COL_ROT), ws.Cells(ro, A_COL_SUB)), "Observação", COR_AMARELO, COR_AZUL_PROF, True
    ws.Cells(ro, A_COL_ROT).Value = "Observação"
    ws.Range(ws.Cells(ro, A_COL_ROT), ws.Cells(ro, A_COL_SUB)).VerticalAlignment = -4108
    Set rg = ws.Range(ws.Cells(ro, A_COL_V1), ws.Cells(ro, cS))
    CelulaDigitar rg
    rg.Merge
    rg.HorizontalAlignment = -4131
    rg.VerticalAlignment = -4160
    rg.WrapText = True
    rg.Font.Bold = False
    rg.IndentLevel = 1
    If ENumero(id) And Not Vazio(obs) Then rg.Cells(1, 1).Value = obs
    With ws.Range(ws.Cells(ro, A_COL_ROT), ws.Cells(ro, cS)).Borders(9)
        .LineStyle = 1
        .Weight = -4138
        .Color = COR_AZUL
    End With
    ws.Range(ws.Columns(A_COL_AUX), ws.Columns(A_COL_AUX + NMAX)).Hidden = True
    Proteger ws
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

Private Sub Juntar(ByVal rg As Range, ByVal texto As String, ByVal fundo As Long, ByVal cor As Long, _
                   ByVal negrito As Boolean)
    Celulas rg, fundo, cor, negrito
    rg.Merge
    rg.Cells(1, 1).Value = texto
    rg.HorizontalAlignment = -4108
End Sub

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

' Modelo em branco nos dois blocos (aba nova, ou depois de limpar)
Public Sub GarantirModeloAcomp()
    GarantirModeloEm shAcomp
End Sub

Private Sub GarantirModeloEm(ByVal ws As Worksheet)
    Dim i As Long, vU() As Variant, vGer() As Variant, vUbu() As Variant
    ReDim vU(1 To NMAX)
    ReDim vGer(1 To NAnalises(), 1 To NMAX)
    ReDim vUbu(1 To NAnalises(), 1 To NMAX)
    For i = 1 To 2
        If CStr(ws.Cells(LinhaBlocoN(i), A_COL_ROT).Value) <> "Batch" Or _
           CStr(ws.Cells(LinhaObs(LinhaBlocoN(i)), A_COL_ROT).Value) <> "Observação" Then
            DesenharBloco ws, i, Empty, 0, vU, vGer, vUbu, Empty
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

Private Sub Rotulo(ByVal c As Range, ByVal texto As String, ByVal fundo As Long, ByVal cor As Long, ByVal negrito As Boolean)
    c.Value = texto
    Celulas c, fundo, cor, negrito
    c.HorizontalAlignment = -4108
End Sub

Private Sub Celulas(ByVal rg As Range, ByVal fundo As Long, ByVal cor As Long, ByVal negrito As Boolean)
    On Error Resume Next
    rg.Interior.Color = fundo
    rg.Font.Color = cor
    rg.Font.Bold = negrito
    rg.Font.Name = "Segoe UI"
    rg.Font.Size = 10
    rg.Borders.LineStyle = 1
    rg.Borders.Color = COR_CINZA_LOGO
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
    fc.Font.Color = RGB(198, 40, 40)
    fc.Font.Bold = True
    fc.Interior.Color = RGB(251, 230, 230)
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
    ws.Columns(1).ColumnWidth = 1.5
    ws.Columns(A_COL_ROT).ColumnWidth = 9
    ws.Columns(A_COL_SUB).ColumnWidth = 18
    For c = A_COL_V1 To A_COL_V1 + NMAX + 3
        ws.Columns(c).ColumnWidth = 8.4
    Next c
    With ws.Range(ws.Cells(2, 2), ws.Cells(2, 17))
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
    LinhaAmarela ws.Range(ws.Cells(2, 2), ws.Cells(2, 17))
    With ws.Range(ws.Cells(3, 2), ws.Cells(3, 17))
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
    ws.Cells(A_ROW0, 4).Select
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
    Dim dObs As Variant, s As Long, k As Long, v As Variant
    For i = 1 To 2
        LerBloco origem, i, id, n, tU, ger, ubu, obs
        If ENumero(id) Then
            For j = 1 To 2
                LerBloco destino, j, dId, dn, dU, dGer, dUbu, dObs
                If ENumero(dId) Then
                    If CDbl(dId) = CDbl(id) Then
                        For s = 1 To dn
                            For k = 1 To NAnalises()
                                v = ValorVelho(tU, ubu, n, k, s, dU(s))
                                If ENumero(v) Then dUbu(k, s) = v
                            Next k
                        Next s
                        If Not Vazio(obs) Then dObs = obs
                        DesenharBloco destino, j, dId, dn, dU, dGer, dUbu, dObs
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
    Dim i As Long, r As Long, cM As Long, area As Range, id As Variant, wb As Workbook, msg As String
    i = 2
    On Error Resume Next
    If ActiveSheet Is shAcomp Then
        If ActiveCell.Row < LinhaBlocoN(2) - 1 Then i = 1
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
    cM = ColunaMedia(shAcomp, r)
    Set area = shAcomp.Range(shAcomp.Cells(r - 1, A_COL_ROT), shAcomp.Cells(r + AlturaBloco() - 1, cM + 2))
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

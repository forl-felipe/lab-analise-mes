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
Private Const COR_OK As Long = 5274634             ' #0A7C50
Private Const COR_OK_BG As Long = 15529187         ' #E3F4EC
Private Const COR_NOK As Long = 2631878            ' #C62828
Private Const COR_NOK_BG As Long = 15132411        ' #FBE6E6

Public Function ModoAcomp() As Boolean
    ModoAcomp = (InStr(1, CfgTxt("bModo"), "ACOMP", vbTextCompare) > 0)
End Function

' ---------------------------------------------------------------- blocos da aba Acompanhamento
' Modelo fixo com dois blocos: 1 = batch ANTERIOR (em cima), 2 = batch EM PROCESSAMENTO (embaixo).
' Buscar no MES so preenche: o batch da aba Batch atualiza o bloco dele; um batch novo (numero maior que o
' em processamento) passa o em processamento para anterior (com tudo o que foi lancado) e entra embaixo.
' Ubu vem SO do que foi lancado no proprio bloco (nunca da aba Batch).

Private Function ChavesAcomp() As String
    ChavesAcomp = ACOMP_Q & "|" & ACOMP_F
End Function

Private Function NQuim() As Long
    NQuim = UBound(Split(ACOMP_Q, "|")) + 1
End Function

Private Function NAnalises() As Long
    NAnalises = UBound(Split(ChavesAcomp(), "|")) + 1
End Function

' Linhas: Batch, Data, Inicio (Ubu) e 3 por analise (Germano, Alimentacao 04, Diferenca)
Public Function AlturaBloco() As Long
    AlturaBloco = 3 + 3 * NAnalises()
End Function

' 1a linha do bloco i (1 = anterior, 2 = em processamento); a linha de cima e o titulo do bloco
Public Function LinhaBlocoN(ByVal i As Long) As Long
    LinhaBlocoN = A_ROW0 + 1 + (i - 1) * (AlturaBloco() + 2)
End Function

' 1a linha (Germano) da analise k do bloco que comeca em r
Private Function LinhaAnalise(ByVal r As Long, ByVal k As Long) As Long
    LinhaAnalise = r + 3 + 3 * (k - 1)
End Function

' Numero do batch do bloco i (Empty = bloco vazio)
Public Function IdBloco(ByVal i As Long) As Variant
    Dim v As Variant
    IdBloco = Empty
    v = shAcomp.Cells(LinhaBlocoN(i), A_COL_V1).Value
    ' so celula numerica ("—" do modelo em branco nao e numero)
    If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        IdBloco = CDbl(v)
    End If
End Function

' Bloco (1 ou 2) do batch; 0 = nao esta na aba
Public Function LinhaBloco(ByVal ws As Worksheet, ByVal id As Variant) As Long
    Dim i As Long
    LinhaBloco = 0
    If Not ENumero(id) Then Exit Function
    For i = 1 To 2
        If ENumero(IdBloco(i)) Then
            If CDbl(IdBloco(i)) = CDbl(id) Then
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

' Le o bloco i: numero de horarios, horarios de Ubu e de Germano, Germano e Ubu por analise
Private Sub LerBloco(ByVal i As Long, ByRef n As Long, tU() As Variant, tG() As Variant, ger() As Variant, _
                     ubu() As Variant)
    Dim r As Long, s As Long, k As Long, c As Long
    r = LinhaBlocoN(i)
    n = ColunaMedia(shAcomp, r) - A_COL_V1
    If n < 0 Then n = 0
    ReDim tU(1 To NMAX)
    ReDim tG(1 To NMAX)
    ReDim ger(1 To NAnalises(), 1 To NMAX)
    ReDim ubu(1 To NAnalises(), 1 To NMAX)
    For s = 1 To n
        c = A_COL_V1 + s - 1
        tU(s) = shAcomp.Cells(r + 2, c).Value
        For k = 1 To NAnalises()
            ger(k, s) = shAcomp.Cells(LinhaAnalise(r, k), c).Value
            ubu(k, s) = shAcomp.Cells(LinhaAnalise(r, k) + 1, c).Value
        Next k
    Next s
End Sub

' Monta/atualiza o bloco do batch da aba Batch; devolve uma linha para a mensagem
Public Function MontarAcompanhamento() As String
    Dim id As Variant, n As Long, s As Long, k As Long, p As Long, ps As Variant, alvo As Long, msg As String
    Dim tU() As Variant, tG() As Variant, ger() As Variant, ubu() As Variant, v As Variant
    Dim nV As Long, vU() As Variant, vG() As Variant, vGer() As Variant, vUbu() As Variant
    Dim id1 As Variant, id2 As Variant, vazios As String, temG As Boolean

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
    ReDim ubu(1 To NAnalises(), 1 To NMAX)
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

    id1 = IdBloco(1)
    id2 = IdBloco(2)
    If ENumero(id2) And CDbl(id2) = id Then
        alvo = 2
    ElseIf ENumero(id1) And CDbl(id1) = id Then
        alvo = 1
    ElseIf Not ENumero(id2) Then
        alvo = 2
    ElseIf id > CDbl(id2) Then
        ' batch novo: o em processamento passa a anterior, com tudo o que foi lancado
        LerBloco 2, nV, vU, vG, vGer, vUbu
        DesenharBloco 1, id2, nV, vU, vG, vGer, vUbu
        DesenharBloco 2, Empty, 0, vU, vG, vGer, vUbu
        msg = "Batch " & id2 & " passou para ""batch anterior"". "
        alvo = 2
    Else
        If Aviso("O batch " & id & " é anterior ao em processamento (" & id2 & ")." & vbCrLf & vbCrLf & _
                 "Colocar o " & id & " como batch anterior" & IIf(ENumero(id1), ", no lugar do " & id1, "") & "?", _
                 vbQuestion + vbYesNo, "Acompanhamento", 6) <> vbYes Then
            MontarAcompanhamento = "Aba Acompanhamento não alterada."
            Exit Function
        End If
        alvo = 1
    End If
    ' Ubu: so o que ja estava lancado no bloco deste batch (por horario de Ubu; sem horario, pela posicao)
    If LinhaBloco(shAcomp, id) > 0 Then
        LerBloco alvo, nV, vU, vG, vGer, vUbu
        For s = 1 To n
            For k = 1 To NAnalises()
                ubu(k, s) = ValorVelho(vU, vUbu, nV, k, s, tU(s))
            Next k
        Next s
    End If
    DesenharBloco alvo, id, n, tU, tG, ger, ubu
    RegistrarAcomp id
    AtualizarUltimos
    msg = msg & "Aba Acompanhamento: batch " & id & IIf(alvo = 2, " em processamento", " (batch anterior)") & _
          ", " & n & " horários."
    If Not EData(tU(1)) Then msg = msg & vbCrLf & _
        "Sem os horários de Ubu: informe a chegada em Ubu (UBU3) e o nível dos tanques (ou o início no processo) " & _
        "e busque de novo."
    If vazios <> "" Then msg = msg & vbCrLf & "Germano sem resultado no MES nas janelas de " & vazios & _
        " (horário de Germano) - conferir no MES."
    MontarAcompanhamento = msg
End Function

' Desenha o bloco i (id vazio = modelo em branco) com os dados passados
Private Sub DesenharBloco(ByVal i As Long, ByVal id As Variant, ByVal n As Long, tU() As Variant, tG() As Variant, _
                          ger() As Variant, ubu() As Variant)
    Dim ws As Worksheet, ps As Variant, h As Long, r As Long, k As Long, s As Long, p As Long, c As Long
    Dim cM As Long, cT As Long, cS As Long, rr As Long, dec As Long, fmt As String, rg As Range, nomeUbu As String
    Dim chave As String, cG As String, cU As String, area As Range, aux As Range

    Set ws = shAcomp
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
    ws.Range(ws.Cells(r, 1), ws.Cells(r + h - 1, 1)).EntireRow.RowHeight = 17
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
    ' canto da Media/Tol./Status: quando o bloco foi atualizado
    Set rg = ws.Range(ws.Cells(r + 1, cM), ws.Cells(r + 2, cS))
    Celulas rg, COR_FUNDO, COR_AZUL_TIT, False
    rg.Merge
    If ENumero(id) Then rg.Cells(1, 1).Value = "Atualizado em " & Format$(Now, "dd\/mm hh:mm")
    rg.Font.Size = 8
    rg.WrapText = True
    rg.HorizontalAlignment = -4108

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
        For s = 1 To n
            c = A_COL_V1 + s - 1
            Celulas ws.Cells(rr, c), COR_BRANCO, COR_AZUL_PROF, True
            Celulas ws.Cells(rr + 1, c), COR_DIGITAR, COR_AZUL_PROF, True
            Celulas ws.Cells(rr + 2, c), COR_FUNDO, COR_AZUL_TIT, False
            ' celula para lancar: amarela com borda dourada
            With ws.Cells(rr + 1, c).Borders
                .LineStyle = 1
                .Color = COR_AMARELO
            End With
            If ENumero(id) Then
                If ENumero(ger(k, s)) Then ws.Cells(rr, c).Value = CDbl(ger(k, s))
                If ENumero(ubu(k, s)) Then ws.Cells(rr + 1, c).Value = CDbl(ubu(k, s))
            End If
            ws.Cells(rr + 1, c).Locked = False
            cG = ws.Cells(rr, c).Address(False, False)
            cU = ws.Cells(rr + 1, c).Address(False, False)
            ws.Cells(rr + 2, c).Formula = "=IF(AND(ISNUMBER(" & cU & "),ISNUMBER(" & cG & "))," & cU & "-" & cG & ","""")"
            ws.Range(ws.Cells(rr, c), ws.Cells(rr + 2, c)).NumberFormat = fmt
            ' diferenca do bi-horario fora da tolerancia: vermelho
            Set aux = ws.Cells(rr + 2, A_COL_AUX + s)
            aux.Formula = ForaTol(ws.Cells(rr + 2, c), dec, chave)
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
        aux.Formula = ForaTol(ws.Cells(rr + 2, cM), dec, chave)
        MarcarForaTol ws.Cells(rr + 2, cM), aux
        Set rg = ws.Range(ws.Cells(rr, cT), ws.Cells(rr + 2, cT))
        Celulas rg, COR_FUNDO, COR_AZUL_TIT, False
        rg.Merge
        rg.Cells(1, 1).Formula = "=padTol_" & chave
        rg.NumberFormat = fmt
        rg.HorizontalAlignment = -4108
        rg.VerticalAlignment = -4108
        Set rg = ws.Range(ws.Cells(rr, cS), ws.Cells(rr + 2, cS))
        Celulas rg, COR_BRANCO, COR_AZUL_TIT, True
        rg.Merge
        rg.Cells(1, 1).Formula = "=IF(" & ws.Cells(rr + 2, cM).Address(False, False) & "="""","""",IF(" & _
            aux.Address(False, False) & ",""NOK"",""OK""))"
        rg.Font.Size = 11
        rg.HorizontalAlignment = -4108
        rg.VerticalAlignment = -4108
        MarcarStatus rg
        ' separacao entre analises
        With ws.Range(ws.Cells(rr + 2, A_COL_ROT), ws.Cells(rr + 2, cS)).Borders(9)
            .LineStyle = 1
            .Weight = -4138
            .Color = COR_AZUL
        End With
    Next k
    ws.Range(ws.Cells(r + 1, A_COL_V1), ws.Cells(r + h - 1, cM)).HorizontalAlignment = -4108
    ws.Range(ws.Columns(A_COL_AUX), ws.Columns(A_COL_AUX + NMAX)).Hidden = True
    Proteger ws
End Sub

Private Function ForaTol(ByVal c As Range, ByVal dec As Long, ByVal chave As String) As String
    ForaTol = "=AND(ISNUMBER(" & c.Address(False, False) & "),ABS(ROUND(" & c.Address(False, False) & "," & dec & _
              "))>padTol_" & chave & "+1E-9)"
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
    Dim i As Long, vU() As Variant, vG() As Variant, vGer() As Variant, vUbu() As Variant
    ReDim vU(1 To NMAX)
    ReDim vG(1 To NMAX)
    ReDim vGer(1 To NAnalises(), 1 To NMAX)
    ReDim vUbu(1 To NAnalises(), 1 To NMAX)
    For i = 1 To 2
        If CStr(shAcomp.Cells(LinhaBlocoN(i), A_COL_ROT).Value) <> "Batch" Then
            DesenharBloco i, Empty, 0, vU, vG, vGer, vUbu
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

Private Sub Faixa(ByVal ws As Worksheet, ByVal r As Long, ByVal cM As Long, ByVal texto As String)
    Dim rg As Range
    Set rg = ws.Range(ws.Cells(r, A_COL_ROT), ws.Cells(r, cM))
    Celulas rg, COR_AMARELO, COR_AZUL_PROF, True
    rg.Merge
    rg.Cells(1, 1).Value = texto
    rg.Font.Size = 8.5
    rg.HorizontalAlignment = -4131
    rg.IndentLevel = 1
    ws.Rows(r).RowHeight = 14
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

' ---------------------------------------------------------------- copiar (e-mail do turno)
' Copia como imagem o bloco onde esta a celula selecionada (fora dos blocos: o batch em processamento)
Public Sub CopiarAcompanhamento()
    Dim i As Long, r As Long, cM As Long, area As Range, id As Variant
    i = 2
    On Error Resume Next
    If ActiveSheet Is shAcomp Then
        If ActiveCell.Row < LinhaBlocoN(2) - 1 Then i = 1
    End If
    On Error GoTo 0
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

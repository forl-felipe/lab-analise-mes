Option Explicit

' ============================================================================
'  ModAcomp - batch em processamento (tipo da busca = Acompanhamento)
'    Buscar no MES traz so Germano e monta o bloco do batch na aba Acompanhamento (quimico e fisico
'    juntos, todos os batches nela, o mais recente em cima). Ubu e preenchido
'    a mao a cada 2 h (planilha no SharePoint); ao refazer a busca os valores de Ubu ficam.
'  Tambem: tabela "Ultimos batches em Ubu" da aba Batch (historico + acompanhamento).
' ============================================================================

' Coluna auxiliar (oculta) com o teste de tolerancia da media: a formatacao condicional so aponta para ela
' (formula de formatacao condicional e escrita no idioma do Excel; referencia simples funciona em todos)
Private Const A_COL_AUX As Long = 30
Private Const A_COL_ROT As Long = 2
Private Const A_COL_V1 As Long = 3
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

Public Function ModoAcomp() As Boolean
    ModoAcomp = (InStr(1, CfgTxt("bModo"), "ACOMP", vbTextCompare) > 0)
End Function

' ---------------------------------------------------------------- bloco do batch na aba Acompanhamento
' Monta/atualiza o bloco do batch da aba Batch; devolve uma linha para a mensagem
Public Function MontarAcompanhamento() As String
    Dim id As Variant, n As Long, s As Long
    Dim tU() As Variant
    id = Nm("bId").Value
    If Not ENumero(id) Then Exit Function
    n = NMAX
    If ENumero(Nm("bNHor").Value) Then n = CLng(Nm("bNHor").Value)
    If n < 1 Then n = 1
    If n > NMAX Then n = NMAX
    ReDim tU(1 To n)
    For s = 1 To n
        tU(s) = shBatch.Cells(B_ROW_G0 + s - 1, B_COL_TU).Value
    Next s
    MontarBloco shAcomp, id, n, tU
    RegistrarAcomp id
    AtualizarUltimos
    MontarAcompanhamento = "Aba Acompanhamento atualizada (batch " & id & ", " & n & " horários)."
    If Not EData(tU(1)) Then MontarAcompanhamento = MontarAcompanhamento & vbCrLf & _
        "Sem os horários de Ubu: informe a chegada em Ubu (UBU3) e o nível dos tanques (ou o início no processo) " & _
        "e busque de novo."
End Function

' Analises do bloco: quimicas e fisicas (nesta ordem)
Private Function ChavesAcomp() As String
    ChavesAcomp = ACOMP_Q & "|" & ACOMP_F
End Function

Private Function NQuim() As Long
    NQuim = UBound(Split(ACOMP_Q, "|")) + 1
End Function

' Linhas: Batch, Data, Inicio, faixa QUIMICO, 3 por analise quimica, faixa FISICO, 3 por analise fisica
Public Function AlturaBloco() As Long
    AlturaBloco = 3 + 2 + 3 * (UBound(Split(ChavesAcomp(), "|")) + 1)
End Function

' 1a linha (Germano) da analise k do bloco que comeca em r
Private Function LinhaAnalise(ByVal r As Long, ByVal k As Long) As Long
    LinhaAnalise = r + 4 + 3 * (k - 1) + IIf(k > NQuim(), 1, 0)
End Function

' Linha do bloco do batch (0 = nao existe)
Public Function LinhaBloco(ByVal ws As Worksheet, ByVal id As Variant) As Long
    Dim r As Long, ult As Long, v As Variant
    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = A_ROW0 To ult
        If CStr(ws.Cells(r, A_COL_ROT).Value) = "Batch" Then
            v = ws.Cells(r, A_COL_V1).Value
            If ENumero(v) Then
                If CDbl(v) = CDbl(id) Then
                    LinhaBloco = r
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

' Coluna "Media" do bloco que comeca em r
Private Function ColunaMedia(ByVal ws As Worksheet, ByVal r As Long) As Long
    Dim c As Long
    For c = A_COL_V1 To A_COL_V1 + NMAX
        If CStr(ws.Cells(r, c).Value) = "Média" Then
            ColunaMedia = c
            Exit Function
        End If
    Next c
End Function

Private Sub MontarBloco(ByVal ws As Worksheet, ByVal id As Variant, ByVal n As Long, tU() As Variant)
    Dim ps As Variant, np As Long, h As Long, r As Long, k As Long, s As Long, p As Long, nVelho As Long
    Dim velhoT() As Variant, velhoU() As Variant, v As Variant, c As Long, cM As Long, rr As Long, dec As Long
    Dim fmt As String, rg As Range, nomeUbu As String, chave As String, cG As String, cU As String

    ps = Split(ChavesAcomp(), "|")
    np = UBound(ps) + 1
    h = AlturaBloco()
    nomeUbu = CfgTxt("padNomeUbuAcomp")
    If nomeUbu = "" Then nomeUbu = "Ubu"
    Desproteger ws
    ReDim velhoU(1 To np, 1 To NMAX)
    ReDim velhoT(1 To NMAX)
    r = LinhaBloco(ws, id)
    If r > 0 Then
        ' guarda o que ja foi lancado de Ubu (por horario; sem horario, pela posicao) e refaz o bloco
        nVelho = ColunaMedia(ws, r) - A_COL_V1
        If nVelho < 0 Then nVelho = 0
        For s = 1 To nVelho
            velhoT(s) = ws.Cells(r + 2, A_COL_V1 + s - 1).Value
            For k = 1 To np
                velhoU(k, s) = ws.Cells(LinhaAnalise(r, k) + 1, A_COL_V1 + s - 1).Value
            Next k
        Next s
        ws.Rows(r & ":" & (r + h)).Delete
    Else
        r = A_ROW0
        If Left$(CStr(ws.Cells(r, A_COL_ROT).Value), 6) = "Nenhum" Then ws.Cells(r, A_COL_ROT).ClearContents
    End If
    ws.Rows(r & ":" & (r + h)).Insert
    ws.Range(ws.Cells(r, 1), ws.Cells(r + h, A_COL_AUX)).Clear
    cM = A_COL_V1 + n
    ws.Rows(r & ":" & (r + h)).RowHeight = 16.5
    ws.Rows(r).RowHeight = 24

    ' cabecalho: Batch (Azul Samarco, linha amarela) / Data / Inicio aproximado no processo
    Rotulo ws.Cells(r, A_COL_ROT), "Batch", COR_AZUL, COR_BRANCO, True
    Set rg = ws.Range(ws.Cells(r, A_COL_V1), ws.Cells(r, cM - 1))
    Celulas rg, COR_AZUL, COR_BRANCO, True
    rg.Merge
    rg.Cells(1, 1).Value = CDbl(id)
    rg.Cells(1, 1).NumberFormat = """Batch ""0"
    rg.Font.Size = 13
    rg.HorizontalAlignment = -4108
    Rotulo ws.Cells(r, cM), "Média", COR_AZUL_MEDIO, COR_BRANCO, True
    LinhaAmarela ws.Range(ws.Cells(r, A_COL_ROT), ws.Cells(r, cM))
    Rotulo ws.Cells(r + 1, A_COL_ROT), "Data", COR_CINZA_CLARO, COR_AZUL_TIT, True
    Rotulo ws.Cells(r + 2, A_COL_ROT), "Início aproximado no processo", COR_CINZA_CLARO, COR_AZUL_TIT, True
    ws.Cells(r + 2, A_COL_ROT).Font.Size = 8.5
    For s = 1 To n
        c = A_COL_V1 + s - 1
        Celulas ws.Cells(r + 1, c), COR_BRANCO, COR_AZUL_TIT, False
        Celulas ws.Cells(r + 2, c), COR_BRANCO, COR_LARANJA, True
        If EData(tU(s)) Then
            ws.Cells(r + 1, c).Value = CDate(Int(CDbl(tU(s))))
            ws.Cells(r + 2, c).Value = CDate(tU(s))
        End If
        ws.Cells(r + 1, c).NumberFormat = "dd/mm/yy"
        ws.Cells(r + 2, c).NumberFormat = "hh:mm"
    Next s
    Celulas ws.Range(ws.Cells(r + 1, cM), ws.Cells(r + 2, cM)), COR_BRANCO, COR_AZUL_TIT, False

    ' faixas QUIMICO / FISICO
    Faixa ws, r + 3, cM, "ANÁLISES QUÍMICAS"
    Faixa ws, LinhaAnalise(r, NQuim() + 1) - 1, cM, "ANÁLISES FÍSICAS"

    ' analises: Germano (MES), Ubu (a mao), Diferenca (Ubu - Germano); Media de todas as leituras
    For k = 1 To np
        chave = CStr(ps(k - 1))
        p = PIndice(chave)
        dec = PDec(p)
        If dec = 0 Then fmt = "0" Else fmt = "0." & String$(dec, "0")
        rr = LinhaAnalise(r, k)
        Rotulo ws.Cells(rr, A_COL_ROT), PRotulo(p) & " Germano", COR_AZUL_MEDIO, COR_BRANCO, True
        Rotulo ws.Cells(rr + 1, A_COL_ROT), PRotulo(p) & " " & nomeUbu, COR_AZUL_CLARO, COR_AZUL_PROF, True
        Rotulo ws.Cells(rr + 2, A_COL_ROT), "Diferença", COR_CINZA_CLARO, COR_AZUL_TIT, True
        For s = 1 To n
            c = A_COL_V1 + s - 1
            Celulas ws.Cells(rr, c), COR_BRANCO, COR_AZUL_PROF, True
            Celulas ws.Cells(rr + 1, c), COR_ENTRADA, COR_AZUL_PROF, False
            Celulas ws.Cells(rr + 2, c), COR_FUNDO, COR_AZUL_TIT, False
            v = shBatch.Cells(B_ROW_G0 + s - 1, ColValor(p, 0)).Value
            If ENumero(v) Then ws.Cells(rr, c).Value = CDbl(v)
            ' Ubu: o que ja estava no bloco; senao, o que houver na aba Batch
            v = ValorVelho(velhoT, velhoU, nVelho, k, s, tU(s))
            If Not ENumero(v) Then v = shBatch.Cells(B_ROW_G0 + s - 1, ColValor(p, 1)).Value
            If ENumero(v) Then ws.Cells(rr + 1, c).Value = CDbl(v)
            ws.Cells(rr + 1, c).Locked = False
            cG = ws.Cells(rr, c).Address(False, False)
            cU = ws.Cells(rr + 1, c).Address(False, False)
            ws.Cells(rr + 2, c).Formula = "=IF(AND(ISNUMBER(" & cU & "),ISNUMBER(" & cG & "))," & cU & "-" & cG & ","""")"
            ws.Range(ws.Cells(rr, c), ws.Cells(rr + 2, c)).NumberFormat = fmt
        Next s
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
        ' media fora da tolerancia (Configuracoes) em vermelho
        ws.Cells(rr + 2, A_COL_AUX).Formula = "=AND(ISNUMBER(" & ws.Cells(rr + 2, cM).Address(False, False) & "),ABS(ROUND(" & _
            ws.Cells(rr + 2, cM).Address(False, False) & "," & dec & "))>padTol_" & chave & "+1E-9)"
        MarcarForaTol ws.Cells(rr + 2, cM), ws.Cells(rr + 2, A_COL_AUX)
    Next k
    ws.Range(ws.Cells(r + 1, A_COL_V1), ws.Cells(r + h - 1, cM)).HorizontalAlignment = -4108
    ws.Columns(A_COL_AUX).Hidden = True
    Proteger ws
End Sub

Private Function ValorVelho(velhoT() As Variant, velhoU() As Variant, ByVal nVelho As Long, ByVal k As Long, _
                            ByVal s As Long, ByVal t As Variant) As Variant
    Dim j As Long
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
        ' horarios mudaram (chegada corrigida): fica a posicao se o bloco antigo nao tinha horarios
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
' Copia como imagem o bloco do batch onde esta a celula selecionada (ou o mais recente, no topo)
Public Sub CopiarAcompanhamento()
    Dim r As Long, rr As Long, cM As Long, area As Range, id As Variant
    On Error Resume Next
    If ActiveSheet Is shAcomp Then
        For rr = ActiveCell.Row To A_ROW0 Step -1
            If CStr(shAcomp.Cells(rr, A_COL_ROT).Value) = "Batch" Then
                r = rr
                Exit For
            End If
        Next rr
        If r > 0 Then
            If ActiveCell.Row > r + AlturaBloco() - 1 Then r = 0
        End If
    End If
    On Error GoTo 0
    If r = 0 Then r = A_ROW0
    If CStr(shAcomp.Cells(r, A_COL_ROT).Value) <> "Batch" Then
        Aviso "Nenhum batch em acompanhamento ainda (aba Batch: tipo da busca = Acompanhamento e Buscar no MES).", _
              vbInformation, "Copiar acompanhamento"
        Exit Sub
    End If
    id = shAcomp.Cells(r, A_COL_V1).Value
    cM = ColunaMedia(shAcomp, r)
    Set area = shAcomp.Range(shAcomp.Cells(r, A_COL_ROT), shAcomp.Cells(r + AlturaBloco() - 1, cM))
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

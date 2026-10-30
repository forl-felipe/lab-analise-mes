Option Explicit

' ============================================================================
'  ModDia - aba Dia e aba Historico
'  Buscar no MES -> grava o dia fechado no historico (um dia por linha) e atualiza o painel.
'  Valor corrigido/lancado a mao na aba Dia -> Salvar no historico. Tambem da para lancar direto no Historico.
' ============================================================================

' ---------------------------------------------------------------- botoes da aba Dia
Public Sub BuscarMES()
    Dim d As Date, vals As Variant, nres As Variant, sit As Variant, ok As Boolean, msg As String, notas As String
    Dim salvo As Boolean
    If Not EData(Nm("dData").Value) Then
        Aviso "Informe a data do report (célula amarela).", vbExclamation, "Buscar no MES"
        Exit Sub
    End If
    d = CDate(DiaDe(Nm("dData").Value))
    Ampulheta True
    ok = BuscarDiaMES(d, vals, nres, sit)
    ModoCarga True
    Desproteger shDia
    PreencherDia vals, nres, sit
    If Not ENumero(Nm("d_fator").Value) Then Nm("d_fator").Value = FatorPadrao(d)
    msg = ResumoBusca(d, vals, sit)
    notas = gNotasMES
    Nm("dStatus").Value = "'" & Format$(Now, "dd\/mm hh:mm") & "  " & Replace(msg & IIf(notas <> "", vbCrLf & notas, ""), _
                          vbCrLf, "  ·  ")
    Proteger shDia
    ModoCarga False
    ' dia fechado: ja grava no historico e leva o painel para este dia (dia em andamento: so mostra)
    salvo = False
    If CDbl(d) < Int(CDbl(Now)) And ContarValores(ValoresDia()) > 0 Then
        salvo = (GravarHist(d, ValoresDia(), Nm("d_fator").Value, Nm("d_obs").Value, False) > 0)
        If salvo Then
            OrdenarHist
            AvancarPainel d
            AtualizarPainel
        End If
    End If
    Ampulheta False
    Recalcular
    ' a caixa de mensagem do Excel corta textos longos: o detalhe completo fica em "Última busca"
    If Len(notas) > 520 Then notas = Left$(notas, 500) & "... (detalhes em Última busca, no fim da aba Dia)"
    If notas <> "" Then msg = msg & vbCrLf & vbCrLf & notas
    If salvo Then
        msg = msg & vbCrLf & vbCrLf & "Dia salvo no histórico e no painel. Se corrigir ou lançar algum valor à mão, " & _
              "clique em Salvar no histórico de novo."
    Else
        msg = msg & vbCrLf & vbCrLf & "Dia ainda não salvo no histórico (dia em andamento ou sem valores)."
    End If
    If ok And InStr(msg, "Sem resultado") = 0 Then
        Aviso msg, vbInformation, "Buscar no MES"
    Else
        Aviso msg & vbCrLf & vbCrLf & "Item sem resultado: confira no MES; se o MES tiver o valor, rode a aba " & _
              "Diagnóstico (ela acha a forma da tag que funciona).", vbExclamation, "Buscar no MES"
    End If
End Sub

Public Sub SalvarDia()
    Dim d As Date, r As Long
    If Not EData(Nm("dData").Value) Then
        Aviso "Informe a data do report (célula amarela).", vbExclamation, "Salvar"
        Exit Sub
    End If
    d = CDate(DiaDe(Nm("dData").Value))
    If ContarValores(ValoresDia()) = 0 Then
        Aviso "Nenhum valor na aba Dia. Busque no MES ou lance os valores antes de salvar.", vbExclamation, "Salvar"
        Exit Sub
    End If
    If CDbl(d) >= Int(CDbl(Now)) Then
        If Aviso(TxtDia(d) & " ainda não terminou: o DVR e o fechamento do batch ainda não estão no MES." & vbCrLf & _
                 vbCrLf & "Salvar assim mesmo? (dá para buscar e salvar de novo amanhã)", vbQuestion + vbYesNo, _
                 "Salvar") <> vbYes Then Exit Sub
    End If
    r = GravarHist(d, ValoresDia(), Nm("d_fator").Value, Nm("d_obs").Value, False)
    If r = 0 Then Exit Sub
    OrdenarHist
    AvancarPainel d
    AtualizarPainel
    Aviso "Dia " & TxtDia(d) & " salvo no histórico. O painel já mostra este dia.", vbInformation, "Salvar"
End Sub

' Comentarios do painel: volta os tres automaticos (formulas), todos "Sim", e limpa os topicos livres.
' Roda sozinho ao trocar a data da aba Dia.
Public Sub RestaurarComentarios()
    Dim j As Long
    ModoCarga True
    Desproteger shDia
    For j = 1 To 3
        Nm("dCom" & j).Formula = "=gCom" & j
        Nm("dComOn" & j).Value = "Sim"
        Nm("dTop" & j).ClearContents
    Next j
    Proteger shDia
    ModoCarga False
    Recalcular
End Sub

' Troca da data na aba Dia: abre o dia do historico (ou limpa para um dia novo)
Public Sub CarregarDia()
    Dim d As Date, r As Long, k As Long, v As Variant
    If Not EData(Nm("dData").Value) Then Exit Sub
    d = CDate(DiaDe(Nm("dData").Value))
    r = LinhaHist(d)
    ModoCarga True
    Desproteger shDia
    For k = 1 To NI
        If r > 0 Then v = shHist.Cells(r, HCol(IChave(k))).Value Else v = Empty
        If Vazio(v) Then Nm("d_" & IChave(k)).ClearContents Else Nm("d_" & IChave(k)).Value = v
        shDia.Cells(D_ROW_IT0 + k - 1, D_COL_NRES).ClearContents
    Next k
    If r > 0 Then
        v = shHist.Cells(r, H_COL_FATOR).Value
        If ENumero(v) Then Nm("d_fator").Value = v Else Nm("d_fator").Value = FatorPadrao(d)
        Nm("d_obs").Value = shHist.Cells(r, H_COL_OBS).Value
    Else
        Nm("d_fator").Value = FatorPadrao(d)
        Nm("d_obs").ClearContents
    End If
    Proteger shDia
    ModoCarga False
    Recalcular
End Sub

' ---------------------------------------------------------------- aba Dia
Private Sub PreencherDia(ByVal vals As Variant, ByVal nres As Variant, ByVal sit As Variant)
    Dim k As Long
    For k = 1 To NI
        If ICfg(k, C_COL_TAG) <> "" Then
            ' item sem resultado no MES: o valor que ja estava (lancado a mao) fica
            If Not IsEmpty(vals(k)) Then Nm("d_" & IChave(k)).Value = Round(CDbl(vals(k)), 2)
            shDia.Cells(D_ROW_IT0 + k - 1, D_COL_NRES).Value = nres(k)
        Else
            shDia.Cells(D_ROW_IT0 + k - 1, D_COL_NRES).ClearContents
        End If
    Next k
End Sub

Private Function ValoresDia() As Variant
    Dim v() As Variant, k As Long, x As Variant
    ReDim v(1 To NI)
    For k = 1 To NI
        x = Nm("d_" & IChave(k)).Value
        If ENumero(x) Then v(k) = CDbl(x)
    Next k
    ValoresDia = v
End Function

Private Function ContarValores(ByVal vals As Variant) As Long
    Dim k As Long
    For k = 1 To NI
        If Not IsEmpty(vals(k)) Then ContarValores = ContarValores + 1
    Next k
End Function

Private Function ResumoBusca(ByVal d As Date, ByVal vals As Variant, ByVal sit As Variant) As String
    Dim k As Long, falta As String, mao As String, n As Long, nT As Long
    For k = 1 To NI
        If ICfg(k, C_COL_TAG) = "" Then
            mao = mao & IIf(mao = "", "", ", ") & IRotulo(k)
        Else
            nT = nT + 1
            If IsEmpty(vals(k)) Then falta = falta & IIf(falta = "", "", ", ") & IRotulo(k) Else n = n + 1
        End If
    Next k
    ResumoBusca = "Dia " & TxtDia(d) & ": " & n & " de " & nT & " itens com resultado no MES."
    If falta <> "" Then ResumoBusca = ResumoBusca & vbCrLf & "Sem resultado: " & falta & "."
    If mao <> "" Then ResumoBusca = ResumoBusca & vbCrLf & "Lançar à mão (sem tag): " & mao & "."
End Function

' ---------------------------------------------------------------- historico
Private Function ColunaDatas() As Variant
    ColunaDatas = shHist.Range(shHist.Cells(H_ROW_HDR + 1, H_COL_DATA), _
                               shHist.Cells(H_ROW_HDR + H_NLIN, H_COL_DATA)).Value
End Function

' Linha do dia d na aba Historico (0 = nao tem)
Public Function LinhaHist(ByVal d As Date) As Long
    Dim arr As Variant, i As Long, alvo As Long
    alvo = CLng(Int(CDbl(d)))
    arr = ColunaDatas()
    For i = 1 To UBound(arr, 1)
        If DiaDe(arr(i, 1)) = alvo Then
            LinhaHist = H_ROW_HDR + i
            Exit Function
        End If
    Next i
End Function

Private Function PrimeiraLivre() As Long
    Dim arr As Variant, i As Long
    arr = ColunaDatas()
    For i = 1 To UBound(arr, 1)
        If Vazio(arr(i, 1)) Then
            PrimeiraLivre = H_ROW_HDR + i
            Exit Function
        End If
    Next i
End Function

' Fator do dia anterior mais proximo com fator no historico (ou o padrao de Configuracoes)
Public Function FatorPadrao(ByVal d As Date) As Double
    Dim arr As Variant, fat As Variant, i As Long, melhor As Long, alvo As Long, dd As Long
    alvo = CLng(Int(CDbl(d)))
    arr = ColunaDatas()
    fat = shHist.Range(shHist.Cells(H_ROW_HDR + 1, H_COL_FATOR), shHist.Cells(H_ROW_HDR + H_NLIN, H_COL_FATOR)).Value
    FatorPadrao = CfgNum("cfgFatorPadrao", 160)
    For i = 1 To UBound(arr, 1)
        dd = DiaDe(arr(i, 1))
        If dd > 0 And dd < alvo And dd > melhor And ENumero(fat(i, 1)) Then
            melhor = dd
            FatorPadrao = CDbl(fat(i, 1))
        End If
    Next i
End Function

' Grava o dia d no historico (substitui a linha do dia, se ja existir). manter = True: item sem valor, Fator e
' Observacao ja lancados ficam como estao. Devolve a linha (0 = sem espaco).
Public Function GravarHist(ByVal d As Date, ByVal vals As Variant, ByVal fator As Variant, ByVal obs As Variant, _
                           ByVal manter As Boolean) As Long
    Dim r As Long, k As Long, c As Range
    r = LinhaHist(d)
    If r = 0 Then
        r = PrimeiraLivre()
        If r = 0 Then
            Aviso "Histórico cheio (" & H_NLIN & " dias). Gere a planilha de novo com mais linhas.", vbCritical, _
                  "Histórico"
            Exit Function
        End If
        manter = False
    End If
    ModoCarga True
    shHist.Cells(r, H_COL_DATA).Value = d
    shHist.Cells(r, H_COL_DATA).NumberFormat = "dd/mm/yyyy"
    For k = 1 To NI
        Set c = shHist.Cells(r, HCol(IChave(k)))
        If Not IsEmpty(vals(k)) Then
            c.Value = vals(k)
        ElseIf Not manter Then
            c.ClearContents
        End If
    Next k
    If Not (manter And ENumero(shHist.Cells(r, H_COL_FATOR).Value)) Then
        If ENumero(fator) Then
            shHist.Cells(r, H_COL_FATOR).Value = CDbl(fator)
        Else
            shHist.Cells(r, H_COL_FATOR).Value = FatorPadrao(d)
        End If
    End If
    If Not manter Then shHist.Cells(r, H_COL_OBS).Value = obs
    shHist.Cells(r, H_COL_ATUAL).Value = Now
    shHist.Cells(r, H_COL_ATUAL).NumberFormat = "dd/mm/yyyy hh:mm"
    ModoCarga False
    GravarHist = r
End Function

' Ordena o historico pela data (so as colunas de entrada; as calculadas ficam no lugar e acompanham a linha)
Public Sub OrdenarHist()
    Dim arr As Variant, n As Long, i As Long, j As Long, ordem() As Long, t As Long, ordenado As Boolean
    Dim cols As Variant, c As Long, col As Variant, novo() As Variant, rg As Range
    arr = ColunaDatas()
    For i = 1 To UBound(arr, 1)
        If Vazio(arr(i, 1)) Then Exit For
        n = i
    Next i
    If n < 2 Then Exit Sub
    ordenado = True
    For i = 2 To n
        If DiaDe(arr(i, 1)) < DiaDe(arr(i - 1, 1)) Then
            ordenado = False
            Exit For
        End If
    Next i
    If ordenado Then Exit Sub
    ReDim ordem(1 To n)
    For i = 1 To n
        ordem(i) = i
    Next i
    ' insercao (o historico quase sempre ja esta em ordem)
    For i = 2 To n
        t = ordem(i)
        j = i - 1
        Do While j >= 1
            If DiaDe(arr(ordem(j), 1)) <= DiaDe(arr(t, 1)) Then Exit Do
            ordem(j + 1) = ordem(j)
            j = j - 1
        Loop
        ordem(j + 1) = t
    Next i
    ModoCarga True
    cols = Split(H_ENTRADAS, "|")
    For c = LBound(cols) To UBound(cols)
        Set rg = shHist.Range(shHist.Cells(H_ROW_HDR + 1, HCol(CStr(cols(c)))), _
                              shHist.Cells(H_ROW_HDR + n, HCol(CStr(cols(c)))))
        col = rg.Value
        ReDim novo(1 To n, 1 To 1)          ' vetor novo (no LibreOffice "novo = col" compartilha o vetor)
        For i = 1 To n
            novo(i, 1) = col(ordem(i), 1)
        Next i
        rg.Value = novo
    Next c
    ModoCarga False
End Sub

' Painel passa a terminar no dia salvo, se ele for mais novo
Private Sub AvancarPainel(ByVal d As Date)
    If Not EData(Nm("rFim").Value) Or CDbl(d) > CDbl(DiaDe(Nm("rFim").Value)) Then
        ModoCarga True
        Desproteger shPainel
        Nm("rFim").Value = d
        Proteger shPainel
        ModoCarga False
    End If
End Sub

Public Function UltimoDiaHist() As Long
    Dim arr As Variant, i As Long
    arr = ColunaDatas()
    For i = 1 To UBound(arr, 1)
        If DiaDe(arr(i, 1)) > UltimoDiaHist Then UltimoDiaHist = DiaDe(arr(i, 1))
    Next i
End Function

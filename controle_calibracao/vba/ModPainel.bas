Option Explicit

' Etapa em andamento (aparece na mensagem de erro)
Private mEtapa As String

' ============================================================================
'  ModPainel - indicadores do periodo a partir da base (BD_Afericoes)
'  Atualiza ao abrir a aba Painel e pelo botao "Atualizar painel".
' ============================================================================

Public Sub PainelMesAtual()
    Nm("painelIni").Value = DateSerial(Year(Now), Month(Now), 1)
    Nm("painelFim").Value = CDate(Int(CDbl(Now)))
    AtualizarPainel
End Sub

Public Sub PainelUltimos30()
    Nm("painelIni").Value = CDate(Int(CDbl(Now)) - 29)
    Nm("painelFim").Value = CDate(Int(CDbl(Now)))
    AtualizarPainel
End Sub

Public Sub AtualizarPainel(Optional ByVal silencioso As Boolean = False)
    Dim ini As Long, fim As Long, ult As Long, a As Variant, i As Long, k As Long, j As Long, n As Long
    Dim reg(1 To 20) As Long, conf(1 To 20) As Long, ncs(1 To 20) As Long, ultData(1 To 20) As Long
    Dim ultNC(1 To 20) As Boolean, prev(1 To 20) As Long, feito(1 To 20) As Long, ocas(1 To 20) As String
    Dim d As Long, t As String, res As String, totReg As Long, totConf As Long, totNC As Long
    Dim totPrev As Long, totFeito As Long, ag As Variant, w As Long, tt As Long, tur As String, ok As Boolean
    Dim ws As Worksheet, sit As String, eqLin() As Long, eqData() As Long, chave As String, eqChave() As String
    Dim idxNC() As Long, nNC As Long, usados() As Boolean, melhor As Long, hoje As Long
    On Error GoTo Falha

    Set ws = shPainel
    If Not (ENumero(Nm("painelIni").Value) Or VarType(Nm("painelIni").Value) = vbDate) Then _
        Nm("painelIni").Value = DateSerial(Year(Now), Month(Now), 1)
    If Not (ENumero(Nm("painelFim").Value) Or VarType(Nm("painelFim").Value) = vbDate) Then _
        Nm("painelFim").Value = CDate(Int(CDbl(Now)))
    ini = CLng(Int(CDbl(Nm("painelIni").Value)))
    fim = CLng(Int(CDbl(Nm("painelFim").Value)))
    If fim < ini Then
        If Not silencioso Then Aviso "O fim do período deve ser depois do início.", vbExclamation
        Exit Sub
    End If
    hoje = CLng(Int(CDbl(Now)))
    Ampulheta True
    Desproteger ws

    ult = UltimaLinhaBD()
    If ult >= 2 Then
        a = shBD.Range("A2:Q" & ult).Value
    Else
        ReDim a(1 To 1, 1 To 17)
    End If
    n = UBound(a, 1)
    ReDim eqLin(1 To PAI_NEQ)
    ReDim eqChave(1 To PAI_NEQ)
    For j = 1 To PAI_NEQ
        eqChave(j) = CStr(ws.Cells(PAI_EQ1 + j - 1, 2).Value) & "#" & CStr(ws.Cells(PAI_EQ1 + j - 1, 4).Value) & "#" & _
                     CStr(ws.Cells(PAI_EQ1 + j - 1, 6).Value)
    Next j
    ReDim eqData(1 To PAI_NEQ)
    ReDim idxNC(1 To n)
    For k = 1 To N_ENSAIOS
        ocas(k) = "|"
    Next k

    mEtapa = "base"
    ' ---- varredura da base
    For i = 1 To n
        If DataNoPeriodo(a(i, 3), ini, fim, d) Then
            k = IndiceEnsaio(CStr(a(i, 1)))
            res = CStr(a(i, 13))
            t = TurnoDaLinha(a(i, 17), a(i, 16))
            If k > 0 Then
                reg(k) = reg(k) + 1
                If res = "Conforme" Then conf(k) = conf(k) + 1
                If res = "Não conforme" Then ncs(k) = ncs(k) + 1
                If d > ultData(k) Then
                    ultData(k) = d
                    ultNC(k) = False
                End If
                If d = ultData(k) And res = "Não conforme" Then ultNC(k) = True
                If t = "" Then t = "*"
                If InStr(1, ocas(k), "|" & d & "#" & t & "|") = 0 Then ocas(k) = ocas(k) & d & "#" & t & "|"
            End If
            If res = "Conforme" Then totConf = totConf + 1
            If res = "Não conforme" Then
                totNC = totNC + 1
                nNC = nNC + 1
                idxNC(nNC) = i
            End If
            totReg = totReg + 1
            ' situacao por equipamento: ultimo resultado de cada combinacao
            chave = CStr(a(i, 1)) & "#" & CStr(a(i, 4)) & "#" & CStr(a(i, 5))
            For j = 1 To PAI_NEQ
                If eqChave(j) = chave Then
                    If d >= eqData(j) Then
                        eqData(j) = d
                        eqLin(j) = i
                    End If
                    Exit For
                End If
            Next j
        End If
    Next i

    mEtapa = "agenda"
    ' ---- aderencia a agenda (turnos previstos x turnos com registro)
    ag = Nm("agenda").Value
    For d = ini To fim
        If d > hoje Then Exit For
        w = Weekday(CDate(d), vbMonday)
        For k = 1 To N_ENSAIOS
            For tt = 1 To 2
                If UCase$(Trim$(CStr(ag(k, (w - 1) * 2 + tt)))) = "X" Then
                    If tt = 1 Then tur = "07x19" Else tur = "19x07"
                    ok = (InStr(1, ocas(k), "|" & d & "#" & tur & "|") > 0) Or (InStr(1, ocas(k), "|" & d & "#*|") > 0)
                    If d < hoje Or ok Then
                        prev(k) = prev(k) + 1
                        If ok Then feito(k) = feito(k) + 1
                    End If
                End If
            Next tt
        Next k
    Next d

    mEtapa = "ensaios"
    ' ---- tabela por ensaio
    For k = 1 To N_ENSAIOS
        i = PAI_ENS1 + k - 1
        ws.Cells(i, 6).Value = IIf(prev(k) > 0, prev(k), "—")
        ws.Cells(i, 8).Value = IIf(prev(k) > 0 Or reg(k) > 0, ContaOcasioes(ocas(k)), "—")
        If prev(k) > 0 Then ws.Cells(i, 9).Value = feito(k) / prev(k) Else ws.Cells(i, 9).Value = "—"
        ws.Cells(i, 10).Value = reg(k)
        ws.Cells(i, 11).Value = conf(k)
        ws.Cells(i, 12).Value = ncs(k)
        If conf(k) + ncs(k) > 0 Then ws.Cells(i, 13).Value = conf(k) / (conf(k) + ncs(k)) Else ws.Cells(i, 13).Value = "—"
        If ultData(k) > 0 Then ws.Cells(i, 14).Value = CDate(ultData(k)) Else ws.Cells(i, 14).Value = "—"
        totPrev = totPrev + prev(k)
        totFeito = totFeito + feito(k)
        If ultNC(k) Then
            sit = "Não conforme (último)"
        ElseIf prev(k) > feito(k) Then
            sit = "Pendente (" & (prev(k) - feito(k)) & ")"
        ElseIf reg(k) > 0 Then
            sit = "Em dia"
        Else
            sit = "Sem registro"
        End If
        ws.Cells(i, 16).Value = sit
    Next k

    mEtapa = "kpis"
    ' ---- KPIs
    Nm("kpiReg").Value = totReg
    Nm("kpiRegSub").Value = "de " & Format$(CDate(ini), "dd\/mm\/yyyy") & " a " & Format$(CDate(fim), "dd\/mm\/yyyy")
    If totConf + totNC > 0 Then
        Nm("kpiConf").Value = totConf / (totConf + totNC)
        Nm("kpiConf").NumberFormat = "0.0%"
        Nm("kpiConfSub").Value = totConf & " de " & (totConf + totNC) & " dentro do critério"
    Else
        Nm("kpiConf").Value = "—"
        Nm("kpiConfSub").Value = "sem resultados no período"
    End If
    Nm("kpiNC").Value = totNC
    Nm("kpiNCSub").Value = IIf(totNC = 0, "nenhum resultado fora do critério", "resultado(s) fora do critério")
    If totPrev > 0 Then
        Nm("kpiAder").Value = totFeito / totPrev
        Nm("kpiAder").NumberFormat = "0%"
        Nm("kpiAderSub").Value = totFeito & " de " & totPrev & " turnos previstos com registro"
    Else
        Nm("kpiAder").Value = "—"
        Nm("kpiAderSub").Value = "sem turnos previstos até hoje"
    End If

    mEtapa = "equip"
    ' ---- situacao por equipamento
    For j = 1 To PAI_NEQ
        i = PAI_EQ1 + j - 1
        If eqLin(j) > 0 Then
            k = eqLin(j)
            ws.Cells(i, 9).Value = a(k, 7)
            ws.Cells(i, 10).Value = a(k, 8)
            ws.Cells(i, 11).Value = a(k, 9)
            ws.Cells(i, 12).Value = TextoCriterio(a(k, 10), a(k, 11), a(k, 12), CStr(a(k, 13)))
            ws.Cells(i, 14).Value = CDate(eqData(j))
            ws.Cells(i, 15).Value = a(k, 14)
            ws.Cells(i, 16).Value = a(k, 13)
        Else
            ws.Range(ws.Cells(i, 9), ws.Cells(i, 11)).ClearContents
            ws.Cells(i, 12).Value = ""
            ws.Cells(i, 14).Value = ""
            ws.Cells(i, 15).Value = ""
            ws.Cells(i, 16).Value = "Sem registro"
        End If
    Next j

    mEtapa = "pen"
    PreencherPeneiradores a, ini, fim
    mEtapa = "nc"
    ' ---- nao conformidades (mais recentes primeiro)
    ReDim usados(1 To IIf(nNC > 0, nNC, 1))
    For j = 1 To PAI_NNC
        i = PAI_NC1 + j - 1
        melhor = 0
        For k = 1 To nNC
            If Not usados(k) Then
                If melhor = 0 Then
                    melhor = k
                ElseIf CDbl(a(idxNC(k), 3)) >= CDbl(a(idxNC(melhor), 3)) Then
                    melhor = k
                End If
            End If
        Next k
        If melhor > 0 Then
            usados(melhor) = True
            k = idxNC(melhor)
            ws.Cells(i, 2).Value = CDate(Int(CDbl(a(k, 3))))
            ws.Cells(i, 3).Value = TurnoDaLinha(a(k, 17), a(k, 16))
            ws.Cells(i, 4).Value = a(k, 1)
            ws.Cells(i, 6).Value = a(k, 4)
            ws.Cells(i, 8).Value = a(k, 5)
            ws.Cells(i, 11).Value = a(k, 7)
            ws.Cells(i, 12).Value = a(k, 8)
            ws.Cells(i, 13).Value = TextoCriterio(a(k, 10), a(k, 11), a(k, 12), "")
            ws.Cells(i, 14).Value = a(k, 14)
            ws.Cells(i, 15).Value = a(k, 16)
        Else
            ws.Cells(i, 2).Value = ""
            ws.Cells(i, 3).Value = ""
            ws.Cells(i, 4).Value = ""
            ws.Cells(i, 6).Value = ""
            ws.Cells(i, 8).Value = ""
            ws.Cells(i, 11).Value = ""
            ws.Cells(i, 12).Value = ""
            ws.Cells(i, 13).Value = ""
            ws.Cells(i, 14).Value = ""
            ws.Cells(i, 15).Value = ""
            If j = 1 Then ws.Cells(i, 4).Value = "Nenhuma não conformidade no período"
        End If
    Next j

    mEtapa = "graficos"
    PreencherGraficos a, ini, fim
    Nm("painelInfo").Value = "Atualizado em " & Format$(Now, "dd\/mm\/yyyy hh:mm") & "   |   base com " & _
        (ult - 1) & " resultado(s)   |   período de " & (fim - ini + 1) & " dia(s)"
    Proteger ws
    Ampulheta False
    Exit Sub
Falha:
    Ampulheta False
    Proteger ws
    If Not silencioso Then Aviso "Erro ao atualizar o painel (" & mEtapa & "): " & Err.Number & " " & Err.Description, vbExclamation
End Sub

Private Function DataNoPeriodo(ByVal v As Variant, ByVal ini As Long, ByVal fim As Long, ByRef d As Long) As Boolean
    If IsError(v) Then Exit Function
    If Not (VarType(v) = vbDate Or ENumero(v)) Then Exit Function
    d = CLng(Int(CDbl(v)))
    DataNoPeriodo = (d >= ini And d <= fim)
End Function

Private Function IndiceEnsaio(ByVal nome As String) As Long
    Dim k As Long
    For k = 1 To N_ENSAIOS
        If NomeEnsaio(k) = nome Then
            IndiceEnsaio = k
            Exit Function
        End If
    Next k
End Function

Private Function ContaOcasioes(ByVal s As String) As Long
    Dim p As Variant
    p = Split(s, "|")
    ContaOcasioes = UBound(p) - LBound(p) - 1
    If ContaOcasioes < 0 Then ContaOcasioes = 0
End Function

Private Function TextoCriterio(ByVal li As Variant, ByVal ls As Variant, ByVal tol As Variant, ByVal res As String) As String
    If ENumero(tol) Then
        TextoCriterio = "± " & Format$(CDbl(tol), "0.###")
    ElseIf ENumero(li) And ENumero(ls) Then
        TextoCriterio = Format$(CDbl(li), "0.###") & " a " & Format$(CDbl(ls), "0.###")
    ElseIf ENumero(li) Then
        TextoCriterio = ">= " & Format$(CDbl(li), "0.###")
    ElseIf ENumero(ls) Then
        TextoCriterio = "<= " & Format$(CDbl(ls), "0.###")
    ElseIf res = "Informativo" Then
        TextoCriterio = "informativo"
    End If
End Function

' Peneiradores: por grupo, a ultima data verificada no periodo e o resumo das telas
Private Sub PreencherPeneiradores(ByVal a As Variant, ByVal ini As Long, ByVal fim As Long)
    Dim g As Long, i As Long, d As Long, grupo As String, ws As Worksheet, ultD As Long
    Dim telas As Long, nok As Long, semTag As Long, resp As String, lin As Long
    Set ws = shPainel
    For g = 1 To PAI_NPEN
        lin = PAI_PEN1 + g - 1
        grupo = CStr(ws.Cells(lin, 2).Value)
        ultD = 0
        For i = 1 To UBound(a, 1)
            If CStr(a(i, 1)) = "Verificação Peneiradores" Then
                If DataNoPeriodo(a(i, 3), ini, fim, d) Then
                    If InStr(1, CStr(a(i, 5)), "(" & grupo & ")") > 0 And d > ultD Then ultD = d
                End If
            End If
        Next i
        telas = 0: nok = 0: semTag = 0: resp = ""
        If ultD > 0 Then
            For i = 1 To UBound(a, 1)
                If CStr(a(i, 1)) = "Verificação Peneiradores" Then
                    If DataNoPeriodo(a(i, 3), ini, fim, d) Then
                        If d = ultD And InStr(1, CStr(a(i, 5)), "(" & grupo & ")") > 0 Then
                            telas = telas + 1
                            If CStr(a(i, 13)) = "Não conforme" Then nok = nok + 1
                            If InStr(1, CStr(a(i, 16)), "Sem TAG", vbTextCompare) > 0 Then semTag = semTag + 1
                            resp = CStr(a(i, 14))
                        End If
                    End If
                End If
            Next i
            ws.Cells(lin, 6).Value = CDate(ultD)
            ws.Cells(lin, 8).Value = telas
            ws.Cells(lin, 10).Value = nok
            ws.Cells(lin, 12).Value = semTag
            ws.Cells(lin, 14).Value = resp
            ws.Cells(lin, 16).Value = IIf(nok > 0, "Não conforme", IIf(semTag > 0, "Conforme · " & semTag & " sem TAG", "Conforme"))
        Else
            ws.Cells(lin, 6).Value = ""
            ws.Cells(lin, 8).Value = ""
            ws.Cells(lin, 10).Value = ""
            ws.Cells(lin, 12).Value = ""
            ws.Cells(lin, 14).Value = ""
            ws.Cells(lin, 16).Value = "Sem registro"
        End If
    Next g
End Sub

' Graficos: ultimas ocasioes (data + turno) de cada ensaio no periodo
Private Sub PreencherGraficos(ByVal a As Variant, ByVal ini As Long, ByVal fim As Long)
    Dim b As Long, r0 As Long, ens As String, campo As Long, criterio As String, ns As Long
    Dim eqs(1 To 10) As String, pas(1 To 10) As String, i As Long, j As Long, d As Long, t As String
    Dim occ() As String, occD() As Long, nocc As Long, chave As String, o As Long, s As Long
    Dim vals() As Variant, li As Variant, ls As Variant, primeiro As Long, npt As Long, tmp As String, tmpD As Long
    Dim ws As Worksheet, lim As String, codigo As String
    Set ws = shGraficos
    npt = GRA_NPT
    For b = 1 To GRA_NBLOCOS
        r0 = 3 + (b - 1) * GRA_LIN
        codigo = CStr(ws.Cells(r0, 1).Value)
        ens = NomeEnsaio(IndiceCodigo(codigo))
        If CStr(ws.Cells(r0, 2).Value) = "I" Then campo = 9 Else campo = 7
        criterio = CStr(ws.Cells(r0, 3).Value)
        ns = 0
        Do While Not Vazio(ws.Cells(r0 + 1, 2 + ns).Value) And CStr(ws.Cells(r0 + 1, 2 + ns).Value) <> "Lim. inferior"
            ns = ns + 1
            eqs(ns) = CStr(ws.Cells(r0 + 1, 1 + ns).Value)
            pas(ns) = CStr(ws.Cells(r0, 3 + ns).Value)
            If ns >= 10 Then Exit Do
        Loop
        ' ocasioes
        nocc = 0
        ReDim occ(1 To UBound(a, 1) + 1)
        ReDim occD(1 To UBound(a, 1) + 1)
        For i = 1 To UBound(a, 1)
            If CStr(a(i, 1)) = ens Then
                If DataNoPeriodo(a(i, 3), ini, fim, d) Then
                    chave = d & "#" & TurnoDaLinha(a(i, 17), a(i, 16))
                    For o = 1 To nocc
                        If occ(o) = chave Then Exit For
                    Next o
                    If o > nocc Then
                        nocc = nocc + 1
                        occ(nocc) = chave
                        occD(nocc) = d * 10 + IIf(InStr(chave, "19x07") > 0, 2, 1)
                    End If
                End If
            End If
        Next i
        ' ordena por data/turno (insercao)
        For i = 2 To nocc
            tmp = occ(i): tmpD = occD(i)
            j = i - 1
            Do While j >= 1
                If occD(j) <= tmpD Then Exit Do
                occ(j + 1) = occ(j): occD(j + 1) = occD(j)
                j = j - 1
            Loop
            occ(j + 1) = tmp: occD(j + 1) = tmpD
        Next i
        primeiro = 1
        If nocc > npt Then primeiro = nocc - npt + 1
        ReDim vals(1 To npt, 1 To ns + 3)
        For o = primeiro To nocc
            vals(o - primeiro + 1, 1) = Format$(CDate(occD(o) \ 10), "dd\/mm") & IIf(occD(o) Mod 10 = 2, " N", " D")
        Next o
        For i = 1 To UBound(a, 1)
            If CStr(a(i, 1)) = ens Then
                If DataNoPeriodo(a(i, 3), ini, fim, d) Then
                    chave = d & "#" & TurnoDaLinha(a(i, 17), a(i, 16))
                    For o = primeiro To nocc
                        If occ(o) = chave Then Exit For
                    Next o
                    If o <= nocc Then
                        For s = 1 To ns
                            If CStr(a(i, 4)) = eqs(s) And CStr(a(i, 5)) = pas(s) Then
                                If ENumero(a(i, campo)) Then vals(o - primeiro + 1, 1 + s) = CDbl(a(i, campo))
                            End If
                        Next s
                    End If
                End If
            End If
        Next i
        ' linhas de referencia
        lim = Mid$(criterio, InStr(criterio, ":") + 1)
        li = Empty: ls = Empty
        If Left$(criterio, 5) = "faixa" Then
            li = Nm("c_" & lim & "_LI").Value
            ls = Nm("c_" & lim & "_LS").Value
        ElseIf Left$(criterio, 3) = "tol" Then
            If ENumero(Nm("c_" & lim & "_TOL").Value) Then
                li = -CDbl(Nm("c_" & lim & "_TOL").Value)
                ls = CDbl(Nm("c_" & lim & "_TOL").Value)
            End If
        End If
        For o = 1 To npt
            If o <= nocc - primeiro + 1 Then
                If ENumero(li) Then vals(o, ns + 2) = CDbl(li)
                If ENumero(ls) Then vals(o, ns + 3) = CDbl(ls)
            End If
        Next o
        ws.Range(ws.Cells(r0 + 2, 1), ws.Cells(r0 + 1 + npt, ns + 3)).ClearContents
        ws.Range(ws.Cells(r0 + 2, 1), ws.Cells(r0 + 1 + npt, ns + 3)).Value = vals
    Next b
End Sub

Private Function IndiceCodigo(ByVal cod As String) As Long
    Select Case cod
        Case "BLA": IndiceCodigo = 1
        Case "TAM": IndiceCodigo = 2
        Case "ALP": IndiceCodigo = 3
        Case "UMI": IndiceCodigo = 4
        Case "COM": IndiceCodigo = 5
        Case "GRA": IndiceCodigo = 6
        Case "T515": IndiceCodigo = 7
        Case "PEN": IndiceCodigo = 8
        Case "FIS": IndiceCodigo = 9
    End Select
End Function

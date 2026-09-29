Option Explicit

' ============================================================================
'  ModResumo - abas "Resumo Dia" e "Resumo Noite" (uma por turno, arquivo diario)
'   Imagem 1 (AreaResultados): resultado, minimo e maximo por usina
'   Imagem 2 (AreaPassagem)  : passagem de turno em texto
'  Os valores sao GRAVADOS (nao sao formulas): o resumo e uma "foto" do turno.
'  Enquanto o turno nao e finalizado, a foto e refeita ao atualizar o MES, pelo
'  botao 3 do Painel e antes de copiar as imagens.
' ============================================================================

Private Const COR_NORMAL As Long = 7092992     ' RGB(0, 59, 108)    Azul Titulo Samarco
Private Const FUNDO_NORMAL As Long = 16644334  ' RGB(238, 248, 253)
Private Const COR_FORA As Long = 2191603       ' RGB(243, 112, 33)  laranja
Private Const FUNDO_FORA As Long = 14214909    ' RGB(253, 230, 216)
Private Const COR_TEXTO As Long = 3549981      ' RGB(29, 43, 54)    texto

' ---------------------------------------------------------------- utilidades
Public Function PrefixoResumo(ByVal turno As String) As String
    If EhDia(turno) Then PrefixoResumo = "rsD" Else PrefixoResumo = "rsN"
End Function

Public Function AbaResumo(ByVal pref As String) As Worksheet
    If pref = "rsD" Then Set AbaResumo = shResumoDia Else Set AbaResumo = shResumoNoite
End Function

Public Function ResumoFinalizado(ByVal pref As String) As Boolean
    ResumoFinalizado = Not Vazio(Nm(pref & "Final").Value)
End Function

' Grava a data/hora de finalizacao do resumo (Empty = reabre o resumo)
Public Sub MarcarFinal(ByVal pref As String, ByVal quando As Variant)
    On Error Resume Next
    AbaResumo(pref).Unprotect Password:=SENHA
    On Error GoTo 0
    Nm(pref & "Final").Value = quando
    ProtegerPlanilhas
End Sub

' Prefixo do turno selecionado no Painel ("" se o Painel estiver incompleto)
Private Function PrefixoPainel() As String
    Dim dt As Date, turno As String, turma As String, resp As String
    If LerTurnoPainel(dt, turno, turma, resp, False) Then PrefixoPainel = PrefixoResumo(turno)
End Function

' ---------------------------------------------------------------- botao 3 do Painel
Public Sub IrResumo()
    Dim dt As Date, turno As String, turma As String, resp As String, pref As String
    If Not LerTurnoPainel(dt, turno, turma, resp, True) Then Exit Sub
    pref = PrefixoResumo(turno)
    If ResumoFinalizado(pref) Then
        If Aviso("O resumo do turno " & turno & " já foi FINALIZADO em " & _
                 Format$(Nm(pref & "Final").Value, "dd\/mm hh:mm") & "." & vbCrLf & vbCrLf & _
                 "Deseja refazer o resumo com os dados atuais?" & vbCrLf & _
                 "(Não = só abrir o resumo como está)", vbQuestion + vbYesNo, "Resumo do turno", vbNo) = vbNo Then
            Mostrar AbaResumo(pref)
            Exit Sub
        End If
        MarcarFinal pref, Empty
    End If
    If Not InformativoConfere(dt, turno) Then
        If Aviso("Os resultados do MES deste turno ainda não foram buscados." & vbCrLf & _
                 "Buscar agora?", vbQuestion + vbYesNo, "Resumo do turno") = vbYes Then
            gSilenciarSucesso = True
            AtualizarMES
            gSilenciarSucesso = False
        End If
    End If
    GerarResumo pref
    Mostrar AbaResumo(pref)
End Sub

' Chamado apos atualizar o MES: refaz o resumo do turno do Painel (se nao finalizado)
Public Sub AtualizarResumoAtual(ByVal exibir As Boolean)
    Dim pref As String
    pref = PrefixoPainel()
    If pref = "" Then Exit Sub
    If Not ResumoFinalizado(pref) Then GerarResumo pref
    If exibir Then Mostrar AbaResumo(pref)
End Sub

' Antes de copiar: refaz somente se for o resumo do turno do Painel e ainda nao finalizado
Public Sub PrepararResumoParaCopia(ByVal pref As String)
    If PrefixoPainel() = pref And Not ResumoFinalizado(pref) Then GerarResumo pref
End Sub

' ---------------------------------------------------------------- navegacao
Public Sub ResumoDiaIrPassagem()
    Rolar shResumoDia, RES_ROW_PASSAGEM
End Sub

Public Sub ResumoDiaIrResultados()
    Rolar shResumoDia, 1
End Sub

Public Sub ResumoNoiteIrPassagem()
    Rolar shResumoNoite, RES_ROW_PASSAGEM
End Sub

Public Sub ResumoNoiteIrResultados()
    Rolar shResumoNoite, 1
End Sub

Private Sub Rolar(ByVal ws As Worksheet, ByVal linha As Long)
    On Error Resume Next
    ws.Activate
    ActiveWindow.ScrollRow = linha
    ActiveWindow.ScrollColumn = 1
End Sub

' ---------------------------------------------------------------- geracao
Public Sub GerarResumo(ByVal pref As String)
    Dim ws As Worksheet, s As String, r As Long, p As Long, k As Long, i As Long, lin As Long, col As Long
    Dim dt As Date, turno As String, turma As String, resp As String, confere As Boolean
    Dim chave As Variant, v As Variant, c As Range, lie As Variant, lse As Variant, fora As Boolean

    Set ws = AbaResumo(pref)
    PrepararEdicao
    On Error Resume Next
    ws.Unprotect Password:=SENHA
    On Error GoTo 0

    If LerTurnoPainel(dt, turno, turma, resp, False) Then confere = InformativoConfere(dt, turno)

    ' Resultados: so sao regravados quando o Informativo e deste turno
    ' (se o Informativo estiver com outro turno/periodo, os resultados ja gravados sao mantidos)
    If confere Then
        ' garante resultado/min/max recalculados (tambem com o Excel em calculo manual)
        On Error Resume Next
        shInformativo.Calculate
        On Error GoTo 0
        For i = 1 To NSLOT
            Nm(pref & "Horas").Cells(1, i).Value = Nm("iHoras").Cells(1, i).Value
        Next i
        For r = 1 To RES_FIM_RESULTADOS
            chave = ws.Cells(r, 1).Value
            If VarType(chave) = vbString Then
                If Len(chave) = 5 And Left$(chave, 1) = "P" Then
                    p = CLng(Mid$(chave, 2, 2))
                    If Right$(chave, 1) = "3" Then k = 1 Else k = 2
                    lin = LinhaInformativo(p, k)
                    lie = CfgCel(p, CFG_COL_LIE)
                    lse = CfgCel(p, CFG_COL_LSE)
                    For i = 1 To NSLOT + 3
                        If i <= NSLOT Then col = INF_COL_H1 + i - 1 Else col = INF_COL_RES + i - NSLOT - 1
                        Set c = ws.Cells(r, 4 + i)
                        v = shInformativo.Cells(lin, col).Value
                        If ENumero(v) Then c.Value = CDbl(v) Else c.Value = "-"
                        ' destaque laranja: janelas de 2 h e resultado fora de LIE/LSE
                        If i <= NSLOT + 1 Then
                            fora = False
                            If ENumero(v) Then
                                If ENumero(lie) Then fora = (CDbl(v) < CDbl(lie))
                                If ENumero(lse) Then fora = fora Or (CDbl(v) > CDbl(lse))
                            End If
                            c.Font.Bold = fora Or (i = NSLOT + 1)
                            If fora Then
                                c.Font.Color = COR_FORA
                                c.Interior.Color = FUNDO_FORA
                            ElseIf i = NSLOT + 1 Then
                                c.Font.Color = COR_NORMAL
                                c.Interior.Color = FUNDO_NORMAL
                            Else
                                c.Font.Color = COR_TEXTO
                                c.Interior.Color = ws.Cells(r, 4).Interior.Color
                            End If
                        End If
                    Next i
                End If
            End If
        Next r
        Nm(pref & "MES").Value = Now
    End If

    s = LinhaDeInformacao(pref, confere)
    Nm(pref & "Info").Value = s
    Nm(pref & "Info2").Value = s
    Nm(pref & "Equipe").Value = Juntar("Equipe", "", False)
    Nm(pref & "Recebe").Value = TextoOuTraco(Nm("ptRecebe").Value)
    Nm(pref & "NaoOper").Value = NaoOperando()
    Nm(pref & "Min").Value = Juntar("Batch", "", True)
    Nm(pref & "Testes").Value = TextoOuTraco(Nm("ptTestes").Value)
    Nm(pref & "Pend").Value = TextoOuTraco(Nm("ptPendencias").Value)
    Nm(pref & "Com3").Value = Juntar("Comentário", "US3", False)
    Nm(pref & "Com4").Value = Juntar("Comentário", "US4", False)
    Nm(pref & "Embarque").Value = Juntar("Embarque", "", False)
    Nm(pref & "Obs").Value = TextoOuTraco(Nm("ptObs").Value)
    Nm(pref & "Gerado").Value = Now
    ProtegerPlanilhas
End Sub

Private Function LinhaDeInformacao(ByVal pref As String, ByVal confere As Boolean) As String
    Dim dt As Date, turno As String, turma As String, resp As String, s As String
    If LerTurnoPainel(dt, turno, turma, resp, False) Then
        s = Format$(dt, "dd\/mm\/yyyy") & "   ·   " & turno
        If turma <> "" Then s = s & "   ·   Turma " & turma
        If resp <> "" Then s = s & "   ·   " & resp
        If Not confere And Vazio(Nm(pref & "MES").Value) Then
            s = s & "   ·   (resultados do MES ainda não atualizados para este turno)"
        End If
    End If
    LinhaDeInformacao = s
End Function

Private Function TextoOuTraco(ByVal v As Variant) As String
    If Vazio(v) Then TextoOuTraco = "-" Else TextoOuTraco = Replace(CStr(v), vbLf, " ")
End Function

Private Function TextoValor(ByVal v As Variant) As String
    If VarType(v) = vbDate Then
        If Int(CDbl(v)) = 0 Then
            TextoValor = Format$(v, "hh:mm")
        Else
            TextoValor = Format$(v, "dd\/mm hh:mm")
        End If
    Else
        TextoValor = Replace(CStr(v), vbLf, " ")
    End If
End Function

' Junta os campos preenchidos de um tipo da aba _Mapa em uma linha de texto
Private Function Juntar(ByVal tipo As String, ByVal secao As String, ByVal comSecao As Boolean) As String
    Dim r As Long, s As String, tx As Variant, sec As String, ultimaSec As String
    r = 2
    Do While Not Vazio(shMapa.Cells(r, 1).Value)
        If CStr(shMapa.Cells(r, 1).Value) = tipo Then
            sec = CStr(LerEspec(shMapa.Cells(r, 2).Value))
            If secao = "" Or sec = secao Then
                tx = LerEspec(shMapa.Cells(r, 5).Value)
                If Not Vazio(tx) Then
                    If s <> "" Then s = s & "   ·   "
                    If comSecao And sec <> ultimaSec Then s = s & sec & " - "
                    s = s & CStr(LerEspec(shMapa.Cells(r, 3).Value)) & ": " & TextoValor(tx)
                    ultimaSec = sec
                End If
            End If
        End If
        r = r + 1
    Loop
    If s = "" Then s = "-"
    Juntar = s
End Function

Private Function NaoOperando() As String
    Dim r As Long, s As String, st As Variant, com As Variant, n As Long, preenchidos As Long, total As Long
    r = 2
    Do While Not Vazio(shMapa.Cells(r, 1).Value)
        If CStr(shMapa.Cells(r, 1).Value) = "Equipamento" Then
            total = total + 1
            st = LerEspec(shMapa.Cells(r, 4).Value)
            If Not Vazio(st) Then preenchidos = preenchidos + 1
            If CStr(st) = "Não operando" Then
                n = n + 1
                If s <> "" Then s = s & "   ·   "
                s = s & CStr(LerEspec(shMapa.Cells(r, 3).Value))
                com = LerEspec(shMapa.Cells(r, 5).Value)
                If Not Vazio(com) Then s = s & " (" & TextoValor(com) & ")"
            End If
        End If
        r = r + 1
    Loop
    If n = 0 Then
        If preenchidos = total And total > 0 Then
            s = "Nenhum - todos operando"
        ElseIf preenchidos = 0 Then
            s = "Status dos equipamentos não preenchido"
        Else
            s = "Nenhum (" & (total - preenchidos) & " sem status)"
        End If
    End If
    NaoOperando = s
End Function

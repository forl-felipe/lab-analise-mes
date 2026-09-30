Option Explicit

' ============================================================================
'  ModRelatorio - abas "Relatorio Dia" e "Relatorio Noite" (arquivo diario)
'  Uma pagina por turno: ocorrencias (aba Ocorrencia) + informativo de qualidade
'  quimico (aba Informativo, de 2 em 2 h).
'  Os valores sao GRAVADOS (nao sao formulas): o relatorio e uma "foto" do turno.
'  Enquanto o turno nao e finalizado, a foto e refeita ao atualizar o MES, pelo
'  botao 3 do Painel, pelo botao "Atualizar este relatorio" e antes de copiar.
' ============================================================================

Private Const COR_NORMAL As Long = 7092992     ' RGB(0, 59, 108)    Azul Titulo Samarco
Private Const FUNDO_NORMAL As Long = 16644334  ' RGB(238, 248, 253)
Private Const COR_FORA As Long = 2191603       ' RGB(243, 112, 33)  laranja
Private Const FUNDO_FORA As Long = 14214909    ' RGB(253, 230, 216)
Private Const COR_TEXTO As Long = 3549981      ' RGB(29, 43, 54)    texto
' Caracteres por linha nos textos do relatorio (para ajustar a altura das linhas)
Private Const CHARS_LINHA As Long = 95
Private Const SEP As String = "   ·   "

' ---------------------------------------------------------------- utilidades
Public Function PrefixoRel(ByVal turno As String) As String
    If EhDia(turno) Then PrefixoRel = "rlD" Else PrefixoRel = "rlN"
End Function

Public Function AbaRel(ByVal pref As String) As Worksheet
    If pref = "rlD" Then Set AbaRel = shRelDia Else Set AbaRel = shRelNoite
End Function

Public Function RelFinalizado(ByVal pref As String) As Boolean
    RelFinalizado = Not Vazio(Nm(pref & "Final").Value)
End Function

' Grava a data/hora de finalizacao do relatorio (Empty = reabre o relatorio)
Public Sub MarcarFinal(ByVal pref As String, ByVal quando As Variant)
    On Error Resume Next
    AbaRel(pref).Unprotect Password:=SENHA
    On Error GoTo 0
    Nm(pref & "Final").Value = quando
    ProtegerPlanilhas
End Sub

' Prefixo do turno selecionado no Painel ("" se o Painel estiver incompleto)
Private Function PrefixoPainel() As String
    Dim dt As Date, turno As String, turma As String, resp As String
    If LerTurnoPainel(dt, turno, turma, resp, False) Then PrefixoPainel = PrefixoRel(turno)
End Function

' ---------------------------------------------------------------- botao 3 do Painel
Public Sub IrRelatorio()
    Dim dt As Date, turno As String, turma As String, resp As String, pref As String
    If Not LerTurnoPainel(dt, turno, turma, resp, True) Then Exit Sub
    pref = PrefixoRel(turno)
    If RelFinalizado(pref) Then
        If Aviso("O relatório do turno " & turno & " já foi FINALIZADO em " & _
                 Format$(Nm(pref & "Final").Value, "dd\/mm hh:mm") & "." & vbCrLf & vbCrLf & _
                 "Deseja refazer o relatório com os dados atuais?" & vbCrLf & _
                 "(Não = só abrir o relatório como está)", vbQuestion + vbYesNo, "Relatório do turno", vbNo) = vbNo Then
            Mostrar AbaRel(pref)
            Exit Sub
        End If
        MarcarFinal pref, Empty
    End If
    If Not InformativoConfere(dt, turno) Then
        If Aviso("Os resultados do MES deste turno ainda não foram buscados." & vbCrLf & _
                 "Buscar agora?", vbQuestion + vbYesNo, "Relatório do turno") = vbYes Then
            gSilenciarSucesso = True
            AtualizarMES
            gSilenciarSucesso = False
        End If
    End If
    GerarRelatorio pref
    Mostrar AbaRel(pref)
End Sub

' Chamado apos atualizar o MES: refaz o relatorio do turno do Painel (se nao finalizado)
Public Sub AtualizarRelatorioAtual(ByVal exibir As Boolean)
    Dim pref As String
    pref = PrefixoPainel()
    If pref = "" Then Exit Sub
    If Not RelFinalizado(pref) Then GerarRelatorio pref
    If exibir Then Mostrar AbaRel(pref)
End Sub

' Antes de copiar: refaz somente se for o relatorio do turno do Painel e ainda nao finalizado
Public Sub PrepararRelatorioParaCopia(ByVal pref As String)
    If PrefixoPainel() = pref And Not RelFinalizado(pref) Then GerarRelatorio pref
End Sub

' Botoes "Atualizar este relatorio" das abas Relatorio Dia / Relatorio Noite
Public Sub AtualizarRelatorioDia()
    AtualizarRelatorioDaAba "rlD", TurnoDia()
End Sub

Public Sub AtualizarRelatorioNoite()
    AtualizarRelatorioDaAba "rlN", TurnoNoite()
End Sub

Private Sub AtualizarRelatorioDaAba(ByVal pref As String, ByVal turno As String)
    If PrefixoPainel() <> pref Then
        Aviso "Este relatório é do turno " & turno & "." & vbCrLf & _
              "No Painel, clique em 'Editar turno " & IIf(pref = "rlD", "Dia", "Noite") & _
              "' para atualizá-lo.", vbExclamation, "Relatório do turno"
        Exit Sub
    End If
    If RelFinalizado(pref) Then
        If Aviso("Este relatório já foi FINALIZADO. Refazer com os dados atuais?", _
                 vbQuestion + vbYesNo, "Relatório do turno", vbNo) = vbNo Then Exit Sub
        MarcarFinal pref, Empty
    End If
    GerarRelatorio pref
    Mostrar AbaRel(pref)
End Sub

' ---------------------------------------------------------------- geracao
Public Sub GerarRelatorio(ByVal pref As String)
    Dim ws As Worksheet, s As String, r As Long, p As Long, k As Long, i As Long, lin As Long, col As Long
    Dim dt As Date, turno As String, turma As String, resp As String, confere As Boolean
    Dim chave As Variant, v As Variant, c As Range, lie As Variant, lse As Variant, fora As Boolean

    Set ws = AbaRel(pref)
    PrepararEdicao
    On Error Resume Next
    ws.Unprotect Password:=SENHA
    On Error GoTo 0

    If LerTurnoPainel(dt, turno, turma, resp, False) Then confere = InformativoConfere(dt, turno)

    ' Resultados: so sao regravados quando o Informativo e deste turno
    ' (se o Informativo estiver com outro turno/periodo, os resultados ja gravados sao mantidos)
    If confere Then
        On Error Resume Next
        shInformativo.Calculate
        On Error GoTo 0
        For i = 1 To NSLOT
            Nm(pref & "Horas").Cells(1, i).Value = Nm("iHoras").Cells(1, i).Value
        Next i
        For r = REL_TAB_INI To REL_FIM
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
                        ' destaque laranja: janelas de 2 h e media fora de LIE/LSE
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
    Escrever pref & "Real", TextoLista("ocReal")
    Escrever pref & "Sol", TextoLista("ocSol")
    Escrever pref & "Equip", TextoLista("ocEquip")
    Escrever pref & "AReal", TextoLista("ocAReal")
    Escrever pref & "Lab", TextoLaboratorio()
    Escrever pref & "Pessoal", TextoPessoal()
    Escrever pref & "Cad", Juntar(Array("Repassados para o turno", "ocCadinhos", "Retirados para reforma", "ocReforma"))
    Escrever pref & "Obs", TextoObservacoes()
    OcultarAnalisesSemTag ws, REL_TAB_INI, REL_FIM
    Nm(pref & "Gerado").Value = Now
    ProtegerPlanilhas
End Sub

' Grava o texto e ajusta a altura da linha ao numero de linhas do texto
Private Sub Escrever(ByVal nome As String, ByVal texto As String)
    Dim c As Range, partes As Variant, i As Long, n As Long
    Set c = Nm(nome)
    c.Value = texto
    partes = Split(texto, vbLf)
    For i = LBound(partes) To UBound(partes)
        n = n + 1 + (Len(partes(i)) - 1) \ CHARS_LINHA
    Next i
    If n < 1 Then n = 1
    On Error Resume Next
    c.EntireRow.RowHeight = Application.WorksheetFunction.Min(409, 5 + 12 * n)
    On Error GoTo 0
End Sub

' Linhas preenchidas de uma lista da Ocorrencia, uma por linha com marcador
Public Function TextoLista(ByVal nome As String) As String
    Dim c As Range, s As String, t As String
    For Each c In Nm(nome).Cells
        t = Trim$(Replace(CStr(c.Value), vbLf, " "))
        If Len(t) > 0 Then
            If Left$(t, 1) = "-" Then t = Trim$(Mid$(t, 2))
            If s <> "" Then s = s & vbLf
            s = s & "•  " & t
        End If
    Next c
    If s = "" Then s = "-"
    TextoLista = s
End Function

Private Function Campo(ByVal nome As String) As String
    Dim v As Variant
    v = Nm(nome).Value
    If Vazio(v) Then Campo = "" Else Campo = Trim$(Replace(CStr(v), vbLf, " "))
End Function

' Junta pares (rotulo, nome) que estiverem preenchidos: "Rotulo: valor   ·   ..."
Private Function Juntar(ByVal pares As Variant) As String
    Dim i As Long, s As String, v As String
    For i = LBound(pares) To UBound(pares) Step 2
        v = Campo(CStr(pares(i + 1)))
        If v <> "" Then
            If s <> "" Then s = s & SEP
            s = s & pares(i) & ": " & v
        End If
    Next i
    If s = "" Then s = "-"
    Juntar = s
End Function

Private Function SimNaoQuem(ByVal rotulo As String, ByVal nome As String) As String
    Dim s As String
    s = Campo(nome)
    If s = "" Then s = "-"
    If Campo(nome & "Quem") <> "" Then s = s & " (" & Campo(nome & "Quem") & ")"
    SimNaoQuem = rotulo & ": " & s
End Function

Private Function TextoLaboratorio() As String
    Dim s As String, l2 As String
    If Campo("ocProg") <> "" Then s = "Em uso: " & Campo("ocProg")
    l2 = "Padrões preparados: " & IIf(Campo("ocPadroes") = "", "-", Campo("ocPadroes"))
    If Campo("ocPadroesQuais") <> "" Then l2 = l2 & " (" & Campo("ocPadroesQuais") & ")"
    If Campo("ocAr") <> "" Then
        l2 = l2 & SEP & "Sistema de ar: " & Campo("ocAr")
        If Campo("ocArStatus") <> "" Then l2 = l2 & " - " & Campo("ocArStatus")
    End If
    If Campo("ocCompressor") <> "" Then l2 = l2 & SEP & "Compressor: " & Campo("ocCompressor")
    If Campo("ocNitrogenio") <> "" Then l2 = l2 & SEP & "Nitrogênio: " & Campo("ocNitrogenio")
    If s <> "" Then s = s & vbLf
    TextoLaboratorio = s & l2
End Function

Private Function TextoPessoal() As String
    Dim s As String
    s = SimNaoQuem("Ausência", "ocAus") & SEP & SimNaoQuem("Troca combinada", "ocTroca") & SEP & _
        SimNaoQuem("Hora extra", "ocHE")
    If Campo("ocRecebe") <> "" Then s = s & SEP & "Letra que recebe: " & Campo("ocRecebe")
    TextoPessoal = s
End Function

Private Function TextoObservacoes() As String
    Dim s As String, o As String
    s = Juntar(Array("Hidrogênio no carvão", "ocH2Carvao", "Hidrogênio no coque", "ocH2Coque", _
                     "Coque Planta 04", "ocCoque04"))
    o = Trim$(CStr(Nm("ocObs").Value))
    If o <> "" Then
        If s = "-" Then s = o Else s = s & vbLf & o
    End If
    TextoObservacoes = s
End Function

Private Function LinhaDeInformacao(ByVal pref As String, ByVal confere As Boolean) As String
    Dim dt As Date, turno As String, turma As String, resp As String, s As String
    If LerTurnoPainel(dt, turno, turma, resp, False) Then
        s = Format$(dt, "dd\/mm\/yyyy") & SEP & turno
        If turma <> "" Then s = s & SEP & "Letra " & turma
        If resp <> "" Then s = s & SEP & "Técnico: " & resp
        If Not confere And Vazio(Nm(pref & "MES").Value) Then
            s = s & SEP & "(resultados do MES ainda não atualizados)"
        End If
    End If
    LinhaDeInformacao = s
End Function

' Oculta as linhas das analises sem tag no MES (e o titulo do grupo que ficar vazio).
' Chaves na coluna A: "G1" = titulo do grupo, "P01U3" = linha da analise.
Public Sub OcultarAnalisesSemTag(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long)
    Dim r As Long, chave As Variant, rGrupo As Long, visiveis As Long, oculta As Boolean
    On Error Resume Next
    For r = r1 To r2 + 1
        chave = ws.Cells(r, 1).Value
        If VarType(chave) = vbString Or r > r2 Then
            If r > r2 Or Left$(CStr(chave), 1) = "G" Then
                If rGrupo > 0 Then ws.Rows(rGrupo).Hidden = (visiveis = 0)
                rGrupo = r
                visiveis = 0
            ElseIf Len(chave) = 5 And Left$(CStr(chave), 1) = "P" Then
                oculta = Not ParamConfigurado(CLng(Mid$(CStr(chave), 2, 2)))
                ws.Rows(r).Hidden = oculta
                If Not oculta Then visiveis = visiveis + 1
            End If
        End If
    Next r
    On Error GoTo 0
End Sub

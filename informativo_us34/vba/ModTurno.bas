Option Explicit

' ============================================================================
'  ModTurno - fechamento do turno, historico e passagem de turno
' ============================================================================

' Botao 4 do Painel
Public Sub FecharTurno()
    Dim dt As Date, turno As String, turma As String, resp As String
    Dim resposta As Long, nSem As Long, quando As Date, pdf As String, msg As String

    PrepararEdicao
    If Not LerTurnoPainel(dt, turno, turma, resp, True) Then Exit Sub
    If turma = "" Or resp = "" Then
        Aviso "Informe a TURMA e o RESPONSÁVEL no Painel antes de fechar o turno.", vbExclamation, "Fechar turno"
        IrPainel
        Exit Sub
    End If

    If Not InformativoConfere(dt, turno) Then
        resposta = Aviso("Os resultados do MES ainda não foram carregados para este turno." & vbCrLf & vbCrLf & _
                          "SIM = buscar os dados do MES agora" & vbCrLf & _
                          "NÃO = fechar o turno sem os resultados" & vbCrLf & _
                          "CANCELAR = voltar", vbQuestion + vbYesNoCancel, "Fechar turno")
        If resposta = vbCancel Then Exit Sub
        If resposta = vbYes Then
            AtualizarMES
            If Not InformativoConfere(dt, turno) Then Exit Sub
        End If
    End If

    nSem = ContarSemStatus()
    If nSem > 0 Then
        If Aviso(nSem & " equipamento(s) estão sem status na Passagem de Turno." & vbCrLf & _
                  "Deseja fechar o turno mesmo assim?", vbQuestion + vbYesNo, "Fechar turno") = vbNo Then
            IrPassagem
            Exit Sub
        End If
    End If

    If ExisteNoHistorico(dt, turno) Then
        If Aviso("Este turno (" & Format$(dt, "dd\/mm\/yyyy") & " - " & turno & ") já foi fechado antes." & vbCrLf & _
                  "Deseja substituir o registro anterior?", vbQuestion + vbYesNo, "Fechar turno") = vbNo Then Exit Sub
        RemoverDoHistorico dt, turno
    End If

    Tela False
    quando = Now
    pdf = ExportarPDF(dt, turno, turma, False)
    GravarHistorico dt, turno, turma, resp, quando, pdf
    GravarRegistro dt, turno, turma, quando
    Tela True

    If UCase$(Left$(CfgTxt("cfgEmail"), 1)) = "S" Then CriarEmail dt, turno, turma, resp, pdf
    SalvarArquivo

    msg = "Turno fechado com sucesso!" & vbCrLf & vbCrLf & "- Resultados e passagem gravados no Histórico"
    If pdf <> "" Then msg = msg & vbCrLf & "- PDF: " & pdf
    msg = msg & vbCrLf & vbCrLf & "Deseja preparar o PRÓXIMO turno agora?" & vbCrLf & _
          "(limpa a passagem, avança Dia/Noite e define a turma que recebe)"
    If Aviso(msg, vbInformation + vbYesNo, "Fechar turno") = vbYes Then PrepararProximoTurno dt, turno
    IrPainel
End Sub

Private Function ContarSemStatus() As Long
    Dim c As Range, n As Long
    For Each c In Nm("ptStatus").Cells
        If Vazio(c.Value) Then n = n + 1
    Next c
    ContarSemStatus = n
End Function

Private Function ContarStatus(ByVal status As String) As Long
    Dim c As Range, n As Long
    For Each c In Nm("ptStatus").Cells
        If UCase$(Trim$(CStr(c.Value))) = UCase$(status) Then n = n + 1
    Next c
    ContarStatus = n
End Function

Private Function ExisteNoHistorico(ByVal dt As Date, ByVal turno As String) As Boolean
    Dim r As Long
    For r = HIST_ROW_HDR + 1 To UltimaLinha(shHistorico, 1)
        If MesmoTurno(shHistorico.Cells(r, 1).Value, shHistorico.Cells(r, 2).Value, dt, turno) Then
            ExisteNoHistorico = True
            Exit Function
        End If
    Next r
End Function

Private Sub RemoverDoHistorico(ByVal dt As Date, ByVal turno As String)
    Dim r As Long
    For r = UltimaLinha(shHistorico, 1) To HIST_ROW_HDR + 1 Step -1
        If MesmoTurno(shHistorico.Cells(r, 1).Value, shHistorico.Cells(r, 2).Value, dt, turno) Then shHistorico.Rows(r).Delete
    Next r
    For r = UltimaLinha(shRegistro, 1) To REG_ROW_HDR + 1 Step -1
        If MesmoTurno(shRegistro.Cells(r, 1).Value, shRegistro.Cells(r, 2).Value, dt, turno) Then shRegistro.Rows(r).Delete
    Next r
End Sub

Private Sub GravarHistorico(ByVal dt As Date, ByVal turno As String, ByVal turma As String, _
                            ByVal resp As String, ByVal quando As Date, ByVal pdf As String)
    Dim ws As Worksheet, lin As Long, p As Long, k As Long, v As Variant

    Set ws = shHistorico
    lin = UltimaLinha(ws, 1) + 1
    If lin <= HIST_ROW_HDR Then lin = HIST_ROW_HDR + 1

    ws.Cells(lin, 1).Value = dt
    ws.Cells(lin, 2).Value = turno
    ws.Cells(lin, 3).Value = turma
    ws.Cells(lin, 4).Value = resp
    ws.Cells(lin, 5).Value = Nm("ptRecebe").Value
    ws.Cells(lin, 6).Value = quando
    If InformativoConfere(dt, turno) Then ws.Cells(lin, 7).Value = Nm("iFonte").Value

    For p = 1 To NPARAM
        For k = 1 To 2
            If InformativoConfere(dt, turno) Then
                v = shInformativo.Cells(LinhaInformativo(p, k), INF_COL_RES).Value
                If ENumero(v) Then ws.Cells(lin, HIST_COL_RES1 + (p - 1) * 2 + (k - 1)).Value = CDbl(v)
            End If
        Next k
    Next p

    ws.Cells(lin, HIST_COL_NOK).Value = ContarStatus("NÃO OK")
    ws.Cells(lin, HIST_COL_NOK + 1).Value = Nm("ptPendencias").Value
    ws.Cells(lin, HIST_COL_NOK + 2).Value = Nm("ptObs").Value
    ws.Cells(lin, HIST_COL_NOK + 3).Value = pdf
End Sub

' Grava cada campo preenchido da passagem em formato "longo" (uma linha por informacao)
Private Sub GravarRegistro(ByVal dt As Date, ByVal turno As String, ByVal turma As String, ByVal quando As Date)
    Dim ws As Worksheet, r As Long, lin As Long, item As Variant, st As Variant, tx As Variant

    Set ws = shRegistro
    lin = UltimaLinha(ws, 1) + 1
    If lin <= REG_ROW_HDR Then lin = REG_ROW_HDR + 1

    r = 2
    Do While Not Vazio(shMapa.Cells(r, 1).Value)
        item = LerEspec(shMapa.Cells(r, 3).Value)
        st = LerEspec(shMapa.Cells(r, 4).Value)
        tx = LerEspec(shMapa.Cells(r, 5).Value)
        If Not (Vazio(st) And Vazio(tx)) Then
            ws.Cells(lin, 1).Value = dt
            ws.Cells(lin, 2).Value = turno
            ws.Cells(lin, 3).Value = turma
            ws.Cells(lin, 4).Value = shMapa.Cells(r, 1).Value
            ws.Cells(lin, 5).Value = LerEspec(shMapa.Cells(r, 2).Value)
            ws.Cells(lin, 6).Value = item
            If Not Vazio(st) Then ws.Cells(lin, 7).Value = st
            If Not Vazio(tx) Then ws.Cells(lin, 8).Value = tx
            ws.Cells(lin, 9).Value = quando
            lin = lin + 1
        End If
        r = r + 1
    Loop
End Sub

' Especificacao usada na aba _Mapa:
'   "@B9"   -> valor da celula B9 da aba Passagem de Turno
'   "#nome" -> valor do nome definido
'   texto   -> o proprio texto
Private Function LerEspec(ByVal espec As Variant) As Variant
    Dim s As String
    If Vazio(espec) Then
        LerEspec = Empty
        Exit Function
    End If
    s = CStr(espec)
    If Left$(s, 1) = "@" Then
        LerEspec = shPassagem.Range(Mid$(s, 2)).Value
    ElseIf Left$(s, 1) = "#" Then
        LerEspec = Nm(Mid$(s, 2)).Value
    Else
        LerEspec = s
    End If
End Function

Private Sub LimparEspec(ByVal espec As Variant)
    Dim s As String
    If Vazio(espec) Then Exit Sub
    s = CStr(espec)
    If Left$(s, 1) = "@" Then
        shPassagem.Range(Mid$(s, 2)).MergeArea.ClearContents
    ElseIf LCase$(Left$(s, 3)) = "#pt" Then
        Nm(Mid$(s, 2)).MergeArea.ClearContents
    End If
End Sub

Private Sub LimparCamposPassagem()
    Dim r As Long
    PrepararEdicao
    r = 2
    Do While Not Vazio(shMapa.Cells(r, 1).Value)
        LimparEspec shMapa.Cells(r, 4).Value
        LimparEspec shMapa.Cells(r, 5).Value
        r = r + 1
    Loop
End Sub

' Botao da aba Passagem de Turno
Public Sub LimparPassagem()
    If Aviso("Limpar TODOS os campos da Passagem de Turno?", vbQuestion + vbYesNo, "Limpar") = vbNo Then Exit Sub
    LimparCamposPassagem
End Sub

Private Sub PrepararProximoTurno(ByVal dt As Date, ByVal turno As String)
    Dim recebe As Variant
    recebe = Nm("ptRecebe").Value
    LimparCamposPassagem
    If EhDia(turno) Then
        Nm("pData").Value = dt
        Nm("pTurno").Value = TurnoNoite()
    Else
        Nm("pData").Value = CDate(CDbl(dt) + 1)
        Nm("pTurno").Value = TurnoDia()
    End If
    Nm("pTurma").Value = recebe
    Nm("pResp").MergeArea.ClearContents
    LimparInformativo
End Sub

Public Sub SalvarArquivo()
    On Error Resume Next
    If ThisWorkbook.Path <> "" Then ThisWorkbook.Save
End Sub

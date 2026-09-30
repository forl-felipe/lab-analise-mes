Option Explicit

' ============================================================================
'  ModTurno - finalizacao do turno e aba Ocorrencia (modelo DIARIO:
'  um arquivo por dia, com o Relatorio Dia e o Relatorio Noite)
' ============================================================================

' Coluna da aba _Mapa com os campos da Ocorrencia (A = nome, B = tipo, C = padrao)
Private Const MAPA_COL_OC As Long = 1

' Botao 4 do Painel: FINALIZAR TURNO
'  (1) garante os dados do MES do turno, (2) refaz e FIXA o relatorio do turno
'  (Relatorio Dia ou Relatorio Noite), (3) salva o arquivo, (4) oferece preparar a
'  Ocorrencia para o proximo turno. O relatorio finalizado nao e mais alterado pelo outro turno.
Public Sub FecharTurno()
    Dim dt As Date, turno As String, turma As String, resp As String, pref As String, msg As String

    PrepararEdicao
    If Not LerTurnoPainel(dt, turno, turma, resp, True) Then Exit Sub
    If turma = "" Or resp = "" Then
        Aviso "Preencha a LETRA e o TÉCNICO no Painel antes de finalizar o turno.", vbExclamation, "Finalizar turno"
        IrPainel
        Exit Sub
    End If
    pref = PrefixoRel(turno)
    If RelFinalizado(pref) Then
        If Aviso("O turno " & turno & " já foi finalizado em " & Format$(Nm(pref & "Final").Value, "dd\/mm hh:mm") & _
                 "." & vbCrLf & "Deseja finalizar de novo (o relatório será refeito com os dados atuais)?", _
                 vbQuestion + vbYesNo, "Finalizar turno", vbNo) = vbNo Then Exit Sub
        MarcarFinal pref, Empty
    End If

    ' Se os resultados deste turno nao estao completos no Informativo (nao carregados, ou
    ' carregados antes do fim do turno), busca agora (sem perguntar)
    If Not InformativoCompleto(dt, turno) Then
        gSilenciarSucesso = True
        AtualizarMES
        gSilenciarSucesso = False
        If Not InformativoConfere(dt, turno) Then
            If Aviso("Não foi possível trazer os resultados do MES para este turno." & vbCrLf & _
                     "Deseja finalizar mesmo assim (relatório sem resultados)?", _
                     vbQuestion + vbYesNo, "Finalizar turno", vbNo) = vbNo Then Exit Sub
        End If
    End If

    GerarRelatorio pref
    MarcarFinal pref, Now
    SalvarArquivo

    msg = "Turno " & turno & " finalizado!" & vbCrLf & vbCrLf & _
          "O relatório está na aba '" & AbaRel(pref).Name & "': use o botão 'Copiar relatório' " & _
          "e cole no e-mail com Ctrl+V." & vbCrLf & vbCrLf & _
          "Deseja PREPARAR a Ocorrência para o próximo turno?" & vbCrLf & _
          "(as 'Tarefas a realizar' deste turno passam para 'Tarefas realizadas' do próximo;" & vbCrLf & _
          " os demais campos são limpos; o relatório deste turno fica guardado)"
    If Aviso(msg, vbInformation + vbYesNo, "Finalizar turno") = vbYes Then
        LimparCampos True
        SalvarArquivo
    End If
    Mostrar AbaRel(pref)
End Sub

' Botao da aba Ocorrencia
Public Sub LimparOcorrencia()
    If Aviso("Limpar TODOS os campos da Ocorrência?", vbQuestion + vbYesNo, "Limpar", vbNo) = vbNo Then Exit Sub
    LimparCampos False
End Sub

' Limpa os campos da Ocorrencia (listados na aba _Mapa) e volta os valores padrao.
' levarPendentes = True: as "Tarefas a realizar" viram o ponto de partida das "Tarefas realizadas".
Private Sub LimparCampos(ByVal levarPendentes As Boolean)
    Dim r As Long, nome As String, padrao As String, c As Range, pend As Collection, i As Long, t As String

    Set pend = New Collection
    If levarPendentes Then
        For Each c In Nm("ocAReal").Cells
            t = Trim$(CStr(c.Value))
            If t <> "" Then pend.Add t
        Next c
    End If

    PrepararEdicao
    On Error Resume Next
    shOcorrencia.Unprotect Password:=SENHA
    On Error GoTo 0
    r = 2
    Do While Not Vazio(shMapa.Cells(r, MAPA_COL_OC).Value)
        nome = CStr(shMapa.Cells(r, MAPA_COL_OC).Value)
        padrao = CStr(shMapa.Cells(r, MAPA_COL_OC + 2).Value)
        For Each c In Nm(nome).Cells
            c.MergeArea.ClearContents
        Next c
        If padrao <> "" Then Nm(nome).Cells(1, 1).Value = padrao
        r = r + 1
    Loop
    i = 0
    For Each c In Nm("ocReal").Cells
        i = i + 1
        If i > pend.Count Then Exit For
        c.Value = pend(i)
    Next c
    ProtegerPlanilhas
End Sub

Public Sub SalvarArquivo()
    On Error Resume Next
    If ThisWorkbook.Path <> "" Then ThisWorkbook.Save
End Sub

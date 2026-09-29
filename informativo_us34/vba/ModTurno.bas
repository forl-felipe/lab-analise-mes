Option Explicit

' ============================================================================
'  ModTurno - finalizacao do turno e passagem de turno (modelo DIARIO:
'  um arquivo por dia, com o Resumo Dia e o Resumo Noite; sem historico acumulado)
' ============================================================================

' Botao 4 do Painel: FINALIZAR TURNO
'  (1) garante os dados do MES do turno, (2) refaz e FIXA o resumo do turno
'  (Resumo Dia ou Resumo Noite), (3) salva o arquivo, (4) oferece limpar a passagem
'  para o proximo turno. O resumo finalizado nao e mais alterado pelo outro turno.
Public Sub FecharTurno()
    Dim dt As Date, turno As String, turma As String, resp As String, pref As String, msg As String

    PrepararEdicao
    If Not LerTurnoPainel(dt, turno, turma, resp, True) Then Exit Sub
    If turma = "" Or resp = "" Then
        Aviso "Preencha a TURMA e o RESPONSÁVEL no Painel antes de finalizar o turno.", vbExclamation, "Finalizar turno"
        IrPainel
        Exit Sub
    End If
    pref = PrefixoResumo(turno)
    If ResumoFinalizado(pref) Then
        If Aviso("O turno " & turno & " já foi finalizado em " & Format$(Nm(pref & "Final").Value, "dd\/mm hh:mm") & _
                 "." & vbCrLf & "Deseja finalizar de novo (o resumo será refeito com os dados atuais)?", _
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
                     "Deseja finalizar mesmo assim (resumo sem resultados)?", _
                     vbQuestion + vbYesNo, "Finalizar turno", vbNo) = vbNo Then Exit Sub
        End If
    End If

    GerarResumo pref
    MarcarFinal pref, Now
    SalvarArquivo

    msg = "Turno " & turno & " finalizado!" & vbCrLf & vbCrLf & _
          "O resumo está na aba '" & AbaResumo(pref).Name & "': use os botões Copiar RESULTADOS e " & _
          "Copiar PASSAGEM para colar no e-mail." & vbCrLf & vbCrLf & _
          "Deseja LIMPAR a Passagem de Turno para o próximo turno?" & vbCrLf & _
          "(o resumo deste turno fica guardado)"
    If Aviso(msg, vbInformation + vbYesNo, "Finalizar turno") = vbYes Then
        LimparCamposPassagem
        SalvarArquivo
    End If
    Mostrar AbaResumo(pref)
End Sub

' Especificacao usada na aba _Mapa:
'   "@B9"   -> valor da celula B9 da aba Passagem de Turno
'   "#nome" -> valor do nome definido
'   texto   -> o proprio texto
Public Function LerEspec(ByVal espec As Variant) As Variant
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

Public Sub SalvarArquivo()
    On Error Resume Next
    If ThisWorkbook.Path <> "" Then ThisWorkbook.Save
End Sub

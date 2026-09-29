Option Explicit

' ============================================================================
'  ModResumo - aba "Resumo do Turno": duas imagens prontas para o e-mail
'   Imagem 1 (rsAreaResultados): resultados do turno - formulas ligadas ao Informativo
'   Imagem 2 (rsAreaPassagem)  : passagem de turno - textos montados aqui
' ============================================================================

Public Sub AtualizarResumo()
    Dim s As String
    PrepararEdicao
    s = LinhaDeInformacao()
    Nm("rsInfo").Value = s
    Nm("rsInfo2").Value = s
    Nm("rsEquipe").Value = Juntar("Equipe", "", False)
    Nm("rsNaoOper").Value = NaoOperando()
    Nm("rsMin").Value = Juntar("Batch", "", True)
    Nm("rsTestes").Value = TextoOuTraco(Nm("ptTestes").Value)
    Nm("rsPend").Value = TextoOuTraco(Nm("ptPendencias").Value)
    Nm("rsCom3").Value = Juntar("Comentário", "US3", False)
    Nm("rsCom4").Value = Juntar("Comentário", "US4", False)
    Nm("rsEmbarque").Value = Juntar("Embarque", "", False)
    Nm("rsObs").Value = TextoOuTraco(Nm("ptObs").Value)
End Sub

' Botoes de navegacao dentro da aba Resumo
Public Sub ResumoIrPassagem()
    On Error Resume Next
    shResumo.Activate
    ActiveWindow.ScrollRow = RES_ROW_PASSAGEM
    ActiveWindow.ScrollColumn = 1
End Sub

Public Sub ResumoIrResultados()
    On Error Resume Next
    shResumo.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

Private Function LinhaDeInformacao() As String
    Dim dt As Date, turno As String, turma As String, resp As String, s As String
    If LerTurnoPainel(dt, turno, turma, resp, False) Then
        s = Format$(dt, "dd\/mm\/yyyy") & "   ·   " & turno
        If turma <> "" Then s = s & "   ·   Turma " & turma
        If resp <> "" Then s = s & "   ·   " & resp
        If Not InformativoConfere(dt, turno) Then
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

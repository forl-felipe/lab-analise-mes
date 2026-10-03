Option Explicit

' ============================================================================
'  ModRegistro - grava o Lancamento na base (BD_Afericoes / tbl_Afericoes)
'
'  A aba _Staging monta, por formula, uma linha para cada resultado possivel
'  (mesmas 17 colunas da base). A macro so copia as linhas preenchidas, com a
'  identificacao do lancamento na coluna Origem:  LCP-000123 | 19x07 | ...
' ============================================================================

Private Const COL_REG As Long = 18      ' _Staging: 1 = linha a registrar
Private Const NCOL As Long = 21         ' colunas da base (17 originais + Turno, Registrado em, Turno do registro, Fora do turno)

' ---------------------------------------------------------------- registrar
Public Sub RegistrarLancamento()
    RegistrarEnsaios ""
End Sub

' Botoes "Enviar so este ensaio" de cada secao
Public Sub EnviarBLA()
    RegistrarEnsaios "BLA"
End Sub

Public Sub EnviarTAM()
    RegistrarEnsaios "TAM"
End Sub

Public Sub EnviarALP()
    RegistrarEnsaios "ALP"
End Sub

Public Sub EnviarUMI()
    RegistrarEnsaios "UMI"
End Sub

Public Sub EnviarCOM()
    RegistrarEnsaios "COM"
End Sub

Public Sub EnviarGRA()
    RegistrarEnsaios "GRA"
End Sub

Public Sub EnviarT515()
    RegistrarEnsaios "T515"
End Sub

Public Sub EnviarPEN()
    RegistrarEnsaios "PEN"
End Sub

Public Sub EnviarFIS()
    RegistrarEnsaios "FIS"
End Sub


' filtro = codigo do ensaio (BLA, TAM...) ou "" para todos os ensaios preenchidos
Public Sub RegistrarEnsaios(ByVal filtro As String)
    Dim d As Variant, turno As String, resp As String, letra As String
    Dim arr As Variant, i As Long, j As Long, ns As Long, k As Long
    Dim sel() As Long, saida() As Variant, origem As String, id As Long, lin As Long
    Dim ensaios As String, dup As String, erros As String, nc As String, nConf As Long, nNC As Long, nInfo As Long
    Dim e As Variant, partes As Variant, sem As String, dReal As Date, tReal As String, fora As Boolean

    d = Nm("pData").Value
    turno = Trim$(CStr(Nm("pTurno").Value))
    resp = Trim$(CStr(Nm("pResp").Value))
    letra = Trim$(CStr(Nm("pLetra").Value))
    If Not (VarType(d) = vbDate Or ENumero(d)) Then
        Aviso "Informe a data do lançamento.", vbExclamation
        Exit Sub
    End If
    If turno <> "07x19" And turno <> "19x07" Then
        Aviso "Informe o turno (07x19 ou 19x07).", vbExclamation
        Exit Sub
    End If
    If letra = "" Then
        Aviso "Informe a letra.", vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False
    NormalizarStatus Nm("penItens")
    Application.EnableEvents = True
    Recalcular
    arr = shStaging.Range("A2").Resize(NSTG, COL_REG + 1).Value
    ReDim sel(1 To NSTG)
    For i = 1 To NSTG
        If filtro <> "" And CStr(arr(i, COL_REG + 1)) <> filtro Then
            ' outro ensaio: fica na tela para ser enviado depois
        ElseIf IsError(arr(i, COL_REG)) Then
            erros = AdicionaUnico(erros, CStr(arr(i, COL_REG + 1)))
        ElseIf NumOu(arr(i, COL_REG), 0) = 1 Then
            For j = 1 To 17
                If IsError(arr(i, j)) Then
                    erros = AdicionaUnico(erros, CStr(arr(i, 1)))
                    Exit For
                End If
            Next j
            ns = ns + 1
            sel(ns) = i
            ensaios = AdicionaUnico(ensaios, CStr(arr(i, 1)))
        End If
    Next i
    If erros <> "" Then
        Aviso "Há valor inválido (texto em campo numérico?) em: " & Replace(Mid$(erros, 2), "|", ", ") & "." & _
              vbCrLf & "Corrija antes de registrar.", vbExclamation
        Exit Sub
    End If
    If ns = 0 Then
        If filtro = "" Then
            Aviso "Nenhum resultado preenchido para registrar.", vbInformation
        Else
            Aviso "Este ensaio não tem resultado preenchido.", vbInformation
        End If
        Exit Sub
    End If
    ' responsavel: o do ensaio ou, se vazio, o do cabecalho
    For k = 1 To ns
        If Trim$(CStr(arr(sel(k), 14))) = "" Then
            Aviso "Informe o responsável (no campo Responsável do ensaio ou no cabeçalho).", vbExclamation
            Exit Sub
        End If
    Next k

    ' resultado Nao conforme exige observacao no ensaio
    For k = 1 To ns
        i = sel(k)
        If CStr(arr(i, 13)) = "Não conforme" Then
            If Trim$(CStr(ObservacaoEnsaio(CStr(arr(i, COL_REG + 1))))) = "" Then sem = AdicionaUnico(sem, CStr(arr(i, 1)))
        End If
    Next k
    If sem <> "" Then
        Aviso "Há resultado NÃO CONFORME sem observação em:" & vbCrLf & "  - " & _
              Replace(Mid$(sem, 2), "|", vbCrLf & "  - ") & vbCrLf & vbCrLf & _
              "Escreva no campo Observação do ensaio o motivo ou a ação tomada (campo em vermelho) e registre de novo.", _
              vbExclamation, "Observação obrigatória"
        Exit Sub
    End If

    ' lancamento fora do turno atual (fica marcado na base)
    TurnoAtual dReal, tReal
    fora = (CLng(Int(CDbl(d))) <> CLng(Int(CDbl(dReal)))) Or (turno <> tReal)
    If fora Then
        If Aviso("Este lançamento é de " & Format$(CDate(d), "dd\/mm\/yyyy") & " " & turno & ", mas o turno atual é " & _
                 Format$(dReal, "dd\/mm\/yyyy") & " " & tReal & "." & vbCrLf & vbCrLf & _
                 "Ele será registrado como FORA DO TURNO. Continuar?", vbYesNo + vbQuestion, "Lançamento fora do turno", _
                 vbYes) <> vbYes Then Exit Sub
    End If

    ' lancamento repetido (mesmo ensaio, data e turno)
    partes = Split(Mid$(ensaios, 2), "|")
    For Each e In partes
        If ExisteRegistro(CStr(e), CLng(Int(CDbl(d))), turno) Then dup = dup & vbCrLf & "  - " & e
    Next e
    If dup <> "" Then
        If Aviso("Já existe registro em " & Format$(CDate(d), "dd\/mm\/yyyy") & " (" & turno & ") para:" & dup & _
                 vbCrLf & vbCrLf & "Registrar mesmo assim?", vbYesNo + vbQuestion, , vbYes) <> vbYes Then Exit Sub
    End If

    id = CLng(NumOu(Nm("cfgUltimoID").Value, 0)) + 1
    origem = "LCP-" & Format$(id, "000000") & " | " & turno & " | registrado em " & Format$(Now, "dd\/mm\/yyyy hh:mm")
    ReDim saida(1 To ns, 1 To NCOL)
    For k = 1 To ns
        i = sel(k)
        For j = 1 To 17
            If Vazio(arr(i, j)) Then saida(k, j) = Empty Else saida(k, j) = arr(i, j)
        Next j
        saida(k, 3) = CDate(Int(CDbl(d)))
        saida(k, 17) = origem
        saida(k, 18) = turno
        saida(k, 19) = Now
        saida(k, 20) = tReal
        saida(k, 21) = IIf(fora, "Sim", "Não")
        Select Case CStr(arr(i, 13))
            Case "Conforme": nConf = nConf + 1
            Case "Não conforme"
                nNC = nNC + 1
                If nNC <= 12 Then nc = nc & vbCrLf & "  - " & arr(i, 1) & " · " & arr(i, 4) & " · " & arr(i, 5)
            Case Else: nInfo = nInfo + 1
        End Select
    Next k

    On Error GoTo Falha
    Ampulheta True
    lin = UltimaLinhaBD() + 1
    shBD.Range("A" & lin).Resize(ns, NCOL).Value = saida
    shBD.Range("C" & lin).Resize(ns, 1).NumberFormat = "dd/mm/yyyy"
    shBD.Range("S" & lin).Resize(ns, 1).NumberFormat = "dd/mm/yyyy hh:mm"
    AjustarTabela
    Nm("cfgUltimoID").Value = id
    Nm("cfgUltimoLanc").Value = "LCP-" & Format$(id, "000000") & "  ·  " & Format$(CDate(d), "dd\/mm\/yyyy") & " " & _
        turno & "  ·  " & resp & "  ·  " & ns & " resultado(s)"
    If filtro = "" Then LimparTela Else LimparSecao filtro
    Ampulheta False
    On Error GoTo 0
    Aviso "Lançamento LCP-" & Format$(id, "000000") & " registrado: " & ns & " resultado(s)." & vbCrLf & vbCrLf & _
          "Conformes: " & nConf & "     Não conformes: " & nNC & IIf(nInfo > 0, "     Informativos: " & nInfo, "") & _
          IIf(nNC > 0, vbCrLf & vbCrLf & "Não conformidades:" & nc & IIf(nNC > 12, vbCrLf & "  ...", ""), ""), _
          IIf(nNC > 0, vbExclamation, vbInformation), "Registrar lançamento"
    Exit Sub
Falha:
    Ampulheta False
    Aviso "Erro ao gravar na base: " & Err.Description, vbCritical
End Sub

' Observacao digitada no ensaio (codigo BLA, TAM...), pelo mapa da aba Config
Private Function ObservacaoEnsaio(ByVal cod As String) As Variant
    Dim m As Variant, i As Long
    m = Nm("mapaInputs").Value
    For i = 1 To UBound(m, 1)
        If CStr(m(i, 1)) = cod Then
            ObservacaoEnsaio = shLancamento.Range(CStr(m(i, 4))).Cells(1, 1).Value
            If IsError(ObservacaoEnsaio) Then ObservacaoEnsaio = ""
            Exit Function
        End If
    Next i
    ObservacaoEnsaio = ""
End Function

Private Function AdicionaUnico(ByVal lista As String, ByVal item As String) As String
    If InStr(1, lista & "|", "|" & item & "|") = 0 Then lista = lista & "|" & item
    AdicionaUnico = lista
End Function

' Turno de uma linha da base: da coluna Origem ou, nos registros antigos de Blaine, da Observacao
Public Function TurnoDaLinha(ByVal origem As Variant, ByVal obs As Variant) As String
    Dim o As String, b As String
    If Not IsError(origem) Then o = CStr(origem)
    If Not IsError(obs) Then b = CStr(obs)
    If InStr(1, o, "07x19", vbTextCompare) > 0 Then
        TurnoDaLinha = "07x19"
    ElseIf InStr(1, o, "19x07", vbTextCompare) > 0 Then
        TurnoDaLinha = "19x07"
    ElseIf LCase$(Left$(b, 5)) = "07x19" Then
        TurnoDaLinha = "07x19"
    ElseIf LCase$(Left$(b, 5)) = "19x07" Then
        TurnoDaLinha = "19x07"
    End If
End Function

Private Function ExisteRegistro(ByVal ensaio As String, ByVal dia As Long, ByVal turno As String) As Boolean
    Dim ult As Long, a As Variant, i As Long, t As String
    ult = UltimaLinhaBD()
    If ult < 2 Then Exit Function
    a = shBD.Range("A2:Q" & ult).Value
    For i = 1 To UBound(a, 1)
        If CStr(a(i, 1)) = ensaio Then
            If ENumero(a(i, 3)) Or VarType(a(i, 3)) = vbDate Then
                If CLng(Int(CDbl(a(i, 3)))) = dia Then
                    t = TurnoDaLinha(a(i, 17), a(i, 16))
                    If t = "" Or t = turno Then
                        ExisteRegistro = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

' ---------------------------------------------------------------- base
Public Function UltimaLinhaBD() As Long
    Dim ur As Range, a As Variant, i As Long, n As Long
    Set ur = shBD.UsedRange
    n = ur.Row + ur.Rows.Count - 1
    UltimaLinhaBD = 1
    If n < 2 Then Exit Function
    a = shBD.Range("A1:A" & n).Value
    For i = n To 2 Step -1
        If Not Vazio(a(i, 1)) Then
            UltimaLinhaBD = i
            Exit Function
        End If
    Next i
End Function

' Estende a tabela tbl_Afericoes ate a ultima linha (o Power BI le a tabela)
Public Sub AjustarTabela()
    Dim lo As Object, ult As Long
    On Error Resume Next
    ult = UltimaLinhaBD()
    If ult < 2 Then ult = 2
    Set lo = shBD.ListObjects("tbl_Afericoes")
    If Not lo Is Nothing Then lo.Resize shBD.Range("A1:U" & ult)
End Sub

' ---------------------------------------------------------------- limpar
Public Sub LimparLancamento()
    If Aviso("Apagar os valores digitados na tela? (o que já foi registrado não é alterado)", _
             vbYesNo + vbQuestion, , vbYes) <> vbYes Then Exit Sub
    LimparTela
End Sub

' Apaga as entradas de um ensaio (depois de enviado sozinho)
Public Sub LimparSecao(ByVal cod As String)
    Dim m As Variant, i As Long, enderecos As Variant, k As Long
    On Error Resume Next
    Application.EnableEvents = False
    Desproteger shLancamento
    m = Nm("mapaInputs").Value
    For i = 1 To UBound(m, 1)
        If CStr(m(i, 1)) = cod Then
            enderecos = Split(CStr(m(i, 3)), ";")
            For k = LBound(enderecos) To UBound(enderecos)
                If Len(enderecos(k)) > 0 Then shLancamento.Range(enderecos(k)).ClearContents
            Next k
        End If
    Next i
    Proteger shLancamento
    Application.EnableEvents = True
End Sub

' Apaga as entradas dos ensaios (mantem data, turno, letra, responsavel e as tags)
Public Sub LimparTela()
    Dim m As Variant, i As Long, enderecos As Variant, k As Long
    On Error Resume Next
    Application.EnableEvents = False
    Desproteger shLancamento
    m = Nm("mapaInputs").Value
    For i = 1 To UBound(m, 1)
        enderecos = Split(CStr(m(i, 3)), ";")
        For k = LBound(enderecos) To UBound(enderecos)
            If Len(enderecos(k)) > 0 Then shLancamento.Range(enderecos(k)).ClearContents
        Next k
    Next i
    Proteger shLancamento
    Application.EnableEvents = True
End Sub

' Peneiradores: preenche com OK os itens ainda vazios
Public Sub MarcarPeneiradoresOK()
    Dim c As Range
    On Error Resume Next
    Application.EnableEvents = False
    For Each c In Nm("penItens").Cells
        If Vazio(c.Value) Then c.Value = "OK"
    Next c
    Application.EnableEvents = True
End Sub

' ---------------------------------------------------------------- desfazer
Public Sub DesfazerUltimo()
    Dim id As Long, pref As String, ult As Long, a As Variant, i As Long, n As Long, info As String
    id = CLng(NumOu(Nm("cfgUltimoID").Value, 0))
    If id <= 0 Then
        Aviso "Não há lançamento para desfazer.", vbInformation
        Exit Sub
    End If
    pref = "LCP-" & Format$(id, "000000") & " |"
    ult = UltimaLinhaBD()
    If ult >= 2 Then
        a = shBD.Range("A2:Q" & ult).Value
        For i = 1 To UBound(a, 1)
            If Left$(CStr(a(i, 17)), Len(pref)) = pref Then
                n = n + 1
                If info = "" Then info = Format$(CDate(a(i, 3)), "dd\/mm\/yyyy") & " · " & TurnoDaLinha(a(i, 17), "") & _
                                        " · " & a(i, 14)
            End If
        Next i
    End If
    If n = 0 Then
        Aviso "O lançamento LCP-" & Format$(id, "000000") & " não está na base (já foi desfeito?).", vbInformation
        Nm("cfgUltimoID").Value = id - 1
        Exit Sub
    End If
    If Aviso("Apagar da base o lançamento LCP-" & Format$(id, "000000") & " (" & info & "), com " & n & _
             " resultado(s)?", vbYesNo + vbExclamation, "Desfazer último lançamento", vbYes) <> vbYes Then Exit Sub
    Ampulheta True
    For i = UBound(a, 1) To 1 Step -1
        If Left$(CStr(a(i, 17)), Len(pref)) = pref Then shBD.Rows(i + 1).Delete
    Next i
    AjustarTabela
    Nm("cfgUltimoID").Value = id - 1
    Nm("cfgUltimoLanc").Value = "LCP-" & Format$(id, "000000") & " desfeito em " & Format$(Now, "dd\/mm\/yyyy hh:mm")
    Ampulheta False
    Aviso n & " resultado(s) apagado(s) da base.", vbInformation
End Sub

' ---------------------------------------------------------------- importar planilha antiga
' Le a aba BD_Afericoes de uma planilha mensal anterior (modelo antigo) e acrescenta na base
' os resultados que ainda nao estao nela.
Public Sub ImportarArquivoAntigo()
    Dim f As Variant, n As Long
    f = Application.GetOpenFilename("Planilhas do Excel (*.xlsm;*.xlsx;*.xls),*.xlsm;*.xlsx;*.xls", , _
                                    "Escolha a planilha de calibração antiga")
    If VarType(f) = vbBoolean Then Exit Sub
    n = ImportarDe(CStr(f))
    If n >= 0 Then Aviso n & " resultado(s) importado(s).", vbInformation, "Importar planilha antiga"
End Sub

Public Function ImportarDe(ByVal caminho As String) As Long
    Dim wbO As Workbook, ws As Worksheet, a As Variant, i As Long, j As Long, c As Long
    Dim col(1 To 21) As Long, cab As String, chaves As String, k As String, n As Long
    Dim saida() As Variant, lin As Long, ult As Long, b As Variant, nomeArq As String
    ImportarDe = -1
    On Error GoTo Falha
    Ampulheta True
    Set wbO = Workbooks.Open(caminho, 0, True)
    nomeArq = wbO.Name
    On Error Resume Next
    Set ws = wbO.Worksheets("BD_Afericoes")
    On Error GoTo Falha
    If ws Is Nothing Then
        wbO.Close False
        Ampulheta False
        Aviso "A planilha escolhida não tem a aba BD_Afericoes.", vbExclamation
        Exit Function
    End If
    a = ws.UsedRange.Value
    wbO.Close False
    Set wbO = Nothing
    ' colunas pelo nome do cabecalho
    For j = 1 To 21
        cab = CStr(shBD.Cells(1, j).Value)
        For c = 1 To UBound(a, 2)
            If StrComp(Trim$(CStr(a(1, c))), cab, vbTextCompare) = 0 Then col(j) = c
        Next c
    Next j
    If col(1) = 0 Or col(3) = 0 Or col(7) = 0 Then
        Ampulheta False
        Aviso "A aba BD_Afericoes da planilha escolhida não tem as colunas esperadas.", vbExclamation
        Exit Function
    End If
    ' chaves do que ja existe
    ult = UltimaLinhaBD()
    If ult >= 2 Then
        b = shBD.Range("A2:Q" & ult).Value
        For i = 1 To UBound(b, 1)
            chaves = chaves & "|" & ChaveLinha(b(i, 1), b(i, 3), b(i, 4), b(i, 5), b(i, 7), b(i, 13))
        Next i
    End If
    chaves = chaves & "|"
    ReDim saida(1 To UBound(a, 1), 1 To 21)
    For i = 2 To UBound(a, 1)
        If (VarType(a(i, col(3))) = vbDate Or ENumero(a(i, col(3)))) And _
           (Not Vazio(a(i, col(7))) Or (col(13) > 0 And Not Vazio(Valor(a, i, col(13))))) Then
            k = ChaveLinha(a(i, col(1)), a(i, col(3)), Valor(a, i, col(4)), Valor(a, i, col(5)), a(i, col(7)), _
                           Valor(a, i, col(13)))
            If InStr(1, chaves, "|" & k & "|") = 0 Then
                chaves = chaves & k & "|"
                n = n + 1
                For j = 1 To 21
                    If col(j) > 0 Then
                        If Vazio(a(i, col(j))) Then saida(n, j) = Empty Else saida(n, j) = a(i, col(j))
                    End If
                Next j
                saida(n, 3) = CDate(Int(CDbl(a(i, col(3)))))
                saida(n, 17) = "Importado: " & nomeArq & IIf(Vazio(saida(n, 17)), "", " | " & saida(n, 17))
                If Vazio(saida(n, 18)) Then saida(n, 18) = TurnoDaLinha("", saida(n, 16))
                If Vazio(saida(n, 18)) Then saida(n, 18) = Empty
            End If
        End If
    Next i
    If n > 0 Then
        lin = UltimaLinhaBD() + 1
        shBD.Range("A" & lin).Resize(n, 21).Value = Recorta(saida, n)
        shBD.Range("C" & lin).Resize(n, 1).NumberFormat = "dd/mm/yyyy"
        AjustarTabela
    End If
    Ampulheta False
    ImportarDe = n
    Exit Function
Falha:
    Ampulheta False
    On Error Resume Next
    If Not wbO Is Nothing Then wbO.Close False
    Aviso "Não foi possível importar: " & Err.Description, vbExclamation
End Function

Private Function Valor(ByVal a As Variant, ByVal i As Long, ByVal c As Long) As Variant
    If c = 0 Then Valor = Empty Else Valor = a(i, c)
End Function

Private Function ChaveLinha(ByVal ens As Variant, ByVal d As Variant, ByVal eq As Variant, ByVal pa As Variant, _
                            ByVal v As Variant, ByVal res As Variant) As String
    Dim sv As String
    If ENumero(v) Then sv = Format$(Round(CDbl(v), 4), "0.0000") Else sv = CStr(v)
    ChaveLinha = CStr(ens) & "#" & CLng(Int(CDbl(d))) & "#" & CStr(eq) & "#" & CStr(pa) & "#" & sv & "#" & CStr(res)
End Function

Private Function Recorta(ByVal a As Variant, ByVal n As Long) As Variant
    Dim r() As Variant, i As Long, j As Long
    ReDim r(1 To n, 1 To 21)
    For i = 1 To n
        For j = 1 To 21
            r(i, j) = a(i, j)
        Next j
    Next i
    Recorta = r
End Function

Option Explicit

' ============================================================================
'  ModGeral - utilidades do Controle de Calibracao e Afericao (Laboratorio Fisico)
' ============================================================================

Public Const XL_VISIBLE As Long = -1
Public Const XL_HIDDEN As Long = 0
Public Const XL_VERYHIDDEN As Long = 2
Public Const XL_UP As Long = -4162
Public Const XL_WAIT As Long = 2
Public Const XL_DEFAULT As Long = -4143

' Quando True, as mensagens nao sao exibidas (testes automaticos)
Public gSilencioso As Boolean
' Ultima mensagem exibida (usada nos testes automaticos)
Public gUltimaMsg As String

' ---------------------------------------------------------------- abertura
Public Sub Inicializar()
    On Error Resume Next
    shBD.Visible = XL_HIDDEN
    shLimites.Visible = XL_HIDDEN
    shConfig.Visible = XL_HIDDEN
    shStaging.Visible = XL_VERYHIDDEN
    shGraficos.Visible = XL_VERYHIDDEN
    ' ao abrir, o cabecalho volta para o turno atual (pode ser alterado; o relogio ao lado nao)
    DefinirTurnoAtual
    shLancamento.Activate
End Sub

' Data e turno pelo relogio (antes das 07h ainda e o turno Noite do dia anterior)
Public Sub DefinirTurnoAtual()
    Dim d As Date, t As String
    TurnoAtual d, t
    Nm("pData").Value = d
    Nm("pTurno").Value = t
End Sub

' Turno real (relogio do computador)
Public Sub TurnoAtual(ByRef d As Date, ByRef t As String)
    Dim h As Double
    h = CDbl(Now) - Int(CDbl(Now))
    d = CDate(Int(CDbl(Now)))
    If h >= 7 / 24 And h < 19 / 24 Then
        t = "07x19"
    ElseIf h >= 19 / 24 Then
        t = "19x07"
    Else
        d = CDate(CDbl(d) - 1)
        t = "19x07"
    End If
End Sub

' ---------------------------------------------------------------- nomes e valores
Public Function Nm(ByVal nome As String) As Range
    Set Nm = ThisWorkbook.Names(nome).RefersToRange
End Function

Public Function CfgTxt(ByVal nome As String) As String
    Dim v As Variant
    v = Nm(nome).Value
    If IsError(v) Then CfgTxt = "" Else CfgTxt = Trim$(CStr(v))
End Function

Public Function Vazio(ByVal v As Variant) As Boolean
    If IsError(v) Then
        Vazio = True
    ElseIf IsEmpty(v) Then
        Vazio = True
    ElseIf VarType(v) = vbString Then
        Vazio = (Len(Trim$(v)) = 0)
    Else
        Vazio = False
    End If
End Function

Public Function ENumero(ByVal v As Variant) As Boolean
    If Vazio(v) Then ENumero = False Else ENumero = IsNumeric(v)
End Function

Public Function NumOu(ByVal v As Variant, ByVal padrao As Double) As Double
    If ENumero(v) Then NumOu = CDbl(v) Else NumOu = padrao
End Function

' Recalcula a planilha (no LibreOffice, aba por aba)
Public Sub Recalcular()
    Dim ws As Worksheet
    On Error Resume Next
    Application.Calculate
    If Err.Number <> 0 Then
        Err.Clear
        For Each ws In ThisWorkbook.Worksheets
            ws.Calculate
        Next ws
    End If
End Sub

' ---------------------------------------------------------------- navegacao
Public Sub IrPainel()
    On Error Resume Next
    shPainel.Activate
End Sub

Public Sub IrLancamento()
    On Error Resume Next
    shLancamento.Activate
End Sub

' ---------------------------------------------------------------- mensagens
Public Function Aviso(ByVal msg As String, Optional ByVal estilo As Long = 64, _
                      Optional ByVal titulo As String = "Controle de Calibração", _
                      Optional ByVal padrao As Long = 6) As Long
    gUltimaMsg = msg
    If gSilencioso Then
        Aviso = padrao
    Else
        Aviso = MsgBox(msg, estilo, titulo)
    End If
End Function

Public Sub Ampulheta(ByVal ligada As Boolean)
    On Error Resume Next
    If ligada Then Application.Cursor = XL_WAIT Else Application.Cursor = XL_DEFAULT
    Application.ScreenUpdating = Not ligada
End Sub

' ---------------------------------------------------------------- protecao
Public Sub Desproteger(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=SENHA
End Sub

Public Sub Proteger(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=SENHA, AllowFormattingColumns:=True, AllowFormattingRows:=True
    If Err.Number <> 0 Then
        Err.Clear
        ws.Protect SENHA
    End If
End Sub

' ---------------------------------------------------------------- backup ao salvar
Public Sub BackupAposSalvar()
    Dim pasta As String
    On Error Resume Next
    pasta = CfgTxt("cfgBackup")
    If pasta = "" Then Exit Sub
    If Right$(pasta, 1) <> "\" Then pasta = pasta & "\"
    If Len(Dir$(pasta, vbDirectory)) = 0 Then Exit Sub
    ThisWorkbook.SaveCopyAs pasta & "Controle de Calibração LCP - backup.xlsm"
End Sub

' ---------------------------------------------------------------- digitacao no Lancamento
' Peneiradores: "ok", "nao ok", "sem tag" digitados de qualquer forma viram o texto padrao
Public Sub AoAlterarLancamento(ByVal alvo As Range)
    Dim area As Range
    On Error GoTo Fim
    Set area = Application.Intersect(alvo, Nm("penItens"))
    If area Is Nothing Then Exit Sub
    Application.EnableEvents = False
    NormalizarStatus area
Fim:
    On Error Resume Next
    Application.EnableEvents = True
End Sub

Public Sub NormalizarStatus(ByVal area As Range)
    Dim c As Range, t As String, n As Long
    On Error Resume Next
    For Each c In area.Cells
        n = n + 1
        If n > 400 Then Exit For
        If Not Vazio(c.Value) Then
            t = UCase$(Trim$(CStr(c.Value)))
            t = Replace(Replace(t, "Ã", "A"), "ã", "A")
            If t = "OK" Or t = "O" Then
                If c.Value <> "OK" Then c.Value = "OK"
            ElseIf t = "NAO OK" Or t = "NOK" Or t = "N" Or t = "NAO" Or t = "NAOOK" Then
                If c.Value <> "NÃO OK" Then c.Value = "NÃO OK"
            ElseIf t = "SEM TAG" Or t = "S" Or t = "SEMTAG" Then
                If c.Value <> "SEM TAG" Then c.Value = "SEM TAG"
            End If
        End If
    Next c
End Sub

Option Explicit

' ============================================================================
'  ModGeral - utilidades, navegacao, protecao e leitura do Painel
' ============================================================================

Public Const XL_VISIBLE As Long = -1
Public Const XL_HIDDEN As Long = 0
Public Const XL_VERYHIDDEN As Long = 2
Public Const XL_UP As Long = -4162
Public Const XL_WAIT As Long = 2
Public Const XL_DEFAULT As Long = -4143

' Quando True, as mensagens nao sao exibidas (usado em testes automaticos)
Public gSilencioso As Boolean
' Quando True, AtualizarMES nao mostra a mensagem de sucesso (usado pelo Salvar turno)
Public gSilenciarSucesso As Boolean
' Minutos apos o inicio de um turno em que o Painel ainda sugere o turno que acabou de terminar
Public Const TOLERANCIA_MIN As Long = 60

' ---------------------------------------------------------------- abertura
Public Sub Inicializar()
    On Error Resume Next
    ProtegerPlanilhas
    shDadosMES.Visible = XL_HIDDEN
    shConfig.Visible = XL_HIDDEN
    shRegistro.Visible = XL_HIDDEN
    shMapa.Visible = XL_VERYHIDDEN
    ' Sempre abre no turno de referencia pelo relogio (pode ser trocado no Painel)
    DefinirTurnoAtual
    shPainel.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

' ---------------------------------------------------------------- nomes
Public Function Nm(ByVal nome As String) As Range
    Set Nm = ThisWorkbook.Names(nome).RefersToRange
End Function

Public Function CfgTxt(ByVal nome As String) As String
    Dim v As Variant
    v = Nm(nome).Value
    If IsError(v) Then
        CfgTxt = ""
    Else
        CfgTxt = Trim$(CStr(v))
    End If
End Function

Public Function CfgCel(ByVal p As Long, ByVal coluna As Long) As Variant
    CfgCel = shConfig.Cells(CFG_ROW1 + p - 1, coluna).Value
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
    If Vazio(v) Then
        ENumero = False
    Else
        ENumero = IsNumeric(v)
    End If
End Function

Public Function UltimaLinha(ByVal ws As Worksheet, ByVal coluna As Long) As Long
    ' Ultima linha preenchida na coluna (sem Range.End, para funcionar tambem sem janela ativa)
    Dim r As Long
    r = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Do While r > 1
        If Not Vazio(ws.Cells(r, coluna).Value) Then Exit Do
        r = r - 1
    Loop
    UltimaLinha = r
End Function

' ---------------------------------------------------------------- turnos
Public Function EhDia(ByVal turno As String) As Boolean
    EhDia = (LCase$(Left$(Trim$(turno), 3)) = "dia")
End Function

Public Function TurnoDia() As String
    TurnoDia = CStr(Nm("lstTurno").Cells(1, 1).Value)
End Function

Public Function TurnoNoite() As String
    TurnoNoite = CStr(Nm("lstTurno").Cells(2, 1).Value)
End Function

Public Function HoraInicioDia() As Double
    Dim v As Variant
    v = Nm("cfgInicioDia").Value
    If ENumero(v) Then
        HoraInicioDia = CDbl(v)
        If HoraInicioDia >= 1 Then HoraInicioDia = HoraInicioDia / 24
    Else
        HoraInicioDia = 7 / 24
    End If
End Function

Public Function InicioTurno(ByVal dt As Date, ByVal turno As String) As Date
    Dim x As Double
    x = Int(CDbl(dt)) + HoraInicioDia()
    If Not EhDia(turno) Then x = x + 0.5
    InicioTurno = CDate(x)
End Function

Public Function MesmoTurno(ByVal vData As Variant, ByVal vTurno As Variant, _
                           ByVal dt As Date, ByVal turno As String) As Boolean
    MesmoTurno = False
    If Vazio(vData) Or Vazio(vTurno) Then Exit Function
    If VarType(vData) = vbDate Or IsNumeric(vData) Then
        If Int(CDbl(CDate(vData))) = Int(CDbl(dt)) Then
            MesmoTurno = (EhDia(CStr(vTurno)) = EhDia(turno))
        End If
    End If
End Function

' Turno de referencia pelo relogio:
'  - durante o turno, o proprio turno;
'  - nos primeiros TOLERANCIA_MIN minutos de um turno, o turno que acabou de terminar
'    (e quando normalmente se fecha o informativo do turno anterior).
Public Sub DefinirTurnoAtual()
    Dim d As Date, t As String
    PrepararEdicao
    TurnoDeReferencia CDbl(Now), d, t
    Nm("pData").Value = d
    Nm("pTurno").Value = t
End Sub

Public Sub TurnoDeReferencia(ByVal agora As Double, ByRef d As Date, ByRef t As String)
    Dim ref As Double, h As Double, hIni As Double, dia As Double
    ref = agora - TOLERANCIA_MIN / 1440#
    hIni = HoraInicioDia()
    dia = Int(ref)
    h = ref - dia
    If h >= hIni And h < hIni + 0.5 Then
        d = CDate(dia)
        t = TurnoDia()
    ElseIf h >= hIni + 0.5 Then
        d = CDate(dia)
        t = TurnoNoite()
    Else
        d = CDate(dia - 1)
        t = TurnoNoite()
    End If
End Sub

' Le e valida a selecao do Painel
Public Function LerTurnoPainel(ByRef dt As Date, ByRef turno As String, ByRef turma As String, _
                               ByRef resp As String, ByVal avisar As Boolean) As Boolean
    Dim v As Variant
    LerTurnoPainel = False
    v = Nm("pData").Value
    If VarType(v) = vbDate Then
        dt = CDate(Int(CDbl(v)))
    ElseIf ENumero(v) Then
        dt = CDate(Int(CDbl(v)))
    ElseIf Not Vazio(v) And IsDate(v) Then
        dt = CDate(Int(CDbl(CDate(v))))
    Else
        If avisar Then
            Aviso "Informe a DATA do turno no Painel.", vbExclamation, "Painel"
            IrPainel
        End If
        Exit Function
    End If
    turno = CfgTxt("pTurno")
    If turno = "" Then
        If avisar Then
            Aviso "Selecione o TURNO (Dia ou Noite) no Painel.", vbExclamation, "Painel"
            IrPainel
        End If
        Exit Function
    End If
    turma = CfgTxt("pTurma")
    resp = CfgTxt("pResp")
    LerTurnoPainel = True
End Function

' ---------------------------------------------------------------- protecao
Public Sub ProtegerPlanilhas()
    Dim ws As Variant
    On Error Resume Next
    For Each ws In Array(shPainel, shInformativo, shPassagem)
        ws.Unprotect Password:=SENHA
        ws.Protect Password:=SENHA, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                   UserInterfaceOnly:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
                   AllowFiltering:=True
    Next ws
End Sub

Public Sub PrepararEdicao()
    ' Reaplica a protecao com UserInterfaceOnly (se perde ao reabrir o arquivo)
    ProtegerPlanilhas
End Sub

' ---------------------------------------------------------------- navegacao
Private Sub Mostrar(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Visible = XL_VISIBLE
    ws.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

Public Sub IrPainel()
    Mostrar shPainel
    ' Configuracoes e dados brutos ficam ocultos; acesso pelo botao "Configuracoes" do Painel
    On Error Resume Next
    shConfig.Visible = XL_HIDDEN
    shDadosMES.Visible = XL_HIDDEN
    shRegistro.Visible = XL_HIDDEN
End Sub

Public Sub IrInformativo()
    Mostrar shInformativo
End Sub

Public Sub IrPassagem()
    Mostrar shPassagem
End Sub

Public Sub IrHistorico()
    Mostrar shHistorico
End Sub

Public Sub IrRegistro()
    Mostrar shRegistro
End Sub

Public Sub IrConfig()
    Mostrar shConfig
End Sub

Public Sub MostrarDadosMES()
    Mostrar shDadosMES
End Sub

Public Sub OcultarDadosMES()
    shDadosMES.Visible = XL_HIDDEN
    Mostrar shConfig
End Sub

Public Sub BtnTurnoAtual()
    DefinirTurnoAtual
    Aviso "Painel ajustado para o turno atual: " & Format$(Nm("pData").Value, "dd\/mm\/yyyy") & _
           " - " & CfgTxt("pTurno") & ".", vbInformation, "Turno atual"
End Sub

Public Function Aviso(ByVal msg As String, Optional ByVal estilo As Long = 64, _
                      Optional ByVal titulo As String = "Informativo US3/US4", _
                      Optional ByVal padrao As Long = 6) As Long
    If gSilencioso Then
        Aviso = padrao
    Else
        Aviso = MsgBox(msg, estilo, titulo)
    End If
End Function

Public Sub Ampulheta(ByVal ligada As Boolean)
    On Error Resume Next
    If ligada Then Application.Cursor = XL_WAIT Else Application.Cursor = XL_DEFAULT
End Sub

Public Sub Tela(ByVal ligada As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = ligada
End Sub

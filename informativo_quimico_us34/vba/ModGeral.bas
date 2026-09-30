Option Explicit

' ============================================================================
'  ModGeral - utilidades (Laboratorio Quimico - Usinas 3 e 4)
'  A planilha funciona quase toda por formulas: o VBA so busca o MES,
'  copia as imagens e oculta as linhas vazias dos resumos.
' ============================================================================

Public Const XL_VISIBLE As Long = -1
Public Const XL_HIDDEN As Long = 0
Public Const XL_VERYHIDDEN As Long = 2
Public Const XL_WAIT As Long = 2
Public Const XL_DEFAULT As Long = -4143

' Quando True, as mensagens nao sao exibidas (usado em testes automaticos)
Public gSilencioso As Boolean

' ---------------------------------------------------------------- abertura
Public Sub Inicializar()
    Dim v As Variant, ref As Double
    On Error Resume Next
    shDadosMES.Visible = XL_HIDDEN
    shConfig.Visible = XL_HIDDEN
    shMapa.Visible = XL_VERYHIDDEN
    ' Arquivo novo (copia do modelo): a data do dia vem pelo relogio.
    ' Antes das 07h ainda e o dia anterior (turno Noite).
    v = Nm("pData").Value
    If Vazio(v) Then
        ref = CDbl(Now) - HoraInicioDia()
        Nm("pData").Value = CDate(Int(ref))
    End If
    shPreenchimento.Activate
End Sub

' ---------------------------------------------------------------- nomes e celulas
Public Function Nm(ByVal nome As String) As Range
    Set Nm = ThisWorkbook.Names(nome).RefersToRange
End Function

Public Function CfgTxt(ByVal nome As String) As String
    Dim v As Variant
    v = Nm(nome).Value
    If IsError(v) Then CfgTxt = "" Else CfgTxt = Trim$(CStr(v))
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
    If Vazio(v) Then ENumero = False Else ENumero = IsNumeric(v)
End Function

' Hora de inicio do turno Dia (Configuracoes; padrao 07:00), em fracao de dia
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

' ---------------------------------------------------------------- navegacao (abas ocultas)
Public Sub IrPreenchimento()
    On Error Resume Next
    shPreenchimento.Activate
    shConfig.Visible = XL_HIDDEN
    shDadosMES.Visible = XL_HIDDEN
End Sub

Public Sub MostrarDadosMES()
    On Error Resume Next
    shDadosMES.Visible = XL_VISIBLE
    shDadosMES.Activate
End Sub

Public Sub OcultarDadosMES()
    On Error Resume Next
    shDadosMES.Visible = XL_HIDDEN
    shConfig.Activate
End Sub

' ---------------------------------------------------------------- mensagens
Public Function Aviso(ByVal msg As String, Optional ByVal estilo As Long = 64, _
                      Optional ByVal titulo As String = "Relatório Químico US3/US4", _
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

' ---------------------------------------------------------------- protecao das abas
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

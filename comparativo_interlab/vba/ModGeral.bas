Option Explicit

' ============================================================================
'  ModGeral - utilidades do Comparativo de Medicoes de Qualidade (Germano x Ubu)
'  O painel funciona por formulas: o VBA busca o MES, guarda/abre batches do
'  historico, ajusta o tamanho das series dos graficos e gera imagem/PDF/e-mail.
' ============================================================================

Public Const XL_VISIBLE As Long = -1
Public Const XL_HIDDEN As Long = 0
Public Const XL_VERYHIDDEN As Long = 2
Public Const XL_WAIT As Long = 2
Public Const XL_DEFAULT As Long = -4143
Public Const TITULO_APP As String = "Comparativo Germano × Ubu"

' Quando True, as mensagens nao sao exibidas (testes automaticos)
Public gSilencioso As Boolean
' Batch da aba Batch alterado desde a ultima gravacao/abertura
Public gAlterado As Boolean
' Escrita feita pelo VBA (nao marca o batch como alterado)
Public gCarregando As Boolean
' Validacao automatica no LibreOffice (graficos nao sao manipulados)
Public gTesteLO As Boolean
' Ultima mensagem exibida (usada na validacao automatica)
Public gUltimaMsg As String

' ---------------------------------------------------------------- abertura
Public Sub Inicializar()
    On Error Resume Next
    shDadosMES.Visible = XL_HIDDEN
    shGraf.Visible = XL_HIDDEN
    shLeituras.Visible = XL_HIDDEN
    shAcompReg.Visible = XL_HIDDEN
    shConfig.Visible = VisConfig()
    ' abas que a variante nao usa: muito ocultas (o codigo e o mesmo nas duas planilhas)
    If VARIANTE = "ACOMP" Then
        shPainel.Visible = XL_VERYHIDDEN
        shHist.Visible = XL_VERYHIDDEN
        ' em Reexibir aparece so Configuracoes
        shDadosMES.Visible = XL_VERYHIDDEN
        shGraf.Visible = XL_VERYHIDDEN
        shLeituras.Visible = XL_VERYHIDDEN
        shAcompReg.Visible = XL_VERYHIDDEN
    ElseIf VARIANTE = "COMP" Then
        shAcomp.Visible = XL_VERYHIDDEN
    End If
    gAlterado = False
    AtualizarUltimos
    AjustarGraficos
    shBatch.Activate
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

Public Function EData(ByVal v As Variant) As Boolean
    If Vazio(v) Then
        EData = False
    ElseIf VarType(v) = vbDate Or IsNumeric(v) Then
        EData = (CDbl(v) > 30000)
    Else
        EData = False
    End If
End Function

' Data/hora devolvida pelo Aspen: pode vir como data do Excel ou como TEXTO ("12/10/2026 15:00:00").
' Texto: le dia/mes/ano (ou mes/dia/ano) sem depender do idioma do Windows; com ref, fica com a
' interpretacao mais proxima de ref.
Public Function ComoData(ByVal v As Variant, ByRef d As Date, Optional ByVal ref As Variant) As Boolean
    Dim t As String, p As Variant, dt As Variant, hm As Variant, a As Long, b As Long, y As Long
    Dim hh As Long, mi As Long, ss As Long, d1 As Date, d2 As Date, ok1 As Boolean, ok2 As Boolean
    On Error GoTo Nao
    If EData(v) Then
        d = CDate(v)
        ComoData = True
        Exit Function
    End If
    If VarType(v) <> vbString Then Exit Function
    t = Trim$(v)
    If Len(t) < 8 Or InStr(t, "/") = 0 Then Exit Function
    p = Split(Application.WorksheetFunction.Trim(t), " ")
    dt = Split(p(0), "/")
    If UBound(dt) <> 2 Then Exit Function
    If Not (IsNumeric(dt(0)) And IsNumeric(dt(1)) And IsNumeric(dt(2))) Then Exit Function
    a = CLng(dt(0)): b = CLng(dt(1)): y = CLng(dt(2))
    If y < 100 Then y = y + 2000
    If UBound(p) >= 1 Then
        hm = Split(p(1), ":")
        hh = CLng(hm(0))
        If UBound(hm) >= 1 Then mi = CLng(hm(1))
        If UBound(hm) >= 2 Then ss = CLng(Val(hm(2)))
        If UBound(p) >= 2 Then
            If UCase$(Left$(p(2), 1)) = "P" And hh < 12 Then hh = hh + 12
            If UCase$(Left$(p(2), 1)) = "A" And hh = 12 Then hh = 0
        End If
    End If
    ok1 = (b >= 1 And b <= 12 And a >= 1 And a <= 31)
    ok2 = (a >= 1 And a <= 12 And b >= 1 And b <= 31)
    If ok1 Then d1 = DateSerial(y, b, a) + TimeSerial(hh, mi, ss)
    If ok2 Then d2 = DateSerial(y, a, b) + TimeSerial(hh, mi, ss)
    If ok1 And ok2 And Not IsMissing(ref) Then
        If EData(ref) Then
            If Abs(CDbl(d2) - CDbl(ref)) < Abs(CDbl(d1) - CDbl(ref)) Then ok1 = False
        End If
    End If
    If ok1 Then
        d = d1
    ElseIf ok2 Then
        d = d2
    Else
        Exit Function
    End If
    ComoData = True
    Exit Function
Nao:
    ComoData = False
End Function

' Numero devolvido pelo Aspen: numero ou TEXTO com ponto decimal ("2052.0"), independente do idioma do Windows
Public Function ComoNumero(ByVal v As Variant, ByRef x As Double) As Boolean
    Dim t As String
    On Error GoTo Nao
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        t = Trim$(v)
        If t = "" Then Exit Function
        If t Like "*[!0-9.,eE+-]*" Then Exit Function
        If InStr(t, ",") > 0 And InStr(t, ".") = 0 Then t = Replace(t, ",", ".")
        x = Val(t)
        ComoNumero = True
    ElseIf IsNumeric(v) Then
        x = CDbl(v)
        ComoNumero = True
    End If
    Exit Function
Nao:
    ComoNumero = False
End Function

' Chave / rotulo / decimais do parametro p (1..NP)
Public Function PChave(ByVal p As Long) As String
    PChave = Split(P_CHAVES, "|")(p - 1)
End Function

Public Function PRotulo(ByVal p As Long) As String
    PRotulo = Split(P_ROTULOS, "|")(p - 1)
End Function

Public Function PDec(ByVal p As Long) As Long
    PDec = CLng(Split(P_DECIMAIS, "|")(p - 1))
End Function

' Parametro pelo rotulo ou pela chave (0 = nao achou)
Public Function PIndice(ByVal texto As String) As Long
    Dim p As Long
    For p = 1 To NP
        If StrComp(texto, PRotulo(p), vbTextCompare) = 0 Or StrComp(texto, PChave(p), vbTextCompare) = 0 Then
            PIndice = p
            Exit Function
        End If
    Next p
End Function

' Coluna (aba Batch) do valor do parametro p no laboratorio lab (0 = A/Germano, 1 = B/Ubu)
Public Function ColValor(ByVal p As Long, ByVal lab As Long) As Long
    ColValor = B_COL_V1 + 2 * (p - 1) + lab
End Function

' Minutos entre o inicio da janela de 2 h e o horario mostrado na linha (padrao 30: 11:00 -> 11:30)
Public Function MinutosRotulo() As Double
    MinutosRotulo = 30
    If ENumero(Nm("cfgMinAntes").Value) Then MinutosRotulo = CDbl(Nm("cfgMinAntes").Value)
End Function

' Recalcula a pasta (o LibreOffice sem janela nao aceita Application.Calculate)
Public Sub Recalcular()
    On Error Resume Next
    Application.Calculate
    If Err.Number <> 0 Then
        Err.Clear
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            ws.Calculate
        Next ws
    End If
End Sub

' ---------------------------------------------------------------- navegacao
Public Sub IrPainel()
    On Error Resume Next
    AjustarGraficos
    shPainel.Activate
    shPainel.Range("A1").Select
End Sub

Public Sub IrBatch()
    On Error Resume Next
    shBatch.Activate
    shDadosMES.Visible = IIf(VARIANTE = "ACOMP", XL_VERYHIDDEN, XL_HIDDEN)
    shConfig.Visible = VisConfig()
End Sub

' Planilha de acompanhamento: Configuracoes e uma aba oculta comum (botao direito numa aba > Reexibir).
' Comparativo: muito oculta (nao aparece em Reexibir); abre so pelo botao Configuracoes, com a senha
Public Function VisConfig() As Long
    If VARIANTE = "ACOMP" Then VisConfig = XL_HIDDEN Else VisConfig = XL_VERYHIDDEN
End Function

Public Sub AbrirConfiguracoes()
    Dim s As Variant
    s = InputBox("Senha da aba Configurações:", "Configurações")
    If CStr(s) = "" Then Exit Sub
    If CStr(s) <> SENHA_CFG Then
        Aviso "Senha incorreta.", vbExclamation, "Configurações"
        Exit Sub
    End If
    On Error Resume Next
    shConfig.Visible = XL_VISIBLE
    shConfig.Activate
End Sub

Public Sub FecharConfiguracoes()
    IrBatch
End Sub


Public Sub MostrarDadosMES()
    On Error Resume Next
    shDadosMES.Visible = XL_VISIBLE
    shDadosMES.Activate
End Sub

Public Sub OcultarDadosMES()
    On Error Resume Next
    shConfig.Visible = XL_VISIBLE
    shConfig.Activate
    shDadosMES.Visible = XL_HIDDEN
End Sub

' ---------------------------------------------------------------- mensagens
Public Function Aviso(ByVal msg As String, Optional ByVal estilo As Long = 64, _
                      Optional ByVal titulo As String = TITULO_APP, _
                      Optional ByVal padrao As Long = 6) As Long
    gUltimaMsg = gUltimaMsg & " // " & msg
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
    ws.Unprotect Password:=SenhaAba(ws)
    If ws.ProtectContents Then
        Err.Clear
        ws.Unprotect SenhaAba(ws)        ' sem argumento nomeado (LibreOffice)
    End If
End Sub

' Configuracoes, a aba de preenchimento e a de acompanhamento tem senha (SENHA_CFG): o desenho delas e por posicao
' (linhas e colunas fixas) e apagar uma linha ou coluna desmonta os blocos. As demais abas: SENHA
Private Function SenhaAba(ByVal ws As Worksheet) As String
    If ws Is shConfig Or ws Is shBatch Or ws Is shAcomp Then SenhaAba = SENHA_CFG Else SenhaAba = SENHA
End Function

Public Sub Proteger(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=SenhaAba(ws), DrawingObjects:=True, Contents:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True
    If Err.Number <> 0 Then
        Err.Clear
        ws.Protect SenhaAba(ws)
    End If
End Sub

Public Sub MarcarAlterado()
    If Not gCarregando Then gAlterado = True
End Sub

' Liga/desliga eventos e atualizacao de tela durante gravacoes do VBA
Public Sub ModoCarga(ByVal ligado As Boolean)
    On Error Resume Next
    gCarregando = ligado
    Application.EnableEvents = Not ligado
    Application.ScreenUpdating = Not ligado
End Sub

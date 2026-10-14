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
    shImportar.Visible = XL_HIDDEN
    gAlterado = False
    AjustarGraficos
    shPainel.Activate
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
    shImportar.Visible = XL_HIDDEN
    shDadosMES.Visible = XL_HIDDEN
End Sub

Public Sub IrEnviar()
    On Error Resume Next
    shEnviar.Activate
End Sub

Public Sub IrImportar()
    On Error Resume Next
    shImportar.Visible = XL_VISIBLE
    shImportar.Activate
    shImportar.Range("B7").Select
End Sub

Public Sub MostrarDadosMES()
    On Error Resume Next
    shDadosMES.Visible = XL_VISIBLE
    shDadosMES.Activate
End Sub

Public Sub OcultarDadosMES()
    On Error Resume Next
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
    ws.Unprotect Password:=SENHA
End Sub

Public Sub Proteger(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=SENHA, DrawingObjects:=True, Contents:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True
    If Err.Number <> 0 Then
        Err.Clear
        ws.Protect SENHA
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

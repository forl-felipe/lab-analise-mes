Option Explicit

' ============================================================================
'  ModGeral - utilidades do Monitoramento S.E. (Espessador x Batch)
'  O Report funciona por formulas (abas Historico e _Graf): o VBA busca o MES,
'  grava o dia no historico, ajusta o tamanho das series dos graficos e copia a imagem.
' ============================================================================

Public Const XL_VISIBLE As Long = -1
Public Const XL_HIDDEN As Long = 0
Public Const XL_VERYHIDDEN As Long = 2
Public Const XL_WAIT As Long = 2
Public Const XL_DEFAULT As Long = -4143
Public Const TITULO_APP As String = "Monitoramento S.E. - Espessador x Batch"

' Quando True, as mensagens nao sao exibidas (testes automaticos)
Public gSilencioso As Boolean
' Escrita feita pelo VBA (eventos das abas nao reagem)
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
    shConfig.Visible = XL_VERYHIDDEN
    If Not ENumero(Nm("rDias").Value) Then
        Desproteger shPainel
        Nm("rDias").Value = CfgNum("cfgDiasPadrao", 15)
        Proteger shPainel
    End If
    ' o report do dia e o de ontem: a aba Dia abre em ontem (se estiver num dia mais antigo)
    If DiaDe(Nm("dData").Value) < Int(CDbl(Now)) - 1 Then
        ModoCarga True
        Desproteger shDia
        Nm("dData").Value = CDate(Int(CDbl(Now)) - 1)
        Proteger shDia
        ModoCarga False
    End If
    CarregarDia
    ' painel abre no ultimo dia salvo no historico
    If UltimoDiaHist() > 0 And DiaDe(Nm("rFim").Value) <> UltimoDiaHist() Then
        ModoCarga True
        Desproteger shPainel
        Nm("rFim").Value = CDate(UltimoDiaHist())
        Proteger shPainel
        ModoCarga False
    End If
    AjustarGraficos True
    shDia.Activate
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

Public Function CfgNum(ByVal nome As String, ByVal padrao As Double) As Double
    Dim v As Variant
    v = Nm(nome).Value
    If ENumero(v) Then CfgNum = CDbl(v) Else CfgNum = padrao
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

' ---------------------------------------------------------------- itens do MES (1..NI)
Public Function IChave(ByVal k As Long) As String
    IChave = Split(I_CHAVES, "|")(k - 1)
End Function

' Item pela chave (0 = nao achou)
Public Function IIndice(ByVal chave As String) As Long
    Dim k As Long
    For k = 1 To NI
        If StrComp(chave, IChave(k), vbTextCompare) = 0 Then
            IIndice = k
            Exit Function
        End If
    Next k
End Function

' Campo do item k na aba Configuracoes (C_COL_TAG, C_COL_JAN...)
Public Function ICfg(ByVal k As Long, ByVal col As Long) As String
    Dim v As Variant
    v = shConfig.Cells(C_ROW_IT0 + k - 1, col).Value
    If IsError(v) Then ICfg = "" Else ICfg = Limpo(CStr(v))
End Function

Public Function IRotulo(ByVal k As Long) As String
    IRotulo = ICfg(k, C_COL_ROT)
End Function

Public Function Limpo(ByVal t As String) As String
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, ChrW(160), "")
    Limpo = Trim$(t)
End Function

' Coluna da aba Historico de cada campo
Public Function HCol(ByVal chave As String) As Long
    Select Case LCase$(chave)
        Case "data": HCol = H_COL_DATA
        Case "prodc3": HCol = H_COL_PRODC3
        Case "prodc2": HCol = H_COL_PRODC2
        Case "espc3": HCol = H_COL_ESPC3
        Case "espc2": HCol = H_COL_ESPC2
        Case "eb4": HCol = H_COL_EB4
        Case "eb6": HCol = H_COL_EB6
        Case "tk01": HCol = H_COL_TK01
        Case "tk02": HCol = H_COL_TK02
        Case "fator": HCol = H_COL_FATOR
        Case "lamac3": HCol = H_COL_LAMAC3
        Case "lamac2": HCol = H_COL_LAMAC2
        Case "concc3": HCol = H_COL_CONCC3
        Case "concc2": HCol = H_COL_CONCC2
        Case "atual": HCol = H_COL_ATUAL
        Case "obs": HCol = H_COL_OBS
    End Select
End Function

' Data como numero inteiro (dia), ou 0
Public Function DiaDe(ByVal v As Variant) As Long
    If EData(v) Then DiaDe = CLng(Int(CDbl(v)))
End Function

Public Function TxtDia(ByVal d As Date) As String
    TxtDia = Format$(d, "dd\/mm\/yyyy")
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
Public Sub IrDiagnostico()
    On Error Resume Next
    If Not EData(Nm("xData").Value) Or DiaDe(Nm("xData").Value) >= Int(CDbl(Now)) Then
        Desproteger shDiag
        Nm("xData").Value = CDate(Int(CDbl(Now)) - 1)
        Proteger shDiag
    End If
    shDiag.Activate
End Sub

Public Sub IrDia()
    On Error Resume Next
    shDia.Activate
    shDadosMES.Visible = XL_HIDDEN
    shConfig.Visible = XL_VERYHIDDEN
End Sub

' Configuracoes fica muito oculta (nao aparece em Reexibir); abre so com a senha
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
    IrDia
End Sub

Public Sub MostrarDadosMES()
    On Error Resume Next
    shDadosMES.Visible = XL_VISIBLE
    shDadosMES.Activate
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

' Configuracoes tem senha propria (SENHA_CFG); as demais abas, SENHA
Private Function SenhaAba(ByVal ws As Worksheet) As String
    If ws Is shConfig Then SenhaAba = SENHA_CFG Else SenhaAba = SENHA
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

' Liga/desliga eventos e atualizacao de tela durante gravacoes do VBA
Public Sub ModoCarga(ByVal ligado As Boolean)
    On Error Resume Next
    gCarregando = ligado
    Application.EnableEvents = Not ligado
    Application.ScreenUpdating = Not ligado
End Sub

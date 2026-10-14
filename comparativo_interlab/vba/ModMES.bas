Option Explicit

' ============================================================================
'  ModMES - resultados de Germano e de Ubu no MES (Aspen IP.21)
'  Botao "Buscar no MES" (aba Batch):
'    Ubu     : janelas de 2 h a partir do "Inicio no processo em Ubu"
'    Germano : as mesmas janelas deslocadas pela defasagem (inicio do batch em Germano)
'  Uma consulta GetCalculationValues por laboratorio, com tags/servidores/mapas em
'  texto literal (mesmo formato dos relatorios do laboratorio). Tags recusadas pelo
'  MES ("Tag Name ... is invalid") saem da consulta, que e refeita com as demais.
' ============================================================================

' Segundos sem mudanca nos resultados para considerar a consulta concluida
Private Const ESTAVEL_SEG As Double = 2
' Consulta que respondeu "Success" sem nenhuma linha: "sem dados" depois deste tempo
Private Const SEM_DADOS_SEG As Double = 15
Private Const MAX_RODADAS As Long = 10
' Celulas da aba Dados_MES (consulta de Germano e de Ubu)
Private Const ANC_G As String = "A10"
Private Const OUT_G As String = "A12"
Private Const ANC_U As String = "A40"
Private Const OUT_U As String = "A42"

Private mOcupado As Boolean
' Tags recusadas pelo MES nesta sessao, no formato |TAG1|TAG2|
Private mInvalidas As String

' ---------------------------------------------------------------- botao
Public Sub BuscarMES()
    Dim n As Long, nG As Long, iniU As Date, iniG As Date
    Dim resG As Variant, resU As Variant, simulado As Boolean, okG As Boolean, okU As Boolean
    Dim qG As Long, qU As Long, msg As String, semTag As String

    If mOcupado Then
        Aviso "Consulta ao MES em andamento. Aguarde.", vbInformation, "MES"
        Exit Sub
    End If
    Recalcular
    If Not EData(Nm("bIniU").Value) Or Not EData(Nm("gJanU").Value) Then
        Aviso "Informe o início do batch em Ubu (fase UBU3 da tela de eventos do MES) na aba Batch.", vbExclamation, "MES"
        On Error Resume Next
        Nm("bIniU").Select
        Exit Sub
    End If
    n = 0
    If ENumero(Nm("bNHor").Value) Then n = CLng(Nm("bNHor").Value)
    If n < 1 Or n > NMAX Then
        Aviso "Nº de horários deve ficar entre 1 e " & NMAX & ".", vbExclamation, "MES"
        Exit Sub
    End If
    simulado = (Left$(UCase$(CfgTxt("cfgFonte")), 3) = "SIM")
    semTag = TagsFaltando()
    If Not simulado And semTag = "TODAS" Then
        Aviso "Nenhuma tag do MES cadastrada." & vbCrLf & vbCrLf & _
              "Preencha as tags de Germano e de Ubu na aba Configurações (ou use Fonte = SIMULAÇÃO para treinar).", _
              vbExclamation, "MES"
        shConfig.Activate
        Exit Sub
    End If

    ' janelas de 2 h: Ubu a partir da 1a janela completa depois do inicio (UBU3); Germano idem a partir
    ' do inicio do bombeamento (sem inicio de Germano = mesmos horarios de Ubu)
    iniU = CDate(Nm("gJanU").Value)
    nG = n
    If EData(Nm("gJanG").Value) Then
        iniG = CDate(Nm("gJanG").Value)
        If ENumero(Nm("gNG").Value) Then nG = CLng(Nm("gNG").Value)
    Else
        iniG = iniU
    End If
    If nG > n Then nG = n

    mOcupado = True
    Ampulheta True
    On Error GoTo Falha
    If simulado Then
        If nG > 0 Then resG = DadosSimulados(0, nG, iniG)
        resU = DadosSimulados(1, n, iniU)
        okG = True
        okU = True
    Else
        Application.StatusBar = "Consultando MES - " & CfgTxt("bLabA") & "..."
        If nG > 0 Then okG = ConsultarLab(0, iniG, nG, resG) Else okG = True
        Application.StatusBar = "Consultando MES - " & CfgTxt("bLabB") & "..."
        okU = ConsultarLab(1, iniU, n, resU)
    End If
    Application.StatusBar = False
    If Not okG And Not okU Then GoTo Fim

    qG = GravarResultados(0, resG, nG, okG)
    qU = GravarResultados(1, resU, n, okU)
    Nm("mesTempo").Value = "'" & Format$(Now, "dd\/mm\/yyyy hh:mm:ss") & IIf(simulado, " (SIMULAÇÃO)", "")
    AjustarGraficos
    msg = "Resultados gravados no batch " & CfgTxt("bId") & ":" & vbCrLf & vbCrLf & _
          "  " & CfgTxt("bLabA") & ": " & qG & IIf(okG, "", " (consulta falhou)") & _
          "   (" & nG & " janelas a partir de " & Format$(iniG, "dd\/mm hh:mm") & ")" & vbCrLf & _
          "  " & CfgTxt("bLabB") & ": " & qU & IIf(okU, "", " (consulta falhou)") & _
          "   (" & n & " janelas a partir de " & Format$(iniU, "dd\/mm hh:mm") & ")"
    If Not simulado And (CDbl(iniU) + n * HORAS_SLOT / 24 > CDbl(Now) Or CDbl(iniG) + nG * HORAS_SLOT / 24 > CDbl(Now)) Then
        msg = msg & vbCrLf & vbCrLf & "Batch ainda em andamento: horários futuros ficam em branco."
    End If
    If semTag <> "" And semTag <> "TODAS" Then
        msg = msg & vbCrLf & vbCrLf & "Sem tag (lançamento manual): " & semTag
    End If
    If simulado Then msg = msg & vbCrLf & vbCrLf & "Fonte: SIMULAÇÃO (dados fictícios)."
    Aviso msg, vbInformation, "MES"
Fim:
    Application.StatusBar = False
    Ampulheta False
    mOcupado = False
    Exit Sub
Falha:
    Application.StatusBar = False
    Ampulheta False
    mOcupado = False
    Aviso "Erro ao consultar o MES: " & Err.Description, vbExclamation, "MES"
End Sub

' ---------------------------------------------------------------- tags (aba Configuracoes)
Public Function TagCfg(ByVal p As Long, ByVal lab As Long, ByVal campo As Long) As String
    Dim v As Variant
    v = shConfig.Cells(C_ROW_TAG0 + 2 * (p - 1) + lab, campo).Value
    If IsError(v) Then TagCfg = "" Else TagCfg = Trim$(CStr(v))
End Function

Private Function TagInvalida(ByVal t As String) As Boolean
    If t = "" Then Exit Function
    TagInvalida = (InStr(1, mInvalidas, "|" & UCase$(t) & "|") > 0)
End Function

' Analise sem tag: "S.E. Germano, P Ubu..." ; "TODAS" se nenhuma tag
Private Function TagsFaltando() As String
    Dim p As Long, lab As Long, r As String, n As Long
    For p = 1 To NP
        For lab = 0 To 1
            If TagCfg(p, lab, C_COL_TAG) = "" Then
                n = n + 1
                If r <> "" Then r = r & ", "
                r = r & PRotulo(p) & " " & IIf(lab = 0, CfgTxt("bLabA"), CfgTxt("bLabB"))
            End If
        Next lab
    Next p
    If n = 2 * NP Then r = "TODAS"
    TagsFaltando = r
End Function

' Parametros do laboratorio com tag valida (lista "1,3,5")
Private Function ParamsLab(ByVal lab As Long) As String
    Dim p As Long, t As String, r As String
    For p = 1 To NP
        t = TagCfg(p, lab, C_COL_TAG)
        If t <> "" And Not TagInvalida(t) Then
            If r <> "" Then r = r & ","
            r = r & p
        End If
    Next p
    ParamsLab = r
End Function

' ---------------------------------------------------------------- consulta de um laboratorio
' res(1..n, 1..NP): valores das janelas de 2 h (Empty = sem resultado)
Private Function ConsultarLab(ByVal lab As Long, ByVal ini As Date, ByVal n As Long, ByRef res As Variant) As Boolean
    Dim anc As Range, sai As Range, ps As String, rodada As Long, limite As Double, nJan As Long
    Dim fim As Date, t0 As Single, d() As Variant, lst As Variant, i As Long, r As Long, s As Long, v As Variant
    Dim estado As Long, nomeLab As String, invalidas As String

    ReDim d(1 To n, 1 To NP)
    res = d
    nomeLab = IIf(lab = 0, CfgTxt("bLabA"), CfgTxt("bLabB"))
    If lab = 0 Then
        Set anc = shDadosMES.Range(ANC_G)
        Set sai = shDadosMES.Range(OUT_G)
    Else
        Set anc = shDadosMES.Range(ANC_U)
        Set sai = shDadosMES.Range(OUT_U)
    End If
    ps = ParamsLab(lab)
    If ps = "" Then
        ConsultarLab = True          ' nada a consultar (sem tags): lancamento manual
        Exit Function
    End If
    ' nunca pede ao MES janelas que ainda nao comecaram
    nJan = n
    If CDbl(ini) + nJan * HORAS_SLOT / 24 > CDbl(Now) Then
        nJan = Int((CDbl(Now) - CDbl(ini)) * 24 / HORAS_SLOT) + 1
        If nJan > n Then nJan = n
    End If
    If nJan < 1 Then
        ConsultarLab = True          ' batch ainda nao chegou a este laboratorio
        Exit Function
    End If
    fim = CDate(CDbl(ini) + nJan * HORAS_SLOT / 24)
    limite = 120
    If ENumero(Nm("cfgTimeout").Value) Then limite = CDbl(Nm("cfgTimeout").Value)

    GravarPeriodo lab, ini, fim
    For rodada = 0 To MAX_RODADAS
        ps = ParamsLab(lab)
        If ps = "" Then Exit For
        LimparSaida sai
        anc.Formula = PrefixoAspen(anc) & ArgumentosLab(lab, ps, anc, sai)
        If Application.Calculation <> -4105 Then anc.Calculate
        estado = Esperar(anc, sai, ini, limite)
        If estado = 3 Then
            ' tag recusada: guarda e consulta de novo sem ela
            If ColetarInvalidas(anc, sai) = 0 Then Exit For
        ElseIf estado = 0 And rodada = 0 Then
            ' nada voltou: recria a matriz de saida no formato original e tenta mais uma vez
            RestaurarSaida anc, sai, nJan, UBound(Split(ps, ",")) + 2
            anc.Calculate
            estado = Esperar(anc, sai, ini, limite)
            Exit For
        Else
            Exit For
        End If
    Next rodada

    invalidas = ListaInvalidas(lab)
    If invalidas <> "" Then
        Nm("mesInvalidas").Value = "'" & Replace(invalidas, vbCrLf, " ")
        Aviso "O MES não reconheceu estas tags de " & nomeLab & " (ficam em branco):" & vbCrLf & vbCrLf & _
              invalidas & vbCrLf & "Corrigir na aba Configurações.", vbExclamation, "MES - tags inválidas"
    End If
    If estado = 0 Then
        DesligarConsulta anc, sai
        If Aviso("O MES não respondeu à consulta de " & nomeLab & " (" & Format$(ini, "dd\/mm\/yyyy hh:mm") & ")." & _
                 vbCrLf & vbCrLf & "Verificar: suplemento Aspen ativo, servidor, tags e formato da data " & _
                 "(aba Configurações)." & vbCrLf & vbCrLf & "Abrir a aba Dados_MES?", vbExclamation + vbYesNo, _
                 "MES sem resposta", 7) = vbYes Then MostrarDadosMES
        Exit Function
    End If

    ' le a matriz: 1a coluna = inicio da janela; demais = analises na ordem de ps
    lst = Split(ParamsLab(lab), ",")
    If estado = 1 Then
        For r = 1 To NMAX + 2
            v = sai.Cells(r, 1).Value
            If Not EData(v) Then Exit For
            s = CLng((CDbl(CDate(v)) - CDbl(ini)) * 24 / HORAS_SLOT) + 1
            If s >= 1 And s <= n Then
                For i = LBound(lst) To UBound(lst)
                    v = sai.Cells(r, 2 + i - LBound(lst)).Value
                    If ValorValido(CLng(lst(i)), lab, v) Then d(s, CLng(lst(i))) = CDbl(v)
                Next i
            End If
        Next r
    End If
    res = d
    DesligarConsulta anc, sai
    ConsultarLab = True
End Function

' 0 = sem resposta, 1 = resultados prontos, 2 = respondeu sem dados, 3 = tag recusada
Private Function Esperar(ByVal anc As Range, ByVal sai As Range, ByVal ini As Date, ByVal limite As Double) As Long
    Dim t0 As Single, tEst As Single, assin As String, ant As String, v As Variant, pronto As Boolean
    t0 = Timer
    tEst = Timer
    Do
        DoEvents
        If Timer < t0 Then t0 = t0 - 86400
        If Timer < tEst Then tEst = tEst - 86400
        If RespostaTagInvalida(anc, sai) Then
            Esperar = 3
            Exit Function
        End If
        v = sai.Cells(1, 1).Value
        pronto = False
        If EData(v) Then pronto = (Abs(CDbl(CDate(v)) - CDbl(ini)) <= HORAS_SLOT / 24 + 1 / 1440)
        assin = Assinatura(anc, sai)
        If assin <> ant Then
            ant = assin
            tEst = Timer
        End If
        If pronto And Timer - tEst >= ESTAVEL_SEG Then
            Esperar = 1
            Exit Function
        End If
        If Not pronto And Timer - t0 >= SEM_DADOS_SEG Then
            If LCase$(Left$(Trim$(anc.Text), 7)) = "success" And Vazio(sai.Cells(1, 1).Value) Then
                Esperar = 2
                Exit Function
            End If
        End If
        If Timer - t0 > limite Then
            If pronto Then Esperar = 1 Else Esperar = 0
            Exit Function
        End If
    Loop
End Function

Private Function Assinatura(ByVal anc As Range, ByVal sai As Range) As String
    Dim arr As Variant, i As Long, j As Long, n As Long, soma As Double
    arr = sai.Resize(NMAX + 1, NP + 1).Value
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            If ENumero(arr(i, j)) Then
                n = n + 1
                soma = soma + CDbl(arr(i, j))
            End If
        Next j
    Next i
    Assinatura = anc.Text & "|" & n & "|" & soma
End Function

Private Function TextoResposta(ByVal anc As Range, ByVal sai As Range) As String
    On Error Resume Next
    TextoResposta = anc.Text & " " & sai.Cells(1, 1).Text
End Function

Private Function RespostaTagInvalida(ByVal anc As Range, ByVal sai As Range) As Boolean
    Dim t As String, i As Long, j As Long, tag As String
    t = TextoResposta(anc, sai)
    i = InStr(1, t, "Tag Name ", vbTextCompare)
    Do While i > 0
        j = InStr(i + 9, t, " is invalid", vbTextCompare)
        If j = 0 Then Exit Do
        tag = UCase$(Trim$(Mid$(t, i + 9, j - i - 9)))
        If tag <> "" And Not TagInvalida(tag) Then
            RespostaTagInvalida = True
            Exit Function
        End If
        i = InStr(j, t, "Tag Name ", vbTextCompare)
    Loop
End Function

Private Function ColetarInvalidas(ByVal anc As Range, ByVal sai As Range) As Long
    Dim t As String, i As Long, j As Long, tag As String
    t = TextoResposta(anc, sai)
    i = InStr(1, t, "Tag Name ", vbTextCompare)
    Do While i > 0
        j = InStr(i + 9, t, " is invalid", vbTextCompare)
        If j = 0 Then Exit Do
        tag = UCase$(Trim$(Mid$(t, i + 9, j - i - 9)))
        If tag <> "" And InStr(tag, " ") = 0 And Not TagInvalida(tag) Then
            If mInvalidas = "" Then mInvalidas = "|"
            mInvalidas = mInvalidas & tag & "|"
            ColetarInvalidas = ColetarInvalidas + 1
        End If
        i = InStr(j, t, "Tag Name ", vbTextCompare)
    Loop
End Function

' Usado so na validacao automatica
Public Function TesteTagsInvalidas(ByVal texto As String) As String
    Dim anc As Range
    Set anc = shDadosMES.Range(ANC_U)
    mInvalidas = ""
    anc.Value = "'" & texto
    ColetarInvalidas anc, shDadosMES.Range(OUT_U)
    TesteTagsInvalidas = ListaInvalidas(1)
    mInvalidas = ""
End Function

Private Function ListaInvalidas(ByVal lab As Long) As String
    Dim p As Long, t As String, r As String
    For p = 1 To NP
        t = TagCfg(p, lab, C_COL_TAG)
        If TagInvalida(t) Then r = r & "  - " & PRotulo(p) & ": " & t & vbCrLf
    Next p
    ListaInvalidas = r
End Function

Private Function ValorValido(ByVal p As Long, ByVal lab As Long, ByVal v As Variant) As Boolean
    Dim x As Double, vmin As String, vmax As String
    If Not ENumero(v) Then Exit Function
    x = CDbl(v)
    vmin = TagCfg(p, lab, C_COL_VMIN)
    vmax = TagCfg(p, lab, C_COL_VMAX)
    If IsNumeric(vmin) And vmin <> "" Then
        If x < CDbl(vmin) Then Exit Function
    End If
    If IsNumeric(vmax) And vmax <> "" Then
        If x > CDbl(vmax) Then Exit Function
    End If
    ValorValido = True
End Function

' ---------------------------------------------------------------- formula do Aspen
' Texto literal para formula, em pedacos de 250 caracteres unidos com & (limite do Excel: 255)
Private Function Literal(ByVal texto As String) As String
    Dim r As String, i As Long
    For i = 1 To Len(texto) Step 250
        If r <> "" Then r = r & "&"
        r = r & """" & Mid$(texto, i, 250) & """"
    Next i
    If r = "" Then r = """"""
    Literal = r
End Function

Public Function ArgumentosLab(ByVal lab As Long, ByVal ps As String, ByVal anc As Range, ByVal sai As Range) As String
    Dim lst As Variant, i As Long, p As Long, tags As String, servs As String, mapas As String
    Dim s As String, m As String, ci As String, cf As String, calc As String
    lst = Split(ps, ",")
    For i = LBound(lst) To UBound(lst)
        p = CLng(lst(i))
        s = TagCfg(p, lab, C_COL_SERV)
        If s = "" Then s = "UBU"
        m = TagCfg(p, lab, C_COL_MAPA)
        If m = "" Then m = "IP_MESVALOR"
        If tags <> "" Then
            tags = tags & ","
            servs = servs & ","
            mapas = mapas & ","
        End If
        tags = tags & TagCfg(p, lab, C_COL_TAG)
        servs = servs & s
        mapas = mapas & m
    Next i
    If lab = 0 Then
        ci = "Dados_MES!$B$3"
        cf = "Dados_MES!$B$4"
    Else
        ci = "Dados_MES!$B$6"
        cf = "Dados_MES!$B$7"
    End If
    calc = CfgTxt("cfgCalc")
    If calc = "" Then calc = "1"
    ArgumentosLab = """time,attribute""," & Literal(tags) & "," & Literal(servs) & "," & Literal(mapas) & _
        "," & ci & "," & cf & ",""2h"",0,"""",0,""" & calc & """,0,1560,0,0,1,1," & _
        "ADDRESS(ROW(Dados_MES!" & anc.Address(False, False) & "),COLUMN(Dados_MES!" & anc.Address(False, False) & _
        "),1,,""Dados_MES""),""Dados_MES!" & sai.Address(False, False) & """,1)"
End Function

' Inicio da formula do Aspen ("=...GetCalculationValues("), como o Excel a mostra
Private Function PrefixoAspen(ByVal anc As Range) As String
    Dim atual As String
    atual = anc.Formula
    If InStr(atual, "(") > 0 And InStr(1, atual, "GetCalculationValues", vbTextCompare) > 0 Then
        PrefixoAspen = Left$(atual, InStr(atual, "("))
        If CfgTxt("mesPrefixo") <> PrefixoAspen Then Nm("mesPrefixo").Value = "'" & PrefixoAspen
    ElseIf Len(CfgTxt("mesPrefixo")) > 0 Then
        PrefixoAspen = CfgTxt("mesPrefixo")
    Else
        PrefixoAspen = "=_xll.AspenTech.PME.ProcessData.Functions.GetCalculationValues("
    End If
End Function

' Depois de ler: a formula vira texto (o Aspen nao refaz a consulta a cada recalculo)
Private Sub DesligarConsulta(ByVal anc As Range, ByVal sai As Range)
    On Error Resume Next
    If anc.HasFormula Then
        Call PrefixoAspen(anc)
        anc.Value = "'Consulta concluída em " & Format$(Now, "dd\/mm\/yyyy hh:mm:ss")
    End If
    LimparSaida sai
End Sub

Private Sub LimparSaida(ByVal sai As Range)
    Dim area As Range, c As Range
    On Error Resume Next
    Set area = sai.Resize(NMAX + 2, NP + 1)
    For Each c In area.Cells
        If c.HasArray Then c.CurrentArray.ClearContents
    Next c
    area.ClearContents
End Sub

Private Sub RestaurarSaida(ByVal anc As Range, ByVal sai As Range, ByVal nLin As Long, ByVal nCol As Long)
    Dim prefixo As String, ender As String
    On Error Resume Next
    prefixo = PrefixoAspen(anc)
    prefixo = Left$(prefixo, InStr(1, prefixo, "GetCalculationValues", vbTextCompare) - 1)
    ender = "Dados_MES!" & anc.Address(False, False)
    sai.Resize(nLin, nCol).FormulaArray = prefixo & "ShowCalculationValues(ADDRESS(ROW(" & ender & "),COLUMN(" & _
        ender & "),1,,""Dados_MES"")," & ender & ", 0)"
End Sub

' Inicio/fim do periodo nas celulas usadas pela formula (Dados_MES!B3:B4 Germano, B6:B7 Ubu)
Private Sub GravarPeriodo(ByVal lab As Long, ByVal ini As Date, ByVal fim As Date)
    Dim ci As Range, cf As Range, modo As String
    If lab = 0 Then
        Set ci = Nm("mesIniG")
        Set cf = Nm("mesFimG")
    Else
        Set ci = Nm("mesIniU")
        Set cf = Nm("mesFimU")
    End If
    modo = UCase$(CfgTxt("cfgFormatoData"))
    If InStr(modo, "MM/DD") > 0 Then
        ci.NumberFormat = "@"
        ci.Value = Format$(ini, "mm\/dd\/yyyy hh:mm:ss")
        cf.NumberFormat = "@"
        cf.Value = Format$(fim, "mm\/dd\/yyyy hh:mm:ss")
    ElseIf InStr(modo, "TEXTO") > 0 Then
        ci.NumberFormat = "@"
        ci.Value = Format$(ini, "dd\/mm\/yyyy hh:mm:ss")
        cf.NumberFormat = "@"
        cf.Value = Format$(fim, "dd\/mm\/yyyy hh:mm:ss")
    Else
        ci.NumberFormat = "dd/mm/yyyy hh:mm:ss"
        ci.Value = ini
        cf.NumberFormat = "dd/mm/yyyy hh:mm:ss"
        cf.Value = fim
    End If
End Sub

' ---------------------------------------------------------------- gravacao na aba Batch
' Grava os valores do laboratorio nas colunas das analises COM tag (as sem tag ficam como estao).
' Devolve quantos valores foram gravados.
Private Function GravarResultados(ByVal lab As Long, ByVal res As Variant, ByVal n As Long, ByVal ok As Boolean) As Long
    Dim p As Long, s As Long, c As Range, qtd As Long, temTag As Boolean, simulado As Boolean
    If Not ok Then Exit Function
    simulado = (Left$(UCase$(CfgTxt("cfgFonte")), 3) = "SIM")
    ModoCarga True
    Desproteger shBatch
    For p = 1 To NP
        temTag = (TagCfg(p, lab, C_COL_TAG) <> "" And Not TagInvalida(TagCfg(p, lab, C_COL_TAG)))
        If temTag Or simulado Then
            For s = 1 To NMAX
                Set c = shBatch.Cells(B_ROW_G0 + s - 1, ColValor(p, lab))
                If s <= n Then
                    If ENumero(res(s, p)) Then
                        c.Value = CDbl(res(s, p))
                        qtd = qtd + 1
                    Else
                        c.ClearContents
                    End If
                Else
                    c.ClearContents
                End If
            Next s
        End If
    Next p
    RestaurarHorarios
    Proteger shBatch
    ModoCarga False
    gAlterado = True
    GravarResultados = qtd
End Function

' Horarios de Ubu voltam a ser calculados pelo inicio e pelo Nº de horarios
Public Sub RestaurarHorarios()
    Dim i As Long
    For i = 1 To NMAX
        shBatch.Cells(B_ROW_G0 + i - 1, B_COL_TU).Formula = _
            "=IF(AND(ISNUMBER(gJanU)," & i & "<=bNHor),gJanU+" & (i - 1) & "/12+cfgMinAntes/1440,"""")"
    Next i
End Sub

' ---------------------------------------------------------------- simulacao
' Em torno do "valor tipico" (aba Configuracoes). Ubu: S.E. a cada 2 h, quimica a cada 4 h,
' <325# a cada 6 h (como nas planilhas de comparacao).
Private Function DadosSimulados(ByVal lab As Long, ByVal n As Long, ByVal ini As Date) As Variant
    Dim d() As Variant, s As Long, p As Long, tip As String, passo As Long
    ReDim d(1 To n, 1 To NP)
    Randomize
    For s = 1 To n
        For p = 1 To NP
            tip = TagCfg(p, lab, C_COL_TIP)
            passo = 1
            If lab = 1 Then
                Select Case PChave(p)
                    Case "325": passo = 3
                    Case "SE": passo = 1
                    Case Else: passo = 2
                End Select
            End If
            If IsNumeric(tip) And tip <> "" And ((s - 1) Mod passo = 0) Then
                d(s, p) = Round(CDbl(tip) * (1 + (Rnd() - 0.5) * 0.06), PDec(p) + 1)
            End If
        Next p
    Next s
    DadosSimulados = d
End Function

Option Explicit

' ============================================================================
'  ModMES - resultados quimicos das Usinas 3 e 4 no MES (Aspen IP.21)
'  Tres botoes:
'    AtualizarMESDia / AtualizarMESNoite  (abas Resumo Dia / Resumo Noite)
'    AtualizarResultados                  (aba Resultados gerais, periodo livre)
'  Os valores de 2 em 2 h sao gravados direto na tabela da propria aba
'  (media, minimo e maximo sao formulas da planilha).
' ============================================================================

' Coluna inicial da tabela de consultas ao MES na aba _Mapa (K)
Private Const MAPA_COL_BLOCO As Long = 11
' Segundos sem mudanca nos resultados para considerar a consulta concluida
Private Const ESTAVEL_SEG As Double = 2
' Consulta que respondeu "Success" mas sem nenhuma linha (inicio do turno, sem analises ainda):
' depois deste tempo e tratada como "sem dados" (sem esperar o limite nem mostrar erro)
Private Const SEM_DADOS_SEG As Double = 15
' Evita duas consultas ao mesmo tempo (clique repetido no botao durante a espera)
Private mOcupado As Boolean

' ---------------------------------------------------------------- botoes
Public Sub AtualizarMESDia()
    AtualizarTurno shResumoDia, "d", False
End Sub

Public Sub AtualizarMESNoite()
    AtualizarTurno shResumoNoite, "n", True
End Sub

' Resultados do turno (6 janelas de 2 h) na data da aba Preenchimento
Private Sub AtualizarTurno(ByVal ws As Worksheet, ByVal t As String, ByVal noite As Boolean)
    Dim v As Variant, ini As Date, dados As Variant, n As Long, simulado As Boolean, nome As String, obs As String
    v = Nm("pData").Value
    If Not (VarType(v) = vbDate Or ENumero(v)) Then
        Aviso "Informe a data na aba Preenchimento.", vbExclamation, "MES"
        Exit Sub
    End If
    ini = CDate(Int(CDbl(v)) + HoraInicioDia())
    If noite Then ini = CDate(CDbl(ini) + 0.5)
    If noite Then nome = "Noite" Else nome = "Dia"
    simulado = FonteSimulada()
    If Not simulado And CDbl(ini) > CDbl(Now) Then
        Aviso "Turno " & nome & " de " & Format$(ini, "dd\/mm\/yyyy") & " ainda não iniciado (" & _
              Format$(ini, "hh:mm") & ").", vbInformation, "MES"
        Exit Sub
    End If
    If Not BuscarMES(ini, NSLOT, dados) Then Exit Sub
    n = EscreverTabela(ws, dados, NSLOT, NSLOT, ini, t & "Horas", t & "Atualizado", simulado)
    CompactarAba ws
    ws.Activate
    If CDbl(ini) + NSLOT * HORAS_SLOT / 24 > CDbl(Now) And Not simulado Then
        obs = vbCrLf & "Turno em andamento: resultados até " & Format$(Now, "hh:mm") & "."
    End If
    If n = 0 And Not simulado Then
        Aviso "Sem resultados no MES para o turno " & nome & "." & obs, vbInformation, "MES"
    Else
        Aviso "Turno " & nome & " atualizado: " & n & " resultados." & obs & _
              IIf(simulado, vbCrLf & "Fonte: SIMULAÇÃO.", ""), _
              vbInformation, "MES"
    End If
End Sub

' Aba Resultados gerais: periodo escolhido (inicio/fim, ate 24 h)
Public Sub AtualizarResultados()
    Dim vi As Variant, vf As Variant, ini As Date, fim As Date, horas As Double, n As Long, qtd As Long
    Dim dados As Variant, simulado As Boolean, s As Long

    vi = Nm("gIni").Value
    vf = Nm("gFim").Value
    If Not DataHoraValida(vi) Or Not DataHoraValida(vf) Then
        Aviso "Informe início e fim do período (dd/mm/aaaa hh:mm).", vbExclamation, "Resultados gerais"
        Exit Sub
    End If
    ini = CDate(Round(CDbl(CDate(vi)) * 1440, 0) / 1440)
    fim = CDate(Round(CDbl(CDate(vf)) * 1440, 0) / 1440)
    horas = (CDbl(fim) - CDbl(ini)) * 24
    If horas <= 0 Then
        Aviso "O fim deve ser posterior ao início.", vbExclamation, "Resultados gerais"
        Exit Sub
    End If
    If horas > NSLOT_MAX * HORAS_SLOT + 0.01 Then
        Aviso "Período máximo: " & NSLOT_MAX * HORAS_SLOT & " horas.", vbExclamation, "Resultados gerais"
        Exit Sub
    End If
    ' janelas de 2 h a partir do inicio (a ultima completa as 2 h)
    n = Int((horas + 0.01) / HORAS_SLOT)
    If n * HORAS_SLOT < horas - 0.01 Then n = n + 1
    simulado = FonteSimulada()
    If Not simulado And CDbl(ini) > CDbl(Now) Then
        Aviso "Período ainda não iniciado.", vbExclamation, "Resultados gerais"
        Exit Sub
    End If
    If Not BuscarMES(ini, n, dados) Then Exit Sub
    qtd = EscreverTabela(shResultados, dados, n, NSLOT_MAX, ini, "gHoras", "gAtualizado", simulado)
    ' so as janelas do periodo aparecem (tambem na imagem)
    On Error Resume Next
    Desproteger shResultados
    For s = 1 To NSLOT_MAX
        shResultados.Columns(COL_SLOT1 + s - 1).Hidden = (s > n)
    Next s
    Proteger shResultados
    On Error GoTo 0
    CompactarAba shResultados
    shResultados.Activate
    Aviso "Período atualizado: " & qtd & " resultados.", vbInformation, "Resultados gerais"
End Sub

Private Function FonteSimulada() As Boolean
    FonteSimulada = (Left$(UCase$(CfgTxt("cfgFonte")), 3) = "SIM")
End Function

' ---------------------------------------------------------------- busca e gravacao
' Busca n janelas de 2 h a partir de ini. Usa a mesma consulta de 12 h (6 janelas) do
' turno; para mais de 12 h consulta mais de uma vez e junta. dados(1..n, 1 + 2*NPARAM).
Private Function BuscarMES(ByVal ini As Date, ByVal n As Long, ByRef dados As Variant) As Boolean
    Dim nPartes As Long, parte As Long, a As Date, fimParte As Date, nAgora As Long, simulado As Boolean
    Dim bloco As Variant, ok As Boolean, s As Long, j As Long, lin As Long
    Dim d() As Variant
    BuscarMES = False
    If mOcupado Then
        Aviso "Consulta ao MES em andamento. Aguarde.", vbInformation, "MES"
        Exit Function
    End If
    mOcupado = True
    simulado = FonteSimulada()
    nPartes = (n + NSLOT - 1) \ NSLOT
    ReDim d(1 To n, 1 To 1 + 2 * NPARAM)
    On Error GoTo Falha
    Ampulheta True
    For parte = 1 To nPartes
        a = CDate(CDbl(ini) + (parte - 1) * NSLOT * HORAS_SLOT / 24)
        ' parte que ainda nao comecou: nada a consultar (fica em branco)
        If simulado Or CDbl(a) <= CDbl(Now) Then
            ' Periodo em andamento: consulta so ate o fim da janela de 2 h atual
            ' (nunca pede horarios futuros ao MES)
            fimParte = CDate(CDbl(a) + NSLOT * HORAS_SLOT / 24)
            If Not simulado And CDbl(fimParte) > CDbl(Now) Then
                nAgora = Int((CDbl(Now) - CDbl(a)) * 24 / HORAS_SLOT) + 1
                If nAgora > NSLOT Then nAgora = NSLOT
                fimParte = CDate(CDbl(a) + nAgora * HORAS_SLOT / 24)
            End If
            Application.StatusBar = "Consultando MES: " & Format$(a, "dd\/mm hh:mm") & " a " & _
                                    Format$(fimParte, "dd\/mm hh:mm")
            If simulado Then
                bloco = DadosSimulados()
                ok = True
            Else
                ok = ConsultarMES(a, fimParte, bloco)
            End If
            If Not ok Then GoTo Fim
            For s = 1 To NSLOT
                lin = (parte - 1) * NSLOT + s
                ' janelas que ainda nao comecaram ficam em branco
                If lin <= n And CDbl(a) + (s - 1) * HORAS_SLOT / 24 <= CDbl(Now) Then
                    For j = 1 To 1 + 2 * NPARAM
                        d(lin, j) = bloco(s, j)
                    Next j
                End If
            Next s
        End If
    Next parte
    dados = d
    BuscarMES = True
Fim:
    Application.StatusBar = False
    Ampulheta False
    mOcupado = False
    Exit Function
Falha:
    Application.StatusBar = False
    Ampulheta False
    mOcupado = False
    Aviso "Erro ao consultar o MES: " & Err.Description, vbExclamation, "MES"
End Function

' Grava as janelas de 2 h na tabela da aba (linhas com chave "P02U3" na coluna A).
' Devolve quantos valores foram gravados.
Private Function EscreverTabela(ByVal ws As Worksheet, ByVal dados As Variant, ByVal n As Long, _
                                ByVal nCols As Long, ByVal ini As Date, ByVal nomeHoras As String, _
                                ByVal nomeAtual As String, ByVal simulado As Boolean) As Long
    Dim r As Long, rFim As Long, chave As String, p As Long, k As Long, s As Long, v As Variant, qtd As Long
    Dim vmin As Variant, vmax As Variant, c As Range, emAndamento As Boolean

    Desproteger ws
    rFim = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = 1 To rFim
        chave = CStr(ws.Cells(r, 1).Value)
        If Len(chave) = 5 And Left$(chave, 1) = "P" Then
            p = CLng(Mid$(chave, 2, 2))
            If Right$(chave, 1) = "3" Then k = 1 Else k = 2
            vmin = CfgCel(p, CFG_COL_VMIN)
            vmax = CfgCel(p, CFG_COL_VMAX)
            For s = 1 To nCols
                Set c = ws.Cells(r, COL_SLOT1 + s - 1)
                v = Empty
                If s <= n Then v = dados(s, 1 + (p - 1) * 2 + k)
                If ValorValido(v, vmin, vmax) Then
                    c.Value = CDbl(v)
                    qtd = qtd + 1
                Else
                    c.ClearContents
                End If
            Next s
        End If
    Next r
    For s = 1 To nCols
        If s <= n Then
            ' horario da amostra bi-horaria (meia hora apos o inicio da janela: 07:30, 09:30...)
            Nm(nomeHoras).Cells(1, s).Value = "'" & Format$(CDate(CDbl(ini) + (s - 1) * HORAS_SLOT / 24 + 1 / 48), "hh:mm")
            ' janela em andamento: media parcial
            If Not simulado And CDbl(ini) + (s - 1) * HORAS_SLOT / 24 <= CDbl(Now) And _
               CDbl(ini) + s * HORAS_SLOT / 24 > CDbl(Now) Then
                Nm(nomeHoras).Cells(1, s).Value = Nm(nomeHoras).Cells(1, s).Value & " *"
                emAndamento = True
            End If
        Else
            Nm(nomeHoras).Cells(1, s).Value = "—"
        End If
    Next s
    Nm(nomeAtual).Value = "Período: " & Format$(ini, "dd\/mm\/yyyy hh:mm") & " a " & _
        Format$(CDate(CDbl(ini) + n * HORAS_SLOT / 24), "dd\/mm\/yyyy hh:mm") & "     |     Atualizado: " & _
        Format$(Now, "dd\/mm\/yyyy hh:mm") & "     |     Fonte: " & IIf(simulado, "SIMULAÇÃO", _
        "MES - servidor " & CfgTxt("cfgServidor"))
    If emAndamento Then Nm(nomeAtual).Value = Nm(nomeAtual).Value & "     |     * janela em andamento"
    ws.Calculate
    Proteger ws
    EscreverTabela = qtd
End Function

' ---------------------------------------------------------------------------
'  Consulta ao MES - mesmo formato da planilha de referencia (out/2014):
'  uma formula GetCalculationValues por grupo (Qualidade, Producao, Ritmo de
'  processo), com tags, servidores e mapas em TEXTO LITERAL dentro da formula.
'  Os grupos estao na aba _Mapa (colunas K:Q).
' ---------------------------------------------------------------------------

Private Function DataHoraValida(ByVal v As Variant) As Boolean
    If Vazio(v) Then
        DataHoraValida = False
    ElseIf VarType(v) = vbDate Or IsNumeric(v) Then
        DataHoraValida = (CDbl(v) > 40000)
    Else
        DataHoraValida = IsDate(v)
    End If
End Function

Public Function NumBlocos() As Long
    Dim r As Long
    r = 2
    Do While Not Vazio(shMapa.Cells(r, MAPA_COL_BLOCO).Value)
        r = r + 1
    Loop
    NumBlocos = r - 2
End Function

' Analise com as duas tags preenchidas em Configuracoes (sem tag = fora da consulta e oculta)
Public Function ParamConfigurado(ByVal p As Long) As Boolean
    ParamConfigurado = Not Vazio(CfgCel(p, CFG_COL_TAG3)) And Not Vazio(CfgCel(p, CFG_COL_TAG4))
End Function

' Parametros do bloco b que entram na consulta (so os configurados), como lista de textos
Private Function ParamsBloco(ByVal b As Long) As Variant
    Dim todos As Variant, i As Long, s As String
    todos = Split(BlocoInfo(b, 5), ",")
    For i = LBound(todos) To UBound(todos)
        If ParamConfigurado(CLng(todos(i))) Then
            If s <> "" Then s = s & ","
            s = s & todos(i)
        End If
    Next i
    ParamsBloco = Split(s, ",")
End Function

Private Function BlocoInfo(ByVal b As Long, ByVal campo As Long) As String
    ' campo: 0 codigo, 1 titulo, 2 calculo, 3 ancora, 4 saida, 5 lista de parametros
    BlocoInfo = CStr(shMapa.Cells(1 + b, MAPA_COL_BLOCO + campo).Value)
End Function

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

' Argumentos da GetCalculationValues do bloco b (tudo o que vem depois do "(")
Public Function ArgumentosBloco(ByVal b As Long) As String
    Dim ps As Variant, i As Long, p As Long, serv As String
    Dim tags As String, mapas As String, servs As String, anc As String

    serv = CfgTxt("cfgServidor")
    If serv = "" Then serv = "UBU"
    ps = ParamsBloco(b)
    For i = LBound(ps) To UBound(ps)
        p = CLng(ps(i))
        If tags <> "" Then
            tags = tags & ","
            mapas = mapas & ","
            servs = servs & ","
        End If
        tags = tags & Trim$(CStr(CfgCel(p, CFG_COL_TAG3))) & "," & Trim$(CStr(CfgCel(p, CFG_COL_TAG4)))
        mapas = mapas & Trim$(CStr(CfgCel(p, CFG_COL_TIPO3))) & "," & Trim$(CStr(CfgCel(p, CFG_COL_TIPO4)))
        servs = servs & serv & "," & serv
    Next i
    anc = BlocoInfo(b, 3)
    ArgumentosBloco = """time,attribute""," & Literal(tags) & "," & Literal(servs) & "," & Literal(mapas) & _
        ",Dados_MES!$B$3,Dados_MES!$B$4,""2h"",0,"""",0,""" & BlocoInfo(b, 2) & """,0,1560,0,0,1,1," & _
        "ADDRESS(ROW(Dados_MES!" & anc & "),COLUMN(Dados_MES!" & anc & "),1,,""Dados_MES"")," & _
        """Dados_MES!" & BlocoInfo(b, 4) & """,1)"
End Function

' Reescreve as formulas das consultas a partir da aba Configuracoes (texto literal).
' O nome da funcao (com ou sem _xll.) e preservado exatamente como o Excel o mostra.
Public Sub RegerarFormulasMES()
    Dim b As Long
    For b = 1 To NumBlocos()
        ' sempre reescreve (mesmo igual): o Excel recalcula a consulta uma unica vez
        shDadosMES.Range(BlocoInfo(b, 3)).Formula = PrefixoAspen(b) & ArgumentosBloco(b)
    Next b
End Sub

' Inicio da formula do Aspen ("=...GetCalculationValues("), exatamente como o Excel a mostra.
' E lido da formula existente e guardado em Dados_MES (mesPrefixo), porque apos cada consulta
' a formula e desligada (vira texto) para nao ser recalculada a toda edicao da planilha.
Private Function PrefixoAspen(ByVal b As Long) As String
    Dim atual As String
    atual = shDadosMES.Range(BlocoInfo(b, 3)).Formula
    If InStr(atual, "(") > 0 And InStr(1, atual, "GetCalculationValues", vbTextCompare) > 0 Then
        PrefixoAspen = Left$(atual, InStr(atual, "("))
        If CStr(Nm("mesPrefixo").Value) <> PrefixoAspen Then Nm("mesPrefixo").Value = "'" & PrefixoAspen
    ElseIf Len(CfgTxt("mesPrefixo")) > 0 Then
        PrefixoAspen = CfgTxt("mesPrefixo")
    Else
        PrefixoAspen = "=_xll.AspenTech.PME.ProcessData.Functions.GetCalculationValues("
    End If
End Function

' Depois de ler os resultados: a formula da consulta vira texto e as matrizes sao apagadas,
' para o Aspen nao refazer a consulta a cada recalculo do Excel (deixava tudo lento).
Private Sub DesligarConsultas()
    Dim b As Long, c As Range
    On Error Resume Next
    For b = 1 To NumBlocos()
        Set c = shDadosMES.Range(BlocoInfo(b, 3))
        If c.HasFormula Then
            Call PrefixoAspen(b)
            c.Value = "'Consulta concluída em " & Format$(Now, "dd\/mm\/yyyy hh:mm:ss")
        End If
    Next b
    LimparSaidasMES
End Sub

' True se alguma consulta do Aspen ainda esta como formula (antes da 1a atualizacao ou apos falha)
Public Function ConsultaAspenAtiva() As Boolean
    Dim b As Long
    On Error Resume Next
    For b = 1 To NumBlocos()
        If InStr(1, shDadosMES.Range(BlocoInfo(b, 3)).Formula, "GetCalculationValues", vbTextCompare) > 0 Then
            ConsultaAspenAtiva = True
            Exit Function
        End If
    Next b
End Function

' Faz o Excel calcular a consulta (no modo automatico, escrever a formula ja basta)
Private Sub CalcularConsultas()
    Dim b As Long
    If Application.Calculation <> -4105 Then
        For b = 1 To NumBlocos()
            shDadosMES.Range(BlocoInfo(b, 3)).Calculate
        Next b
    End If
End Sub

Private Function SaidaBloco(ByVal b As Long) As Range
    Dim ps As Variant
    ps = ParamsBloco(b)
    Set SaidaBloco = shDadosMES.Range(BlocoInfo(b, 4)).Resize(NSLOT, 1 + 2 * (UBound(ps) - LBound(ps) + 1))
End Function

' A consulta respondeu ("Success"), mas o MES ainda nao tem nenhum resultado no periodo
Private Function RespondeuSemDados(ByVal b As Long) As Boolean
    RespondeuSemDados = (LCase$(Left$(Trim$(shDadosMES.Range(BlocoInfo(b, 3)).Text), 7)) = "success") And _
                        Vazio(SaidaBloco(b).Cells(1, 1).Value)
End Function

Private Function BlocoPronto(ByVal b As Long, ByVal ini As Date) As Boolean
    Dim v As Variant
    v = SaidaBloco(b).Cells(1, 1).Value
    BlocoPronto = False
    If Vazio(v) Then Exit Function
    If VarType(v) = vbDate Or IsNumeric(v) Or IsDate(v) Then
        BlocoPronto = (Abs(CDbl(CDate(v)) - CDbl(ini)) <= HORAS_SLOT / 24 + 1 / 1440)
    End If
End Function

' Grava inicio/fim do periodo nas celulas usadas pelas formulas (Dados_MES!B3:B4).
' Padrao: valor de data do Excel (independe do idioma). Alternativas em Configuracoes.
Public Sub GravarPeriodoMES(ByVal ini As Date, ByVal fim As Date)
    Dim modo As String
    modo = UCase$(CfgTxt("cfgFormatoData"))
    If InStr(modo, "MM/DD") > 0 Then
        EscreverTexto Nm("mesInicio"), Format$(ini, "mm\/dd\/yyyy hh:mm:ss")
        EscreverTexto Nm("mesFim"), Format$(fim, "mm\/dd\/yyyy hh:mm:ss")
    ElseIf InStr(modo, "TEXTO") > 0 Then
        EscreverTexto Nm("mesInicio"), Format$(ini, "dd\/mm\/yyyy hh:mm:ss")
        EscreverTexto Nm("mesFim"), Format$(fim, "dd\/mm\/yyyy hh:mm:ss")
    Else
        Nm("mesInicio").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        Nm("mesInicio").Value = ini
        Nm("mesFim").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        Nm("mesFim").Value = fim
    End If
End Sub

Private Sub EscreverTexto(ByVal c As Range, ByVal texto As String)
    c.NumberFormat = "@"
    c.Value = texto
End Sub

' Resumo do conteudo atual das saidas (texto da ancora + qtde e soma dos numeros)
Private Function AssinaturaSaidas(ByVal nb As Long) As String
    Dim b As Long, arr As Variant, i As Long, j As Long, n As Long, soma As Double, r As String
    For b = 1 To nb
        arr = SaidaBloco(b).Value
        n = 0
        soma = 0
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                If ENumero(arr(i, j)) Then
                    n = n + 1
                    soma = soma + CDbl(arr(i, j))
                End If
            Next j
        Next i
        r = r & shDadosMES.Range(BlocoInfo(b, 3)).Text & "|" & n & "|" & soma & ";"
    Next b
    AssinaturaSaidas = r
End Function

' Espera ate todas as consultas responderem E os valores pararem de mudar por
' ESTAVEL_SEG segundos (o Aspen preenche a matriz aos poucos; ler antes disso
' trazia dados incompletos). Devolve quantas consultas responderam.
Private Function EsperarBlocos(ByVal ini As Date, ByVal nb As Long, ByRef pronto() As Boolean, _
                               ByVal limite As Double) As Long
    Dim t0 As Single, tEstavel As Single, assin As String, assinAnt As String, b As Long, nOk As Long
    t0 = Timer
    tEstavel = Timer
    Do
        DoEvents
        If Timer < t0 Then t0 = t0 - 86400
        If Timer < tEstavel Then tEstavel = tEstavel - 86400
        nOk = 0
        For b = 1 To nb
            pronto(b) = BlocoPronto(b, ini)
            If Not pronto(b) And Timer - t0 >= SEM_DADOS_SEG Then pronto(b) = RespondeuSemDados(b)
            If pronto(b) Then nOk = nOk + 1
        Next b
        assin = AssinaturaSaidas(nb)
        If assin <> assinAnt Then
            assinAnt = assin
            tEstavel = Timer
        End If
        If nOk = nb And Timer - tEstavel >= ESTAVEL_SEG Then Exit Do
        If Timer - t0 > limite Then Exit Do
    Loop
    EsperarBlocos = nOk
End Function

' Apaga as matrizes de resultados das consultas (inteiras, nunca "parte" delas)
Public Sub LimparSaidasMES()
    Dim b As Long, area As Range, c As Range
    On Error Resume Next
    For b = 1 To NumBlocos()
        Set area = shDadosMES.Range(BlocoInfo(b, 4)).Resize(NSLOT + 1, LarguraSaida(b))
        For Each c In area.Cells
            If c.HasArray Then c.CurrentArray.ClearContents
        Next c
        area.ClearContents
    Next b
    On Error GoTo 0
End Sub

' Recria as matrizes ShowCalculationValues no formato original (NSLOT linhas)
Public Sub RestaurarSaidasMES()
    Dim b As Long, anc As Range, prefixo As String, ender As String
    On Error Resume Next
    For b = 1 To NumBlocos()
        Set anc = shDadosMES.Range(BlocoInfo(b, 3))
        prefixo = PrefixoAspen(b)
        prefixo = Left$(prefixo, InStr(1, prefixo, "GetCalculationValues", vbTextCompare) - 1)
        ender = "Dados_MES!" & anc.Address(False, False)
        shDadosMES.Range(BlocoInfo(b, 4)).Resize(NSLOT, LarguraSaida(b)).FormulaArray = _
            prefixo & "ShowCalculationValues(ADDRESS(ROW(" & ender & "),COLUMN(" & ender & "),1,,""Dados_MES""),"  & _
            ender & ", 0)"
    Next b
    On Error GoTo 0
End Sub

Private Function LarguraSaida(ByVal b As Long) As Long
    Dim ps As Variant
    ps = ParamsBloco(b)
    LarguraSaida = 1 + 2 * (UBound(ps) - LBound(ps) + 1)
End Function

Private Function SegundosDesde(ByVal t0 As Single) As Double
    SegundosDesde = Timer - t0
    If SegundosDesde < 0 Then SegundosDesde = SegundosDesde + 86400
End Function

' Escreve o periodo, regera as formulas, recalcula e espera o retorno de cada consulta
Private Function ConsultarMES(ByVal ini As Date, ByVal fim As Date, ByRef dados As Variant) As Boolean
    Dim tIni As Single
    Dim limite As Double, msg As String, nb As Long, b As Long, nOk As Long
    Dim pronto() As Boolean, d() As Variant, arr As Variant, ps As Variant
    Dim i As Long, k As Long, s As Long, p As Long, falhas As String

    ConsultarMES = False
    tIni = Timer
    nb = NumBlocos()
    ReDim pronto(1 To nb)
    limite = 60
    If ENumero(Nm("cfgTimeout").Value) Then limite = CDbl(Nm("cfgTimeout").Value)

    ' O suplemento Aspen cria/redimensiona sozinho a matriz de resultados de cada consulta.
    ' Se sobrar a matriz de uma consulta anterior com outro numero de linhas, o Excel recusa
    ' ("Nao e possivel alterar parte de uma matriz"). Por isso as saidas sao limpas antes.
    LimparSaidasMES
    GravarPeriodoMES ini, fim
    On Error GoTo FalhaFormula
    RegerarFormulasMES
    On Error GoTo 0
    CalcularConsultas
    nOk = EsperarBlocos(ini, nb, pronto, limite)

    ' Seguranca: se nada voltou, recria as matrizes de saida no formato original e tenta de novo
    If nOk = 0 Then
        RestaurarSaidasMES
        RegerarFormulasMES
        CalcularConsultas
        nOk = EsperarBlocos(ini, nb, pronto, limite)
    End If

    For b = 1 To nb
        If Not pronto(b) Then
            falhas = falhas & "  - " & BlocoInfo(b, 1) & ": " & _
                     Left$(shDadosMES.Range(BlocoInfo(b, 3)).Text & " " & SaidaBloco(b).Cells(1, 1).Text, 150) & vbCrLf
        End If
    Next b

    If nOk = 0 Then
        msg = "O MES não respondeu (" & Format$(ini, "dd\/mm\/yyyy hh:mm") & ")." & vbCrLf & vbCrLf & _
              falhas & vbCrLf & _
              "Verificar: suplemento Aspen ativo, servidor '" & CfgTxt("cfgServidor") & "', tags e formato da data " & _
              "(aba Configurações)." & vbCrLf & vbCrLf & "Abrir a aba Dados_MES?"
        If Aviso(msg, vbExclamation + vbYesNo, "MES sem resposta") = vbYes Then MostrarDadosMES
        Exit Function
    End If
    If falhas <> "" Then
        Aviso "Consultas sem resposta (ficam em branco):" & vbCrLf & falhas, vbExclamation, "MES"
    End If

    ReDim d(1 To NSLOT, 1 To 1 + 2 * NPARAM)
    For b = 1 To nb
        If pronto(b) Then
            arr = SaidaBloco(b).Value
            ps = ParamsBloco(b)
            For i = LBound(ps) To UBound(ps)
                p = CLng(ps(i))
                For k = 1 To 2
                    For s = 1 To NSLOT
                        d(s, 1 + (p - 1) * 2 + k) = arr(s, 1 + (i - LBound(ps)) * 2 + k)
                    Next s
                Next k
            Next i
        End If
    Next b
    dados = d
    DesligarConsultas
    Nm("mesTempo").Value = "'" & Format$(Now, "dd\/mm\/yyyy hh:mm:ss") & "  -  " & _
        Format$(SegundosDesde(tIni), "0") & " s"
    ConsultarMES = True
    Exit Function

FalhaFormula:
    Aviso "Erro na fórmula de consulta ao MES (" & Err.Description & ")." & vbCrLf & _
          "Verificar as tags na aba Configurações.", vbExclamation, "MES"
End Function

' Gera valores ficticios em torno do "valor tipico" (modo de treinamento/teste)
Private Function DadosSimulados() As Variant
    Dim d() As Variant, s As Long, p As Long, k As Long, col As Long
    Dim tip As Variant, tipo As String

    ReDim d(1 To NSLOT, 1 To 1 + 2 * NPARAM)
    Randomize
    For s = 1 To NSLOT
        For p = 1 To NPARAM
            For k = 1 To 2
                col = 1 + (p - 1) * 2 + k
                tip = CfgCel(p, CFG_COL_TIP3 + k - 1)
                tipo = UCase$(CStr(CfgCel(p, CFG_COL_TIPO3 + (k - 1) * 2)))
                If Not ENumero(tip) Or Not ParamConfigurado(p) Then
                    d(s, col) = Empty
                ElseIf tipo = "IP_MESVALOR" And Rnd() < 0.2 Then
                    d(s, col) = Empty
                Else
                    d(s, col) = Round(CDbl(tip) * (1 + (Rnd() - 0.5) * 0.06), 4)
                End If
            Next k
        Next p
    Next s
    DadosSimulados = d
End Function

Private Function ValorValido(ByVal v As Variant, ByVal vmin As Variant, ByVal vmax As Variant) As Boolean
    Dim x As Double
    ValorValido = False
    If Not ENumero(v) Then Exit Function
    x = CDbl(v)
    If ENumero(vmin) Then
        If x < CDbl(vmin) Then Exit Function
    End If
    If ENumero(vmax) Then
        If x > CDbl(vmax) Then Exit Function
    End If
    ValorValido = True
End Function

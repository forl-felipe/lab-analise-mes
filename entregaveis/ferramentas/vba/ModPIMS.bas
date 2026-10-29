Attribute VB_Name = "ModPIMS"
Option Explicit

' ============================================================================
'  ModPIMS - busca no PIMS (Aspen IP.21) da prensa automática e do laboratório (MES)
'  Mesmo mecanismo do Relatório Físico: a macro escreve a fórmula GetCalculationValues
'  do suplemento Aspen Process Data na aba Dados_PIMS, espera o retorno, lê os valores
'  e grava nas abas Prensa e MES. As abas Correlacao e Painel são fórmulas.
'  Botões da aba Painel:
'    AtualizarPeriodo       período das células Início e Fim do Painel
'    AtualizarDesdeInicio   desde o início de operação da prensa (aba Configuracoes)
'    LimparDados            apaga os dados das abas Prensa e MES
'    TestarConexao          consulta curta das duas tags, com o resultado de cada etapa na aba Log
'  Cada busca registra as etapas na aba Log (fórmula enviada, resposta do Aspen, linhas lidas).
' ============================================================================

' Segundos sem mudança nos resultados para considerar a consulta concluída
Private Const ESTAVEL_SEG As Double = 2
' Consulta que respondeu "Success" sem nenhuma linha: depois deste tempo é tratada como "sem dados"
Private Const SEM_DADOS_SEG As Double = 15
' Linhas por consulta (o suplemento limita a quantidade de pontos; consultas maiores são divididas)
Private Const LINHAS_CONSULTA As Long = 1440
' Capacidade das abas de dados (precisa bater com as fórmulas das abas Prensa e MES)
Private Const MAX_PRENSA As Long = 5000
Private Const MAX_MES As Long = 1500
' Blocos de consulta na aba Dados_PIMS
Private Const BLOCO_PRENSA As Long = 1
Private Const BLOCO_LAB As Long = 2
Private Const PREFIXO_XLL As String = "=_xll.AspenTech.PME.ProcessData.Functions.GetCalculationValues("

Private mOcupado As Boolean

' ---------------------------------------------------------------- botões
Public Sub AtualizarPeriodo()
    Dim ini As Date, fim As Date
    If Not LerData(Nm("pnlInicio").Value, ini) Then
        Aviso "Preencha a data e a hora de início na aba Painel.", vbExclamation
        Exit Sub
    End If
    If Vazio(Nm("pnlFim").Value) Then
        fim = Now
    ElseIf Not LerData(Nm("pnlFim").Value, fim) Then
        Aviso "A data e a hora de fim na aba Painel não são válidas." & vbCrLf & _
              "Deixe em branco para buscar até o momento atual.", vbExclamation
        Exit Sub
    End If
    Atualizar ini, fim
End Sub

Public Sub AtualizarDesdeInicio()
    Dim ini As Date
    If Not LerData(Nm("cfgInicioPrensa").Value, ini) Then
        Aviso "Preencha o início de operação da prensa na aba Configuracoes.", vbExclamation
        Exit Sub
    End If
    Nm("pnlInicio").Value = ini
    Nm("pnlFim").ClearContents
    Atualizar ini, Now
End Sub

Public Sub LimparDados()
    If Aviso("Apagar os dados das abas Prensa e MES?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    Aba("Prensa").Range("A2:B" & MAX_PRENSA + 1).ClearContents
    Aba("MES").Range("A2:B" & MAX_MES + 1).ClearContents
    Nm("pnlStatus").Value = "Sem dados. Use um dos botões de atualização."
End Sub

Public Sub MostrarDadosPIMS()
    Aba("Dados_PIMS").Visible = xlSheetVisible
    Aba("Dados_PIMS").Activate
End Sub

' ---------------------------------------------------------------- busca completa
Private Sub Atualizar(ByVal ini As Date, ByVal fim As Date)
    Dim iniJan As Date, ok As Boolean, tIni As Single
    Dim tP() As Date, vP() As Variant, nP As Long
    Dim tL() As Date, vL() As Variant, nL As Long
    Dim nCortes As Long, nRes As Long, msg As String

    If mOcupado Then
        Aviso "Busca em andamento. Aguarde.", vbInformation
        Exit Sub
    End If
    If CDbl(fim) > CDbl(Now) Then fim = Now
    If CDbl(fim) <= CDbl(ini) Then
        Aviso "O fim do período precisa ser depois do início.", vbExclamation
        Exit Sub
    End If

    mOcupado = True
    tIni = Timer
    On Error GoTo Falha
    Ampulheta True
    LimparLog
    Registrar "Início", "Período " & Format$(ini, "dd\/mm\/yyyy hh:mm") & " a " & Format$(fim, "dd\/mm\/yyyy hh:mm") & _
              " | Excel " & Application.Version & " | cálculo " & IIf(Application.Calculation = xlCalculationAutomatic, "automático", "manual")

    ' As janelas de 2 h começam nas horas pares (ou conforme o deslocamento configurado).
    ' A prensa é buscada desde 1 h antes, para completar a primeira janela.
    iniJan = AlinharJanela(ini)
    ok = BuscarSerie(BLOCO_PRENSA, CDate(CDbl(iniJan) - 1 / 24), fim, tP, vP, nP)
    If ok Then ok = BuscarSerie(BLOCO_LAB, iniJan, fim, tL, vL, nL)
    DesligarConsultas
    If Not ok Then GoTo Encerrar

    Application.ScreenUpdating = False
    nCortes = GravarPrensa(tP, vP, nP)
    nRes = GravarMES(tL, vL, nL)
    Registrar "Gravação", nCortes & " resultados da prensa (" & nP & " leituras) e " & nRes & " do laboratório (" & nL & " leituras)"
    Application.Calculate
    Application.ScreenUpdating = True

    Nm("pnlStatus").Value = "Atualizado em " & Format$(Now, "dd\/mm\/yyyy hh:mm") & _
        "  |  período de " & Format$(ini, "dd\/mm\/yyyy hh:mm") & " a " & Format$(fim, "dd\/mm\/yyyy hh:mm") & _
        "  |  " & nCortes & " resultados da prensa e " & nRes & " do laboratório  |  " & _
        Format$(SegundosDesde(tIni), "0") & " s"
    Aba("Painel").Activate

    msg = "Busca concluída." & vbCrLf & vbCrLf & _
          "Prensa automática: " & nCortes & " resultados" & vbCrLf & _
          "Laboratório (MES): " & nRes & " resultados"
    If nCortes = 0 Then msg = msg & vbCrLf & vbCrLf & "Nenhum resultado da prensa no período. Confira a tag e o período."
    If nRes = 0 Then msg = msg & vbCrLf & vbCrLf & "Nenhum resultado do laboratório no período. Confira a tag e o período."
    Aviso msg, IIf(nCortes = 0 Or nRes = 0, vbExclamation, vbInformation)

Encerrar:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Ampulheta False
    mOcupado = False
    Exit Sub
Falha:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Ampulheta False
    mOcupado = False
    Registrar "Erro", Err.Number & ": " & Err.Description
    Aviso "Erro na busca: " & Err.Description & vbCrLf & vbCrLf & "Detalhes na aba Log.", vbExclamation
End Sub

' Consulta curta das duas tags (laboratório nas últimas 24 h, prensa nas últimas 2 h), para conferir a ligação com o PIMS
Public Sub TestarConexao()
    Dim tt() As Date, vv() As Variant, nP As Long, nL As Long, okP As Boolean, okL As Boolean, msg As String
    Dim lista As String
    If mOcupado Then Exit Sub
    mOcupado = True
    On Error GoTo Falha
    Ampulheta True
    LimparLog
    Registrar "Teste", "Excel " & Application.Version & " | cálculo " & _
              IIf(Application.Calculation = xlCalculationAutomatic, "automático", "manual")
    lista = ListarSuplementos()
    Registrar "Suplementos ativos", IIf(lista = "", "nenhum encontrado", lista)
    okL = BuscarSerie(BLOCO_LAB, CDate(Int(CDbl(Now) * 48) / 48 - 1), Now, tt, vv, nL)
    If okL Then Registrar "Laboratório", nL & " leituras; " & ResumoSerie(tt, vv, nL)
    okP = BuscarSerie(BLOCO_PRENSA, CDate(Int(CDbl(Now) * 1440) / 1440 - 2 / 24), Now, tt, vv, nP)
    If okP Then Registrar "Prensa", nP & " leituras; " & ResumoSerie(tt, vv, nP)
    DesligarConsultas
    msg = "Teste de conexão com o PIMS" & vbCrLf & vbCrLf & _
          "Laboratório (últimas 24 h): " & IIf(okL, nL & " leituras", "falhou") & vbCrLf & _
          "Prensa (últimas 2 h): " & IIf(okP, nP & " leituras", "falhou") & vbCrLf & vbCrLf & _
          "O detalhe de cada etapa está na aba Log."
    Ampulheta False
    mOcupado = False
    Aviso msg, IIf(okL And okP And nL > 0 And nP > 0, vbInformation, vbExclamation)
    Exit Sub
Falha:
    Ampulheta False
    mOcupado = False
    Registrar "Erro", Err.Number & ": " & Err.Description
    Aviso "Erro no teste: " & Err.Description & vbCrLf & vbCrLf & "Detalhes na aba Log.", vbExclamation
End Sub

' Suplementos ativos no Excel (para conferir se o Aspen Process Data está carregado)
Private Function ListarSuplementos() As String
    Dim a As Object, r As String
    On Error GoTo Fim
    For Each a In Application.COMAddIns
        If a.Connect Then r = r & a.Description & "; "
    Next a
    For Each a In Application.AddIns
        If a.Installed Then r = r & a.Name & "; "
    Next a
Fim:
    ListarSuplementos = r
End Function

Private Function ResumoSerie(ByRef tt() As Date, ByRef vv() As Variant, ByVal n As Long) As String
    Dim i As Long, nNum As Long, r As String
    For i = 1 To n
        If ENumero(vv(i)) Then
            nNum = nNum + 1
            If nNum <= 3 Then r = r & Format$(tt(i), "dd\/mm hh:mm") & " = " & CStr(vv(i)) & "; "
        End If
    Next i
    ResumoSerie = nNum & " com valor numérico" & IIf(r = "", "", " (" & r & "...)")
End Function

' Busca uma tag de ini a fim, dividindo em consultas de até LINHAS_CONSULTA pontos.
Private Function BuscarSerie(ByVal bloco As Long, ByVal ini As Date, ByVal fim As Date, _
                             ByRef tt() As Date, ByRef vv() As Variant, ByRef n As Long) As Boolean
    Dim passo As Double, a As Date, b As Date, dados As Variant, nLin As Long, i As Long
    Dim t As Date, ultimo As Boolean, cap As Long

    passo = PassoDias(Cfg(bloco, "periodo"))
    If passo <= 0 Then
        Aviso "Intervalo de consulta inválido na aba Configuracoes: " & Cfg(bloco, "periodo") & vbCrLf & _
              "Use, por exemplo, 1m, 5m, 30m ou 2h.", vbExclamation
        Exit Function
    End If
    cap = 2000
    ReDim tt(1 To cap)
    ReDim vv(1 To cap)
    n = 0
    a = ini
    Do While CDbl(a) < CDbl(fim)
        b = CDate(CDbl(a) + LINHAS_CONSULTA * passo)
        ultimo = (CDbl(b) >= CDbl(fim))
        If ultimo Then b = fim
        Application.StatusBar = "Consultando " & NomeBloco(bloco) & ": " & Format$(a, "dd\/mm hh:mm") & _
                                " a " & Format$(b, "dd\/mm hh:mm")
        If Not ConsultarBloco(bloco, a, b, dados, nLin) Then Exit Function
        For i = 1 To nLin
            t = dados(i, 1)
            ' cada consulta cobre [a, b); pontos repetidos na emenda das consultas ficam de fora
            If CDbl(t) >= CDbl(a) - 1 / 86400 And (CDbl(t) < CDbl(b) - 1 / 86400 Or ultimo) Then
                If n = 0 Then
                    n = 1
                ElseIf CDbl(t) > CDbl(tt(n)) + 1 / 86400 Then
                    n = n + 1
                Else
                    GoTo proximo
                End If
                If n > cap Then
                    cap = cap * 2
                    ReDim Preserve tt(1 To cap)
                    ReDim Preserve vv(1 To cap)
                End If
                tt(n) = t
                vv(n) = dados(i, 2)
            End If
proximo:
        Next i
        a = b
    Loop
    BuscarSerie = True
End Function

' ---------------------------------------------------------------- consulta ao Aspen
' Escreve o período e a fórmula do bloco, espera o retorno e devolve dados(1..nLin, 1..2) = data/hora, valor
Private Function ConsultarBloco(ByVal bloco As Long, ByVal ini As Date, ByVal fim As Date, _
                                ByRef dados As Variant, ByRef nLin As Long) As Boolean
    Dim anc As Range, limite As Double, t0 As Single, tEst As Single
    Dim assin As String, assinAnt As String, pronto As Boolean, semDados As Boolean, txt As String
    Dim passo As Double, r As Long, saida As Range, d() As Variant, v As Variant, t As Date

    Set anc = Aba("Dados_PIMS").Range(Cfg(bloco, "ancora"))
    Set saida = Aba("Dados_PIMS").Range(Cfg(bloco, "saida"))
    passo = PassoDias(Cfg(bloco, "periodo"))
    limite = 300
    If ENumero(Nm("cfgTimeout").Value) Then limite = CDbl(Nm("cfgTimeout").Value)

    ' O suplemento cria a matriz de resultados sozinho; sobras de uma consulta anterior
    ' com outro tamanho fazem o Excel recusar ("Não é possível alterar parte de uma matriz").
    ' A consulta anterior vira texto antes de mudar o período: as duas usam Dados_PIMS!B3:B4 e,
    ' se continuasse como fórmula, o Aspen a refaria junto com a nova.
    DesligarConsultas
    GravarPeriodo ini, fim
    If Not EscreverFormula(bloco) Then
        Registrar NomeBloco(bloco), "Fórmula recusada pelo Excel: " & Left$(anc.Formula, 200)
        Exit Function
    End If
    Registrar NomeBloco(bloco), "Consulta " & Format$(ini, "dd\/mm\/yyyy hh:mm") & " a " & Format$(fim, "dd\/mm\/yyyy hh:mm") & _
              " | fórmula: " & anc.Formula
    If Application.Calculation <> xlCalculationAutomatic Then anc.Calculate

    t0 = Timer
    tEst = Timer
    Do
        DoEvents
        txt = TextoResposta(bloco)
        If InStr(1, txt, " is invalid", vbTextCompare) > 0 Then
            Registrar NomeBloco(bloco), "Tag recusada: " & Left$(txt, 200)
            Aviso "O PIMS não reconheceu a tag de " & NomeBloco(bloco) & ": " & Cfg(bloco, "tag") & vbCrLf & vbCrLf & _
                  "Resposta: " & Left$(txt, 200) & vbCrLf & vbCrLf & _
                  "Corrigir a tag ou o tipo (mapa) na aba Configuracoes.", vbExclamation
            Exit Function
        End If
        pronto = TemData(saida)
        semDados = False
        If Not pronto And SegundosDesde(t0) >= SEM_DADOS_SEG Then
            semDados = (LCase$(Left$(Trim$(anc.Text), 7)) = "success") And Vazio(saida.Cells(1, 1).Value)
        End If
        assin = Assinatura(saida) & "|" & txt
        If assin <> assinAnt Then
            assinAnt = assin
            tEst = Timer
        End If
        If (pronto Or semDados) And SegundosDesde(tEst) >= ESTAVEL_SEG Then Exit Do
        If SegundosDesde(t0) > limite Then
            Registrar NomeBloco(bloco), "Sem resposta em " & Format$(limite, "0") & " s | âncora: " & Left$(anc.Text, 150) & _
                      " | saída: " & Left$(saida.Cells(1, 1).Text & " " & saida.Cells(1, 2).Text, 100)
            If Aviso("O PIMS não respondeu à consulta de " & NomeBloco(bloco) & " (" & _
                     Format$(ini, "dd\/mm\/yyyy hh:mm") & " a " & Format$(fim, "dd\/mm\/yyyy hh:mm") & ")." & vbCrLf & vbCrLf & _
                     "Resposta: " & Left$(txt, 200) & vbCrLf & vbCrLf & _
                     "Verificar: suplemento Aspen Process Data ativo, servidor, tag e formato da data (aba Configuracoes)." & _
                     vbCrLf & vbCrLf & "Abrir a aba Dados_PIMS?", vbExclamation + vbYesNo) = vbYes Then MostrarDadosPIMS
            Exit Function
        End If
    Loop

    ' Lê todas as linhas com data e hora (linhas de cabeçalho ou mensagem ficam de fora)
    nLin = 0
    If pronto Then
        Dim arr As Variant
        arr = saida.Resize(LINHAS_CONSULTA + 20, 2).Value
        ReDim d(1 To UBound(arr, 1), 1 To 2)
        For r = 1 To UBound(arr, 1)
            If LerData(arr(r, 1), t) Then
                v = arr(r, 2)
                If IsError(v) Then v = Empty
                nLin = nLin + 1
                d(nLin, 1) = t
                d(nLin, 2) = v
            End If
        Next r
        OrdenarPorData d, nLin
        dados = d
    End If
    txt = IIf(pronto, nLin & " linhas lidas", "Respondeu sem dados") & " em " & _
          Format$(SegundosDesde(t0), "0") & " s | âncora: " & Left$(anc.Text, 120)
    If nLin > 0 Then txt = txt & " | primeira: " & Format$(d(1, 1), "dd\/mm hh:mm") & " = " & CStr(d(1, 2)) & _
                           " | última: " & Format$(d(nLin, 1), "dd\/mm hh:mm") & " = " & CStr(d(nLin, 2))
    Registrar NomeBloco(bloco), txt
    ConsultarBloco = True
End Function

' Argumentos da GetCalculationValues (mesma montagem do Relatório Físico)
Private Function Argumentos(ByVal bloco As Long) As String
    Dim anc As String
    anc = Cfg(bloco, "ancora")
    Argumentos = """time,attribute""," & Q(Cfg(bloco, "tag")) & "," & Q(Servidor()) & "," & Q(Cfg(bloco, "mapa")) & _
        ",Dados_PIMS!$B$3,Dados_PIMS!$B$4," & Q(Cfg(bloco, "periodo")) & ",0,"""",0," & Q(Cfg(bloco, "calculo")) & _
        ",0,1560,0,0,1,1," & _
        "ADDRESS(ROW(Dados_PIMS!" & anc & "),COLUMN(Dados_PIMS!" & anc & "),1,,""Dados_PIMS"")," & _
        """Dados_PIMS!" & Cfg(bloco, "saida") & """,1)"
End Function

' Escreve a fórmula de consulta. Tenta o nome da função como configurado e, se o Excel não
' reconhecer, o nome completo do suplemento.
Private Function EscreverFormula(ByVal bloco As Long) As Boolean
    Dim anc As Range, prefixo As String, k As Long, v As Variant
    Set anc = Aba("Dados_PIMS").Range(Cfg(bloco, "ancora"))
    prefixo = Trim$(CStr(Nm("cfgPrefixo").Value))
    If prefixo = "" Then prefixo = "=GetCalculationValues("
    For k = 1 To 2
        Err.Clear
        On Error Resume Next
        anc.Formula = prefixo & Argumentos(bloco)
        If Err.Number = 0 Then
            v = anc.Value
            If Not IsError(v) Then
                EscreverFormula = True
            ElseIf v <> CVErr(xlErrName) Then
                EscreverFormula = True
            End If
        End If
        On Error GoTo 0
        If EscreverFormula Then
            If k = 2 Then Nm("cfgPrefixo").Value = "'" & prefixo
            Exit Function
        End If
        prefixo = PREFIXO_XLL
    Next k
    anc.Value = "'Fórmula recusada pelo Excel"
    Aviso "O Excel não reconheceu a função GetCalculationValues." & vbCrLf & vbCrLf & _
          "Verificar se o suplemento Aspen Process Data está ativo (Arquivo > Opções > Suplementos) " & _
          "e se a mesma consulta funciona no Relatório Físico.", vbExclamation
End Function

' Depois de ler os resultados, a fórmula vira texto e as matrizes são apagadas, para o Aspen
' não refazer a consulta a cada recálculo do Excel.
Private Sub DesligarConsultas()
    Dim bloco As Long, c As Range
    On Error Resume Next
    For bloco = BLOCO_PRENSA To BLOCO_LAB
        Set c = Aba("Dados_PIMS").Range(Cfg(bloco, "ancora"))
        If c.HasFormula Then c.Value = "'Consulta concluída em " & Format$(Now, "dd\/mm\/yyyy hh:mm:ss")
        LimparSaida bloco
    Next bloco
    On Error GoTo 0
End Sub

Private Sub LimparSaida(ByVal bloco As Long)
    Dim zona As Range, c As Range
    On Error Resume Next
    Set zona = Aba("Dados_PIMS").Range(Cfg(bloco, "saida")).Resize(LINHAS_CONSULTA + 20, 2)
    For Each c In zona.Cells
        If c.HasArray Then c.CurrentArray.ClearContents
    Next c
    zona.ClearContents
    On Error GoTo 0
End Sub

' Início e fim do período nas células usadas pelas fórmulas (Dados_PIMS!B3:B4)
Private Sub GravarPeriodo(ByVal ini As Date, ByVal fim As Date)
    Dim modo As String, c3 As Range, c4 As Range
    Set c3 = Aba("Dados_PIMS").Range("B3")
    Set c4 = Aba("Dados_PIMS").Range("B4")
    modo = UCase$(CStr(Nm("cfgFormatoData").Value))
    If InStr(modo, "MM/DD") > 0 Then
        EscreverTexto c3, Format$(ini, "mm\/dd\/yyyy hh:mm:ss")
        EscreverTexto c4, Format$(fim, "mm\/dd\/yyyy hh:mm:ss")
    ElseIf InStr(modo, "TEXTO") > 0 Then
        EscreverTexto c3, Format$(ini, "dd\/mm\/yyyy hh:mm:ss")
        EscreverTexto c4, Format$(fim, "dd\/mm\/yyyy hh:mm:ss")
    Else
        c3.NumberFormat = "dd/mm/yyyy hh:mm:ss"
        c3.Value = ini
        c4.NumberFormat = "dd/mm/yyyy hh:mm:ss"
        c4.Value = fim
    End If
End Sub

Private Sub EscreverTexto(ByVal c As Range, ByVal texto As String)
    c.NumberFormat = "@"
    c.Value = texto
End Sub

' Texto da âncora e da primeira célula da saída, onde o Aspen escreve as mensagens
Private Function TextoResposta(ByVal bloco As Long) As String
    Dim v As Variant
    On Error Resume Next
    v = Aba("Dados_PIMS").Range(Cfg(bloco, "ancora")).Value
    If Not IsError(v) Then TextoResposta = CStr(v)
    v = Aba("Dados_PIMS").Range(Cfg(bloco, "saida")).Value
    If Not IsError(v) Then TextoResposta = TextoResposta & " " & CStr(v)
End Function

' A saída já tem data e hora em alguma das primeiras linhas (o Aspen começou a devolver resultados)
Private Function TemData(ByVal saida As Range) As Boolean
    Dim r As Long, t As Date
    For r = 1 To 3
        If LerData(saida.Cells(r, 1).Value, t) Then
            TemData = True
            Exit Function
        End If
    Next r
End Function

' Ordena as linhas 1..n de d por data e hora (o Aspen pode devolver do mais novo para o mais antigo)
Private Sub OrdenarPorData(ByRef d() As Variant, ByVal n As Long)
    Dim i As Long, j As Long, t As Variant, v As Variant, crescente As Boolean
    If n < 2 Then Exit Sub
    crescente = True
    For i = 2 To n
        If CDbl(d(i, 1)) < CDbl(d(i - 1, 1)) Then
            crescente = False
            Exit For
        End If
    Next i
    If crescente Then Exit Sub
    For i = 2 To n
        t = d(i, 1)
        v = d(i, 2)
        j = i - 1
        Do While j >= 1
            If CDbl(d(j, 1)) <= CDbl(t) Then Exit Do
            d(j + 1, 1) = d(j, 1)
            d(j + 1, 2) = d(j, 2)
            j = j - 1
        Loop
        d(j + 1, 1) = t
        d(j + 1, 2) = v
    Next i
End Sub

' Resumo do conteúdo atual da saída (quantidade e soma dos números), para saber quando parou de mudar
Private Function Assinatura(ByVal saida As Range) As String
    Dim arr As Variant, i As Long, n As Long, soma As Double
    arr = saida.Resize(LINHAS_CONSULTA + 10, 2).Value
    For i = 1 To UBound(arr, 1)
        If ENumero(arr(i, 2)) Then
            n = n + 1
            soma = soma + Num(arr(i, 2))
        End If
        If Not IsError(arr(i, 1)) Then
            If Not IsEmpty(arr(i, 1)) Then n = n + 1000000
        End If
    Next i
    Assinatura = n & "|" & soma
End Function

' ---------------------------------------------------------------- gravação nas abas de dados
' Prensa: cada resultado novo vira uma linha (data e hora em que apareceu no PIMS, valor).
' A tag volta a zero entre os ensaios ou mantém o valor até o próximo ensaio; nos dois casos
' só a primeira leitura de cada resultado é gravada.
Private Function GravarPrensa(ByRef tt() As Date, ByRef vv() As Variant, ByVal n As Long) As Long
    Dim d() As Variant, i As Long, k As Long, x As Double, ant As Double
    Aba("Prensa").Range("A2:B" & MAX_PRENSA + 1).ClearContents
    If n = 0 Then Exit Function
    ReDim d(1 To MAX_PRENSA, 1 To 2)
    ant = 0
    For i = 1 To n
        If ENumero(vv(i)) Then
            x = Num(vv(i))
            If x > 0 Then
                If ant = 0 Or Abs(x - ant) > 0.000001 Then
                    k = k + 1
                    If k > MAX_PRENSA Then
                        k = MAX_PRENSA
                        Aviso "Mais de " & MAX_PRENSA & " resultados da prensa. Reduza o período.", vbExclamation
                        Exit For
                    End If
                    d(k, 1) = tt(i)
                    d(k, 2) = x
                End If
                ant = x
            Else
                ant = 0
            End If
        End If
    Next i
    If k > 0 Then Aba("Prensa").Range("A2").Resize(k, 2).Value = d
    GravarPrensa = k
End Function

' MES: um resultado por composto. Com a consulta em intervalos menores que o composto, o mesmo
' resultado pode aparecer em intervalos seguidos: fica só o primeiro.
Private Function GravarMES(ByRef tt() As Date, ByRef vv() As Variant, ByVal n As Long) As Long
    Dim d() As Variant, i As Long, k As Long, x As Double, temAnt As Boolean, ant As Double
    Aba("MES").Range("A2:B" & MAX_MES + 1).ClearContents
    If n = 0 Then Exit Function
    ReDim d(1 To MAX_MES, 1 To 2)
    For i = 1 To n
        If ENumero(vv(i)) Then
            x = Num(vv(i))
            If x > 0 And Not (temAnt And Abs(x - ant) < 0.000001) Then
                k = k + 1
                If k > MAX_MES Then
                    k = MAX_MES
                    Aviso "Mais de " & MAX_MES & " resultados do laboratório. Reduza o período.", vbExclamation
                    Exit For
                End If
                d(k, 1) = tt(i)
                d(k, 2) = x
            End If
            temAnt = (x > 0)
            ant = x
        Else
            temAnt = False
        End If
    Next i
    If k > 0 Then Aba("MES").Range("A2").Resize(k, 2).Value = d
    GravarMES = k
End Function

' ---------------------------------------------------------------- configuração
' Parâmetros de cada bloco de consulta (aba Configuracoes e posições na aba Dados_PIMS)
Private Function Cfg(ByVal bloco As Long, ByVal campo As String) As String
    Dim s As String
    s = IIf(bloco = BLOCO_PRENSA, "Prensa", "Lab")
    Select Case campo
        Case "ancora": Cfg = IIf(bloco = BLOCO_PRENSA, "A7", "E7")
        Case "saida": Cfg = IIf(bloco = BLOCO_PRENSA, "A9", "E9")
        Case Else: Cfg = Trim$(CStr(Nm("cfg" & UCase$(Left$(campo, 1)) & Mid$(campo, 2) & s).Value))
    End Select
End Function

Private Function NomeBloco(ByVal bloco As Long) As String
    NomeBloco = IIf(bloco = BLOCO_PRENSA, "prensa automática", "laboratório (MES)")
End Function

Private Function Servidor() As String
    Servidor = Trim$(CStr(Nm("cfgServidor").Value))
    If Servidor = "" Then Servidor = "UBU"
End Function

' Intervalo da consulta ("1m", "30m", "2h", "1d", "30s") em dias
Private Function PassoDias(ByVal p As String) As Double
    Dim num As String, uni As String
    p = LCase$(Trim$(p))
    If Len(p) < 2 Then Exit Function
    num = Left$(p, Len(p) - 1)
    uni = Right$(p, 1)
    If Not IsNumeric(num) Then Exit Function
    Select Case uni
        Case "s": PassoDias = Val(num) / 86400
        Case "m": PassoDias = Val(num) / 1440
        Case "h": PassoDias = Val(num) / 24
        Case "d": PassoDias = Val(num)
    End Select
End Function

' Início da janela de 2 h que contém t (janelas nas horas pares, mais o deslocamento em meias horas)
Private Function AlinharJanela(ByVal t As Date) As Date
    Dim desl As Double, k As Double
    desl = 0
    If ENumero(Nm("cfgDeslocamento").Value) Then desl = CDbl(Nm("cfgDeslocamento").Value) / 48
    k = Int((CDbl(t) - desl) * 12 + 0.0000001)
    AlinharJanela = CDate(Round((k / 12 + desl) * 1440, 0) / 1440)
End Function

' ---------------------------------------------------------------- utilitários
Private Function Aba(ByVal nome As String) As Worksheet
    Set Aba = ThisWorkbook.Worksheets(nome)
End Function

Private Sub LimparLog()
    On Error Resume Next
    Aba("Log").Range("A3:C1000").ClearContents
End Sub

' Registra uma etapa na aba Log (data e hora, etapa, detalhe)
Private Sub Registrar(ByVal etapa As String, ByVal detalhe As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = Aba("Log")
    r = 3
    Do While Not IsEmpty(ws.Cells(r, 2).Value) And r < 1000
        r = r + 1
    Loop
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = etapa
    ws.Cells(r, 3).Value = "'" & Left$(detalhe, 1000)
End Sub

Private Function Nm(ByVal nome As String) As Range
    Set Nm = ThisWorkbook.Names(nome).RefersToRange
End Function

Private Function Q(ByVal texto As String) As String
    Q = """" & Replace(texto, """", """""") & """"
End Function

Private Function Vazio(ByVal v As Variant) As Boolean
    If IsError(v) Then
        Vazio = True
    ElseIf IsEmpty(v) Then
        Vazio = True
    Else
        Vazio = (Trim$(CStr(v)) = "")
    End If
End Function

Private Function ENumero(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        ENumero = IsNumeric(v) And Trim$(v) <> ""
    Else
        ENumero = IsNumeric(v)
    End If
End Function

' Número de uma célula; texto com vírgula ou ponto decimal (independe do idioma do Windows)
Private Function Num(ByVal v As Variant) As Double
    If VarType(v) = vbString Then
        Num = Val(Replace(Trim$(v), ",", "."))
    Else
        Num = CDbl(v)
    End If
End Function

' Lê data e hora de uma célula: valor de data do Excel, número ou texto dd/mm/aaaa hh:mm[:ss]
Private Function LerData(ByVal v As Variant, ByRef d As Date) As Boolean
    Dim s As String, p() As String, dt() As String, hr() As String
    On Error GoTo Nao
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        d = v
        LerData = True
    ElseIf VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
        If CDbl(v) > 1 Then
            d = CDate(CDbl(v))
            LerData = True
        End If
    Else
        s = Trim$(CStr(v))
        If s = "" Then Exit Function
        p = Split(s, " ")
        dt = Split(p(0), "/")
        If UBound(dt) = 2 Then
            d = DateSerial(CInt(dt(2)), CInt(dt(1)), CInt(dt(0)))
            If UBound(p) >= 1 Then
                hr = Split(p(UBound(p)), ":")
                d = d + TimeSerial(CInt(hr(0)), CInt(hr(1)), IIf(UBound(hr) >= 2, Val(hr(UBound(hr))), 0))
            End If
            LerData = True
        ElseIf IsDate(s) Then
            d = CDate(s)
            LerData = True
        End If
    End If
    Exit Function
Nao:
    LerData = False
End Function

Private Function SegundosDesde(ByVal t0 As Single) As Double
    SegundosDesde = Timer - t0
    If SegundosDesde < 0 Then SegundosDesde = SegundosDesde + 86400
End Function

Private Function Aviso(ByVal msg As String, Optional ByVal estilo As Long = 64) As Long
    Aviso = MsgBox(msg, estilo, "Correlação prensa x laboratório")
End Function

Private Sub Ampulheta(ByVal ligada As Boolean)
    Application.Cursor = IIf(ligada, xlWait, xlDefault)
End Sub

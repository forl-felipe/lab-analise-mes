Option Explicit

' ============================================================================
'  ModMES - busca dos resultados das Usinas 3 e 4 no MES (Aspen IP.21)
'  e preenchimento da aba Informativo
' ============================================================================

Private mLinhas() As Long
Private mCarregado As Boolean

' Botao 1 do Painel
Public Sub AtualizarMES()
    Dim dt As Date, turno As String, turma As String, resp As String
    Dim ini As Date, fim As Date, dados As Variant, fonte As String
    Dim ok As Boolean, simulado As Boolean, n As Long

    PrepararEdicao
    If Not LerTurnoPainel(dt, turno, turma, resp, True) Then Exit Sub
    ini = InicioTurno(dt, turno)
    fim = CDate(CDbl(ini) + NSLOT * HORAS_SLOT / 24)

    fonte = UCase$(CfgTxt("cfgFonte"))
    simulado = (Left$(fonte, 3) = "SIM")

    Application.StatusBar = "Consultando resultados das Usinas 3 e 4 (" & _
        Format$(ini, "dd\/mm hh:mm") & " a " & Format$(fim, "dd\/mm hh:mm") & ")..."
    If simulado Then
        dados = DadosSimulados()
        ok = True
    Else
        ok = ConsultarMES(ini, fim, dados)
    End If
    Application.StatusBar = False
    If Not ok Then Exit Sub

    n = PreencherInformativo(dados, dt, turno, turma, resp, ini, fim, simulado)
    IrInformativo
    If simulado Then
        Aviso "Informativo preenchido com DADOS SIMULADOS (" & n & " valores)." & vbCrLf & _
               "Para usar o MES, altere 'Fonte dos dados' para MES na aba Configurações.", _
               vbInformation, "Simulação"
    ElseIf n = 0 Then
        Aviso "O MES respondeu, mas nenhum resultado válido foi encontrado para o turno." & vbCrLf & _
               "Confira as tags na aba Configurações ou a aba Dados_MES.", vbExclamation, "MES"
    Else
        Aviso n & " resultados carregados do MES para " & Format$(dt, "dd\/mm\/yyyy") & _
               " - " & turno & ".", vbInformation, "MES"
    End If
End Sub

' ---------------------------------------------------------------------------
'  Consulta ao MES - mesmo formato da planilha de referencia (out/2014):
'  uma formula GetCalculationValues por grupo (Qualidade, Producao, Ritmo de
'  processo), com tags, servidores e mapas em TEXTO LITERAL dentro da formula.
'  Os grupos estao na aba _Mapa (colunas K:Q).
' ---------------------------------------------------------------------------
Private Const MAPA_COL_BLOCO As Long = 11

Public Function NumBlocos() As Long
    Dim r As Long
    r = 2
    Do While Not Vazio(shMapa.Cells(r, MAPA_COL_BLOCO).Value)
        r = r + 1
    Loop
    NumBlocos = r - 2
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
    ps = Split(BlocoInfo(b, 5), ",")
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
    Dim b As Long, c As Range, atual As String, prefixo As String, nova As String
    For b = 1 To NumBlocos()
        Set c = shDadosMES.Range(BlocoInfo(b, 3))
        atual = c.Formula
        If InStr(atual, "(") > 0 And InStr(1, atual, "GetCalculationValues", vbTextCompare) > 0 Then
            prefixo = Left$(atual, InStr(atual, "("))
        Else
            prefixo = "=_xll.AspenTech.PME.ProcessData.Functions.GetCalculationValues("
        End If
        nova = prefixo & ArgumentosBloco(b)
        If nova <> atual Then c.Formula = nova
    Next b
End Sub

Private Function SaidaBloco(ByVal b As Long) As Range
    Dim ps As Variant
    ps = Split(BlocoInfo(b, 5), ",")
    Set SaidaBloco = shDadosMES.Range(BlocoInfo(b, 4)).Resize(NSLOT, 1 + 2 * (UBound(ps) - LBound(ps) + 1))
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

' Escreve o periodo, regera as formulas, recalcula e espera o retorno de cada consulta
Private Function ConsultarMES(ByVal ini As Date, ByVal fim As Date, ByRef dados As Variant) As Boolean
    Dim t0 As Single, limite As Double, msg As String, nb As Long, b As Long, nOk As Long
    Dim pronto() As Boolean, d() As Variant, arr As Variant, ps As Variant
    Dim i As Long, k As Long, s As Long, p As Long, falhas As String

    ConsultarMES = False
    nb = NumBlocos()
    ReDim pronto(1 To nb)
    Nm("mesInicio").Value = Format$(ini, "dd\/mm\/yyyy hh:mm:ss")
    Nm("mesFim").Value = Format$(fim, "dd\/mm\/yyyy hh:mm:ss")
    On Error GoTo FalhaFormula
    RegerarFormulasMES
    On Error GoTo 0
    Application.CalculateFull

    limite = 60
    If ENumero(Nm("cfgTimeout").Value) Then limite = CDbl(Nm("cfgTimeout").Value)
    t0 = Timer
    Do
        DoEvents
        nOk = 0
        For b = 1 To nb
            If Not pronto(b) Then pronto(b) = BlocoPronto(b, ini)
            If pronto(b) Then nOk = nOk + 1
        Next b
        If nOk = nb Then Exit Do
        If Timer < t0 Then t0 = t0 - 86400
        If Timer - t0 > limite Then Exit Do
    Loop

    For b = 1 To nb
        If Not pronto(b) Then
            falhas = falhas & "  - " & BlocoInfo(b, 1) & ": " & _
                     Left$(shDadosMES.Range(BlocoInfo(b, 3)).Text & " " & SaidaBloco(b).Cells(1, 1).Text, 150) & vbCrLf
        End If
    Next b

    If nOk = 0 Then
        msg = "O MES não retornou dados para " & Format$(ini, "dd\/mm\/yyyy hh:mm") & "." & vbCrLf & vbCrLf & _
              "Resposta de cada consulta:" & vbCrLf & falhas & vbCrLf & _
              "Verifique:" & vbCrLf & _
              "  - Suplemento Aspen Process Data ativo (#NOME? = suplemento ausente);" & vbCrLf & _
              "  - Fonte de dados '" & CfgTxt("cfgServidor") & "' (aba Configurações);" & vbCrLf & _
              "  - Tags e mapas na aba Configurações." & vbCrLf & vbCrLf & _
              "Deseja abrir a aba Dados_MES para conferir?"
        If Aviso(msg, vbExclamation + vbYesNo, "MES sem resposta") = vbYes Then MostrarDadosMES
        Exit Function
    End If
    If falhas <> "" Then
        Aviso "Algumas consultas do MES não responderam e ficarão em branco:" & vbCrLf & falhas, _
              vbExclamation, "MES - resposta parcial"
    End If

    ReDim d(1 To NSLOT, 1 To 1 + 2 * NPARAM)
    For b = 1 To nb
        If pronto(b) Then
            arr = SaidaBloco(b).Value
            ps = Split(BlocoInfo(b, 5), ",")
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
    ConsultarMES = True
    Exit Function

FalhaFormula:
    Aviso "Não foi possível montar as fórmulas de consulta ao MES:" & vbCrLf & Err.Description & vbCrLf & _
          "Confira as tags na aba Configurações.", vbExclamation, "MES"
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
                If Not ENumero(tip) Then
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

Private Function PreencherInformativo(ByVal dados As Variant, ByVal dt As Date, ByVal turno As String, _
                                      ByVal turma As String, ByVal resp As String, ByVal ini As Date, _
                                      ByVal fim As Date, ByVal simulado As Boolean) As Long
    Dim ws As Worksheet, p As Long, k As Long, s As Long, nLin As Long, n As Long
    Dim vmin As Variant, vmax As Variant, v As Variant, lin As Long

    Set ws = shInformativo
    nLin = UBound(dados, 1) - LBound(dados, 1) + 1
    If nLin > NSLOT Then nLin = NSLOT

    For p = 1 To NPARAM
        vmin = CfgCel(p, CFG_COL_VMIN)
        vmax = CfgCel(p, CFG_COL_VMAX)
        For k = 1 To 2
            lin = LinhaInformativo(p, k)
            ws.Cells(lin, INF_COL_H1).Resize(1, NSLOT).ClearContents
            For s = 1 To nLin
                v = dados(LBound(dados, 1) + s - 1, LBound(dados, 2) + (p - 1) * 2 + k)
                If ValorValido(v, vmin, vmax) Then
                    ws.Cells(lin, INF_COL_H1 + s - 1).Value = CDbl(v)
                    n = n + 1
                End If
            Next s
        Next k
    Next p

    Nm("iData").Value = dt
    Nm("iTurno").Value = turno
    Nm("iTurma").Value = turma
    Nm("iResp").Value = resp
    Nm("iAtualizado").Value = Now
    If simulado Then
        Nm("iFonte").Value = "SIMULAÇÃO (dados fictícios)"
    Else
        Nm("iFonte").Value = "MES - servidor " & CfgTxt("cfgServidor")
    End If
    Nm("iPeriodo").Value = Format$(ini, "dd\/mm\/yyyy hh:mm") & "  a  " & Format$(fim, "dd\/mm\/yyyy hh:mm")
    For s = 1 To NSLOT
        Nm("iHoras").Cells(1, s).Value = Format$(CDate(CDbl(ini) + (s - 1) * HORAS_SLOT / 24), "hh") & "h-" & _
                                         Format$(CDate(CDbl(ini) + s * HORAS_SLOT / 24), "hh") & "h"
    Next s
    PreencherInformativo = n
End Function

Public Sub LimparInformativo()
    Dim p As Long, k As Long
    On Error Resume Next
    shInformativo.Unprotect Password:=SENHA
    On Error GoTo 0
    For p = 1 To NPARAM
        For k = 1 To 2
            shInformativo.Cells(LinhaInformativo(p, k), INF_COL_H1).Resize(1, NSLOT).ClearContents
        Next k
    Next p
    Nm("iData").ClearContents
    Nm("iTurno").ClearContents
    Nm("iTurma").ClearContents
    Nm("iResp").ClearContents
    Nm("iAtualizado").ClearContents
    Nm("iFonte").ClearContents
    Nm("iPeriodo").ClearContents
    Nm("iObs").MergeArea.ClearContents
    ProtegerPlanilhas
End Sub

Public Function InformativoConfere(ByVal dt As Date, ByVal turno As String) As Boolean
    InformativoConfere = MesmoTurno(Nm("iData").Value, Nm("iTurno").Value, dt, turno)
End Function

' Linha da aba Informativo para o parametro p (1..NPARAM) e usina k (1 = US3, 2 = US4)
Public Function LinhaInformativo(ByVal p As Long, ByVal k As Long) As Long
    If Not mCarregado Then CarregarLinhas
    LinhaInformativo = mLinhas(p, k)
End Function

Private Sub CarregarLinhas()
    Dim r As Long
    ReDim mLinhas(1 To NPARAM, 1 To 2)
    For r = 2 To 1 + NPARAM * 2
        mLinhas(CLng(shMapa.Cells(r, 7).Value), CLng(shMapa.Cells(r, 8).Value)) = CLng(shMapa.Cells(r, 9).Value)
    Next r
    mCarregado = True
End Sub

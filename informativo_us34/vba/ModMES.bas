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

' Escreve o periodo na aba Dados_MES, recalcula as formulas Aspen e espera o retorno
Private Function ConsultarMES(ByVal ini As Date, ByVal fim As Date, ByRef dados As Variant) As Boolean
    Dim t0 As Single, limite As Double, v As Variant, pronto As Boolean, msg As String

    ConsultarMES = False
    Nm("mesInicio").Value = Format$(ini, "dd\/mm\/yyyy hh:mm:ss")
    Nm("mesFim").Value = Format$(fim, "dd\/mm\/yyyy hh:mm:ss")
    Application.CalculateFull

    limite = 60
    If ENumero(Nm("cfgTimeout").Value) Then limite = CDbl(Nm("cfgTimeout").Value)
    t0 = Timer
    Do
        DoEvents
        v = Nm("mesSaida").Cells(1, 1).Value
        If Not Vazio(v) Then
            If VarType(v) = vbDate Or IsNumeric(v) Or IsDate(v) Then
                If Abs(CDbl(CDate(v)) - CDbl(ini)) <= HORAS_SLOT / 24 + 1 / 1440 Then pronto = True
            End If
        End If
        If pronto Then Exit Do
        If Timer < t0 Then t0 = t0 - 86400
        If Timer - t0 > limite Then Exit Do
    Loop

    If Not pronto Then
        msg = "O MES não retornou dados para " & Format$(ini, "dd\/mm\/yyyy hh:mm") & "." & vbCrLf & vbCrLf & _
              "Resposta do MES: " & Left$(Nm("mesConsulta").Text & " " & Nm("mesSaida").Cells(1, 1).Text, 180) & vbCrLf & vbCrLf & _
              "Verifique:" & vbCrLf & _
              "  - Suplemento Aspen Process Explorer / Excel Add-in ativo" & vbCrLf & _
              "    (Arquivo > Opções > Suplementos). #NOME? indica suplemento ausente;" & vbCrLf & _
              "  - Nome da fonte de dados (servidor) '" & CfgTxt("cfgServidor") & "' na aba Configurações." & vbCrLf & _
              "    'Nome de fonte de dados inválido' = use o nome exibido no suplemento Aspen;" & vbCrLf & _
              "  - Tags e tipos na aba Configurações." & vbCrLf & vbCrLf & _
              "Deseja abrir a aba Dados_MES para conferir?"
        If Aviso(msg, vbExclamation + vbYesNo, "MES sem resposta") = vbYes Then MostrarDadosMES
        Exit Function
    End If
    dados = Nm("mesSaida").Value
    ConsultarMES = True
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

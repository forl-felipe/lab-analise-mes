Option Explicit

' Teste automatico (LibreOffice headless): grava o resultado em _Graf!BA1 (texto, linhas separadas por |)
Private mLog As String
Private mFalhas As Long

Private Sub Ok(ByVal cond As Boolean, ByVal desc As String)
    If cond Then
        mLog = mLog & "OK   " & desc & "|"
    Else
        mLog = mLog & "FALHA " & desc & "|"
        mFalhas = mFalhas + 1
    End If
End Sub

Private Function V(ByVal nome As String) As Variant
    V = Nm(nome).Value
End Function

Private Function Celula(ByVal i As Long, ByVal p As Long, ByVal lab As Long) As Variant
    Celula = shBatch.Cells(B_ROW_G0 + i - 1, ColValor(p, lab)).Value
End Function

Public Sub TesteLO()
    Dim n As Long, id As Long, r As Long, nL As Long, i As Long, a As Variant, b As Variant, s As String
    Dim pSE As Long, pSi As Long, salvo1 As Variant, salvo2 As Variant
    On Error GoTo Erro
    gSilencioso = True
    gTesteLO = True
    mLog = ""
    mFalhas = 0
    pSE = PIndice("SE")
    pSi = PIndice("SiO2")
    Recalcular

    ' --- estado inicial (batch 288 importado)
    Ok CStr(V("bId")) = "288", "batch inicial 288 (" & V("bId") & ")"
    Ok Celula(1, pSE, 0) = 2052 And Celula(1, pSE, 1) = 1787, "1a leitura S.E. 2052/1787"
    Ok CStr(shGraf.Cells(G_ROW_ST + pSE - 1, G_COL_ST).Value) = "NOK", "S.E. NOK no farol"
    Ok LinhaHistorico(288) > 0 And LinhaHistorico(290) > 0, "historico inicial com 288 e 290"
    Ok InStr(CStr(V("bFrase")), "Superfície Específica") > 0, "frase NOK automatica: " & V("bFrase")
    Ok InStr(CStr(V("eResumo")), "Farol: 4 de 5") > 0, "resumo do e-mail (4 de 5 OK)"
    Ok PRotulo(pSi) = "SiO" & ChrW(8322), "rotulo SiO2 com subscrito"

    ' --- simulacao: novo batch + buscar no MES
    shConfig.Unprotect
    Nm("cfgFonte").Value = "SIMULAÇÃO"
    NovoBatch
    id = CLng(V("bId"))
    Ok id = 291, "novo batch = 291 (" & id & ")"
    Ok Not ENumero(Celula(1, pSE, 0)), "novo batch sem leituras"
    Nm("bIniU").Value = DateSerial(2026, 10, 10) + TimeSerial(7, 30, 0)
    Nm("bIniG").Value = DateSerial(2026, 10, 8) + TimeSerial(5, 30, 0)
    Nm("bNHor").Value = 12
    Recalcular
    Ok Abs(CDbl(V("bDef")) - 50) < 0.01, "defasagem 50 h (" & V("bDef") & ")"
    Ok Abs(CDbl(shBatch.Cells(B_ROW_G0 + 1, B_COL_TU).Value) - CDbl(DateSerial(2026, 10, 10) + TimeSerial(9, 30, 0))) < 0.0001, _
       "2o horario Ubu 09:30"
    Ok Abs(CDbl(shBatch.Cells(B_ROW_G0, B_COL_TG).Value) - CDbl(DateSerial(2026, 10, 8) + TimeSerial(5, 30, 0))) < 0.0001, _
       "1o horario Germano 08/10 05:30"
    Ok shBatch.Cells(B_ROW_G0 + 12, B_COL_TU).Text = "", "13o horario vazio (Nº de horarios = 12)"
    BuscarMES
    Recalcular
    n = 0
    For i = 1 To 12
        If ENumero(Celula(i, pSE, 0)) Then n = n + 1
    Next i
    Ok n = 12, "simulacao: 12 S.E. Germano (" & n & ")"
    Ok ENumero(Celula(1, pSE, 1)) And Not ENumero(Celula(2, PIndice("325"), 1)), "Ubu: S.E. a cada 2 h, <325# a cada 6 h"
    Ok CLng(shGraf.Cells(G_ROW_ST + pSE - 1, G_COL_NP).Value) = 12, "12 pares S.E. no grafico"
    Ok CStr(shGraf.Cells(2 + 12, (pSE - 1) * G_BLOCO + 1).Value) = "Média", "rotulo Media apos os pares"
    salvo1 = Celula(3, pSi, 0)
    salvo2 = Celula(5, pSE, 1)
    SalvarHistorico
    r = LinhaHistorico(id)
    Ok r > 0, "291 salvo no historico (linha " & r & ")"
    nL = 0
    For i = 2 To 2000
        If Vazio(shLeituras.Cells(i, 1).Value) Then Exit For
        If CLng(shLeituras.Cells(i, 1).Value) = id Then nL = nL + 1
    Next i
    Ok nL = 12 * 5, "leituras do 291 em BD_Leituras (" & nL & ")"
    Ok InStr(gUltimaMsg, "Erro") = 0, "nenhuma mensagem de erro: " & gUltimaMsg
    Ok CStr(shHist.Cells(r, H_COL1 + 2).Value) = "03", "mineroduto gravado como texto 03 (" & shHist.Cells(r, H_COL1 + 2).Value & ")"

    ' --- abrir outro batch e voltar
    AbrirBatch 288
    Ok CStr(V("bId")) = "288" And Celula(1, pSE, 0) = 2052, "abrir 288"
    Ok CStr(V("bMin")) = "03", "mineroduto aberto como 03 (" & V("bMin") & ")"
    Ok CStr(V("pSel")) = "288", "painel em exibicao = 288"
    AbrirBatch id
    Ok Celula(3, pSi, 0) = salvo1 And Celula(5, pSE, 1) = salvo2, "reabrir 291 com os mesmos valores"
    Ok Abs(CDbl(V("bDef")) - 50) < 0.01, "reabrir 291 com a defasagem"
    ' selecao pelo painel
    Nm("pSel").Value = 290
    SelecionarBatch
    Ok CStr(V("bId")) = "290", "selecionar 290 pelo painel"
    Ok ENumero(Celula(14, pSE, 1)), "290 com 14 horarios"

    ' --- excluir
    ExcluirDoHistorico id
    Ok LinhaHistorico(id) = 0, "291 excluido"
    nL = 0
    For i = 2 To 2000
        If Vazio(shLeituras.Cells(i, 1).Value) Then Exit For
        If CLng(shLeituras.Cells(i, 1).Value) = id Then nL = nL + 1
    Next i
    Ok nL = 0, "leituras do 291 excluidas"

    ' --- importar da planilha antiga
    LimparImportacao
    With Nm("impArea")
        .Cells(1, 1).Value = "Batch"
        .Cells(1, 2).Value = 295
        .Cells(2, 1).Value = "Data"
        .Cells(3, 1).Value = "Inicio aproximado no processo"
        .Cells(4, 1).Value = "S.E Germano"
        .Cells(5, 1).Value = "S.E Alimentação 04"
        .Cells(6, 1).Value = "Diferença"
        .Cells(7, 1).Value = "< 325 # Germano"
        .Cells(8, 1).Value = "< 325 # Alimentação 04"
        For i = 1 To 4
            .Cells(2, 1 + i).Value = DateSerial(2026, 10, 12) + IIf(i = 4, 1, 0)
            .Cells(3, 1 + i).Value = TimeSerial(IIf(i = 4, 1, 17 + 2 * i), 30, 0)
            .Cells(4, 1 + i).Value = 2000 + i
            .Cells(5, 1 + i).Value = 1900 + i
            .Cells(6, 1 + i).Value = -100
            .Cells(7, 1 + i).Value = 86 + i / 10
        Next i
        .Cells(8, 2).Value = 85.5
        .Cells(10, 1).Value = "Batch"
        .Cells(10, 2).Value = 295
        .Cells(11, 1).Value = "Data"
        .Cells(12, 1).Value = "Inicio aproximado no processo"
        .Cells(13, 1).Value = "SiO2 Germano"
        .Cells(14, 1).Value = "SiO2 Alimentação 04"
        .Cells(15, 1).Value = "P Germano"
        .Cells(16, 1).Value = "PPC Alimentação 04"
        For i = 1 To 4
            .Cells(11, 1 + i).Value = DateSerial(2026, 10, 12) + IIf(i = 4, 1, 0)
            .Cells(12, 1 + i).Value = TimeSerial(IIf(i = 4, 1, 17 + 2 * i), 30, 0)
            .Cells(13, 1 + i).Value = 1.8 + i / 100
            .Cells(15, 1 + i).Value = 0.06
        Next i
        .Cells(14, 3).Value = 1.95
        .Cells(16, 5).Value = 3.5
    End With
    ImportarColado
    Recalcular
    Ok CStr(V("bId")) = "295", "importado como batch 295"
    Ok CLng(V("bNHor")) = 4, "importado com 4 horarios (" & V("bNHor") & ")"
    Ok Celula(1, pSE, 0) = 2001 And Celula(4, pSE, 1) = 1904, "importado S.E."
    Ok Celula(1, PIndice("325"), 1) = 85.5, "importado <325# Ubu"
    Ok Celula(2, pSi, 1) = 1.95 And Celula(4, PIndice("PPC"), 1) = 3.5, "importado SiO2/PPC Ubu"
    Ok Abs(CDbl(shBatch.Cells(B_ROW_G0 + 3, B_COL_TU).Value) - CDbl(DateSerial(2026, 10, 13) + TimeSerial(1, 30, 0))) < 0.0001, _
       "importado 4o horario 13/10 01:30"
    Ok ENumero(Celula(1, PIndice("P"), 0)) And Not ENumero(Celula(1, PIndice("P"), 1)), "importado P so Germano"

    ' --- MES: formula e tags recusadas
    shConfig.Cells(C_ROW_TAG0 + 1, C_COL_TAG).Value = "TAG-UBU-SE"
    shConfig.Cells(C_ROW_TAG0 + 2 * (pSi - 1) + 1, C_COL_TAG).Value = "TAG-UBU-SIO2"
    s = ArgumentosLab(1, "1," & pSi, shDadosMES.Range("A40"), shDadosMES.Range("A42"))
    Ok InStr(s, """TAG-UBU-SE,TAG-UBU-SIO2"",""UBU,UBU"",""IP_MESVALOR,IP_MESVALOR"",Dados_MES!$B$6,Dados_MES!$B$7,""2h""") > 0, _
       "argumentos da consulta Ubu: " & Left$(s, 140)
    s = TesteTagsInvalidas("Error: Tag Name TAG-UBU-SIO2 is invalid")
    Ok InStr(s, "TAG-UBU-SIO2") > 0, "tag recusada identificada"

    GoTo Fim
Erro:
    mLog = mLog & "ERRO " & Err.Number & " " & Err.Description & "|"
    mFalhas = mFalhas + 1
Fim:
    shGraf.Range("BA1").Value = "'" & IIf(mFalhas = 0, "TUDO OK", mFalhas & " FALHA(S)") & "|" & mLog
End Sub

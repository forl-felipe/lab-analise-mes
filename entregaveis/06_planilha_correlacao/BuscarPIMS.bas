Attribute VB_Name = "BuscarPIMS"
' Busca no PIMS (Aspen InfoPlus.21, via ODBC do Aspen SQLplus) o histórico da prensa automática e do MES
' e grava nas abas Prensa e MES da planilha de correlação.
'
' Como instalar: Alt+F11 > Arquivo > Importar arquivo > BuscarPIMS.bas; salve a planilha como .xlsm.
' Como usar: preencha na aba Parametros o início, o fim, a fonte ODBC (DSN) e as tags; execute BuscarDadosPIMS
' (Alt+F8). Dica: insira um botão na aba Painel e associe a esta macro.
'
' Pré-requisito: driver ODBC "AspenTech SQLplus" instalado e uma fonte de dados (DSN) apontando para o servidor
' do InfoPlus.21. A TI informa o nome da DSN. A consulta usa a tabela HISTORY do SQLplus (valores reais, REQUEST = 4).

Option Explicit

Private Const LIN_INICIO As String = "B10"
Private Const LIN_FIM As String = "B11"
Private Const LIN_DSN As String = "B12"
Private Const LIN_TAG_PRENSA As String = "B13"
Private Const LIN_TAG_MES As String = "B14"

Public Sub BuscarDadosPIMS()
    Dim par As Worksheet
    Set par = ThisWorkbook.Worksheets("Parametros")

    Dim ini As Date, fim As Date
    ini = par.Range(LIN_INICIO).Value
    fim = par.Range(LIN_FIM).Value
    If fim <= ini Then
        MsgBox "O fim do período precisa ser depois do início (aba Parametros).", vbExclamation
        Exit Sub
    End If

    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    On Error GoTo Falha
    cn.ConnectionTimeout = 30
    cn.CommandTimeout = 300
    cn.Open "DSN=" & par.Range(LIN_DSN).Value & ";"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' A busca da prensa traz uma janela extra de 2 h antes do início, para completar a primeira janela.
    Dim nPrensa As Long, nMes As Long
    nPrensa = Gravar(cn, par.Range(LIN_TAG_PRENSA).Value, ini - 2 / 24, fim + 1 / 24, ThisWorkbook.Worksheets("Prensa"), 20000)
    nMes = Gravar(cn, par.Range(LIN_TAG_MES).Value, ini, fim, ThisWorkbook.Worksheets("MES"), 5000)

    cn.Close
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("Painel").Activate
    MsgBox "Busca concluída." & vbCrLf & _
           "Prensa: " & nPrensa & " registros" & vbCrLf & _
           "MES: " & nMes & " resultados", vbInformation
    Exit Sub

Falha:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Não foi possível buscar os dados no PIMS." & vbCrLf & vbCrLf & _
           "Erro: " & Err.Description & vbCrLf & vbCrLf & _
           "Confira a fonte ODBC (DSN) e as tags na aba Parametros. Se o problema continuar, " & _
           "busque os dados pelo Aspen Process Data no Excel e cole nas abas Prensa e MES.", vbCritical
End Sub

' Consulta uma tag e grava data/hora e valor a partir da linha 2 das colunas A e B. Retorna a quantidade.
Private Function Gravar(cn As Object, tag As String, ini As Date, fim As Date, ws As Worksheet, limite As Long) As Long
    Dim sql As String
    sql = "SELECT TS, VALUE FROM HISTORY WHERE NAME = '" & Replace(tag, "'", "''") & "'" & _
          " AND TS BETWEEN '" & DataSQLplus(ini) & "' AND '" & DataSQLplus(fim) & "'" & _
          " AND REQUEST = 4 ORDER BY TS"

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, cn

    ws.Range("A2:B" & (limite + 1)).ClearContents
    Dim dados() As Variant, n As Long
    ReDim dados(1 To limite, 1 To 2)
    Do While Not rs.EOF And n < limite
        n = n + 1
        dados(n, 1) = CDate(rs.Fields(0).Value)
        If IsNull(rs.Fields(1).Value) Then dados(n, 2) = Empty Else dados(n, 2) = CDbl(rs.Fields(1).Value)
        rs.MoveNext
    Loop
    If n > 0 Then ws.Range("A2").Resize(n, 2).Value = dados
    If Not rs.EOF Then
        MsgBox "A tag " & tag & " tem mais registros do que a planilha comporta (" & limite & "). " & _
               "Reduza o período ou estenda as fórmulas da aba " & ws.Name & ".", vbExclamation
    End If
    rs.Close
    Gravar = n
End Function

' Formato de data aceito pelo SQLplus: 28-OCT-26 13:00:00 (mês sempre em inglês).
Private Function DataSQLplus(d As Date) As String
    Dim meses As Variant
    meses = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    DataSQLplus = Format(Day(d), "00") & "-" & meses(Month(d) - 1) & "-" & Format(Year(d) Mod 100, "00") & " " & _
                  Format(Hour(d), "00") & ":" & Format(Minute(d), "00") & ":" & Format(Second(d), "00")
End Function

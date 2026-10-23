Option Explicit

' Gerado por build_interlab.py - posicoes fixas das abas (nao editar a mao)
Public Const NP As Long = 5
Public Const NMAX As Long = 24
Public Const HORAS_SLOT As Double = 2
' aba Batch (1-based)
Public Const B_ROW_G0 As Long = 38
Public Const B_COL_N As Long = 2
Public Const B_COL_TU As Long = 3
Public Const B_COL_TG As Long = 4
Public Const B_COL_V1 As Long = 5
Public Const B_COL_D1 As Long = 16
' aba Configuracoes: tabela de tags (1-based)
Public Const C_ROW_TAG0 As Long = 8
' aba Configuracoes: tags do pellet feed ultra fino (1-based; mesmas colunas)
Public Const C_ROW_PF0 As Long = 21
Public Const NPF As Long = 4
' aba Batch: ultimos batches (1-based) e tipo da busca
Public Const B_ROW_ULT As Long = 18
Public Const N_ULT As Long = 3
Public Const U_COL_BATCH As Long = 2
Public Const U_COL_NIVEL As Long = 4
Public Const U_COL_CHEG As Long = 5
Public Const U_COL_FIM As Long = 7
Public Const U_COL_INIP As Long = 9
Public Const U_COL_DUR As Long = 11
Public Const U_COL_ORIGEM As Long = 13
' abas de acompanhamento: 1a linha dos blocos e analises (chaves) de cada aba
Public Const A_ROW0 As Long = 5
Public Const ACOMP_Q As String = "SiO2|P|PPC"
Public Const ACOMP_F As String = "SE|325"
Public Const SENHA_CFG As String = "Samarco@Lab"
Public Const REG_NCOL As Long = 9
' Historico: coluna (na tabela, 1-based) de campos usados nos ultimos batches
Public Const H_COL_INIU As Long = 4
Public Const H_COL_FIMU As Long = 5
Public Const H_COL_NIVEL As Long = 6
Public Const H_COL_INIP As Long = 56
Public Const PF_CHAVES As String = "Fe|SiO2|P|SE"
Public Const PF_DECIMAIS As String = "2|2|3|0"
Public Const C_COL_PARAM As Long = 2
Public Const C_COL_LAB As Long = 3
Public Const C_COL_TAG As Long = 4
Public Const C_COL_SERV As Long = 5
Public Const C_COL_MAPA As Long = 6
Public Const C_COL_VMIN As Long = 7
Public Const C_COL_VMAX As Long = 8
Public Const C_COL_TIP As Long = 9
Public Const C_COL_FREQ As Long = 10
Public Const C_COL_HORA1 As Long = 11
' aba _Graf
Public Const G_BLOCO As Long = 9
Public Const G_ROW_ST As Long = 31
Public Const G_COL_NP As Long = 7
Public Const G_COL_MA As Long = 2
Public Const G_COL_MB As Long = 3
Public Const G_COL_D As Long = 4
Public Const G_COL_ST As Long = 6
' Historico: colunas da tabela tbl_Batches (ordem)
Public Const H_NCOL As Long = 61
Public Const H_ROW_HDR As Long = 7
Public Const H_COL1 As Long = 2
' coluna (na tabela) da 1a estatistica (media A do 1o parametro)
Public Const H_COL_STAT1 As Long = 36
Public Const H_NCAMPOS As Long = 35
Public Const L_NCOL As Long = 10
' chaves e rotulos dos parametros (ordem das colunas da aba Batch)
Public Const P_CHAVES As String = "SE|325|SiO2|P|PPC"
Public Const P_DECIMAIS As String = "0|1|2|3|2"
' ordem dos graficos no Painel (ChartObjects 1..5)
Public Const GRAF_ORDEM As String = "P|SiO2|PPC|325|SE"
' Senha de protecao das abas (vazio = sem senha)
Public Const SENHA As String = ""

' campos da aba Batch gravados em tbl_Batches (mesma ordem das 1as colunas)
Public Function CAMPOS_BATCH() As String
    CAMPOS_BATCH = "bId|bData|bMin|bIniU|bFimU|bNivel|bIniPMan|bIniG|bFimG|bNHor|bLabA|bLabB|bRef|bTit|bSub|bTol_SE|bTol_325|bTol_SiO2|bTol_P|bTol_PPC|bIncFrase|bIncAcomp|bFrase|bAcomp|bCom1|bCom2|bCom3|bCom4|bCom5|bCom6|bPfData|bPfFe|bPfSiO2|bPfP|bPfSE"
End Function

' rotulos dos parametros (com caracteres fora do cp1252 via ChrW)
Public Function P_ROTULOS() As String
    P_ROTULOS = "S.E.|<325#|SiO" & ChrW(8322) & "|P|PPC"
End Function

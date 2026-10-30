Option Explicit

' Gerado por build_monitor.py - posicoes fixas das abas (nao editar a mao)
Public Const NI As Long = 12
Public Const NMAXD As Long = 60
' aba Dia (1-based)
Public Const D_ROW_IT0 As Long = 11
Public Const D_COL_ITEM As Long = 2
Public Const D_COL_VAL As Long = 3
Public Const D_COL_NRES As Long = 4
Public Const D_COL_JAN As Long = 5
' aba Diagnostico (1-based)
Public Const X_ROW0 As Long = 9
Public Const X_COL_TAG As Long = 3
Public Const X_COL_MAPA As Long = 4
Public Const X_COL_RES As Long = 5
Public Const X_COL_VAL As Long = 6
Public Const X_COL_LISTA As Long = 7
' aba Configuracoes (1-based)
Public Const C_ROW_IT0 As Long = 8
Public Const C_COL_CHAVE As Long = 2
Public Const C_COL_ROT As Long = 3
Public Const C_COL_TAG As Long = 4
Public Const C_COL_SERV As Long = 5
Public Const C_COL_MAPA As Long = 6
Public Const C_COL_JAN As Long = 7
Public Const C_COL_AGR As Long = 8
Public Const C_COL_VMIN As Long = 9
Public Const C_COL_VMAX As Long = 10
Public Const C_COL_DEC As Long = 11
Public Const C_COL_ORIG As Long = 12
' aba Historico (1-based)
Public Const H_ROW_HDR As Long = 6
Public Const H_NLIN As Long = 3000
Public Const H_COL_DATA As Long = 2
Public Const H_COL_PRODC3 As Long = 3
Public Const H_COL_PRODC2 As Long = 4
Public Const H_COL_ESPC3 As Long = 5
Public Const H_COL_ESPC2 As Long = 6
Public Const H_COL_ESPP As Long = 7
Public Const H_COL_EB4 As Long = 8
Public Const H_COL_EB4P As Long = 9
Public Const H_COL_EB6 As Long = 10
Public Const H_COL_DIFJ As Long = 11
Public Const H_COL_DIFK As Long = 12
Public Const H_COL_TK01 As Long = 13
Public Const H_COL_TK02 As Long = 14
Public Const H_COL_FATOR As Long = 15
Public Const H_COL_LAMAC3 As Long = 16
Public Const H_COL_LAMAC2 As Long = 17
Public Const H_COL_LAMA As Long = 18
Public Const H_COL_CONCC3 As Long = 19
Public Const H_COL_CONCC2 As Long = 20
Public Const H_COL_PCTLAMA As Long = 21
Public Const H_COL_ATUAL As Long = 22
Public Const H_COL_OBS As Long = 23
Public Const H_ENTRADAS As String = "data|prodC3|prodC2|espC3|espC2|eb4|eb6|tk01|tk02|fator|lamaC3|lamaC2|concC3|concC2|atual|obs"
' itens do MES (ordem da aba Dia e de Configuracoes)
Public Const I_CHAVES As String = "prodC3|prodC2|espC3|espC2|eb4|eb6|tk01|tk02|lamaC3|lamaC2|concC3|concC2"
' colunas da aba _Graf (1-based, nesta ordem); linha 1 = titulos, dia k na linha k + 1
Public Const G_CHAVES As String = "data|rot|espP|eb6|difK|difKmed|tolKP|tolKN|eb4P|difJ|difJmed|tolJP|tolJN|tk01|tk02|pctLama|pctMed|pct0"
' series de cada grafico do Painel (barras primeiro) e quantas sao barras
Public Const GRAF_SERIES As String = "espP,eb6,difK,difKmed,tolKP,tolKN|eb4P,eb6,difJ,difJmed,tolJP,tolJN|eb6,tk01,tk02,pctLama,pctMed"
Public Const GRAF_NBARRAS As String = "2|2|3"
Public Const COR_BASE As String = "15980963|14670805|13353913"
Public Const COR_MEDIA As String = "7356416|10192993|9997437"
Public Const SENHA_CFG As String = "Samarco@Lab"
Public Const SENHA As String = ""


Option Explicit

' Gerado por build_calibracao.py - posicoes fixas (nao editar a mao)
Public Const NSTG As Long = 67
Public Const N_ENSAIOS As Long = 9
Public Const PAI_ENS1 As Long = 14
Public Const PAI_EQ1 As Long = 79
Public Const PAI_NEQ As Long = 40
Public Const PAI_PEN1 As Long = 122
Public Const PAI_NPEN As Long = 4
Public Const PAI_NC1 As Long = 129
Public Const PAI_NNC As Long = 25
Public Const GRA_NBLOCOS As Long = 5
Public Const GRA_LIN As Long = 40
Public Const GRA_NPT As Long = 30
' Coluna da 1a celula de cada tabela do Painel
Public Const PAI_COL1 As Long = 2
Public Const SENHA As String = ""

Public Function NomeEnsaio(ByVal i As Long) As String
    Select Case i
        Case 1: NomeEnsaio = "Blaine"
        Case 2: NomeEnsaio = "Tambor de Abrasão"
        Case 3: NomeEnsaio = "Gran. Fina Alpine"
        Case 4: NomeEnsaio = "Umidade"
        Case 5: NomeEnsaio = "Compressão"
        Case 6: NomeEnsaio = "Granulometria"
        Case 7: NomeEnsaio = "Tamb 5 kg x 15 kg"
        Case 8: NomeEnsaio = "Verificação Peneiradores"
        Case 9: NomeEnsaio = "Comparativo Fisher"
    End Select
End Function

' Criterio de cada linha da tabela 'Situacao por equipamento' (faixa:CHAVE, tol:CHAVE ou info)
Public Function CriterioEquip(ByVal i As Long) As String
    Select Case i
        Case 1: CriterioEquip = "faixa:BLAM"
        Case 2: CriterioEquip = "faixa:BLAS"
        Case 3: CriterioEquip = "faixa:BLAA"
        Case 4: CriterioEquip = "faixa:TAMB"
        Case 5: CriterioEquip = "faixa:TAMB"
        Case 6: CriterioEquip = "faixa:TAMB"
        Case 7: CriterioEquip = "faixa:TAMB"
        Case 8: CriterioEquip = "faixa:TAMB"
        Case 9: CriterioEquip = "faixa:ALP"
        Case 10: CriterioEquip = "faixa:ALP"
        Case 11: CriterioEquip = "faixa:ALP"
        Case 12: CriterioEquip = "tol:UMI"
        Case 13: CriterioEquip = "tol:UMI"
        Case 14: CriterioEquip = "tol:UMI"
        Case 15: CriterioEquip = "tol:COM16"
        Case 16: CriterioEquip = "faixa:VEL"
        Case 17: CriterioEquip = "tol:COM16"
        Case 18: CriterioEquip = "faixa:VEL"
        Case 19: CriterioEquip = "tol:COM16"
        Case 20: CriterioEquip = "faixa:VEL"
        Case 21: CriterioEquip = "tol:COM16"
        Case 22: CriterioEquip = "tol:COM16"
        Case 23: CriterioEquip = "faixa:VEL"
        Case 24: CriterioEquip = "tol:COM12"
        Case 25: CriterioEquip = "faixa:VEL"
        Case 26: CriterioEquip = "tol:COM12"
        Case 27: CriterioEquip = "faixa:VEL"
        Case 28: CriterioEquip = "tol:COM12"
        Case 29: CriterioEquip = "faixa:VEL"
        Case 30: CriterioEquip = "tol:COM12"
        Case 31: CriterioEquip = "tol:COM12"
        Case 32: CriterioEquip = "faixa:VEL"
        Case 33: CriterioEquip = "tol:G63"
        Case 34: CriterioEquip = "tol:GRG"
        Case 35: CriterioEquip = "info"
        Case 36: CriterioEquip = "info"
        Case 37: CriterioEquip = "info"
        Case 38: CriterioEquip = "tol:T515"
        Case 39: CriterioEquip = "info"
        Case 40: CriterioEquip = "tol:FIS"
    End Select
End Function
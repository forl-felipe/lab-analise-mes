Option Explicit

Private Sub Workbook_Open()
    Inicializar
End Sub

' Ctrl+V da planilha (ModBatch.ColarInteligente) so enquanto este arquivo esta ativo
Private Sub Workbook_Activate()
    LigarColar True
End Sub

Private Sub Workbook_Deactivate()
    LigarColar False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    LigarColar False
End Sub

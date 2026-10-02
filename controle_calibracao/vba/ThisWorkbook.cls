Option Explicit

Private Sub Workbook_Open()
    Inicializar
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If Success Then BackupAposSalvar
End Sub

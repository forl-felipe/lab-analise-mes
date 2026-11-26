Attribute VB_Name = "ModToleranciaSE"
Option Explicit

' Atualizacao da planilha do comparativo (Germano x Ubu) ja em uso:
' tolerancia da Superficie Especifica (S.E.) por faixa, definida pela media de Germano e Ubu do batch:
'     ate 2000 cm2/g        -> 150
'     de 2000 a 2200 cm2/g  -> 200
'     acima de 2200 cm2/g   -> 250
' Rode a macro AjustarToleranciaSE uma vez (F5), salve a planilha e depois pode remover este modulo.

Public Sub AjustarToleranciaSE()
    Const SENHA As String = "fislab"
    Dim cel As Range, wsCfg As Worksheet, wsGraf As Worksheet, wsBat As Worksheet
    Dim linSE As Long, r As Long, med As String, f As String

    On Error GoTo Falha
    Set cel = ThisWorkbook.Names("padTol_SE").RefersToRange
    Set wsCfg = cel.Worksheet
    Set wsGraf = ThisWorkbook.Worksheets("_Graf")

    ' linha da S.E. na tabela de calculo do farol (aba oculta _Graf)
    For r = 1 To 80
        If Trim$(CStr(wsGraf.Cells(r, 1).Value)) = "S.E." And wsGraf.Cells(r, 2).HasFormula Then
            linSE = r
            Exit For
        End If
    Next r
    If linSE = 0 Then
        MsgBox "Nao encontrei a linha da S.E. na aba _Graf. Nada foi alterado.", vbExclamation
        Exit Sub
    End If

    med = "_Graf!$B$" & linSE & ":$C$" & linSE          ' media de Germano e media de Ubu do batch
    f = "=IF(COUNT(" & med & ")=0,150,IF(AVERAGE(" & med & ")<=2000,150," & _
        "IF(AVERAGE(" & med & ")<=2200,200,250)))"

    wsCfg.Unprotect SENHA
    cel.Formula = f
    cel.NumberFormat = "General"
    cel.Offset(0, -1).Value = "Toler" & ChrW(226) & "ncia S.E. (" & ChrW(177) & ") por faixa"
    cel.Offset(0, 4).Value = "at" & ChrW(233) & " 2000 " & ChrW(8594) & " 150 " & ChrW(183) & " 2000 a 2200 " & _
        ChrW(8594) & " 200 " & ChrW(183) & " acima de 2200 " & ChrW(8594) & " 250 (pela m" & ChrW(233) & "dia Germano/Ubu)"
    wsCfg.Protect Password:=SENHA, DrawingObjects:=True, Contents:=True, AllowFormattingCells:=True, _
        AllowFormattingColumns:=True, AllowFormattingRows:=True

    ' a tolerancia do batch em andamento passa a seguir a da faixa
    Set wsBat = ThisWorkbook.Names("bTol_SE").RefersToRange.Worksheet
    wsBat.Unprotect SENHA
    ThisWorkbook.Names("bTol_SE").RefersToRange.Formula = "=padTol_SE"
    wsBat.Protect Password:=SENHA, DrawingObjects:=True, Contents:=True, AllowFormattingCells:=True, _
        AllowFormattingColumns:=True, AllowFormattingRows:=True

    Application.CalculateFull
    MsgBox "Toler" & ChrW(226) & "ncia da S.E. agora " & ChrW(233) & " por faixa:" & vbCrLf & vbCrLf & _
           "   at" & ChrW(233) & " 2000 cm" & ChrW(178) & "/g: 150" & vbCrLf & _
           "   de 2000 a 2200 cm" & ChrW(178) & "/g: 200" & vbCrLf & _
           "   acima de 2200 cm" & ChrW(178) & "/g: 250" & vbCrLf & vbCrLf & _
           "Toler" & ChrW(226) & "ncia do batch atual: " & ThisWorkbook.Names("bTol_SE").RefersToRange.Value & vbCrLf & _
           "Salve a planilha (Ctrl+S). Depois pode remover este m" & ChrW(243) & "dulo.", vbInformation
    Exit Sub
Falha:
    MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel aplicar: " & Err.Description, vbExclamation
End Sub

' Eventos ao fechar a planilha
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    SalvarAntesdeFechar
End Sub



' Funcao Salvar antes de fechar
Public Sub SalvarAntesdeFechar()

     Dim resultado As VbMsgBoxResult
     resultado = MsgBox("Deseja Salvar as Alterações Feitas?", vbYesNo + vbQuestion, "Salvar")

     If resultado = vbYes Then
        ThisWorkbook.Save
     End If
End Sub



' Atualizar preços ao abrir a planilha
Private Sub Workbook_Open()
    Call historicoBTC_BRL("1d")
End Sub


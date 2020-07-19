Attribute VB_Name = "Módulo1"
'Esse primeiro projeto tem a disponibilidade de retirar o espaço desnecessario e deixar o texto mais organizado.
Sub ConvEspac()
Application.ScreenUpdating = False
Do While ActiveCell <> ""
If ActiveCell <> "" Then
       numero = Trim(Application.WorksheetFunction.Trim(ActiveCell))
      ActiveCell = numero
      ActiveCell.Offset(1, 0).Select
Else
      ActiveCell.Offset(1, 0).Select
End If
Loop
Application.ScreenUpdating = True
End Sub
'Este projeto tem a disponibilidade para converter as linha de número em texto para o formato de de número.
Sub ConvTXTValor()
Application.ScreenUpdating = False
Do While ActiveCell.Offset(0, -(ActiveCell.Column - 1)) <> ""
If ActiveCell <> "" Then
       numero = Str(ActiveCell.Value)
      ActiveCell.FormulaR1C1 = numero
      ActiveCell.Offset(1, 0).Select
Else
      ActiveCell.Offset(1, 0).Select
End If
Loop
Application.ScreenUpdating = True
End Sub
'Já esse projeto tem a disponibilidade e converter a celula de data em data no formato original (ex. 01/200 para 01/01/2000).
Sub ConvData()
Application.ScreenUpdating = False
Do While ActiveCell <> ""
If ActiveCell <> "" Then
       numero = DateAdd("d", 0, ActiveCell)
      ActiveCell = numero
      ActiveCell.Offset(1, 0).Select
Else
      ActiveCell.Offset(1, 0).Select
End If
Loop
Application.ScreenUpdating = True
End Sub
'Já esse último projeto tem a disponibilidade de percorrer todas as guia da planilha para utilizar as funcições assima para sua cada particularidade.
Sub localiza()
Dim W As Worksheet
For Each W In ThisWorkbook.Worksheets
W.Select
W.Range("a2").Select
Módulo1.ConvEspac
W.Range("b2").Select
Módulo1.ConvData
W.Range("c2").Select
Módulo1.ConvTXTValor
Next W
End Sub

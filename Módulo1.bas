Attribute VB_Name = "Módulo1"
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
Sub localiza()
Dim W As Worksheet
For Each W In ThisWorkbook.Worksheets
'W.Select
H = A01.Cells(Rows.Count, 1).End(xlUp).Row + 1
A01.Cells(H, 1) = W.Range("A2")
A01.Cells(H, 2) = W.Name
        A01.Cells(H, 2).Hyperlinks.Add Anchor:=A01.Cells(H, 2), _
        Address:="", _
        SubAddress:="'" & W.Name & "'!A2", _
        TextToDisplay:=A01.Cells(H, 2).Value
        

'W.Range("a2").Select
'Módulo1.ConvEspac
'W.Range("b2").Select
'Módulo1.ConvData
'W.Range("c2").Select
'Módulo1.ConvTXTValor
Next W
End Sub

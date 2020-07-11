Sub Metodo2()
'Definimos variables
Dim i%, Fin%
With Sheets("resumen")
Fin = Application.CountA(.Range("I:I")) ' RANGO DE I 
'Mediante un bucle indicamos que si una celda está vacía
'el valor sea el de la celda anterior.
For i = 2 To Fin
If .Cells(i, 1) = "" Then .Cells(i, 1) = .Cells(i - 1, 1)
Next
End With
End Sub

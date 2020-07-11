Private Sub CommandButton2_Click()
Dim fila As Integer
Dim columna As Integer

        


'Columna cantidades

        columna = 13
        For fila = 14 To 200 Step 1
            If Worksheets("Cotizador").Cells(fila, columna).Value > 0 Then
    'buscar espacio en la hoja exportar
            filafinal = Worksheets("Exportar").Cells(Rows.Count, 4).End(xlUp).Row + 1
    ' copiamos la columna que es mayor a cero
            Worksheets("Cotizador").Cells(fila, columna).Copy
    'pegamos en la hoja exportar
            Worksheets("Exportar").Cells(filafinal, 4).PasteSpecial Paste:=xlPasteValues
            End If
            Next fila


'Columna EAN
  
        columna = 13
             For fila = 14 To 200 Step 1
       'buscar espacio en la hoja exportar
            filafinal = Worksheets("Exportar").Cells(Rows.Count, 2).End(xlUp).Row + 1
             If Worksheets("Cotizador").Cells(fila, columna).Value > 0 Then
                Worksheets("Cotizador").Cells(fila, columna - 9).Copy
                Worksheets("Exportar").Cells(filafinal, 2).PasteSpecial Paste:=xlPasteValues
            End If
            Next fila


'Columna precio
  
            columna = 13
            For fila = 14 To 200 Step 1
        'buscar espacio en la hoja exportar
            filafinal = Worksheets("Exportar").Cells(Rows.Count, 5).End(xlUp).Row + 1
            If Worksheets("Cotizador").Cells(fila, columna).Value > 0 Then
                Worksheets("Cotizador").Cells(fila, columna - 7).Copy
                Worksheets("Exportar").Cells(filafinal, 5).PasteSpecial Paste:=xlPasteValues
            End If
           Next fila


'Crear Carrito

    filacotiza = Worksheets("Cotizador").Cells(Rows.Count, 14).End(xlUp).Row
            
            For x = 14 To filacotiza
            valor = Worksheets("cotizador").Cells(x, 14).Value

                If valor > 0 Then
                Worksheets("cotizador").Cells(x, 1).Resize(2, 14).Copy
                carritofila = Worksheets("Carrito").Cells(Rows.Count, 1).End(xlUp).Row + 1
                Worksheets("Carrito").Cells(carritofila, 1).PasteSpecial Paste:=xlPasteValues
              End If
              Next x



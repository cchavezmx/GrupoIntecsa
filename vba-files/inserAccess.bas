Private Sub CommandButton1_Click()
    'funcion a llamar principal
    'verificamos si existe la hoja resumen
    
    Dim existe As Boolean
    Dim msg As String
    Dim title As String
    Dim respuesta As VbMsgBoxResult
    
        On Error Resume Next
        existe = (Worksheets("resumen").Name <> "")
            
        If Not existe Then
        
        Call crearDB
    
        ElseIf existe = True Then
        
        
        msg = "¿Esta acción actualizara el requerimiento, quieres continuar?"
        title = "Actulizacion terminada"
        respuesta = msgbox(msg, vbYesNo, title)
      
       
       
      If respuesta = vbYes Then
        
        Call crearDB
        msgbox "Actualizacion de requerimientos exitoso"
        
        Else
        
       
               
        End If
        End If
    
    End Sub
              
    Private Sub relleno()
    'Definimos variables
    Dim i%, Fin%
    With Sheets("resumen")
    Fin = Application.CountA(.Range("A:A"))
    'Mediante un bucle indicamos que si una celda está vacía
    'el valor sea el de la celda anterior.
    For i = 3 To Fin
    If .Cells(i, 9) = "" Then .Cells(i, 9) = .Cells(i - 1, 8)
    If .Cells(i, 10) = "" Then .Cells(i, 10) = .Cells(i - 1, 9)
    If .Cells(i, 11) = "" Then .Cells(i, 11) = .Cells(i - 1, 10)
    Next
    End With
    
    End Sub
    
    Sub crearDB()
    
        Worksheets.Add.Name = "resumen"
        Worksheets("resumen").Range("A1").FormulaR1C1 = "cod"
        Worksheets("resumen").Range("B1").FormulaR1C1 = "partida"
        Worksheets("resumen").Range("C1").FormulaR1C1 = "item"
        Worksheets("resumen").Range("D1").FormulaR1C1 = "codigo"
        Worksheets("resumen").Range("E1").FormulaR1C1 = "concepto"
        Worksheets("resumen").Range("F1").FormulaR1C1 = "unidad"
        Worksheets("resumen").Range("G1").FormulaR1C1 = "cantidad"
        Worksheets("resumen").Range("H1").FormulaR1C1 = "control"
        Worksheets("resumen").Range("I1").FormulaR1C1 = "proyecto"
        Worksheets("resumen").Range("J1").FormulaR1C1 = "tablero"
        Worksheets("resumen").Range("k1").FormulaR1C1 = "fecha"
        
        
     'Buscamos hojas con en encabezado requerimiento y copiamos la celdas a la hoja de resumen
     
    nhojas = Sheets.Count
    Dim control As String
    Dim tablero As String
    Dim proyecto As String
    Dim rangotab As Range
    
        For x = 1 To nhojas Step 1
      
        encabezado = Worksheets(x).Range("B8").Value ' buscar encabezado
        control = Worksheets(x).Range("I5").Value    ' busca valor de control
        tablero = Worksheets(x).Range("I6").Value    ' busca valor de tablero
        proyecto = Worksheets(x).Range("c4").Value   ' busca valor de proyecto
        
                
        If encabezado = "REQUERIMIENTO DE MATERIAL" Then
        
        espacioResumen = Worksheets("resumen").Cells(Rows.Count, 1).End(xlUp).Row + 1 ' la ultima celda para escribir en resumen
        filafinal = Worksheets(x).Cells(Rows.Count, 7).End(xlUp).Row ' Encontrar el tamaño de la lista
        
        Worksheets(x).Range("B11:B" & filafinal).Copy ' Copiamos partida
        Worksheets("resumen").Cells(espacioResumen, 2).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
            Worksheets(x).Range("C11:C" & filafinal).Copy ' Copiamos partida
        Worksheets("resumen").Cells(espacioResumen, 3).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
            Worksheets(x).Range("D11:D" & filafinal).Copy ' Copiamos partida
        Worksheets("resumen").Cells(espacioResumen, 4).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
            Worksheets(x).Range("E11:E" & filafinal).Copy ' Copiamos partida
        Worksheets("resumen").Cells(espacioResumen, 5).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
            Worksheets(x).Range("H11:H" & filafinal).Copy ' Copiamos partida
        Worksheets("resumen").Cells(espacioResumen, 6).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
            Worksheets(x).Range("I11:I" & filafinal).Copy ' Copiamos partida
        Worksheets("resumen").Cells(espacioResumen, 7).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
            
        Worksheets("resumen").Cells(espacioResumen, 8).FormulaR1C1 = control
        Worksheets("resumen").Cells(espacioResumen, 9).FormulaR1C1 = tablero
        Worksheets("resumen").Cells(espacioResumen, 10).FormulaR1C1 = proyecto
        
        
         End If
         Next x
    
    
        Call relleno ' rellenar
        Call datostoDb ' mandamos datos a la base de datos
       
       
       'FILTRA CELDAS VACIAS
       
        'Worksheets("resumen").Range("A1:L1").Select
        'Selection.AutoFilter
        'ActiveSheet.Range(Selection, Selection.End(xlDown)).AutoFilter Field:=3, Criteria1:="<>"
        'Close
        
    End Sub
    
    
    Private Sub datostoDb()
    
    'declaramos las variables para la consulta con la hoja
    
    Dim cn As Object
    Dim datos As Object
    Dim consultaDB As String
    Dim conectar As String
    Dim identificacion As String
    Dim final As String
    
    
    ' Le damos valor a cn para la coneccion
    
      Set cn = CreateObject("ADODB.connection")
          conexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=C:\Users\Saul\Documents\html5\invazoronia\vba-files\almacen.accdb"
    
    
    ' vamos a buscar algun si el codigo unico existe
    
    
        With Sheets("resumen")
        final = Application.CountA(.Range("A:A"))
                  
         For i = 2 To final
                
        identificacion = .Range("A" & i)
        consultaDB = "Select * from requerimientos where cod = " & Chr(34) & identificacion & Chr(34)
        
        
        cn.Open conexion ' abre la conexion
        
        Set datos = cn.Execute(consultaDB) ' le asigna el valor de la conexion ejecuntado la consulta en el string de consultaDB
               
        If Not datos.EOF Then  ' si datos regresa con el valor TRUE ahora la accion despues del then
                                'msgbox "El codigo ya existe", vbCritical, "consulta persona"
           
        Else
        
        'msgbox "its Works" ' que si no, ejecuta lo que sigue despues del else
    
        Call insertaREG
        'msgbox "Se ha actualizado el requerimiento", vbOKOnly + vbInformation
        
        End If
        
        cn.Close   ' cerramos la conexcion
        Set datos = Nothing  ' borramos el contenido de datos
           
        Next ' e iniciamos el ciclo for
        
        End With
        
    End Sub
    
    
    Sub insertaREG()
    
    'declaramos las variables para la consulta con la hoja
    
    Dim cn As Object
    Dim datos As Object
    Dim consultaDB As String
    Dim conexion As String
    Dim identificacion As String
    Dim final As String
    
    
    ' Le damos valor a cn para la coneccion
    
      Set cn = CreateObject("ADODB.connection")
          conexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=C:\Users\Saul\Documents\html5\invazoronia\vba-files\almacen.accdb"
    
    
    With Sheets("resumen") ' activamos la hoja para trabajar con ella
    final = Application.CountA(.Range("A:A")) ' calculamos el tamaño de las filas en las que va a trabajar.
    
        For i = 2 To final
        
        On Error Resume Next
        
      ' declaramos cada variable de que componenen la consulta de la hoja.
     
     
        Dim cod As String
        Dim partida As String
        Dim item As String
        Dim codigo As String
        Dim concepto As String
        Dim unidad As String
        Dim cantidad As String
        Dim control As String
        Dim proyecto As String
        Dim tablero As String
        Dim fecha As Date
        
     ' les asginamos un valor.
    
        cod = Range("A" & i)
        partida = Range("B" & i)
        item = Range("C" & i)
        codigo = Range("D" & i)
        concepto = Range("E" & i)
        unidad = Range("F" & i)
        cantidad = Range("G" & i)
        control = Range("H" & i)
        proyecto = Range("I" & i)
        tablero = Range("J" & i)
        fecha = Range("K" & i)
        
        
        ' creamos la conexcion
    
        consultaSql = "insert into requerimientos values(" & Chr(34) & cod & Chr(34) & "," & Chr(34) & partida & Chr(34) & "," & Chr(34) & item & Chr(34) & "," & Chr(34) & codigo & Chr(34) & "," & Chr(34) & concepto & Chr(34) & "," & Chr(34) & unidad & Chr(34) & "," & Chr(34) & cantidad & Chr(34) & "," & Chr(34) & control & Chr(34) & ", " & Chr(34) & proyecto & Chr(34) & ", " & Chr(34) & tablero & Chr(34) & ", " & Chr(34) & fecha & Chr(34) & ")"
        
        'abrimos la conexcion
        
        cn.Open conexion
        
        'pasamos el string para que se ejcute y despues contolar con la variable datos
        
        Set datos = cn.Execute(consultaSql)
        cn.Close   ' cerramos la conexcion
        Set datos = Nothing  ' borramos el contenido de datos
       
        
        Next ' e iniciamos el ciclo for
        
        End With
        
      
        
    
    End Sub
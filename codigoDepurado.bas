'MACRO PARA COMPARAR DOS COLUMNAS DE DATOS EN EXCXEL 

Sub buscador()

Dim filaOrigen As Integer
Dim filaBuscar As Integer

Dim col_origen As Integer
Dim col_destino As Integer

' definimos la columna que queremos comprar
    
    col_origen = Application.InputBox(prompt:="Seleccione la columna de la hoja Origen", Type:=1)
    col_destino = Application.InputBox(prompt:="Seleccione la columna de la hoja Destino", Type:=1)
        
'Buscar tamaño de las filas

    filaOrigen = Worksheets("Origen").Cells(Rows.Count, col_origen).End(xlUp).Row
    filaBuscar = Worksheets("Destino").Cells(Rows.Count, col_destino).End(xlUp).Row
    
' Pasamos el ciclo for
    
    For x = 2 To filaOrigen Step 1
    
    ' setamos el valor a buscar
    Set curCell = Worksheets("Origen").Cells(x, col_origen)
    
    
    For i = 2 To filaBuscar Step 1
    Set seekCell = Worksheets("Destino").Cells(i, col_destino)
    
    If curCell.Value = seekCell.Value Then
    Worksheets("Destino").Cells(i, col_destino).Interior.Color = 65535
    
    'revisa en cada iteracion si ya termino la lista para liberar el for y terminar la macro
    If i = filaBuscar Then Exit For
    
    End If
    
    Next i
    Next x
    
End Sub


'MODULO DE NUEVOS ELEMENTOS A LA BASE DE DATOS.
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub btn_agregar_Click()
Dim sql As String
Dim Connected As Boolean
Dim proyectos As String, lugar As String, cliente As String, fecha As Date, sn As String, tablero As String



With Datos_alta

 proyecto = Me.txt_proyecto.Value
 lugar = Me.txt_lugar.Value
 cliente = Me.txt_cliente.Value
 sn = Me.txt_sn.Value
 tablero = Me.txt_tablero.Value
 req = Me.txt_req.Value



    sql = "INSERT INTO dbo.proyectos (nserie, proyecto, lugar, residente, fecha, tablero, req) VALUES ('" & sn & "', '" & proyecto & "',  '" & lugar & "', '" & cliente & "', CURRENT_TIMESTAMP, '" & tablero & "','" & req & "');"
    'Connected = SqlConnect("localhost", "master")
    'Connected = SqlConnect("ADMINPAQ-SERVER\COMPAC", "sa", "mOON020106", "almacenNB")
     
     Connected = SqlConnect("ADMINPAQ-SERVER\COMPAC", "sa", "mOON020106", "almacenNB")
    
    
End With

    If Connected Then
    
        Call Query(sql)
        Call Disconnect
        
        
       ActiveSheet.Range("C4") = proyecto
       Me.txt_proyecto.Value = ""
       ActiveSheet.Range("C5") = lugar
       Me.txt_lugar.Value = ""
       ActiveSheet.Range("C6") = cliente
       Me.txt_cliente.Value = ""
       ActiveSheet.Range("I5") = sn
       Me.txt_sn.Value = ""
       ActiveSheet.Range("I6") = tablero
       Me.txt_tablero.Value = ""
       ActiveSheet.Range("I7") = req
       Me.txt_req.Value = ""
       'ActiveSheet.Range("C4") = proyecto
       
       
        
    Else
        MsgBox "No se puedo realizar la conexión"
    End If

End Sub

'MODULO DE NUEVOS ELEMENTOS A LA BASE DE DATOS.
'//////////////////////////////////////////////----FINAL----///////////////////////////////////////////////////////

'MODULO DE BUSQUEDA DE LOS ELEMENTOS A LA BASE DE DATOS.
'/////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub btn_find_Click()

Dim SQLB As String
Dim RSB As ADODB.Recordset
Dim i, j
Dim CNB As ADODB.Connection
Dim toFind  As String


Set CNB = New ADODB.Connection
    'On Error Resume Next
    ' En caso de error, lo saltamos con esta linea.
        
    CNB.Open "Provider=SQLOLEDB.1;Persist security info=True; User ID=sa;Password =mOON020106; Initial Catalog =almacenNB;Data Source=ADMINPAQ-SERVER\COMPAC;"
    
    
toFind = Me.txt_find.Value


SQLB = "SELECT nserie, proyecto, lugar FROM proyectos WHERE nserie LIKE '%" & toFind & "%';"
 'SQLB = "SELECT nserie, proyecto, lugar, residente, fecha, tablero, req FROM proyectos"
 
Set RSB = New ADODB.Recordset

'Recuerda, una vez hecha la query, creada la coneccion y el nuevo objeto recordset, hacemos esta linea, esta ejecuta todo... dije todoooodo
RSB.Open SQLB, CNB

With ListBox1
    .ColumnCount = RSB.Fields.Count
End With

RSB.MoveFirst
i = 1

    With Me.ListBox1
        .Clear
        .AddItem
        
        For j = 0 To 2
         .List(0, j) = RSB.Fields(j).Name
         Next j
    Do
        .AddItem
        .List(i, 0) = RSB![nserie]
        .List(i, 1) = RSB![proyecto]
        .List(i, 2) = RSB![lugar]
        '.List(i, 3) = RSB![residente]
        '.List(i, 4) = RSB![fecha]
        '.List(i, 5) = RSB![tablero]
        '.List(i, 6) = RSB![req]
     i = i + 1
RSB.MoveNext

Loop Until RSB.EOF
End With

RSB.Close
CNB.Close
Set RSB = Nothing
Set CNB = Nothing

End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim columna As Long
Dim sql As String
Dim abuscar As String


abuscar = Datos_busqueda.ListBox1.Value

Set cn = New ADODB.Connection
'se crea el objeto para la conexión
cn.Open "Provider=SQLOLEDB.1;Persist security info=True; User ID=sa;Password =mOON020106; Initial Catalog =almacenNB;Data Source=ADMINPAQ-SERVER\COMPAC;"
'se añaden los parametros para el objeto cn


Set rs = New ADODB.Recordset
'creamos el objetos para la guardar la consulta

sql = "Select * from proyectos where nserie ='" & abuscar & "';"
' creamos la consulta


rs.Open sql, cn
' mandamos la consulta(sql) al objeto rs (recordset), con la conexion establecida en CN

If rs.State Then
        
  With ActiveSheet
 'mandamos los datos de la cabezera del rs a la hoja electronica
 'usamos el with para no escribir tantas lineas
 
 
    .Range("I5") = rs![nserie]
    .Range("C4") = rs![proyecto]
    .Range("C5") = rs![lugar]
    .Range("C6") = rs![residente]
    .Range("I4") = rs![fecha]
    .Range("I6") = rs![tablero]
    .Range("I7") = rs![req]

          
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

 End With
 

End If
End Sub
Private Sub txt_nuevo_Click()
Datos_alta.Show
End Sub
Private Sub UserForm_open()

Dim SQLB As String
Dim RSB As ADODB.Recordset
Dim i, j
Dim CNB As ADODB.Connection
Dim toFind  As String


Set CNB = New ADODB.Connection
    'On Error Resume Next
    ' En caso de error, lo saltamos con esta linea.
        
    CNB.Open "Provider=SQLOLEDB.1;Persist security info=True; User ID=sa;Password =mOON020106; Initial Catalog =almacenNB;Data Source=ADMINPAQ-SERVER\COMPAC;"
    
    
toFind = Me.txt_find.Value


'SQLB = "SELECT * FROM proyectos WHERE nserie LIKE '%" & toFind & "%';"
 'SQLB = "SELECT nserie, proyecto, lugar, residente, fecha, tablero, req FROM proyectos"
 SQLB = "SELECT nserie, proyecto, lugar FROM proyectos"
 
Set RSB = New ADODB.Recordset

'Recuerda, una vez hecha la query, creada la coneccion y el nuevo objeto recordset, hacemos esta linea, esta ejecuta todo... dije todoooodo
RSB.Open SQLB, CNB

With ListBox1
    .ColumnCount = RSB.Fields.Count
End With

RSB.MoveFirst
i = 1

    With Me.ListBox1
        .Clear
        .AddItem
        
        For j = 0 To 2
         .List(0, j) = RSB.Fields(j).Name
         Next j
    Do
        .AddItem
        .List(i, 0) = RSB![nserie]
        .List(i, 1) = RSB![proyecto]
        .List(i, 2) = RSB![lugar]
        .List(i, 3) = RSB![residente]
        '.List(i, 4) = RSB![fecha]
        '.List(i, 5) = RSB![tablero]
        '.List(i, 6) = RSB![req]
     i = i + 1
RSB.MoveNext

Loop Until RSB.EOF
End With

RSB.Close
CNB.Close
Set RSB = Nothing
Set CNB = Nothing

End Sub

Private Sub UserForm_Click()

End Sub


'MODULO DE BUSCADOR ELEMENTOS A LA BASE DE DATOS.
'//////////////////////////////////////////////----FINAL----///////////////////////////////////////////////////////


'MODULO DE CONSTRUYE LA BASE DE DATOS Y CARGA LOS DATOS DE LA HOJA RESUMEN A LA BASE DE DATOS.
'/////////////////////////////////////////////////////////////////////////////////////////////////////


Private Sub CommandButton1_Click()
Datos_busqueda.Show
End Sub

Private Sub ok_reg_Click()
'funcion a llamar principal
'verificamos si existe la hoja resumen

Dim existe As Boolean
Dim msg As String
Dim title As String
Dim respuesta As VbMsgBoxResult

Application.ScreenUpdating = False


    'On Error Resume Next
    existe = (Worksheets("resumen").Name <> "")
        
    If Not existe Then
  
    
    Call crearDB
    MsgBox "Se genero un requerimiento de materiales"
    

    ElseIf existe = True Then
    
    
    msg = "¿Esta acción actualizara el requerimiento, quieres continuar?"
    title = "Actulizacion terminada"
    respuesta = MsgBox(msg, vbYesNo, title)
  
   
   
    If respuesta = vbYes Then
    
    
    Call BORRAS
    Call crearDB
    
    MsgBox "Actualizacion de requerimientos exitoso"
    
    Else
    
   
           
    End If
    End If

End Sub
          
Private Sub relleno()
'Definimos variables
Dim i%, Fin%
With Sheets("resumen")
Fin = Application.CountA(.Range("B:B"))
'Mediante un bucle indicamos que si una celda está vacía
'el valor sea el de la celda anterior.
For i = 3 To Fin
If .Cells(i, 8) = "" Then .Cells(i, 8) = .Cells(i - 1, 8)
If .Cells(i, 9) = "" Then .Cells(i, 9) = .Cells(i - 1, 9)
If .Cells(i, 10) = "" Then .Cells(i, 10) = .Cells(i - 1, 10)
If .Cells(i, 11) = "" Then .Cells(i, 11) = .Cells(i - 1, 11)
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
    Worksheets("resumen").Range("H1").FormulaR1C1 = "ns"
    Worksheets("resumen").Range("I1").FormulaR1C1 = "proyecto"
    Worksheets("resumen").Range("J1").FormulaR1C1 = "tablero"
    Worksheets("resumen").Range("k1").FormulaR1C1 = "fecha"
    
    
 'Buscamos hojas con en encabezado requerimiento y copiamos la celdas a la hoja de resumen
 
nhojas = Sheets.Count
Dim ns As String
Dim tablero As String
Dim proyecto As String
Dim rangotab As Range

    For X = 1 To nhojas Step 1
  
    encabezado = Worksheets(X).Range("B8").Value ' buscar encabezado
    ns = Worksheets(X).Range("I5").Value    ' busca valor de control
    tablero = Worksheets(X).Range("I6").Value    ' busca valor de tablero
    proyecto = Worksheets(X).Range("c4").Value   ' busca valor de proyecto
    
            
    If encabezado = "REQUERIMIENTO DE MATERIAL" Then
    
    espacioResumen = Worksheets("resumen").Cells(Rows.Count, 1).End(xlUp).Row + 1 ' la ultima celda para escribir en resumen
    filafinal = Worksheets(X).Cells(Rows.Count, 9).End(xlUp).Row ' Encontrar el tamaño de la lista
           
    
    Worksheets(X).Range("A11:B" & filafinal).Copy ' Copiamos partida
    Worksheets("resumen").Cells(espacioResumen, 1).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
    
    Worksheets(X).Range("B11:B" & filafinal).Copy ' Copiamos partida
    Worksheets("resumen").Cells(espacioResumen, 2).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
    
        Worksheets(X).Range("C11:C" & filafinal).Copy ' Copiamos partida
    Worksheets("resumen").Cells(espacioResumen, 3).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
    
        Worksheets(X).Range("D11:D" & filafinal).Copy ' Copiamos partida
    Worksheets("resumen").Cells(espacioResumen, 4).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
    
        Worksheets(X).Range("E11:E" & filafinal).Copy ' Copiamos partida
    Worksheets("resumen").Cells(espacioResumen, 5).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
    
        Worksheets(X).Range("H11:H" & filafinal).Copy ' Copiamos partida
    Worksheets("resumen").Cells(espacioResumen, 6).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
    
        Worksheets(X).Range("I11:I" & filafinal).Copy ' Copiamos partida
    Worksheets("resumen").Cells(espacioResumen, 7).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
        
    Worksheets("resumen").Cells(espacioResumen, 8).FormulaR1C1 = ns
    Worksheets("resumen").Cells(espacioResumen, 9).FormulaR1C1 = tablero
    Worksheets("resumen").Cells(espacioResumen, 10).FormulaR1C1 = proyecto
    Worksheets("resumen").Cells(espacioResumen, 11).FormulaR1C1 = "=TODAY()"
    
    
     End If
     Next X
    
    
   Call relleno ' rellenar
   Call datostoDb ' mandamos datos a la base de datos
   
   'FILTRA CELDAS VACIAS
   
    'Worksheets("resumen").Range("A1:L1").Select
    'Selection.AutoFilter
    'ActiveSheet.Range(Selection, Selection.End(xlDown)).AutoFilter Field:=3, Criteria1:="<>"
    'Close
    
End Sub

Function datostoDb()

'declaramos las variables para la consulta con la hoja

    Dim SQLdb As String
    Dim RSdb As ADODB.Recordset
    Dim CNdb As ADODB.Connection
    Dim cod As String, partida As Integer, item As Integer, codigo As String, concepto As Variant, unidad As String, cantidad As Integer, ns As String, proyecto As String, tablero As String, fecha As Date
    Dim fila
 
'creamos la conexion con esto
    Set CNdb = New ADODB.Connection
    CNdb.Open "Provider=SQLOLEDB.1;Persist security info=True; User ID=sa;Password =mOON020106; Initial Catalog =almacenNB;Data Source=ADMINPAQ-SERVER\COMPAC;"


'creamos el objeto recordset
    
    Set RSdb = New ADODB.Recordset
    
    
With Sheets("resumen")
    On Error Resume Next
    final = Application.CountA(.Range("A:A"))

    For fila = 2 To final
    
    cod = .Cells(fila, 1)
    partida = .Cells(fila, 2)
    item = .Cells(fila, 3)
    codigo = .Cells(fila, 4)
    concepto = .Cells(fila, 5)
    unidad = .Cells(fila, 6)
    cantidad = .Cells(fila, 7)
    ns = .Cells(fila, 8)
    proyecto = .Cells(fila, 9)
    tablero = .Cells(fila, 10)
    fecha = .Cells(fila, 11)
 


    sql = "INSERT INTO requerimientos (cod, partida, item, codigo, concepto, unidad, cantidad, ns, proyecto, tablero, fecha) VALUES ('" & cod & "', '" & partida & "',  '" & item & "', '" & codigo & "', '" & concepto & "' , '" & unidad & "','" & cantidad & "','" & ns & "','" & proyecto & "','" & tablero & "', CURRENT_TIMESTAMP);"
    RSdb.Open sql, CNdb
    
Next fila

    End With
    
RSdb.Close
CNdb.Close
Set RSdb = Nothing
Set CNdb = Nothing

End Function

Private Sub UserForm_Click()

End Sub

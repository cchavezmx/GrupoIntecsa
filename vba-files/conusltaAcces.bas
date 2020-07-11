Private Sub CommandButton2_Click()

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
    
    msgbox "its Works" ' que si no, ejecuta lo que sigue despues del else

    End If
    
    cn.Close   ' cerramos la conexcion
    Set datos = Nothing  ' borramos el contenido de datos
   'Set consultaDB = Nothing
    
    Next ' e iniciamos el ciclo for
    
    End With
     nmj
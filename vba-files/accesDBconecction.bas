Sub CommandButton1_Click()

Dim cn As Object
Dim datos As Object
Dim consultasql As String
Dim conexion As String


 'declaramos los datos a subir.
 


' creamos la coneccion

Set cn = CreateObject("ADODB.connection")
    conexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=C:\Users\Saul\Documents\html5\invazoronia\vba-files\requerimientos.mdb"
    

Application.ScreenUpdating = False

'usamos with para no escribir tanto el Worbook.

'On Error Resume Next

With Sheets("resumen")
largo = Application.CountA(.Range("A:A")) ' encontramos el largo de la lista


Dim partida As String
Dim item As String
Dim codigo As String
Dim concepto As String
Dim numeroUnico As String
Dim unidad As String
Dim cantidad As String
Dim control As String
Dim proyecto As String
Dim tablero As String
    
    For i = 2 To largo
    
    
    'partida = .Range("A" & i)
    partida = Range("A2")
   ' item = .Range("B" & i)
   ' codigo = .Range("C" & i)
   ' concepto = .Range("D" & i)
   ' numeroUnico = .Range("E" & i)
   ' unidad = .Range("G" & i)
   ' cantidad = .Range("H" & i)
   ' control = .Range("I" & i)
   ' proyecto = .Range("J" & i)
   ' tablero = .Range("K" & i)
    
    consultasql = "insert into requerimientos values(" & Chr(34) & partida & Chr(34) & ")"
    cn.Open conexion
    Set datos = cn.Execute(consultasql)
    cn.Close conexion
    
    Next


End With


msgbox consultasql

End Sub
option Explicit
Public CN as ADODB.Connection


function SqlConnect (server as string, usuario as string, pwd as string, database as string) as Boolean
'creamos una funcion para poder conectarse a la base de datos. 

    'creamos el objeto de adodb y la preparamos para conectar. 
    set CN = new ADODB.Conecction
    On Error Resume Next 
    ' En caso de error, lo saltamos con esta linea. 


    whit CN
    .ConnectionString = "Provider=SQLOLEDB.1;" & _
                        "Password=" & Pass & ";" & _
                        "Persist Security Info=True;" & _
                        "User ID=" & User & ";" & _
                        "Initial Catalog=" & Database & ";" & _
                        "Data Source=" & Server
       
        .Open

    End With

    if CN.state = 0 then
        connect = False 
    else 
        connect = True
    end If 
End function 

' function query(sql as string)

'     dim RS as ADODB.recordset 
'     dim field as ADODB.field 

'     dim Col as Long

'     set RS = new ADODB.recordset
'     RS.Open sql, CN
'     RS.Open "SELECT * FROM PRODUCTOS, CN




' PARA ENVIAR LOS DATOS A LA BASE DE DATOS DE UN FORMULARION


Private Sub btn_agregar_Click()
Dim SQL As String
Dim Connected As Boolean
Dim proyectos As String, lugar As String, cliente As String, fecha As String, sn As String, tablero As String, req As String


with Datos_alta

 proyecto = me.txt_proyecto.Value
 lugat =  me.txt_lugar.Value
 cliente  =  me.txt_cliente.Value
 fecha = me.txt_fecha.Value
 sn = me.txt_sn.Value
 tablero = me.txt_tablero.Value 
 req = me.txt_req.Value


    SQL = "INSERT INTO dbo.proyectos (nserie, proyecto, lugar, residente, fecha, tablero, req) VALUES ('"& nserie & "', '" & proyecto & "',  '" & lugar & "', '" & cliente & "', '" & fecha & "', '" & tablero & "' , '" & req & "');"

 
    
    Connected = Connect("192.168.0.12", "usuario1", "12345", "inventario")

end with


    If Connected Then
        
        Call Query(SQL)
        Call Disconnect
    Else
        
        MsgBox "Could Not Connect!"
    End If
End Sub




'PARA EL EL FORMATO DE HOJA DE ALMACEN 


Function Query(SQL As String)
 
    Dim RS As ADODB.Recordset
    Dim Field As ADODB.Field
 
    Dim Col As Long
 
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, CN
    RS.Open "select * from dbo.proyectos", CN
 
    If RS.State Then
        Col = 1
        
        For Each Field In RS.Fields
            Cells(1, Col) = Field.Name
            Col = Col + 1
        Next Field
        
        Cells(2, 1).CopyFromRecordset RS
        Set RS = Nothing
    End If
End Function
Function Disconnect()
    CN.Close
End Function

Sub RunForm()
UserForm1.Show
End Sub




'BUSQUEDA DE  ACCES ESPERO ME SIRVA PARA EXCEL 


Option Explicit

'EXCELeINFO
'MVP Sergio Alejandro Campos
'http://www.exceleinfo.com
'https://www.youtube.com/user/sergioacamposh
'http://blogs.itpro.es/exceleinfo

Private Sub CommandButton1_Click()
Dim Conn As ADODB.Connection
Dim MiConexion
Dim Rs As ADODB.Recordset
Dim MiBase As String
Dim Query As String
Dim i, j

MiBase = "MiBase.accdb"

Query = "SELECT * FROM MiTabla WHERE nombre LIKE '%" & Me.TextBox1.Value & "%'"

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Rs.Open Source:=Query, _
    ActiveConnection:=Conn

'Valir si la consulta devuelve resultados
If Rs.EOF And Rs.BOF Then
    'Borrar la conexión al Recordset
    Rs.Close
    Conn.Close
    'Borrar la memoria
    Set Rs = Nothing
    Set Conn = Nothing
    
    MsgBox "No hay resultados para la consulta", vbInformation, "EXCELeINFO"
    Me.ListBox1.Clear
    Exit Sub
End If

'Asignar número de columnas
With Me.ListBox1
    .ColumnCount = Rs.Fields.Count
End With

'Recorrer el Recordset
Rs.MoveFirst
i = 1

With Me.ListBox1
        .Clear
    
    'Asignar los encabezados
        .AddItem
        
        For j = 0 To 4
            .List(0, j) = Rs.Fields(j).Name
        Next j
    
    Do
        .AddItem
        .List(i, 0) = Rs![ID]
        .List(i, 1) = Rs![Fecha]
        .List(i, 2) = Rs![Nombre]
        .List(i, 3) = Rs![Ventas]
        .List(i, 4) = Rs![Comentarios]
    i = i + 1
Rs.MoveNext

Loop Until Rs.EOF
End With

'Cerrar la conexión
Rs.Close
Conn.Close
Set Rs = Nothing
Set Conn = Nothing

'yo creo que con esto puedo hacer lo del enviar los datos a sql 
End Sub


    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim Field As ADODB.Field
 
    Dim Col As Long
 
    
    Set RS = New ADODB.Recordset
    
Final = GetUltimoR(Hoja1)

For Fila = 2 To Final
    Cod_Prod = Hoja1.Cells(Fila, 2)
    Nombre = Hoja1.Cells(Fila, 3)
    Existencia = Hoja1.Cells(Fila, 4)

 
    
    SQL = "insert into productos values('" & Cod_Prod & "','" & Nombre & "'," & Existencia & ");"
    RS.Open SQL, CN
Next
    




pista para la fecha 
    Range("N12").Select
    Selection.NumberFormat = "yyyy-mm-dd;@"

    cambiar a tipo de dato formato de fecha gregoriano. 
    
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


Private Sub CommandButton1_Click()
 
    Dim SQL As String
    Dim Connected As Boolean
 
    
    SQL = "insert into productos values('" & UCase(TextBox1) & "','" & TextBox2 & "'," & TextBox3 & ");"

 
    
    Connected = Connect("192.168.0.12", "usuario1", "12345", "inventario")
 
    If Connected Then
        
        Call Query(SQL)
        Call Disconnect
    Else
        
        MsgBox "Could Not Connect!"
    End If
End Sub

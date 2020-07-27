Dim Connected As Boolean
Dim proyectos As String, lugar As String, cliente As String, fecha As Date, sn As String, tablero As String



With Datos_alta

 proyecto = Me.txt_proyecto.Value
 lugar = Me.txt_lugar.Value
 cliente = Me.txt_cliente.Value
 sn = Me.txt_sn.Value
 tablero = Me.txt_tablero.Value
 req = Me.txt_req.Value



    SQL = "INSERT INTO dbo.proyectos (nserie, proyecto, lugar, residente, fecha, tablero, req) VALUES ('" & sn & "', '" & proyecto & "',  '" & lugar & "', '" & cliente & "', CURRENT_TIMESTAMP, '" & tablero & "','" & req & "');"
    'Connected = SqlConnect("localhost", "master")
    'Connected = SqlConnect("ADMINPAQ-SERVER\COMPAC", "sa", "mOON020106", "almacenNB")
     
     Connected = SqlConnect("DESKTOP-HIU9GUK\COMPAC", "sa", "mOON020106", "almacenNB")
    
    
End With

    If Connected Then
    
        Call Query(SQL)
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
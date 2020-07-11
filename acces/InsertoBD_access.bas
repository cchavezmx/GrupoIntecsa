Private Sub modificadatos()

Dim datoentregado As String
Dim idp As String
    
    datoentregado = Me.sumaEntregado.Value + Me.txtentregado.Value
    idp = Me.txtId.Value

    If Not IsNull(Me.txtClave) And Not IsNull(Me.txtNombre) And Not IsNull(Me.txtDescripcion) And Not IsNull(Me.txtCantidad) Then
        Select Case Me.OpenArgs
            Case 1 ' esta parte del codigo es para dar de alta productos
                CurrentDb.Execute "INSERT INTO Tbl_CatProductos(Clave,Nombre,Descripcion,Cantidad)VALUES('" _
                                & Me.txtClave & "','" & Me.txtNombre & "', '" & Me.txtDescripcion & "'," _
                                & Me.txtCantidad & ")", dbFailOnError
                MsgBox "DATOS GUARDADOS", vbInformation, "Avíso"
                Forms!Frm_CatProductos!SubFrm_CatProductos.Form.Requery
                DoCmd.Close acForm, "Frm_Datos"
            Case 2
                CurrentDb.Execute "UPDATE carrito SET entregado =" & Chr(34) & datoentregado & Chr(34) & "WHERE id = " & Chr(34) & idp & Chr(34) & "", dbFailOnError
                MsgBox "DATOS MODIFICADOS", vbInformation, "Avíso"
                Forms!Frm_CatProductos!SubFrm_CatProductos.Form.Requery
                DoCmd.Close acForm, "Frm_Datos"
        End Select
    Else
        MsgBox "ES NECESARIO LLENAR TODOS LOS CAMPOS" & vbLf & vbLf & "Verifique.", vbInformation, "Avíso"
    End If

End Sub

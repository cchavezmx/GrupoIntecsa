Option Compare Database
Option Explicit

Private Sub btnNuevo_Click()
    DoCmd.OpenForm "Frm_Datos", , , , , , 1
End Sub

Private Sub txtBusqueda_Change()
    Dim Consulta As String
    
    'Consulta = "SELECT Id,Clave,Nombre,Descripcion,Cantidad"
    'Consulta = Consulta & " FROM Tbl_CatProductos"
    'Consulta = Consulta & " WHERE Clave LIKE '*" & Replace(Me.txtbusqueda.Text, "'", "''") & "*'"
    'Me.SubFrm_CatProductos.Form.RecordSource = Consulta
    
    'A mi DB
    
    Consulta = "SELECT control,nproducto," ' le damos los primeros parametros a la busqueda, que partes de carrito queremos mostrar
    Consulta = Consulta & "FROM carrito"   ' seleccionamos la base de datos
    Consulta = Consulta & "where control like '*" & Replace(Me.txtbusqueda.Text, "'", "'") & "*'" ' creamos el string que hara la peticion sql
    Me.SubFrm_CatProductos.Form.RecordSource = Consulta ' mostramos el contenido la busqueda.
    
    
    
    
End Sub



Consulta = Consulta "SELECT control, nproducto from carrito where carrito like " & "'" & Replace(Me.txtbusqueda.Text) & "*" & "'"
Consulta = "SELECT control, id, nproducto FROM carrito WHERE control LIKE '*" & Replace(Me.box_busqueda.Value, "'", "''") & "*' and psurtir>0"




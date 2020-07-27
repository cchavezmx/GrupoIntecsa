
Private Sub Label7_Click()

End Sub




End Sub

Private Sub btn_find_Click()

Dim SQLB As String
Dim RSB As ADODB.Recordset
Dim i, j
Dim CNB As ADODB.Connection
Dim toFind  As String


Set CNB = New ADODB.Connection
    'On Error Resume Next
    ' En caso de error, lo saltamos con esta linea.
        
    CNB.Open "Provider=SQLOLEDB.1;Persist security info=True; User ID=sa;Password =mOON020106; Initial Catalog =almacenNB;Data Source=DESKTOP-HIU9GUK\COMPAC;"
    
    
toFind = Me.txt_find.Value


SQLB = "SELECT * FROM proyectos WHERE nserie LIKE '%" & toFind & "%';"
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
        
        For j = 0 To 6
         .List(0, j) = RSB.Fields(j).Name
         Next j
    Do
        .AddItem
        .List(i, 0) = RSB![nserie]
        .List(i, 1) = RSB![proyecto]
        .List(i, 2) = RSB![lugar]
        .List(i, 3) = RSB![residente]
        .List(i, 4) = RSB![fecha]
        .List(i, 5) = RSB![tablero]
        .List(i, 6) = RSB![req]
     i = i + 1
RSB.MoveNext

Loop Until RSB.EOF
End With

RSB.Close
CNB.Close
Set RSB = Nothing
Set CNB = Nothing

End Sub

Private Sub btn_seleccionar_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub seleccionar_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)


MsgBox Datos_busqueda.ListBox1.Value










End Sub

Private Sub txt_find_Change()

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
        
    CNB.Open "Provider=SQLOLEDB.1;Persist security info=True; User ID=sa;Password =mOON020106; Initial Catalog =almacenNB;Data Source=DESKTOP-HIU9GUK\COMPAC;"
    
    
toFind = Me.txt_find.Value


'SQLB = "SELECT * FROM proyectos WHERE nserie LIKE '%" & toFind & "%';"
 SQLB = "SELECT nserie, proyecto, lugar, residente, fecha, tablero, req FROM proyectos"
 
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
        
        For j = 0 To 6
         .List(0, j) = RSB.Fields(j).Name
         Next j
    Do
        .AddItem
        .List(i, 0) = RSB![nserie]
        .List(i, 1) = RSB![proyecto]
        .List(i, 2) = RSB![lugar]
        .List(i, 3) = RSB![residente]
        .List(i, 4) = RSB![fecha]
        .List(i, 5) = RSB![tablero]
        .List(i, 6) = RSB![req]
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
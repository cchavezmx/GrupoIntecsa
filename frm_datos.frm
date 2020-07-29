VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_datos 
   Caption         =   "Requerimientos"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7485
   OleObjectBlob   =   "frm_datos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_datos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_tablero_Enter()
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim SQL As String
Dim i, j
Dim elemento As String

elemento = frm_datos.ListBox1.Value

Set cn = New ADODB.Connection
'creamos el objeto llamado conexion

cn.Open "Provider=SQLOLEDB.1;Persist security info=True; User ID=sa;Password =mOON020106; Initial Catalog =almacenNB;Data Source=localhost\COMPAC;"

Set rs = New ADODB.Recordset
'creamos el objeto para el query

SQL = "SELECT tablero from proyectos where nserie = '" & elemento & "';"

rs.Open SQL, cn


rs.MoveFirst

i = 0
    
    With frm_datos.cmb_tablero
    .Clear
    .AddItem
    
    .List(0, 1) = "Proyectos"
Do
    .AddItem
    .List(i, 0) = rs![tablero]
    i = i + 1

rs.MoveNext


Loop Until rs.EOF
End With

rs.Close
cn.Close
Set rs = Nothing
Set cn = Nothing
End Sub

Private Sub CommandButton1_Click()
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim SQL As String
Dim elemento As String
Dim final As Integer
Dim colRS As Integer
Dim partida As Integer

'/ configurando las conexiones /

elemento = frm_datos.ListBox1.Value
partida = frm_datos.txt_partida.Value

Set cn = New ADODB.Connection
cn.Open "Provider=SQLOLEDB.1;Persist security info=True; User ID=sa;Password =mOON020106; Initial Catalog =almacenNB;Data Source=localhost\COMPAC;"

Set rs = New ADODB.Recordset

SQL = "Select partida, codigo, concepto, unidad, cantidad from requerimientos where ns = '" & elemento & "' and partida like '%" & partida & "%';"



'// Mandamos la consulta al RecordSet //

rs.Open SQL, cn

'/// The magic is begun ///

    Call borrar

    With Worksheets("Requerimiento") ' usamos with para no escribir tantas veces la hoja donde nos hubicamos
      
    final = .Cells(Rows.Count, 2).End(xlUp).Row 'tamaño de la lista para empezar a copiar.
    
        'For colRS = 0 To rs.Fields.Count - 1
        
        .Cells(11, 2).CopyFromRecordset rs
        
     End With


End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim columna As Long
Dim SQL As String
Dim abuscar As String


abuscar = frm_datos.ListBox1.Value

Me.cmb_tablero = ""
Me.txt_partida = ""

Set cn = New ADODB.Connection
'se crea el objeto para la conexión
cn.Open "Provider=SQLOLEDB.1;Persist security info=True; User ID=sa;Password =mOON020106; Initial Catalog =almacenNB;Data Source=localhost\COMPAC;"
'se añaden los parametros para el objeto cn


Set rs = New ADODB.Recordset
'creamos el objetos para la guardar la consulta

SQL = "Select * from proyectos where nserie ='" & abuscar & "';"
' creamos la consulta


rs.Open SQL, cn
' mandamos la consulta(sql) al objeto rs (recordset), con la conexion establecida en CN

If rs.State Then
        
  With ActiveSheet
 'mandamos los datos de la cabezera del rs a la hoja electronica
 'usamos el with para no escribir tantas lineas
 
 
    .Range("M5") = rs![nserie]
    .Range("C4") = rs![proyecto]
    .Range("C5") = rs![lugar]
    .Range("C6") = rs![residente]
    .Range("M4") = rs![fecha]
    .Range("M6") = rs![tablero]
    .Range("M7") = rs![req]

          
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

 End With
 

End If
End Sub

Private Sub UserForm_Initialize()

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim SQL As String
Dim i, j


Set cn = New ADODB.Connection
'creamos el objeto llamado conexion

cn.Open "Provider=SQLOLEDB.1;Persist security info=True; User ID=sa;Password =mOON020106; Initial Catalog =almacenNB;Data Source=localhost\COMPAC;"

Set rs = New ADODB.Recordset
'creamos el objeto para el query

SQL = "SELECT nserie from proyectos"

rs.Open SQL, cn

With ListBox1
    .ColumnCount = rs.Fields.Count
End With

rs.MoveFirst
i = 0
    
    With Me.ListBox1
    .Clear
    .AddItem
    
    .List(0, 1) = "Serie"
Do
    .AddItem
    .List(i, 0) = rs![nserie]
    i = i + 1

rs.MoveNext


Loop Until rs.EOF
End With

rs.Close
cn.Close
Set rs = Nothing
Set cn = Nothing

End Sub

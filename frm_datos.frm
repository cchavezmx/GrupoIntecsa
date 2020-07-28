VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_datos 
   Caption         =   "Requerimientos"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5940
   OleObjectBlob   =   "frm_datos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_datos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim columna As Long
Dim sql As String
Dim abuscar As String


abuscar = frm_datos.ListBox1.Value

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


Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim i, j

Set cn = New ADODB.Connection
'creamos el objeto llamado conexion

cn.Open "Provider=SQLOLEDB.1;Persist security info=True; User ID=sa;Password =mOON020106; Initial Catalog =almacenNB;Data Source=ADMINPAQ-SERVER\COMPAC;"

Set rs = New ADODB.Recordset
'creamos el objeto para el query

sql = "SELECT nserie from proyectos"

rs.Open sql, cn

With ListBox1
    .ColumnCount = rs.Fields.Count
End With

rs.MoveFirst
i = 0
    
    With Me.ListBox1
    .Clear
    .AddItem
    
    .List(0, 1) = "Proyectos"
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

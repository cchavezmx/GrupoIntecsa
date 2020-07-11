Private Sub CommandButton1_Click()

'verificamos si existe la hoja resumen

Dim existe As Boolean
Dim msg As String
Dim title As String
Dim respuesta As VbMsgBoxResult

    On Error Resume Next
    existe = (Worksheets("resumen").Name <> "")
        
    If Not existe Then
    
    Call crearDB

    ElseIf existe = True Then
    
    
    msg = "¿Esta acción actualizara el requerimiento, quieres continuar?"
    title = "Actulizacion terminada"
    respuesta = msgbox(msg, vbYesNo, title)
  
   
   
  If respuesta = vbYes Then
    msgbox "Puchste si"
    
    Else
    
   msgbox "no se como pararlo"
           
    End If
    End If

End Sub
          
Private Sub relleno()
'Definimos variables
Dim i%, Fin%
With Sheets("resumen")
Fin = Application.CountA(.Range("A:A"))
'Mediante un bucle indicamos que si una celda está vacía
'el valor sea el de la celda anterior.
For i = 3 To Fin
If .Cells(i, 9) = "" Then .Cells(i, 9) = .Cells(i - 1, 9)
If .Cells(i, 10) = "" Then .Cells(i, 10) = .Cells(i - 1, 10)
If .Cells(i, 11) = "" Then .Cells(i, 11) = .Cells(i - 1, 11)
Next
End With

End Sub




Sub crearDB()

    Worksheets.Add.Name = "resumen"
    Worksheets("resumen").Range("A1").FormulaR1C1 = "PARTIDA"
    Worksheets("resumen").Range("B1").FormulaR1C1 = "ITEM"
    Worksheets("resumen").Range("C1").FormulaR1C1 = "CODIGO"
    Worksheets("resumen").Range("D1").FormulaR1C1 = "CONCEPTO"
    Worksheets("resumen").Range("G1").FormulaR1C1 = "UNIDAD"
    Worksheets("resumen").Range("H1").FormulaR1C1 = "CANTIDAD"
    Worksheets("resumen").Range("I1").FormulaR1C1 = "CONTROL"
    Worksheets("resumen").Range("J1").FormulaR1C1 = "PROYECTO"
    Worksheets("resumen").Range("K1").FormulaR1C1 = "TABLERO"
    Worksheets("resumen").Range("L1").FormulaR1C1 = "FECHA"
    
    
 'Buscamos hojas con en encabezado requerimiento y copiamos la celdas a la hoja de resumen
 
nhojas = Sheets.Count
Dim control As String
Dim tablero As String
Dim proyecto As String
Dim rangotab As Range

    For x = 1 To nhojas Step 1
  
    encabezado = Worksheets(x).Range("B8").Value ' buscar encabezado
    control = Worksheets(x).Range("I5").Value    ' busca valor de control
    tablero = Worksheets(x).Range("I6").Value    ' busca valor de tablero
    proyecto = Worksheets(x).Range("c4").Value   ' busca valor de proyecto
    
            
    If encabezado = "REQUERIMIENTO DE MATERIAL" Then
    
    espacioResumen = Worksheets("resumen").Cells(Rows.Count, 1).End(xlUp).Row + 1 ' la ultima celda para escribir en resumen
    filafinal = Worksheets(x).Cells(Rows.Count, 9).End(xlUp).Row ' Encontrar el tamaño de la lista
    
    Worksheets(x).Range("B11:J" & filafinal).Copy ' copiamos rango de cada hoja
    Worksheets("resumen").Cells(espacioResumen, 1).PasteSpecial Paste:=xlPasteValues  'pegamos solo valores
    Worksheets("resumen").Cells(espacioResumen, 9).FormulaR1C1 = control
    Worksheets("resumen").Cells(espacioResumen, 10).FormulaR1C1 = tablero
    Worksheets("resumen").Cells(espacioResumen, 11).FormulaR1C1 = proyecto
    
    
     End If
     Next x

        
   'RELLENAR NOMBRE TABLERO
   
    Call relleno
   
   
   'FILTRA CELDAS VACIAS
   
    'Worksheets("resumen").Range("A1:L1").Select
    'Selection.AutoFilter
    'ActiveSheet.Range(Selection, Selection.End(xlDown)).AutoFilter Field:=3, Criteria1:="<>"
    'Close
    
End Sub

Sub accesDB()

Dim cn As Object
Dim datos As Object
Dim consultasql As String
Dim conexion As String


 'declaramos los datos a subir.
 
Dim partida As String
Dim item As String
Dim codigo As String
Dim concepto As String
Dim numeroUnico As String
Dim unidad As String
Dim cantidad As Integer
Dim control As String
Dim proyecto As String
Dim tablero As String

' creamos la coneccion

Set cn = CreateObject("ADODB.connection")
    conexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=C:\Users\Saul\Documents\html5\invazoronia\vba-files\requerimientos.mdb"

'usamos with para no escribir tanto el Worbook.

With Sheets("resumen")

largo = Application.CountA(.Range("A:A")) ' encontramos el largo de la lista


    For i = 1 To largo
    
    partida = Range(i, 1)
    item = Range(i, 2)
    codigo = Range(i, 3)
    concepto = Range(i, 4)
    numeroUnico = Range(i, 5)
    unidad = Range(i, 7)
    cantidad = Range(i, 8)
    control = Range(i, 9)
    proyecto = Range(i, 10)
    tablero = Range(i, 111)
    
    consultasql = "insert into requerimientos values(" & Chr(34) & partida & Chr(34) & "," & Chr(34) & item & Chr(34) & "," & Chr(34) & codigo & Chr(34) & ", " & Chr(34) & concepto & Chr(34) & "," & Chr(34) & numeroUnico & Chr(34) & "," & Chr(34) & unidad & Chr(34) & "," & Chr(34) & cantidad & Chr(34) & "," & Chr(34) & control & Chr(34) & "," & Chr(34) & proyecto & Chr(34) & "," & Chr(34) & tablero & Chr(34) & ")"
    cn.Open conexion
    Set datos = cn.Execute(consultasql)
    msgbox "Registro realizado exitosamente", vbInformation, "Nueva persona"
    
    Next i


End Sub



Sub RunForm()
    Load UserForm1
    UserForm1.Show
End Sub
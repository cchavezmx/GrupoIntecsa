VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7020
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

'verificamos si existe la hoja resumen

Dim existe As Boolean
    
    On Error Resume Next
    existe = (Worksheets("resumen").Name <> "")
    
    If Not existe Then
    Worksheets.Add.Name = "resumen"
    Worksheets("resumen").Range("A1").FormulaR1C1 = "PARTIDA"
    Worksheets("resumen").Range("B1").FormulaR1C1 = "ITEM"
    Worksheets("resumen").Range("C1").FormulaR1C1 = "CODIGO"
    Worksheets("resumen").Range("D1").FormulaR1C1 = "CONCEPTO"
    Worksheets("resumen").Range("G1").FormulaR1C1 = "UNIDAD"
    Worksheets("resumen").Range("H1").FormulaR1C1 = "CANTIDAD"
    Worksheets("resumen").Range("I1").FormulaR1C1 = "ID"
    End If
    
'Buscamos hojas con en encabezado requerimiento y copiamos la celdas a la hoja de resumen
nhojas = Sheets.Count
Dim tablero As String
Dim rangotab As Range

    For X = 1 To nhojas Step 1
  
    encabezado = Worksheets(X).Range("B8").Value
    tablero = Worksheets(X).Range("I5").Value
            
    If encabezado = "REQUERIMIENTO DE MATERIAL" Then
    
    espacioResumen = Worksheets("resumen").Cells(Rows.Count, 1).End(xlUp).Row + 1 ' la ultima celda para escribir
   
       
       Worksheets(X).Range("B11:J357").Copy
       Worksheets("resumen").Cells(espacioResumen, 1).PasteSpecial Paste:=xlPasteValues
       Worksheets("resumen").Cells(espacioResumen, 9).FormulaR1C1 = tablero

     End If
     Next X

        
   'RELLENAR NOMBRE TABLERO
   
    Call FillCellsFromAbove
   
   
   'FILTRA CELDAS VACIAS
   
    Worksheets("resumen").Range("A1:I1").Select
    Selection.AutoFilter
    ActiveSheet.Range(Selection, Selection.End(xlDown)).AutoFilter Field:=3, Criteria1:="<>"
   

End Sub
Sub FillCellsFromAbove()
    ' Turn off screen updating to improve performance
    Application.ScreenUpdating = False
    On Error Resume Next
    Worksheets("resumen").Select
    ' Look in column A
    With Columns(9)
        ' For blank cells, set them to equal the cell above
        .SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        'Convert the formula to a value
        .Value = .Value
    End With
    Err.Clear
    Application.ScreenUpdating = True
End Sub


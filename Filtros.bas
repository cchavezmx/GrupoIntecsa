Attribute VB_Name = "Filtros"
Sub borrar()
Attribute borrar.VB_ProcData.VB_Invoke_Func = " \n14"
'
' borrar Macro
'

'
    Range("B11:N11").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B11").Select
End Sub

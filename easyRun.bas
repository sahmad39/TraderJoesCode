Attribute VB_Name = "Module4"
Sub makeIt()
Attribute makeIt.VB_ProcData.VB_Invoke_Func = " \n14"
'
' makeIt Macro
'

'
    ActiveWorkbook.RefreshAll
    Application.Run "tryThis.xlsm!Module1.ProdArray"
End Sub

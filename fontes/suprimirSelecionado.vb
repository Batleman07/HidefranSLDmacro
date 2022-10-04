' ******************************************************************************
' C:\Users\Projetos4\AppData\Local\Temp\swx11936\Macro1.swb - macro recorded on 10/04/22 by projetos4
' ******************************************************************************
Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

Set swApp = Application.SldWorks

Set Part = swApp.ActiveDoc
Part.EditSuppress2
Part.ClearSelection2 True
End Sub
'

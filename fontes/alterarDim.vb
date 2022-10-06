



' ******************************************************************************
' C:\Users\Projetos4\AppData\Local\Temp\swx15560\Macro1.swb - macro recorded on 10/06/22 by projetos4
' ******************************************************************************
Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

Set swApp = Application.SldWorks

Set Part = swApp.ActiveDoc
boolstatus = Part.Extension.SelectByID2("D1@Esboço8@" & Part.GetTitle, "DIMENSION", 5.54982694156027E-02, 5.00000000000234E-05, -0.101727442102208, False, 0, Nothing, 0)
Dim myDimension As Object
Set myDimension = Part.Parameter("D1@Esboço8")
myDimension.SystemValue = 205 / 1000
Part.ClearSelection2 True
End Sub

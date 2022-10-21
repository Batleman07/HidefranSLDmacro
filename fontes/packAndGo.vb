Dim swModelDocExt As SldWorks.ModelDocExtension
Dim swPackAndGo As SldWorks.PackAndGo
Dim openFile As String
Dim pgFileNames As Variant
Dim pgFileStatus As Variant
Dim pgGetFileNames As Variant
Dim pgDocumentStatus As Variant
Dim status As Boolean
Dim warnings As Long
Dim errors As Long
Dim i As Long
Dim namesCount As Long
Dim myPath As String
Dim statuses As Variant

Dim partDocExt As SldWorks.ModelDocExtension

Sub PackAndGo()

Set swApp = GetObject(, "SldWorks.Application")
Set swModelDoc = swApp.OpenDoc("E:\FORMAT\FormatSketch.SLDPRT", swDocPART)
Set swModelDocExt = swModelDoc.Extension

'Open Part
openFile = "E:\FORMAT\FormatSketch.SLDPRT"

'Get Pack and Go object
Set swPackAndGo = swModelDocExt.GetPackAndGo

'Include any drawings
swPackAndGo.IncludeDrawings = True

'Set folder where to save the files
myPath = "E:\FORMAT\Temp\"
status = swPackAndGo.SetSaveToName(True, myPath)

'Flatten the Pack and Go folder structure; save all files to the root directory
swPackAndGo.FlattenToSingleFolder = True

'Pack and Go
statuses = swModelDocExt.SavePackAndGo(swPackAndGo)
        
End Sub
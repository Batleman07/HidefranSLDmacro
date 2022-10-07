
Dim swApp As Object

Dim swModel As ModelDoc2
Dim cusPropMgr As CustomPropertyManager
Dim config As Configuration
Dim propertyName As String
Dim valOut As String
Dim resolvedValOut As String
Dim newName As String
Dim Os As String
Dim strDelPath As String

Dim wasResolved As Boolean
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

	Set swApp = Application.SldWorks
	Set swModel = swApp.ActiveDoc
	Set config = swModel.GetActiveConfiguration
	Set cusPropMgr = config.CustomPropertyManager


	propertyName = "Posição"
	Os = "OS-3073 - POS-"
	strDelPath = swModel.GetPathName

	' Get the property value
	longstatus = cusPropMgr.Get5(propertyName, False, valOut, resolvedValOut, wasResolved)

	' Create the new name with full path and extension
	newName = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\")) + Os + valOut + Right(swModel.GetPathName, 7)

	Dim fso As FileSystemObject
	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FileExists(newName) Then
		If MsgBox("SUBSTITUIR:  " & newName, vbYesNo) = vbYes Then
			' Save As
			boolstatus = swModel.Extension.SaveAs(newName, 0, 0, Nothing, longstatus, longwarnings)
	'        Debug.Print ("Saved file:  " & boolstatus)
		End If
	Else
		If MsgBox("RENOMEAR PARA:  " & newName, vbYesNo) = vbYes Then
			' Save As
			boolstatus = swModel.Extension.SaveAs(newName, 0, 0, Nothing, longstatus, longwarnings)
			
			Kill strDelPath
			
		End If
	End If

	Set swApp = Nothing
	Set swModel = Nothing
	Set config = Nothing
	Set cusPropMgr = Nothing
	propertyName = Clear
	Os = Clear
	strDelPath = Clear
	longstatus = Clear
	longstatus = Clear
	boolstatus = Clear
	Set fso = Nothing

End Sub

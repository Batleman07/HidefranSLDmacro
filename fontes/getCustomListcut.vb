' Preconditions:
' 1. Open a new SOLIDWORKS session.
' 2. Open public_documents\samples\tutorial\api\weldment_box3.sldprt.
' 3. Click Tools > Options > Document Properties > Weldments >
'    Rename cut list folders with Description property value > OK.
' 4. Right-click Cut list(31) in the FeatureManager design tree
'    and click Update.
' 5. Verify that the specified assembly template exists.
' 6. Open the Immediate window.
'
' Postconditions:
' 1. Creates an assembly document and inserts the part document.
' 2. Traverses the part's FeatureManager design tree and gets
'    the names of custom properties, values, and evaluated values
'    for the cut-list folders in the part document.
' 3. Examine the Immediate window.
'
' NOTE: Because the part document is used elsewhere, do not
' save changes.
'-------------------------------------------------------
Option Explicit
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swModelDocExt As SldWorks.ModelDocExtension
Dim swAssem As SldWorks.AssemblyDoc
Dim swSelMgr As SldWorks.SelectionMgr
Dim status As Boolean
Dim errors As Long
Dim warnings As Long
Dim fileName As String
 
Sub VisitFeatureCustomProperties(swDocFeat As SldWorks.Feature)
    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Dim propNames As Variant
    Dim vName As Variant
    Dim propName As String
    Dim Value As String
    Dim resolvedValue As String
    Set swCustPropMgr = swDocFeat.CustomPropertyManager
    If Not swCustPropMgr Is Nothing Then
        propNames = swCustPropMgr.GetNames
        If Not IsEmpty(propNames) Then
            Debug.Print swDocFeat.Name, swDocFeat.GetTypeName
            For Each vName In propNames
                propName = vName
                Call swCustPropMgr.Get2(propName, Value, resolvedValue)
                Debug.Print "", "", propName, Value, resolvedValue
            Next vName
        End If
    End If
End Sub
 
Sub VisitDocWeldmentProperties(swCompDoc As SldWorks.ModelDoc2)
    Dim swFeature As SldWorks.Feature
    Dim swSubFeature As SldWorks.Feature
    Dim swCutFolder As SldWorks.BodyFolder
    Set swFeature = swCompDoc.FirstFeature
    Do While Not swFeature Is Nothing
        Set swSubFeature = swFeature.GetFirstSubFeature
        Do While Not swSubFeature Is Nothing
            If swSubFeature.GetTypeName2 = "CutListFolder" Then
                Set swCutFolder = swSubFeature.GetSpecificFeature2
            End If
            If Not swCutFolder Is Nothing Then
                    Call VisitFeatureCustomProperties(swSubFeature)
            End If
            Set swSubFeature = swSubFeature.GetNextSubFeature
        Loop
        Set swFeature = swFeature.GetNextFeature
    Loop
End Sub
 
Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2016\templates\assembly.asmdot", 0, 0, 0)
    swApp.ActivateDoc2 "Assem1", False, errors
    Set swModel = swApp.ActiveDoc
    fileName = "C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS 2018\samples\tutorial\api\weldment_box3.sldprt"
    Set swAssem = swModel
    swApp.ActivateDoc2 "Assem1", False, errors
    swAssem.AddComponent5 fileName, swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0.508489013092717, 0.724898979334123, 0.550645508621615
    Set swSelMgr = swModel.SelectionManager
    Set swModelDocExt = swModel.Extension
    status = swModelDocExt.SelectByID2("weldment_box3-1@Assem1", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    Dim swSelComp As SldWorks.Component2
    Dim refConfig As String
    Dim swCompDoc As SldWorks.ModelDoc2
    Set swSelComp = swSelMgr.GetSelectedObject6(1, -1)
    Set swCompDoc = swSelComp.GetModelDoc
    Dim configNames As Variant
    Dim vName As Variant
    Dim configName As String
    configNames = swCompDoc.GetConfigurationNames()
    For Each vName In configNames
        configName = vName
        Debug.Print "-----------------------------------------------"
        Debug.Print "Configuration: " + configName
        status = swCompDoc.ShowConfiguration2(configName)
        Call VisitDocWeldmentProperties(swCompDoc)
    Next vName
End Sub
This example shows how to select both solid and surface bodies in either a part or an assembly.

'----------------------------------------
' Preconditions:
' 1. Open a part or assembly.
' 2. Open the Immediate window.
'
' Postconditions:
' 1. Selects all solid and surface bodies.
' 2. Examine the Immediate window.
'----------------------------------------
Option Explicit
Sub SelectBodies(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, vBody As Variant, sPadStr As String)
    Dim swModExt As SldWorks.ModelDocExtension
    Dim swBody As SldWorks.Body2
    Dim sBodySelStr As String
    Dim sBodyTypeSelStr As String
    Dim i As Long
    Dim bRet As Boolean
    If IsEmpty(vBody) Then Exit Sub
    Set swModExt = swModel.Extension
    For i = 0 To UBound(vBody)
        Set swBody = vBody(i)
        sBodySelStr = swBody.GetSelectionId
        Debug.Print "  " & sPadStr & sBodySelStr
        Select Case swBody.GetType
            Case swSolidBody
                sBodyTypeSelStr = "SOLIDBODY"
            Case swSheetBody
                sBodyTypeSelStr = "SURFACEBODY"
            Case Else
                Debug.Assert False
        End Select
        bRet = swModExt.SelectByID2(sBodySelStr, sBodyTypeSelStr, 0#, 0#, 0#, True, 0, Nothing, swSelectOptionDefault): Debug.Assert bRet
    Next i
End Sub
Sub ProcessComponent(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swComp As SldWorks.Component2, nLevel As Long)
    Dim vChildComp As Variant
    Dim swChildComp As SldWorks.Component2
    Dim sPadStr As String
    Dim vBody As Variant
    Dim i As Long
    For i = 0 To nLevel - 1
        sPadStr = sPadStr + "  "
    Next i
    Debug.Print sPadStr & swComp.Name2 & " <" & swComp.ReferencedConfiguration & ">"
    ' Solid bodies
    vBody = swComp.GetBodies2(swSolidBody)
    SelectBodies swApp, swModel, vBody, sPadStr
    ' Surface bodies
    vBody = swComp.GetBodies2(swSheetBody)
    SelectBodies swApp, swModel, vBody, sPadStr
    vChildComp = swComp.GetChildren
    For i = 0 To UBound(vChildComp)
        Set swChildComp = vChildComp(i)
        ProcessComponent swApp, swModel, swChildComp, nLevel + 1
    Next i
End Sub
Sub ProcessAssembly(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2)
    Dim swConfigMgr As SldWorks.ConfigurationManager
    Dim swConf As SldWorks.Configuration
    Dim swRootComp As SldWorks.Component2
    Set swConfigMgr = swModel.ConfigurationManager
    Set swConf = swConfigMgr.ActiveConfiguration
    Set swRootComp = swConf.GetRootComponent3(True)
    ProcessComponent swApp, swModel, swRootComp, 1
End Sub
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swPart As SldWorks.PartDoc
    Dim vBody As Variant
    Dim i As Long
    Dim bRet As Boolean
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    swModel.ClearSelection2 True
    Debug.Print "File = " & swModel.GetPathName
    Select Case swModel.GetType
        Case swDocPART
            Set swPart = swModel
            ' Solid bodies
            vBody = swPart.GetBodies2(swSolidBody, True)
            SelectBodies swApp, swModel, vBody, ""
            ' Surface bodies
            vBody = swPart.GetBodies2(swSheetBody, True)
            SelectBodies swApp, swModel, vBody, ""
        Case swDocASSEMBLY
            ProcessAssembly swApp, swModel
        Case Else
            Exit Sub
    End Select
End Sub
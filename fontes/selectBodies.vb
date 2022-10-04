'Select Bodies Example (VB)
'This example shows how to select both solid and surface bodies in either a part or an assembly.
'----------------------------------------
'
' Preconditions: Part or assembly is open.
'
' Postconditions: All solid and surface bodies are selected.
'
'----------------------------------------

Option Explicit

Public Enum swDocumentTypes_e
    swDocNONE = 0       '  Used to be TYPE_NONE
    swDocPART = 1       '  Used to be TYPE_PART
    swDocASSEMBLY = 2   '  Used to be TYPE_ASSEMBLY
    swDocDRAWING = 3    '  Used to be TYPE_DRAWING
End Enum

Public Enum swBodyType_e
    swSolidBody = 0
    swSheetBody = 1
    swWireBody = 2
    swMinimumBody = 3
    swGeneralBody = 4
    swEmptyBody = 5
End Enum

Sub SelectBodies(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, vBody As Variant, sPadStr As String)
    Dim swModExt                    As SldWorks.ModelDocExtension
    Dim swBody                      As SldWorks.Body2
    Dim sBodySelStr                 As String
    Dim sBodyTypeSelStr             As String
    Dim i                           As Long
    Dim bRet                        As Boolean
    
    If IsEmpty(vBody) Then Exit Sub
    Set swModExt = swModel.Extension
    
    For i = 0 To UBound(vBody)
        swModel.ClearSelection2 True '*ADDED*
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
        If bRet Then '*ADDED*
            InsertWeldmentCutList swApp, swModel '*ADDED*
        End If '*ADDED*
    Next i
End Sub

Sub ProcessComponent(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swComp As SldWorks.Component2, nLevel As Long)
    Dim vChildComp                  As Variant
    Dim swChildComp                 As SldWorks.Component2
    Dim swCompConfig                As SldWorks.Configuration
    Dim sPadStr                     As String
    Dim vBody                       As Variant
    Dim i                           As Long
    
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
    Dim swConfigMgr                 As SldWorks.ConfigurationManager
    Dim swConf                      As SldWorks.Configuration
    Dim swRootComp                  As SldWorks.Component2
    
    Set swConfigMgr = swModel.ConfigurationManager
    Set swConf = swConfigMgr.ActiveConfiguration
    Set swRootComp = swConf.GetRootComponent
    ProcessComponent swApp, swModel, swRootComp, 1
End Sub

Sub InsertWeldmentCutList(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2) '*CHANGED* Renamed and added variables
'Insert Weldment Cut List Example (VB)
'This example shows how to insert a weldment cut list in a weldment part document.
'
' Preconditions:
'       (1) Weldment part is open
'       (2) At least one solid body is selected in FeatureManager design tree.
'
' Postconditions:
'       (1) New weldment Cut-List-Item folder is created in the FeatureManager
'           design tree.
'       (2) Selected bodies are placed in the new folder.
'
'*DON'T NEED* since it's a subroutine    Dim swApp                           As SldWorks.SldWorks
'*DON'T NEED* since it's a subroutine    Dim swModel                         As SldWorks.ModelDoc2
    Dim swSelMgr                        As SldWorks.SelectionMgr
    Dim swFeatMgr                       As SldWorks.FeatureManager
    Dim swCutListFeat                   As SldWorks.Feature
    Dim nSelCount                       As Long
    Dim swBody()                        As SldWorks.Body2
    Dim i                               As Long
    Dim bRet                            As Boolean
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swSelMgr = swModel.SelectionManager
    Set swFeatMgr = swModel.FeatureManager
    
    nSelCount = swSelMgr.GetSelectedObjectCount
    ReDim swBody(nSelCount - 1)
    
    For i = 1 To nSelCount
        Set swBody(i - 1) = swSelMgr.GetSelectedObject5(i)
    Next i
    
    Set swCutListFeat = swFeatMgr.InsertWeldmentCutList
    
    Debug.Print "File = " & swModel.GetPathName
    Debug.Print "  " & swCutListFeat.Name
    
    For i = 0 To nSelCount - 1
        Debug.Print "    " & swBody(i).GetSelectionId
    Next i
    
End Sub

Sub main()
    Dim swApp                       As SldWorks.SldWorks
    Dim swModel                     As SldWorks.ModelDoc2
    Dim swPart                      As SldWorks.PartDoc
    Dim vBody                       As Variant
    Dim i                           As Long
    Dim bRet                        As Boolean
    
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
            
'*DON'T NEED*           ' Surface bodies
'*DON'T NEED*            vBody = swPart.GetBodies2(swSheetBody, True)
'*DON'T NEED*            SelectBodies swApp, swModel, vBody, ""
        
        Case swDocASSEMBLY
'*DON'T NEED*            ProcessAssembly swApp, swModel
            
        Case Else
            Exit Sub
    End Select
    swModel.ClearSelection2 True '*ADDED*
End Sub
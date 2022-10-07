' Preconditions:
' 1. Open public_documents\samples\tutorial\api\weldment_box3.sldprt.
' 2. Click Tools > Options > Document Properties > Weldments >
'    Rename cut list folders with Description property value > OK.
' 3. Right-click Cut list(31) in the FeatureManager design tree
'    and click Update.
' 4. Expand Cut list(31) and click TUBE, SQUARE 30 X 30 X 2.60<1>.
' 5. Open the Immediate window
'
' Postconditions:
' 1. Gets the custom properties for the selected cut-list item.
' 2. Examine the Immediate window.
'
' NOTE: Because the part document is used elsewhere, do not
' save changes.
'----------------------------------------
Option Explicit
Sub main() 
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swCutlistItem As SldWorks.Feature
    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Dim names As Variant
    Dim name As Variant
    Dim textexp As String
    Dim evalval As String
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swSelMgr = swModel.SelectionManager
    Set swCutlistItem = swSelMgr.GetSelectedObject6(1, 0)    
    Set swCustPropMgr = swCutlistItem.CustomPropertyManager
    Debug.Print "Custom properties for selected weldment cut-list item"
    Debug.Print "Number of custom properties = " + CStr(swCustPropMgr.Count)
    Debug.Print "Name", "Text Expression", "Value", "Type"    
    names = swCustPropMgr.GetNames
    For Each name In names
        swCustPropMgr.Get2 name, textexp, evalval
        Debug.Print name, textexp, evalval, swCustPropMgr.GetType(name)
    Next name
End Sub
This example shows how to insert a Body-Delete/Keep feature into a multibody part.

'----------------------------------------------------------------------------
' Preconditions:
' 1. Open install_dir\samples\tutorial\multibody\multi_local.sldprt.
' 2. Open an Immediate window.
'
' Postconditions:
' 1. Creates Body-Delete/Keep 1 in the FeatureManager design tree.
' 2. Inspect the Immediate window.
'
' NOTE: Because the model is used elsewhere, do not save changes.
' ---------------------------------------------------------------------------

Dim swApp As SldWorks.SldWorks
Dim Part As SldWorks.ModelDoc2
Dim myFeature As SldWorks.Feature
Dim insDelBody As SldWorks.DeleteBodyFeatureData
Dim boolstatus As Boolean
Option Explicit

Sub main()

    Set swApp = Application.SldWorks
   

    Set Part = swApp.ActiveDoc
    ' Select body to delete
    boolstatus = Part.Extension.SelectByID2("Fillet5", "SOLIDBODY", 5.92851881957586E-02, 4.09115950836281E-02, -1.97275812591329E-02, True, 0, Nothing, 0)
    ' Create a Body-Delete/Keep feature
    Set myFeature = Part.FeatureManager.InsertDeleteBody2(False)
   

    Set insDelBody = myFeature.GetDefinition
    Debug.Print "Number of bodies in this Body-Delete/Keep feature: " & insDelBody.GetBodiesCount
    Debug.Print "Keep bodies: " & insDelBody.KeepBodies

End Sub
This example shows how to obtain the contents of feature folders.

'---------------------------------------------------------------
' Preconditions:
' 1. Open a part document.
' 2. Right-click a feature in the FeatureManager design tree.
' 3. Select Add to New Folder, which moves the feature to
'    the new folder.
' 4. Open the Immediate Window.
'
' Postconditions: Inspect the Immediate Window.
'---------------------------------------------------------------
Option Explicit

Sub main()

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swFeatMgr As SldWorks.FeatureManager
    Dim swFeat As SldWorks.Feature
    Dim swFeatFolder As SldWorks.FeatureFolder
    Dim swFtrFolder As SldWorks.Feature
    Dim FeatType As String
    Dim FeatTypeName As String
    Dim NbrOfFeatures As Long
    Dim Features As Variant
    Dim i As Long
    Dim featureArr(2) As SldWorks.Feature
    Dim featureObj As Variant
   

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swFeatMgr = swModel.FeatureManager
   

    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        FeatType = swFeat.Name
        FeatTypeName = swFeat.GetTypeName2
        Dim temp As Long
        temp = InStr(1, FeatType, "EndTag", vbTextCompare)
        If (FeatTypeName = "FtrFolder" And temp = 0) Then
            Debug.Print FeatType & " [" & FeatTypeName & "]"
            Set swFeatFolder = swFeat.GetSpecificFeature2
            Features = swFeatFolder.GetFeatures
            Debug.Print "    Number of Features: " & swFeatFolder.GetFeatureCount
            For i = 0 To (swFeatFolder.GetFeatureCount - 1)
                 Set swFtrFolder = Features(i)
                 Debug.Print "     Name of feature: " & swFtrFolder.Name
                 Debug.Print "     Type of feature: " & swFtrFolder.GetTypeName2
            Next i
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
 

End Sub
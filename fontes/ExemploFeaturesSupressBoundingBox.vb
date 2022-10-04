
Dim swApp As SldWorks.SldWorks

Dim Part As SldWorks.PartDoc
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim swFeatData As Object
Sub main()

Set swApp = Application.SldWorks

Set Part = swApp.ActiveDoc
Dim swFeat As SldWorks.Feature
Dim swFeatMgr As SldWorks.FeatureManager

Set swFeatMgr = Part.FeatureManager

Set swFeatData = swFeatMgr.CreateDefinition(swFeatureNameID_e.swFmBoundingBox)
swFeatData.IncludeHiddenBodies = False
swFeatData.IncludeSurfaces = False
swFeatData.ReferenceFaceOrPlane = 1
Set swFeat = swFeatMgr.CreateFeature(swFeatData)
Part.ClearSelection2 True
Set swApp = Nothing
Set Part = Nothing
Set swFeatMgr = Nothing
Set swFeatData = Nothing
End Sub

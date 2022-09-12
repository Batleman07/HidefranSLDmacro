Function criarCaixa()
'////////////////////////////variaveis//////////////////////////////////////////
    Dim swApp As SldWorks.SldWorks
    
    Dim Part As SldWorks.PartDoc
    Dim swModel As SldWorks.ModelDoc2
    
    Dim boolstatus As Boolean
    Dim longstatus As Long, longwarnings As Long
    Dim swFeatData As Object
'////////////////////////////variaveis//////////////////////////////////////////

        Set swApp = Application.SldWorks
        Set Part = swApp.ActiveDoc
        
'////////////////////////////variaveis//////////////////////////////////////////
    Dim swFeat As SldWorks.Feature
    Dim swFeatMgr As SldWorks.FeatureManager
'////////////////////////////variaveis//////////////////////////////////////////
'////////////////////////////caixa delimitadora/////////////////////////////////
        Set swFeatMgr = Part.FeatureManager
        
        Set swFeatData = swFeatMgr.CreateDefinition(swFeatureNameID_e.swFmBoundingBox)
            swFeatData.IncludeHiddenBodies = False
            swFeatData.IncludeSurfaces = False
            swFeatData.ReferenceFaceOrPlane = 1
        Set swFeat = swFeatMgr.CreateFeature(swFeatData)
        
        ' ocultar caixa
        Set Part = swApp.ActiveDoc
        boolstatus = Part.Extension.SelectByID2("Caixa delimitadora", "BBOXSKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Part.BlankSketch
        
    Part.ClearSelection2 True
            Set swApp = Nothing
            Set Part = Nothing
            Set swFeatMgr = Nothing
            Set swFeatData = Nothing
'////////////////////////////caixa delimitadora//////////////////////////////////////////
End Function
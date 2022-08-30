
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim swExt As SldWorks.ModelDocExtension
Dim oLocal As swUserPreferenceIntegerValue_e
Dim oUnidade As swUnitsMassPropMass_e
Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swExt = swModel.Extension
    oLocal = swUserPreferenceIntegerValue_e.swUnitsMassPropMass
    oUnidade = swUnitsMassPropMass_e.swUnitsMassPropMass_Kilograms
    'swUnitsMassPropMass_e.swUnitsMassPropMass_Kilograms
    
        boolstatus = swExt.SetUserPreferenceInteger(oLocal, 0, oUnidade)
        
            If boolstatus = False Then
                 swExt.SetUserPreferenceInteger swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_Custom
                 swExt.SetUserPreferenceInteger oLocal, 0, oUnidade
            End If
        
        boolstatus = swModel.EditRebuild3()


Set swApp = Nothing
Set swModel = Nothing
Set swExt = Nothing
oLocal = Clear

oUnidade = Clear

End Sub

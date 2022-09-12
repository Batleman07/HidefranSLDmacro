Function setPref(qualInfo As Object, unidade As String)
'////////////////////////////possibilidade de colocar um case com as possibilidades de alteração//////////////////////////////////////////
Dim swApp As SldWorks.SldWorks

Dim swModel As SldWorks.ModelDoc2
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim oLocal As swUserPreferenceIntegerValue_e
oLocal = swUnitsMassPropMass


    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
        boolstatus = swModel.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.oLocal, 0, swUnitsMassPropMass_e.swUnitsMassPropMass_Kilograms)
            If boolstatus = False Then
                 swModel.Extension.SetUserPreferenceInteger swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_Custom
                 swModel.Extension.SetUserPreferenceInteger swUserPreferenceIntegerValue_e.swUnitsMassPropMass, 0, swUnitsMassPropMass_e.swUnitsMassPropMass_Kilograms
            End If
        
        boolstatus = swModel.EditRebuild3()

End Function
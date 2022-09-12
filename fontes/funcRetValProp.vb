
Function getPropVal(strValor As String) As String
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
       
    Dim swPropMgr As SldWorks.CustomPropertyManager
    Dim swExt As SldWorks.ModelDocExtension
    
    Dim boolstatus As Boolean
    
    Dim strSaida As String
    Dim strCalculado As String
    
    
    
        Set swApp = Application.SldWorks
        Set swModel = swApp.ActiveDoc
        Set swExt = swModel.Extension
        Set swPropMgr = swExt.CustomPropertyManager(swApp.GetActiveConfigurationName(swModel.GetPathName))
            boolstatus = swPropMgr.Get3(strValor, True, strSaida, strCalculado)
                       
            
          
            getPropVal = strCalculado
              

End Function
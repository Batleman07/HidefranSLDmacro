' Preconditions:
' 1. Open public_documents\samples\tutorial\api\2012-sm.sldprt.
' 2. Open the Immediate window.
'
' Postconditions:
' 1. Creates two configurations, Config1 and Config1 Derived.
' 2. Gets whether each of these configurations are derived and the
'    number of children configurations for Config1.
' 3. Gets the number and names of the configurations in the part.
' 4. Examine the Immediate window.
'
' NOTE: Because the part is used elsewhere, do not save changes.
'---------------------------------------------------------------
Option Explicit
Dim swApp As SldWorks.SldWorks
Dim Model As ModelDoc2
Dim ConfigMgr As ConfigurationManager
Dim C1a As Configuration
Dim SelMgr As SelectionMgr
Dim CDerived As Configuration
Dim CParent As Configuration
Dim V As Variant
Dim i As Long

Sub main()
    Set swApp = Application.SldWorks
    Set Model = swApp.ActiveDoc
    Set SelMgr = Model.SelectionManager
    Set ConfigMgr = Model.ConfigurationManager    
    ' Create a new configuration named Config1
    ConfigMgr.AddConfiguration2 "Config1", "Config1 comment", "alternateName", 1, "", "no description", True

    ' Create a derived configuration called "Config1 Derived" whose parent configuration is "Config1"
    ConfigMgr.AddConfiguration2 "Config1 Derived", "Config1 Derived Comment", "Alternate Name", 1, "Config1", "no description", True
    ' Add a rebuild and save mark to all configurations
    bRet = ConfigMgr.AddRebuildSaveMark(2, Nothing)

    ' Display all configuration nodes expanded in the top pane
    ConfigMgr.SetExpanded 0, True

    ' Show Config1 and make it the active configuration
    Model.ShowConfiguration2 ("Config1")

    ' Select the configuration to preview in the bottom pane
    boolstatus = Model.Extension.SelectByID2("Config1", "CONFIGURATIONS", 0, 0, 0, False, 0, Nothing, 0)
    ConfigMgr.ShowPreview = True

    ConfigMgr.ShowConfigurationDescriptions = True
    ConfigMgr.ShowConfigurationNames = True

    ' Get Config1
    Set C1a = Model.GetActiveConfiguration
    ' Determine if the active configuration is a derived configuration
    Debug.Print C1a.Name & " configuration derived? " & C1a.IsDerived
    Dim VChildren As Variant
    
    ' Determine the number of children configurations
    Debug.Print "  Number of children configurations: " & C1a.GetChildrenCount

    ' Get all of the children configurations
    VChildren = C1a.GetChildren
    Set CDerived = VChildren(0)

    ' Determine if the active configuration is a derived configuration
    Debug.Print CDerived.Name & " configuration derived? " & CDerived.IsDerived
    
    ' Get the parent configuration of the derived configuration
    Set CParent = CDerived.GetParent
    ' Determine the number of configurations in this document
    Debug.Print "Number of configurations in part: " & swApp.GetConfigurationCount("C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS 2018\samples\tutorial\api\2012-sm.sldprt")
 
    ' Get the names of these configurations
    V = swApp.GetConfigurationNames("C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS 2018\samples\tutorial\api\2012-sm.sldprt")
    
    Debug.Print "Names of configurations in part:"
    For i = 0 To UBound(V)
        ' Print the names of these configurations
        Debug.Print "  " & V(i)
    Next

End Sub
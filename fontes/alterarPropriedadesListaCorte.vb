Andrzej Kurlapski - 2018-06-20
Do you need something like below?

If yes, you can change value of property "Description". This property is combine with name of elements in cut list. Here you have my code for this operation:

Option Explicit

    Dim swApp               As SldWorks.SldWorks

    Dim swModel             As SldWorks.ModelDoc2

    Dim swFeat              As SldWorks.Feature

    Dim swCustPropMgr       As SldWorks.CustomPropertyManager

    Dim strValue0           As String

    Dim strValue1           As String

    Dim strValue2           As String

    Dim bool                As Boolean

    Dim Name                As String

    Dim z, x                As Integer

    Dim boolstatus          As Boolean

   

Sub main()

   

    On Error Resume Next

     

    Set swApp = Application.SldWorks

    Set swModel = swApp.ActiveDoc

    Name = swModel.GetPathName

    Name = Dir(Name)

    Name = Left(Name, Len(Name) - 7)

    If Right(Name, 2) = "00" Then

    Name = Left(Name, Len(Name) - 2)

    Else: Name = Name & "."

    End If

    Set swFeat = swModel.FirstFeature

    z = 1

    x = 0

        Do While Not swFeat Is Nothing

            If swFeat.GetTypeName() = "CutListFolder" Then

                x = x + 1

            End If

        Set swFeat = swFeat.GetNextFeature

        Loop

    Set swFeat = swModel.FirstFeature

    If x > 1 Then

        Do While Not swFeat Is Nothing

            If swFeat.GetTypeName() = "CutListFolder" Then

           

                    Set swCustPropMgr = swFeat.CustomPropertyManager

                        If z < 10 Then

                        swCustPropMgr.Add3 "Description", 30, Name & "0" & z, 1

                        ElseIf z >= 10 Then

                        swCustPropMgr.Add3 "Description", 30, Name & z, 1

                        End If

     'for metalsheet                  

                        If UCase(swFeat.Name) Like "*SHEET*" Then

                        swCustPropMgr.Add3 "Description", 30, "Plate", 1

                        End If                     

                        z = z + 1

            End If

        Set swFeat = swFeat.GetNextFeature

        Loop

    End If

boolstatus = swModel.ForceRebuild3(True)

End Sub
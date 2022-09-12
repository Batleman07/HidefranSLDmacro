Function conexBda(propIndex As String) As String()
Const WM_QUIT = &H12
Dim xlApp As Excel.Application
Dim xlWB As Excel.Workbook
Dim intLinha As Integer
Dim intCol As Integer
Dim teste As String


Set xlApp = New Excel.Application
xlApp.Visible = False
Set xlWB = xlApp.Workbooks.Open("C:\Users\Projetos4\Desktop\ALESSANDRO\MACRO\Criando\BdaC_ProAtiv.xlsx")
intLinha = 2
intCol = 2
    
With xlWB.Worksheets(1)
    teste = Cells(intLinha, intCol).Text
     MsgBox teste
    While Cells(intLinha, intCol).Value <> ""
        teste = Cells(intLinha, intCol).Text
       
        If (teste = "BASE INFERIOR") Then
            teste = Cells(intLinha, intCol).Text
            MsgBox teste
            intLinha = intLinha + 1
        Else
            MsgBox "novo"
        End If
    Wend
End With
    
    
    Set xlWB = Nothing
    Set xlApp = Nothing
End Function
Attribute VB_Name = "Module1"
Sub Spinner2_Code()
Attribute Spinner2_Code.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveSheet.PivotTables("PivotTable13").PivotFields("Emp Name").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable13").PivotFields("Emp Name").PivotFilters. _
        Add2 Type:=xlTopCount, DataField:=ActiveSheet.PivotTables("PivotTable13"). _
        PivotFields("Total working Hours"), Value1:=Sheet6.Range("C2").Value
        
    Call FixBackground
End Sub
Sub Show_Chart1()
Attribute Show_Chart1.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveSheet.Shapes.Range(Array("Chart 1")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Picture 9")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Picture 11")).Visible = msoTrue
End Sub
Sub Hide_Chart1()

    ActiveSheet.Shapes.Range(Array("Chart 1")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Picture 11")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Picture 9")).Visible = msoTrue
End Sub
Sub Spinner8_Code()

    ActiveSheet.PivotTables("PivotTable22").PivotFields("Emp Name").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable22").PivotFields("Emp Name").PivotFilters. _
        Add2 Type:=xlTopCount, DataField:=ActiveSheet.PivotTables("PivotTable22"). _
        PivotFields("Active With Customer %"), Value1:=Sheet6.Range("C3").Value

    Call FixBackground
End Sub
Sub Show_Chart2()

    ActiveSheet.Shapes.Range(Array("Chart 22")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Picture 18")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Picture 19")).Visible = msoTrue
End Sub
Sub Hide_Chart2()

    ActiveSheet.Shapes.Range(Array("Chart 22")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Picture 19")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Picture 18")).Visible = msoTrue
End Sub

Sub FixBackground()
    Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("Dashboard")
        ' Colour only the body area, below the pivot headers
        .Range("A6:E100").Interior.Color = RGB(230, 230, 230)
    End With
    Application.ScreenUpdating = True
End Sub

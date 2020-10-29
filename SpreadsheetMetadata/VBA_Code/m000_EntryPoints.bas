Attribute VB_Name = "m000_EntryPoints"
Option Explicit


Sub AdjustNumberOfModelListRowsEntry()
'Adjusted the number of rows in the model table based on number of required rows.
'This sub is activated when item is selected in drop down

    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    AdjustNumberOfModelListRows
    Sheets("Model").ListObjects("tbl_Model").DataBodyRange.Cells(1).Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True


End Sub




Sub GenerateAllOutput()
'Loops through all items / scenarios for model and writes consolidated values to the output sheet

    Dim i As Integer
    Dim dblFirstEmptyCell As Double
    Dim sht As Worksheet
    Dim lo As ListObject
    
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    

    On Error Resume Next
    Sheets("Output").Delete
    On Error GoTo 0
    
    Set sht = Sheets.Add(Before:=Sheets("Validations"))
    sht.Name = "Output"
    Set lo = Sheets("Model").ListObjects("tbl_Model")
    lo.HeaderRowRange.Copy
    sht.Range("A1").PasteSpecial xlPasteValues
    
    'To be cautious in case model is not correctly refreshed before running
    AdjustNumberOfModelListRows
    Application.Calculate
    
    For i = 1 To Sheets("Inputs").ListObjects("tbl_Inputs").DataBodyRange.Rows.Count
        Sheets("Model").Range("ItemIndex").Value = i
        Application.StatusBar = "Writing " & Sheets("Model").Range("ItemName")
        AdjustNumberOfModelListRows
        Application.Calculate
        Do While Application.CalculationState = xlCalculating
        Loop
        Application.Wait (Now + TimeValue("0:00:01"))
        dblFirstEmptyCell = WorksheetFunction.CountA(Sheets("Output").Range("A:A")) + 1
        lo.DataBodyRange.Copy
        Sheets("Output").Cells(dblFirstEmptyCell, 1).PasteSpecial xlPasteValues
    Next i
    
    Application.StatusBar = False
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    


End Sub

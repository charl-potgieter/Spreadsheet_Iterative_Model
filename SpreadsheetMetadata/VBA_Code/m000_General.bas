Attribute VB_Name = "m000_General"
Option Explicit
Option Private Module



Sub AdjustNumberOfModelListRows()

    Dim shtModel As Worksheet
    Dim lo As ListObject
    Dim iNumOfRowsRequired As Integer
    Dim iNumOfRowsCurrent As Integer
    Dim iFirstRowToDelete As Integer
    Dim iLastRowToDelete As Integer

    
    Set shtModel = ThisWorkbook.Sheets("Model")
    Set lo = shtModel.ListObjects("tbl_Model")
    iNumOfRowsRequired = shtModel.Range("NumberOfModelRows").Value
    
    
    With lo
        iNumOfRowsCurrent = .DataBodyRange.Rows.Count
        
        If iNumOfRowsRequired < iNumOfRowsCurrent Then
            iFirstRowToDelete = .HeaderRowRange.Row + iNumOfRowsRequired + 1
            iLastRowToDelete = .HeaderRowRange.Row + iNumOfRowsCurrent
            Sheets("Model").Range(iFirstRowToDelete & ":" & iLastRowToDelete).EntireRow.Delete
        End If
        
        If iNumOfRowsRequired > iNumOfRowsCurrent Then
            lo.Resize lo.Range.Resize(iNumOfRowsRequired + 1, lo.Range.Columns.Count)
        End If
        
    End With

End Sub

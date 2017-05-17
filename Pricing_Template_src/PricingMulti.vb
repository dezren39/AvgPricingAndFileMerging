'This module contains the sub to prepare an array _
'for runAllProds, execute runAllProds, and clean-up the data _
'that was generated for review. This module also contains the _
'sub runAllProds, which executes EDIT_ACCESS_QUERY from _
'TRF_PricingLogic against each element in the array.

Sub Show_DB_Form()
    dbSelect.Show
End Sub

Sub PrepThenRunAllProdsThenCleanUp(db_name As String)
On Error GoTo Err1
    Application.ScreenUpdating = False
    
    'Initialize startRows, differences, myRow, and db_name.
    Const conPricingSheetStartRow As Integer = 7
    Const conRevisedSheetStartRow As Integer = 4
    Dim conStartRowsRevisedMinusPricing As Integer: conStartRowsRevisedMinusPricing = conRevisedSheetStartRow - conPricingSheetStartRow
    Dim myRow As Long: myRow = conPricingSheetStartRow
    Dim PricingFirstBlankRowMinusStartRow As Long
    Dim rowDiffPricingToRevised As Long
    
    'While myRow not blank, add one to myRow.'
    While ThisWorkbook.Sheets("Pricing").Cells(myRow, 1).Value <> ""
        myRow = myRow + 1
    Wend
    
    'Calculate difference between appropriate Pricing row and Revised Description row.
    PricingFirstBlankRowMinusStartRow = conPricingSheetStartRow - myRow
    rowDiffPricingToRevised = PricingFirstBlankRowMinusStartRow + conStartRowsRevisedMinusPricing
    
    'Call Module 1 -> runAllProds'
    runAllProds (db_name)
    
    'If runAllProds succeeded, do basic manipulations on the new records.
    'Only done for new records created by previous runAllProds call.
    ' \/ While Column and Row (master_description, myRow) isn't empty \/
    While ThisWorkbook.Sheets("Pricing").Cells(myRow, 1).Value <> ""
        'If Median isn't empy..
        If ThisWorkbook.Sheets("Pricing").Cells(myRow, 2).Value <> "" Then
            'If Count is greater than 10..
            If CLng(ThisWorkbook.Sheets("Pricing").Cells(myRow, 5).Value) > 10 Then
                'Copy Median to Suggested Price.
                ThisWorkbook.Sheets("Pricing").Range("B" & myRow).Copy _
                Destination:=ThisWorkbook.Sheets("Pricing").Range("G" & myRow)
            End If
        
            'Overwrite Column D, STD, with Median / # of Units. ($30 Median, 3 pack = $10)
            'If pack, count, or ml.
            Dim lowerDesc As String: lowerDesc = LCase(ThisWorkbook.Sheets("Pricing").Cells(myRow, 1).Value)
            
            If lowerDesc Like "* pack" Or lowerDesc Like "* count" Or _
                lowerDesc Like "* ml" Or lowerDesc Like "* oz" Then
                    'Split description by spaces.
                    Dim desc: desc = Split(ThisWorkbook.Sheets("Pricing").Cells(myRow, 1).Value, " ")
                    
                    'If second to last split is a number. ("....4 Pack" Yes, "...per mL" No)
                    If IsNumeric(desc(UBound(desc) - 1)) Then
                        'Copy Median/Number to Column D.
                        ThisWorkbook.Sheets("Pricing").Cells(myRow, 4).Value = _
                            Round(ThisWorkbook.Sheets("Pricing").Cells(myRow, 2).Value / desc(UBound(desc) - 1), 2)
                    'Else blank out the Standard Deviation
                    Else
                        ThisWorkbook.Sheets("Pricing").Cells(myRow, 4).Value = ""
                    End If
            'If single.
            ElseIf LCase(ThisWorkbook.Sheets("Pricing").Cells(myRow, 1).Value) Like "* single" Then
                'Copy Median to Column D.
                ThisWorkbook.Sheets("Pricing").Cells(myRow, 4).Value = _
                    ThisWorkbook.Sheets("Pricing").Cells(myRow, 2).Value
            'Else blank out the Standard Deviation
            Else
                ThisWorkbook.Sheets("Pricing").Cells(myRow, 4).Value = ""
            End If
        End If
             
        'If relative line in REVISED DESCRIPTION suggested_retail_price column is not empty.
        If ThisWorkbook.Sheets("REVISED DESCRIPTION").Cells((myRow + rowDiffPricingToRevised), 17).Value <> "" Then
                'Copy Revised Description.suggested_retail_price to Pricing.Original Suggested Price
                ThisWorkbook.Sheets("Pricing").Cells(myRow, 8).Value = _
                    ThisWorkbook.Sheets("REVISED DESCRIPTION").Cells((myRow + rowDiffPricingToRevised), 17).Value
        End If
        
        'Move to next row.
        myRow = myRow + 1
    Wend

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & Err.Description
    Resume Done
End Sub

Sub runAllProds(file As String)
On Error GoTo Err1

    Application.ScreenUpdating = False
    
    'Initialize i, startRow, worksheet var, and an array based on amount of columns.
    Dim i As Integer, startRow As Integer, iRow As Integer: startRow = 4: iRow = startRow
    Dim rd As Worksheet: Set rd = ThisWorkbook.Sheets("REVISED DESCRIPTION")
    ReDim rng(rd.Range("F" & Rows.Count).End(xlUp).row) As String

    'From F4 down, until you encounter a blank cell, record the string in the cell to the array.'
    While rd.Cells(iRow, 6).Value <> ""
        Dim currentRow As String: currentRow = rd.Cells(iRow, 6).Value
        
        'If there is a single apostrophe, escape it by using double apostrophe.
        If InStr(currentRow, "'") <> 0 Then
            currentRow = Replace(currentRow, "'", "''")
        End If
        
        rng(iRow - startRow) = currentRow
        iRow = iRow + 1
    Wend

    'Call all the prods.'
    For i = 0 To iRow - startRow
        Call EDIT_ACCESS_QUERY(file, rng(i), "", "", "", "", False, True)
    Next

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & Err.Description
    Resume Done
End Sub

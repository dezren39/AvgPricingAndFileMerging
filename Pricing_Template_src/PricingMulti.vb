Option Explicit

'This module contains the sub to prepare an array _
for runAllProds, execute runAllProds, and clean-up the data _
that was generated for review. This module also contains the _
sub runAllProds, which executes EDIT_ACCESS_QUERY from _
TRF_PricingLogic against each element in the array.

Sub PrepThenRunAllProdsThenCleanUp(db_name As String)
    On Error GoTo Err1
    Application.ScreenUpdating = False
    
    Dim myRow As Long
    Dim PricingFirstBlankRowMinusStartRow As Long
    Dim rowDiffPricingToRevised As Long
    Dim lowerDesc As String, desc() As String
    
    myRow = conPricingSheetStartRow
    
    With ThisWorkbook.Sheets(conPricingSheetName)
        'While myRow not blank, add one to myRow.'
        While .Cells(myRow, conPricingSheetMDescColumn).Value <> ""
            myRow = myRow + 1
        Wend
        
        'Calc diff between appropriate Pricing row and Revised Description row.
        PricingFirstBlankRowMinusStartRow = conPricingSheetStartRow - myRow
        rowDiffPricingToRevised = PricingFirstBlankRowMinusStartRow + _
                                                conStartRowsRevisedMinusPricing
        
        'Call -> runAllProds'
        RunAllProds (db_name)
        
        'If runAllProds succeeded, do basic manipulations on the new records.
        'Only done for new records created by previous runAllProds call.
        ' \/ While Column and Row (master_description, myRow) isn't empty \/
        While .Cells(myRow, conPricingSheetMDescColumn).Value <> ""
            'If Median isn't empy..
            If .Cells(myRow, conPricingSheetMedianColumn).Value <> "" Then
                'If Count is greater than 10..
                If CLng(.Cells(myRow, _
                                conPricingSheetCountColumn).Value) > 10 Then
                
                    'Copy Median to Suggested Price.
                    .Cells(myRow, conPricingSheetMedianColumn).Copy _
                    Destination:=.Cells(myRow, conPricingSheetSuggPriceColumn)
                End If
            
                'Overwrite Column D, STD, with Median / # of Units.
                'If pack, count, or ml. ($30 Median, 3 pack = $10)
                lowerDesc = LCase(.Cells(myRow, _
                                            conPricingSheetMDescColumn).Value)
                
                If lowerDesc Like "* pack" Or lowerDesc Like "* count" Or _
                    lowerDesc Like "* ml" Or lowerDesc Like "* oz" Then
                        
                        'Split description by spaces.
                        desc = Split(.Cells(myRow, _
                                            conPricingSheetMDescColumn).Value _
                                            , " ")
                        
                        'If second to last split is a number.
                        '("....4 Pack" Yes, "...per mL" No)
                        If IsNumeric(desc(UBound(desc) - 1)) Then
                            'Copy Median/Number to Column D.
                            .Cells(myRow, _
                                    conPricingSheetPerCountColumn).Value = _
                                Round(.Cells(myRow, _
                                        conPricingSheetMedianColumn).Value _
                                                / desc(UBound(desc) - 1), 2)
                        'Else blank out the Standard Deviation
                        Else
                            .Cells(myRow, _
                                    conPricingSheetPerCountColumn).Value = ""
                        End If
                
                'If single.
                ElseIf LCase(.Cells(myRow, _
                        conPricingSheetMDescColumn).Value) Like "* single" Then
                
                    'Copy Median to Column D.
                    .Cells(myRow, conPricingSheetPerCountColumn).Value = _
                        .Cells(myRow, conPricingSheetMedianColumn).Value
                
                Else 'Else blank out the Standard Deviation
                    .Cells(myRow, conPricingSheetPerCountColumn).Value = ""
                End If
            End If
                 
            'If relative line in REVISED DESCRIPTION suggested_retail_price column is not empty.
            If ThisWorkbook.Sheets(conRevisedSheetName).Cells((myRow + rowDiffPricingToRevised), _
                            conRevisedSheetSuggPriceColumn).Value <> "" Then
                    
                    'Copy Revised Description.suggested_retail_price to Pricing.Original Suggested Price
                    .Cells(myRow, conPricingSheetOSuggPriceColumn).Value = _
                        ThisWorkbook.Sheets(conRevisedSheetName).Cells((myRow + rowDiffPricingToRevised), _
                                        conRevisedSheetSuggPriceColumn).Value
            End If
            
            'Move to next row.
            myRow = myRow + 1
        Wend
    End With
    
Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & err.Description
    Resume Done
End Sub

Sub RunAllProds(file As String)
    On Error GoTo Err1
    Application.ScreenUpdating = False

    Dim rd As Worksheet, iRow As Integer
    Dim currentRow As String, i As Integer
    
    Set rd = ThisWorkbook.Sheets(conRevisedSheetName)
    
    With rd
        ReDim rng(.Cells(.Rows.Count, _
                        conRevisedSheetMDescColumn).End(xlUp).Row) As String
        
        iRow = conRevisedSheetStartRow
        
        'From F4 down, until you encounter a blank cell,
        'record the string in the cell to the array.
        While .Cells(iRow, conRevisedSheetMDescColumn).Value <> ""
            currentRow = .Cells(iRow, conRevisedSheetMDescColumn).Value
            
            'If there is a single apostrophe,
            If InStr(currentRow, "'") <> 0 Then
                'escape it by using double apostrophe.
                currentRow = Replace(currentRow, "'", "''")
            End If
            
            rng(iRow - conRevisedSheetStartRow) = currentRow
            iRow = iRow + 1
        Wend
    End With
    
    'Call all the prods.'
    For i = 0 To iRow - conRevisedSheetStartRow
        Call EDIT_ACCESS_QUERY(file, rng(i), "", "", "", "", False, True)
    Next

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & err.Description
    Resume Done
End Sub







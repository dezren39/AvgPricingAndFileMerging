'This module copies final suggested prices from the pricing sheet to the revised sheet.

Sub Finalize_Pricing()
On Error GoTo Err1

  If (vbYes = MsgBox("CAUTION: Are you sure you wish to finalize?" + vbCrLf + _
        "This will overwrite 'Suggested Pricing' values in the 'Revised Description' sheet." + vbCrLf + _
        "Non-empty prices are copied to the first matching 'master_description'." _
        , vbYesNo, "PLEASE CONFIRM TO CONTINUE:")) Then
    Const conPricingSheetStartRow As Integer = 7
    Const conRevisedSheetStartRow As Integer = 4
    Const conStartRowsRevisedMinusPricing As Integer = conRevisedSheetStartRow - conPricingSheetStartRow
    Dim startTime As Single: startTime = Timer()
    Dim pricingRow As Long: pricingRow = conPricingSheetStartRow
    Dim revisedRow As Long
    Dim rd As Worksheet: Set rd = ThisWorkbook.Sheets("REVISED DESCRIPTION")
    Dim pr As Worksheet: Set pr = ThisWorkbook.Sheets("Pricing")
    While pr.Cells(pricingRow, 1).Value <> ""
        revisedRow = conRevisedSheetStartRow
        While rd.Cells(revisedRow, 6).Value <> pr.Cells(pricingRow, 1) And rd.Cells(revisedRow, 6).Value <> ""
            revisedRow = revisedRow + 1
        Wend
        If rd.Cells(revisedRow, 6).Value <> "" Then
            If pr.Cells(pricingRow, 7).Value <> "" Then
                rd.Cells(revisedRow, 17).Value = pr.Cells(pricingRow, 7).Value
                rd.Cells(revisedRow, 17).Font.Bold = True
            End If
        End If
        pricingRow = pricingRow + 1
    Wend
    MsgBox "Finalized in: " + CStr(Timer() - startTime) + " seconds." + vbCrLf + "Be sure to double-check before submitting!"
  End If

Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & Err.Description
    Resume Done
End Sub

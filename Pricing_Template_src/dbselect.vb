'This form contains the 2 functions:
' UserForm_Initialize configures the combo box on formload.
' get_data_btn_click is ran when the run query button is pressed.
'    PrepThenRunAllProdsThenCleanUp is executed. This code _
'     is ran to prepare for running MTRF_PricingLogics's runAllProds sub.
'    runAllProds is executed, averaging pricing for each item on the pricing sheet.
'    code is ran after runAllProds to copy specific values if they meet certain requirements.

Private Sub UserForm_Initialize()
On Error GoTo Err1
    'Populate Combo Box.'
    DB_Combo_Box.AddItem "Client Connection"
    DB_Combo_Box.AddItem "Last Dose"
    DB_Combo_Box.AddItem "GFK Nutritional"
    DB_Combo_Box.AddItem "GFK Pharma"
    DB_Combo_Box.Value = "Client Connection"
Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & Err.Description
    Resume Done
End Sub

Private Sub get_data_btn_Click()
On Error GoTo Err1
    'Hide this form, turn off screen updating.'
    dbSelect.Hide
    Application.ScreenUpdating = False
    Dim startTime As Single: startTime = Timer()
    Dim db_name As String
    'Match Combo Box items to DB names. Assign one to db_name'
    If DB_Combo_Box.Value = "Client Connection" Then db_name = "CC_P2.mdb"
    If DB_Combo_Box.Value = "Last Dose" Then db_name = "Last_Dose.mdb"
    If DB_Combo_Box.Value = "GFK Pharma" Then db_name = "Pharma.mdb"
    If DB_Combo_Box.Value = "GFK Nutritional" Then
        db_name = "Nutritional.mdb"
        nutria = True
    End If

    PrepThenRunAllProdsThenCleanUp db_name

    'Turn screen updating back on, then unload this form.'
    ThisWorkbook.Sheets("Pricing").Activate
    ThisWorkbook.Sheets("Pricing").Cells(2, 2).Select
    Application.ScreenUpdating = True
    MsgBox "Completed pricing in: " + CStr(Timer() - startTime) + " seconds." + vbCrLf + "Now it's time to evaluate the results." _
            + vbCrLf + "Use good judgement and don't forget to finalize!"
    Unload Me
Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & Err.Description
    Resume Done
End Sub


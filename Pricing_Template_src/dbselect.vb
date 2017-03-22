Private Sub UserForm_Initialize()
    'Populate Combo Box.'
    DB_Combo_Box.AddItem "Client Connection"
    DB_Combo_Box.AddItem "Last Dose"
    DB_Combo_Box.AddItem "GFK Nutritional"
    DB_Combo_Box.AddItem "GFK Pharma"
    DB_Combo_Box.Value = "Client Connection"
End Sub

Private Sub get_data_btn_Click()
    'Init db_name var, hide this form, then turn off screen updating.'
    Dim db_name As String
    dbSelect.Hide
    Application.ScreenUpdating = False
    
    'Match Combo Box items to DB names.'
    If DB_Combo_Box.Value = "Client Connection" Then db_name = "CC_P2.mdb"
    If DB_Combo_Box.Value = "Last Dose" Then db_name = "Last_Dose.mdb"
    If DB_Combo_Box.Value = "GFK Pharma" Then db_name = "Pharma.mdb"
    If DB_Combo_Box.Value = "GFK Nutritional" Then
        db_name = "Nutritional.mdb"
        nutria = True
    End If
    
    'Call Module 1 -> runAllProds'
    runAllProds (db_name)
    
    'Turn screen updating back on, then unload this form.'
    Application.ScreenUpdating = True
    Unload Me
End Sub


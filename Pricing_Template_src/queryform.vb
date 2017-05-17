'This form contains the text boxes and triggers to manually _
run a single product for pricing.

Private Sub get_data_btn_Click()
On Error GoTo Err1
    'Init all vars, Assign Unset to nutria flag, Assign name textbox to prod_name after trimming whitespace.'
    Dim db_name As String
    Dim prod_name As String: prod_name = LTrim(RTrim(Name_Text.Value))
    Dim qty As String
    Dim max As String
    Dim min As String
    Dim contains As String
    Dim nutria As Boolean: nutria = False

    'Either select only 1, or any value not = to 1.'
    If qty_one_btn.Value = True Then qty = " AND ((All_qa.[Change Quantity]) = 1) "
    If qty_more_btn.Value = True Then qty = " AND ((All_qa.[Change Quantity]) <> 1) "

    'If there is a min or max value in respective textbox, add that to the query.'
    If Max_Price_Text.Value = True Then max = " AND ((All_qa.[Average Price_Per])<" & Max_Price_Text.Value & ") "
    If Min_Price_Text.Value = True Then min = " AND ((All_qa.[Average Price_Per])>" & Min_Price_Text.Value & ") "

    'If there is a specific clinic description in respective textbox, add that constraint to query.'
    'Original, commented out method: If Clinic Description exactly equals the text of the respective textbox.
    'New, currently active mthod: If Clinic Description contains the text of the respective textbox.
    If contains_Text.Value = True Then contains = " AND ((All_qa.[Clinic Description]) LIKE '%" & LTrim(RTrim(contains_Text.Value)) & "%')" '(InStr(1,All_qa.[Clinic Description],'" & LTrim(RTrim(contains_Text.Value)) & "') > 0) "

    'Match Combo Box items to DB names.'
    If DB_Combo_Box.Value = "Client Connection" Then db_name = "CC_P2.mdb"
    If DB_Combo_Box.Value = "Last Dose" Then db_name = "Last_Dose.mdb"
    If DB_Combo_Box.Value = "GFK Pharma" Then db_name = "Pharma.mdb"
    If DB_Combo_Box.Value = "GFK Nutritional" Then
       db_name = "Nutritional.mdb"
       nutria = True
    End If
    
    'Run Calculations.'
    Call EDIT_ACCESS_QUERY(db_name, prod_name, qty, max, min, contains, nutria, False)
    
    'Unload This Form.'
    Unload Me
Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & Err.Description
    Resume Done
End Sub

'Clear all values in Textboxes and the Combobox.'
Private Sub Clear_Btn_Click()
    Name_Text.Value = ""
    Max_Price_Text.Value = ""
    Min_Price_Text.Value = ""
    DB_Combo_Box.Value = ""
End Sub

Private Sub UserForm_Initialize()
On Error GoTo Err1
    'Populate Combo Box
    DB_Combo_Box.AddItem "Client Connection"
    DB_Combo_Box.AddItem "GFK Pharma"
    DB_Combo_Box.AddItem "GFK Nutritional"
    DB_Combo_Box.AddItem "Last Dose"
    DB_Combo_Box.Value = "Client Connection"
Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & Err.Description
    Resume Done
End Sub

Public Sub DB_Combo_Box_Change()
On Error GoTo Err1
    'Init db_name var and match Combo Box items to DB names.'
    Dim db_name As String
    If DB_Combo_Box.Value = "Client Connection" Then db_name = "CC_P2.mdb"
    If DB_Combo_Box.Value = "Last Dose" Then db_name = "Last_Dose.mdb"
    If DB_Combo_Box.Value = "GFK Nutritional" Then db_name = "Nutritional.mdb"
    
    'If DB Name is unavailable, inform user.'
    If Dir(ThisWorkbook.Path & "\" & db_name) = "" Then MsgBox (db_name & " - This database is currently unavalible.")
Done:
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "The following error occurred: " & Err.Description
    Resume Done
End Sub




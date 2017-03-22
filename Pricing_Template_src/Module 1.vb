Sub runAllProds(file As String)
    'Dim array of string to contain list of prod descriptions to be autoran.'
    Dim rng(100) As String, i As Integer, row As Integer: row = 4
    
    'Activate Sheet containing list of prod classes.'
    Sheets("REVISED DESCRIPTION").Activate
    
    'From F4 down, until you encounter a blank cell, record the string in the cell to the array.'
    While Cells(row, 6).Value <> ""
        rng(row - 4) = Cells(row, 6).Value
        row = row + 1
    Wend
    
    'Call all the prods.'
    For i = 0 To row - 4
        Call EDIT_ACCESS_QUERY(file, rng(i), "", "", "", "", False, True)
    Next
End Sub

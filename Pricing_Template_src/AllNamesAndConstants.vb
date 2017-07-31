Option Explicit

Public Const conRevisedSheetName As String = "REVISED DESCRIPTION"
Public Const conRevisedSheetStartRow As Integer = 2
Public Const conRevisedSheetMDescColumn As Integer = 6
Public Const conRevisedSheetSuggPriceColumn As Integer = 17

'Pricing functions rely on difference between MDesc & Count Columns (1:5)
'Cannot move the relative location of these columns without refactor.
'Absolute location can easily change so long as (Count - MDesc) = 4
'(the columns between mdesc and count are also used by pricing functions.)
Public Const conPricingSheetName As String = "Pricing"
Public Const conPricingSheetStartRow As Integer = 7
Public Const conPricingSheetMDescColumn As Integer = 1 'MDesc
Public Const conPricingSheetMedianColumn As Integer = 2
Public Const conPricingSheetPerCountColumn As Integer = 4
Public Const conPricingSheetCountColumn As Integer = 5 'Count
Public Const conPricingSheetSuggPriceColumn As Integer = 7
Public Const conPricingSheetOSuggPriceColumn As Integer = 8

Public Const conStartRowsRevisedMinusPricing As Integer = _
                        conRevisedSheetStartRow - conPricingSheetStartRow

Function GetSentinelColName(Col As Collection) As String
    'This column must be found in each sheet that is merged with the Import button.
    'The import step copies all rows within the import sheet from the start until _
    the last row containing the sentinel column. Any data after the last row in the
    'sentinel column will not be copied. Any data on the originating workbook that
    'is after the last row in it's respective sentinel column runs the risk
    'of being overwritten by incoming import data.

    Const sentinelA As String = "category_id (see drop down)"
    Const sentinelB As String = "category_id"
    
    If Contains(Col, sentinelA) Then
        GetSentinelColName = sentinelA
    ElseIf Contains(Col, sentinelB) Then
        GetSentinelColName = sentinelB
    Else
        GetSentinelColName = vbNull
    End If
End Function

Function RetrieveImportSheetsAsVariant() As Variant()
    RetrieveImportSheetsAsVariant = Array("NEW", _
                                            conRevisedSheetName, _
                                            "OFF THE MARKET-DO NOT USE")
End Function


Function CollectResearchColumns() As Collection
    'Not Used, Could be used to pull specific column #s for data manipulation
    'ex: CId_Collection(THIS_COLLECTION_RETURNED_AS_VARIABLE("m_desc"))
    'will pull the column ID which has "master_data_set_id" as a value.
    '"ex: If mdesc contains "pack" and uom not = pack then...
    
    Dim colColl As New Collection
    
    colColl.Add Item:="OTM", key:="otm"
    colColl.Add Item:="DNU", key:="dnu"
    colColl.Add Item:="product_id", key:="prod_id"
    colColl.Add Item:="master_data_set_id", key:="mds_id"
    colColl.Add Item:="category_id (see drop down)", key:="cat_id"
    colColl.Add Item:="manufacturer", key:="mfr"
    colColl.Add Item:="brand", key:="brand"
    colColl.Add Item:="master_ description", key:="m_desc"
    colColl.Add Item:="upc", key:="upc"
    colColl.Add Item:="species", key:="species"
    colColl.Add Item:="life_stage (Nutrition Only)", key:="life_stage"
    colColl.Add Item:="items_per_pkg", key:="items_per"
    colColl.Add Item:="item_size  (number only)", key:="item_size"
    colColl.Add Item:="unit_of_measure", key:="uom"
    colColl.Add Item:="Bucket_Type", key:="bucket"
    colColl.Add Item:="Package_Type", key:="pack_type"
    colColl.Add Item:="Formula_Adjective_1", key:="form1"
    colColl.Add Item:="Formula_Adjective_2", key:="form2"
    colColl.Add Item:="Suggested_Retail_Price (Pharma Only)", _
                        key:="sugg_price"
    colColl.Add Item:="Revised description as it will appear in PIT" & _
                                    "     (Use only when revised.)", _
                        key:="rev_desc"
    colColl.Add Item:="Revised UPC as it will appear", key:="rev_upc"
    colColl.Add Item:="Blank", key:="blank"
    colColl.Add Item:="Comments", key:="comments"
    colColl.Add Item:="Mfg Website", key:="Mfg Website"
    colColl.Add Item:="website #1", key:="web1"
    colColl.Add Item:="website #2", key:="web2"
    colColl.Add Item:="website #3 (add more after 'status' if needed)", _
                        key:="web3"
    colColl.Add Item:="CREATE_DATETIME", key:="createTime"
    colColl.Add Item:="OFF_THE_MARKET_DATETIME", key:="otmTime"
    colColl.Add Item:="DO_NOT_USE_DATETIME", key:="dnuTime"
    colColl.Add Item:="STATUS", key:="stat"
    
    Set CollectResearchColumns = colColl
End Function





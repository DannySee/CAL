Attribute VB_Name = "btn_Item_Lookup"

'Declare private module constants
Private Const varShp As Variant = Array("Item_Lookup_Pane", "Item_Lookup_MPC", _
    "Item_Lookup_GTIN", "Item_Lookup_Search", "Item_Lookup_Cancel", _
    "Item_Lookup_List")
Private Const varHeaders As Variant = Array("SUPC", "PACK/SIZE", "BRAND", _
    "DESCRIPTION", "MPC", "GTIN")


'*******************************************************************************
'Show all utility elements, update and resize listbox.
'*******************************************************************************
Private Sub Item_Lookup_Initialize()

    'Hide any visible shapes
    Utility.ClearShapes

    'Show utility elements
    Utility.Show(varShp)
End Sub


'*******************************************************************************
'Pull in selected customer OpCo data. Output formatted report on new workbook.
'*******************************************************************************
Private Sub Item_Lookup_Select()

    'Declare sub variables
    Dim strMPC As String
    Dim strGTIN As String

    'Get user input in SQL syntax (quote/comma delimited)
    strMPC = Utility.GetItmSearch("P")
    strGTIN = Utility.GetItmSearch("Q")

    'If there is user input
    If strMPC <> strGTIN Then

        'Create new workbook for report
        Utility.CreateWorkbook("Report")

        'Add report headers to workbook
        Utility.AddHeaders(varHeaders)

        'Search for SUPC and paste query results to A2
        Range("A2").CopyFromRecordset Pull.GetSUPC(strGTIN, strMPC)

        'If query returned results
        If Range("A2").Value <> "" Then

            'Format report
            Utility.AddBorders

        'If query did not return results
        Else

            'Alert user of no results & Close workbook
            ActiveWorkbook.Close SaveChanges:=False
            MsgBox "No items were found."
        End If

        'Clear Control Panel shapes
        Utility.ClearShapes

    'If missing data
    Else

        'Alert user of missing data
        msgbox "You must enter at least one GTIN/MPC to search."
    End if
End Sub


'*******************************************************************************
'Unhide list columns.
'*******************************************************************************
Sub Item_Lookup_List()

    'Unlock sheet
    Utility.ShtUnlock("Control Panel")

    'Focus on control panel sheet
    With Sheets("Control Panel")
        .Shapes("Item_Lookup_List").Visible = False
        .Columns("P:Q").Hidden = False
        .Columns("P:Q").Locked = False
    End With

    'Lock Sheet
    Utility.ShtLock("Control Panel")
End Sub

Attribute VB_Name = "btn_Pull_OpCo_List"

'Declare private module constants
Private Const varShp As Variant = _Array("Listbox_Pane", _"Multiuse_Listbox", _
    "Listbox_Cancel","Listbox_Select")


'*******************************************************************************
'Show all utility elements, update and resize listbox.
'*******************************************************************************
Private Sub Pull_OpCo_List_Initialize()

    'Hide any visible shapes
    Utility.ClearShapes

    'Set default listbox view
    Utility.ListboxByCst(True)

    'Show utility elements
    Utility.Show(varShp)

    'Assign Select button to correct routine
    Sheets("Control Panel").Shapes("Listbox_Select").OnAction = _
        "Pull_Unassigned_Customers_Select"
End Sub


'*******************************************************************************
'Pull in selected customer OpCo data. Output formatted report on new workbook.
'*******************************************************************************
Private Sub Pull_OpCo_List_Select()

    'Declare sub variables
    Dim varCst As Variant
    Dim strCst As String

    'Get Select customer
    varCst = Utility.GetSelection

    'If a customer was selected from list
    If Not IsEmpty(varCst) Then

        'Get string of selected customer (comma and quote delimited)
        strCst = GetStr(varCst, True)

        'Update/insert all new recordset






        'Add/remove customers from DropDowns sheet
        Utility.ReviseDropDwns(varCst)

        'Populate CAL sheets with refreshed data
        Utility.PopulatePages

        'Save all sheet data set to static dictionary
        Set tempDct = oPrgms.GetSaveData(True)
        Set tempDct = oCst.GetSaveData(True)
        Set tempDct = oDev.GetSaveData(True)

        'Clear utility shapes
        Utility.ClearShapes

    'If no customers were selected from list
    Else

        'Alert user and exit sub if no customers were selected
        msgbox "You must make at least one selection."
    End If
End Sub

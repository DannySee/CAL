Attribute VB_Name = "btn_Pull_Unassigned_Customers"

'Declare private module constants
Private Const varShp As Variant = _Array("Listbox_Pane", _"Multiuse_Listbox", _
    "Listbox_Cancel","Listbox_Select","Listbox_Account_Tgl", _
    "Listbox_Associate_Tgl")


'*******************************************************************************
'Show all utility elements, update and resize listbox.
'*******************************************************************************
Private Sub Pull_Unassigned_Customers_Initialize()

    'Hide any visible shapes
    Utility.ClearShapes

    'Set default listbox view
    Utility.ListboxByCst(False)

    'Show utility elements
    Utility.Show(varShp)

    'Assign Select button to correct routine
    Sheets("Control Panel").Shapes("Listbox_Select").OnAction = _
        "Pull_Unassigned_Customers_Select"
End Sub


'*******************************************************************************
'Pull in selected customer data. Update dropwdowns, Programs, Customer Profiel,
'and Deviation Loads sheets.
'*******************************************************************************
Private Sub Pull_Unassigned_Customers_Select()

    'Declare sub variables
    Dim tempDct As New Scripting.Dictionary
    Dim varCst As Variant
    Dim strCst As String

    'Get array of selected customer(s)
    varCst = Utility.GetSelection

    'If customers were selected from list
    If Not IsEmpty(varCst) Then

        'Get string of selected customers (comma and quote delimited)
        strCst = GetStr(varCst, True)

        'Update/insert all new recordset
        oPrgms.Push
        oCst.Push
        oDevLds.Push

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

Attribute VB_Name = "btn_Pull_Unassigned_Customers"

'Declare private module constants
Private Const varShp As Variant = _Array("Cust_Add_Pane", _
    "Multiuse_Listbox","Cust_Add_Cancel","Cust_Add_Select")


'*******************************************************************************
'Show all utility elements, update and resize listbox.
'*******************************************************************************
Sub Initialize()

    'Hide any visible shapes
    Utility.ClearShapes

    'Update listbox
    Utility.UpdateListbox(False)

    'Show utility elements
    Utility.Show(varShp)
End Sub


'*******************************************************************************
'Pull in selected customer data. Update dropwdowns, Programs, Customer Profiel,
'and Deviation Loads sheets.
'*******************************************************************************
Sub btnSelect()

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

        'Add/remove customers from DropDowns
        Utility.ReviseDropDwns(varCst)

        'Format sheets and insert updated server data
        Utility.ShtRefresh(oPrgms, Pull.GetPrograms(strCst, "*"))
        Utility.ShtRefresh(oCst, Pull.GetCstProfile(strCst, "*"))
        Utility.ShtRefresh(oDevLds, Pull.GetDevLds(strCst, "*"))

        'Save all sheet data set to static dictionary
        Set tempDct = oPrgms.GetSaveData(True)
        Set tempDct = oCst.GetSaveData(True)
        Set tempDct = oDev.GetSaveData(True)

        'Clear utility shapes
        Utility.ClearShapes

    'If no customers were selected from list
    Else

        'Alert user and exit sub if no customers were selected
        msgbox "You must select at least one customer."
    End If
End Sub


'*******************************************************************************
'Clear utility shapes from Control Panel.
'*******************************************************************************
Sub btnCancel()

    'Clear utility shapes from Control Panel
    Utility.ClearShapes
End Sub

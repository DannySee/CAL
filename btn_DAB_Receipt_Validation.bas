Attribute VB_Name = "btn_DAB_Receipt_Validation"

'*******************************************************************************
'Show all utility elemenst, pdate and resize listbox. Testing
'*******************************************************************************
Sub Initialize()

    'Hide any visible shapes
    Utility.ClearShapes

    'Update listbox
    Utility.UpdateListbox(False)

    'Show utility elements
    Utility.Show(oBtnPull.GetShapes)

End Sub


'*******************************************************************************
'Pull in selected customer data. Update dropwdowns, Programs, Customer Profiel,
'and Deviation Loads sheets.
'*******************************************************************************
Sub SelectCst()

    'Declare sub variables
    Dim tempDct As New Scripting.Dictionary
    Dim varCst As Variant
    Dim strCst As String

    'Get array of selected customer(s)
    varCst = Utility.GetSelection

    'Alert user and exit sub if no customers were selected
    If IsEmpty(varCst) Then
        msgbox "You must select at least one customer."
        Goto NoCst
    End If

    'Get string of selected customers (comma and quote delimited)
    strCst = GetStr(varCst, True)

    'Update/insert all new recordset
    oPrgms.Push
    oCst.Push
    oDevLds.Push

    'Add/remove customers from DropDowns
    Format.ReviseDropDwns(varCst)

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Format sheets and insert updated server data
    Format.ShtRefresh(oPrgms, Pull.GetPrograms(strCst))
    Format.ShtRefresh(oCst, Pull.GetCstProfile(strCst))
    Format.ShtRefresh(oDevLds, Pull.GetDevLds(strCst))

    'Close connection
    cnn.Close

    'Save all sheet data set to static dictionary
    Set tempDct = oPrgms.GetSaveData(True)
    Set tempDct = oCst.GetSaveData(True)
    Set tempDct = oDev.GetSaveData(True)

    'Clear utility shapes
    Utility.ClearShapes

'Label to alert user of missing selection
NoCst:

    'Free objects
    Set cnn = Nothing
    Set rst = Nothing
End Sub


'*******************************************************************************
'Clear utility shapes from Control Panel.
'*******************************************************************************
Sub Cancel()

    'Clear utility shapes from Control Panel
    Utility.ClearShapes
End Sub

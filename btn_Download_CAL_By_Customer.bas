Attribute VB_Name = "btn_Download_CAL_By_Customer"

'Declare private module constants TESTING3
Private Const varShp As Variant = _Array("Listbox_Pane", _"Multiuse_Listbox", _
    "Listbox_Cancel","Listbox_Select","Listbox_All")


'*******************************************************************************
'Show all utility elements, update and resize listbox.
'*******************************************************************************
Private Sub Download_CAL_Initialize()

    'Hide any visible shapes
    Utility.ClearShapes

    'Set default listbox view
    Utility.ListboxByCst(False)

    'Show utility elements
    Utility.Show(varShp)

    'Assign Select button to correct routine
    Sheets("Control Panel").Shapes("Listbox_Select").OnAction = _
        "Download_CAL_Select"
End Sub


'*******************************************************************************
'Download CAL for account selections.
'*******************************************************************************
Sub Download_CAL_Select()

    'Declare sub variables
    Dim varCst As Variant
    Dim strCst As String
    Dim strPth As String
    Dim wb As Workbook

    'Get folder path from user selection
    strPth = Utility.SelectFolder

    'Get array of selected customer(s)
    varCst = Utility.GetSelection

    'Alert user and exit sub if no assigned customers/selected folder
    If Not IsEmpty(varCst) And strPth <> "" Then

        'Loop through all assigned customers
        For each cst In varCst

            'Create customer friendly CAL Workbook and set to variable
            Set wb = Utility.DwnCstCAL(cst)

            'Save and close workbook
            wb.Close SaveChanges:=True, Filename:= _
                strPth & cst & " CUSTOMER AGREEMENT LIST.xlsx"
        Next

        'Clear utility shapes
        Utility.ClearShapes

    'If no customers were selected from list
    Else

        'Alert user and exit sub if no customers were selected
        msgbox "You must make at least one selection."
    End If
End Sub

Attribute VB_Name = "btn_Listbox"


'*******************************************************************************
'Select all elements from multiuse listbox
'*******************************************************************************
Sub Listbox_All_Initialize()

    'Declare sub variables
    Dim blToggle As Boolean

    'Determine if button is toggled on/off
    blToggle = Utility.IsToggle("Listbox_All_Tgl")

    'Actavte object shape
    With Sheets("Control Panel").Shapes("Listbox_All_Tgl")

        'If button is toggled on
        If blToggle Then
            .Fill.ForeColor.RGB = RGB(89, 89, 89)

        'If button is toggled off
        Else
            .Fill.ForeColor.RGB = RGB(64, 64, 64)
        End If
    End With

    'Loop through all dropdowns55to create string
    With Sheets("Control Panel").Multiuse_Listbox

        'Loop through list and toggle on/off according to boolean operator
        For i = 0 To .ListCount - 1
            .Selected(i) = blToggle
        Next
    End With
End Sub


'*******************************************************************************
'Update multiuse listbox with list of customers. Boolean operator indicates
'if listbox should conatain assigned customers or unassigned customers.
'*******************************************************************************
Sub Listbox_Customer_Initialize()

    'Setup listbox to display customers (always unassigned)
    Utility.ListboxByCst(False)
End Sub


'*******************************************************************************
'Update multiuse listbox with list of account holders. Boolean operator
'indicates if listbox should conatain assigned/unassigned customers.
'*This function will only work for other user's account assignments
'*******************************************************************************
Sub Listbox_Associate_Initialize()

    'Focus on dropdowns sheet
    With Sheets("DropDowns")

        'Find last row of Dropowns sheet (custom column)
        iLRow = Cells(.Rows.Count, "J").End(xlUp).Row + 1

        'Update customer list be just unassigned customers
        Multiuse_Listbox.List = .Range("J1:J" & iLRow).Value
    End With

    'Update toggle color
    Utility.ResetToggle
    Sheets("Control Panel").Shapes( _
        "Listbox_Associate_Tgl").Fill.ForeColor.RGB = RGB(64,64,64)

    'Correct listbox sizing
    ResizeListbox
End Sub


'*******************************************************************************
'Hide any shapes on Control Panel sheetthat are not constant UI elements.
'*******************************************************************************
Sub Cancel_Listbox_Initialize()

    'Loop through each shape in Control panel
    For Each shp In Sheets("Control Panel").Shapes

        'Hide shape if it is not a constant UI element
        If InStr(shp.Name, "Const") = 0 Then shp.Visible = False
    Next
End Sub

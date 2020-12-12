Attribute VB_Name = "Format"


'*******************************************************************************
'Hide any shapes on Control Panel sheetthat are not constant UI elements.
'*******************************************************************************
Sub Clear_Shapes()

    'Loop through each shape in Control panel
    For Each shp In Sheets("Control Panel").Shapes

        'Hide shape if it is not a constant UI element
        If InStr(shp.Name, "Const") = 0 Then shp.Visible = False
    Next
End Sub


'*******************************************************************************
'Update multiuse listbox with list of customers. Boolean operator indicates
'if listbox should conatain assigned customers or unassigned customers.
'*******************************************************************************
Sub Update_Listbox(blMyCst As Boolean)

    'Focus on dropdowns sheet
    With Sheets("DropDowns")

        'If listbox should be populated with assigned customers
        If blMyCst = True Then

            'Find last row of Dropowns sheet
            iLRow = Cells(.Rows.Count, "H").End(xlUp).Row + 1

            'Update customer list be just unassigned customers
            Multiuse_Listbox.List = .Range("H1:H" & iLRow).Value

        'If listbox should be populated with unassigned customers
        Else

            'Find last row of Dropowns sheet
            iLRow = Cells(.Rows.Count, "I").End(xlUp).Row + 1

            'Update customer list be just unassigned customers
            Multiuse_Listbox.List = .Range("I1:I" & iLRow).Value
        End If
    End With

    'Correct listbox sizing
    Resize_Listbox
End Sub


'*******************************************************************************
'Resize multiuse listbox to help accomodate different screen aspect ratios.
'*******************************************************************************
Sub Resize_Listbox()

    'Focus on Control Panel Sheet
    With Sheets("Control Panel")
        .Shapes("Cust_Add_Listbox").Width = .Range("N3:S3").Width
        .Shapes("Cust_Add_Listbox").Height = .Range("N3:N22").Height - 3
        .Shapes("Cust_Add_Listbox").Top = .Range("N3").Top
        .Shapes("Cust_Add_Listbox").Left = .Range("N3").Left - 6
    End With
End Sub


'*******************************************************************************
'Show all elements of selected Control Panel utility. Variant array includes
'all shapes to unhide.
'*******************************************************************************
Sub Show(varShapes As Variant)

    'Focus on Control Panel sheet
    With Sheets("Control Panel")

        'Loop through all shapes in object library
        For Each shp In varShapes

            'Unhide shape
            .Shapes(shp).Visible = True
        Next
    End With
End Sub

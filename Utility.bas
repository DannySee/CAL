Attribute VB_Name = "Format"


'*******************************************************************************
'Hide any shapes on Control Panel sheetthat are not constant UI elements.
'*******************************************************************************
Sub ClearShapes()

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
Sub UpdateListbox(blMyCst As Boolean)

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
    ResizeListbox
End Sub


'*******************************************************************************
'Resize multiuse listbox to help accomodate different screen aspect ratios.
'*******************************************************************************
Sub ResizeListbox()

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


'*******************************************************************************
'Return string of user selections (wrapped in quotes).
'*******************************************************************************
Function GetSelection() As Variant

    'Declare function variables
    Dim strCst As String

    'Loop through all dropdowns to create string
    With Sheets("Control Panel").Cust_Add_Listbox
        For i = 0 To .ListCount - 1
            If .Selected(i) Then strCst = Append(strCst, ",", .List(i))
        Next
    End With

    'Return string of customers
    If strCst <> "" Then GetSelection = Split(strCst, ",")
End Function


'******************************************************************************
'Prompt user to select a folder, create new folder, return folder path.
'******************************************************************************
Function SelectFolder() As String

    'Declare function variables
    Dim FldrPicker As FileDialog
    Dim strPth As String

    'Prompt user to select a folder path
    With FldrPicker
        .Title = "Select Folder Location"
        .AllowMultiSelect = False
        If .Show = -1 Then strPth = _
            .SelectedItems(1) & "\CAL by Customer " & Format(Now(), mm.dd.yy))
    End With

    'Create folders for Update letters to be saved to
    MkDir(strPth)

    'Return selected folder path
    SelectFolder = strPath
End Function


'******************************************************************************
'Create new workbook and return object.
'******************************************************************************
Function CreateWorkbook(strName As String) As Workbook

    'Create new workbook
    Workbooks.Add

    'Rename active sheet of new workbook
    ActiveSheet.Name = strName

    'Return workbook name
    Set CreateWorkbook = ActiveWorkbook
End Function

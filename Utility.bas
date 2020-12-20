Attribute VB_Name = "Utility"


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

            'Find last row of Dropowns sheet (custom column)
            iLRow = Cells(.Rows.Count, "H").End(xlUp).Row + 1

            'Update customer list be just unassigned customers
            Multiuse_Listbox.List = .Range("H1:H" & iLRow).Value

        'If listbox should be populated with unassigned customers
        Else

            'Find last row of Dropowns sheet (custom column)
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


'*******************************************************************************
'Prompt user to select a folder, create new folder, return folder path.
'*******************************************************************************
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
    SelectFolder = strPth & "\"
End Function


'*******************************************************************************
'Create new workbook with white formatting.
'*******************************************************************************
Sub CreateWorkbook(strName As String)

    'Create new workbook
    Workbooks.Add

    'Add blank formatting to workbook
    Cells.Interior.Color = vbWhite

    'Rename active sheet of new workbook
    ActiveSheet.Name = strName
End Function


'*******************************************************************************
'Add headers to active workbook with formatting.
'*******************************************************************************
Sub AddHeaders(varHeaders As Variant)

    'Declare sub variables
    Dim i As Integer

    'Loop through all elements in passthrough array
    For i = 0 To UBound(varHeaders)

        'Add value and formatting to cell
        With Cells(1,i + 1)
            .Value = varHeaders(i)
            .Interior.Color = RBG(255,255,255)
            .Font.Color = vbWhite
            .Font.Bold = True
        End With
    Next
End Sub


'*******************************************************************************
'Add borders to any range, assuming data starts in cell A1 and contains headers
'*******************************************************************************
Sub AddBorders()

    'Get last row and last column
    iLRow = LastRow(ActiveSheet.Name)
    iLCol = LastRow(ActiveSheet.Name)

    'Set borders to data range
    Range(Cells(iLRow,1), Cells(1,iLCol)).Borders.Linestyle = xlContinuous
End Sub


'*******************************************************************************
'Download customer friendly CAL to folder
'*******************************************************************************
Function DwnCstCAL(strCst As String) As Workbook

    'Declare function variables
    Dim strCst As String

    'Create new Workbook w/ formatting
    CreateWorkbook(strCst)

    'Add headers to Workbook
    AddHeaders(oPrgms.CstHeaders)

    'Setup customer string with quote delimiters
    strCst = "'" & cst & "'"

    'Query customer records and paste to sheet
    Cells(2,1).CopyFromRecordset Pull.GetPrograms(strCst, oPrgms.CstFlds)

    'Add borders to Workbook
    AddBorders

    'Add conditional formatting (indeces are for Excel start/end date fields)
    Format.AddCondFormatting(2,3)

    'Return Workbook
    Set DwnCstCAL = ActiveWorkbook
End Function


'*******************************************************************************
'Assemble body of reminder emails (bulleted list of expiring agreements).
'*******************************************************************************
Function GetReminderBody(strCst As String) As String

    'Declare function variables
    Dim exp As ADODB.Recordset
    Dim strExp As String
    Dim strDelmtr As String

    'Get recordset of expiring progrmas
    Set exp = Pull.GetExpPrograms(strCst)

    'Setup string delimiter variable
    strDelmtr = vbLf & "   " & Chr(149) & " "

    'Assemble string of program descriptions (bulleted & n\)
    Do While exp.EOF = False
        strExp = strExp & strDelmtr & exp.Fields("PROGRAM_DESCRIPTION").value
        rst.MoveNext
    Loop

    'Return string of expiring agreements
    GetReminderBody = strExp
End Sub


'*******************************************************************************
'Send reminder to DPM hotline Salesforce queue.
'*******************************************************************************
Sub SendReminder(strSubject As String, strTxt As String, strFile As String)

    'Declare sub variables
    Dim olOutlook As New Outlook.Application
    Dim olEmail As Object

    'Set email object
    Set olEmail = olOutlook.CreateItem(olMailItem)

    'Send email to inquiries queue
     With olEmail
        .To = "DPMHotline@corp.sysco.com"
        .Subject = strSubject
        .Body = strTxt
        .Attachments.Add strFile
        .Send
    End With

    'Free objects
    Set olEmail = Nothing
    Set olOutlook = Nothing
End Sub

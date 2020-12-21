Attribute VB_Name = "Utility"

'Declare private module constants
Private Const shtPWD As String = "Dac123am"
Private Const shtProperties As String = _
    "UserInterFaceOnly:=True, " & _
    "AllowFormattingCells:=True, " & _
    "AllowDeletingRows:=True, " & _
    "AllowFormattingRows:=True, " & _
    "AllowInsertingRows:=True, " & _
    "AllowSorting:=True, " & _
    "AllowFiltering:=True"


'*******************************************************************************
'Unlocks sheet using password constant. Parameter is sheet to unlock.
'*******************************************************************************
Sub ShtUnlock(strSht As String)

    'unprotect sheet
    Sheets(strSht).Unprotect shtPWD
End Sub


'*******************************************************************************
'Lock sheet using password constant. Parameter is sheet to lock.
'*******************************************************************************
Sub ShtLock(strSht As String)

    'Protect sheet with constant variables
    Sheets(strSht).Protect Password:=strPWD, shtProperties
End Sub


'*******************************************************************************
'Clear data from 3 data pages (Programs, Customer Profile & Deviation Loads).
'*******************************************************************************
Sub ClearShts()

    'Declare variables
    Dim varSht As Variant

    'Setup array of all Sheets
    varSht = Array("Programs","Customer Profile","Deviation Loads")

    'Loop through Sheets
    For Each sht In varSht

        'Unlock sheets
        shtUnlock(sht)

        'Get last row
        iLRow = LastRow(sht) + 1

        'Focus on sheet and clear all data
        With Sheets(sht)
            .ShowAll
            .Range("A2:A" & iLRow).EntireRow.Delete
        End With

        'Lock sheets
        shtLock(sht)
    Next
End Sub


'*******************************************************************************
'Delete old sheet detail and paste new. Parameters are sheet name and open
'recordset.
'*******************************************************************************
Sub ShtRefresh(obj As Object, upd As ADODB.Recordset)

    'Unlock sheet
    ShtUnlock(obj.Sht)

    'Get last row of sheet
    iLRow = LastRow(obj.Sht) + 1

    'Clear previous sheet content and paste new
    With Sheets(obj.Sht)
        .ShowAll
        .Cells(iLRow, 1).CopyFromRecordset upd
    End With

    'Format sheet
    AddFormat(obj)

    'Lock sheet
    ShtLock(obj.Sht)

    'Close recordset
    upd.Close
End Sub


'*******************************************************************************
'Format CAL sheet. Resize rows, correct borders, lock white space etc.
'Parameter is the sheet to be formatted.
'*******************************************************************************
Sub AddFormat(obj As Object)

    'Activate sheet
    With Sheets(obj.Sht)

        'Get last row and last COLUMN_NUM
        iLRow = LastRow(obj.Sht)
        iLCol = LastCol(obj.Sht)

        'Format Columns (width/height)
        .Columns.ColumnWidth = 100
        .Rows.RowHeight = 100
        .Rows.AutoFit
        .Columns.AutoFit

        'Format cell borders
        .Cells.Borders.LineStyle = xlNone
        .Range(.Cells(1, 1), .Cells(iLRow, iLCol)).Borders.LineStyle = _
            xlContinuous

        'Unlock all cells
        .Cells.Locked = False

        'Lock Primary Key & Customer ID column
        .Range(.Cells(1,1), .Cells(iLRow, 2)).Locked = True

        'Lock all fields to the right of data range
        .Range(.Cells(1, iLCol + 1), .Cells(iLRow, .Columns.Count)).Locked= True

        'Lock all fields under data range
        .Range(.Cells(iLRow + 1, 1), _
            .Cells(.Rows.Count, 1)).EntireRow.Locked = True
    End With

    'Add data validation (drop down options)
    obj.AddDataValidation

    'If active sheet is Programs tab then add end date condition formatting
    If obj.Sht = "Programs" Then AddCondFormatting( _
        oPrgms.ColIndex("START_DATE") + 1, _
        oPrgms.ColIndex("END_DATE") + 1)
End Sub


'*******************************************************************************
'Add conditional formatting to the Programs tab. Weekly programs are highlighted
'green and programs expiring EOM are highlighted red.
'*******************************************************************************
Sub AddCondFormatting(iStrt As Integer, iEnd As Integer)

    'Declare sub variables
    Dim rngFrmt As Range
    Dim strDteRng As String

    'Get last row of sheet
    iLRow = LastRow("Programs")

    'Activate Programs tab
    With Sheets("Programs")

        'Find last row
        iLRow = .Cells(.Rows.Count, 1).End(xlUp).Row

        'Set range to be formatted
        Set rngFrmt = .Range(.Cells(2,iEnd), .Cells(iLRow, iEnd))

        'Get formula string for weekly highlight (end-iStart)
        strDteRng = .Cells(2, iEnd).Address & "-" & .Cells(2, iStrt).Address

        'Clear conditional formatting from range
        rngFrmt.FormatConditions.Delete

        'Set conditional formatting for weekly programs
        rngFrmt.Add(xlExpression, xlEqual, Formula1:="=(" & strDteRng _
            & ")=6").Interior.Color = RGB(137, 191, 101)

        'Set conditional formatting for standard programs
        rngFrmt.FormatConditions.Add(xlCellValue, xlLess, "=" & _
            CLng(DateSerial(Year(Now), Month(Now) + 1, 11))).Interior.Color = _
            RGB(250, 120, 120)
    End With
End Sub


'*******************************************************************************
'Add drop down items to excel file to be referenced for dropdown formatting.
'*******************************************************************************
Sub AddDropDwns()

    'Declare sub variables
    Dim dropDwns As Variant
    Dim myCst As Variant
    Dim othCst As Variant
    Dim iRow As Integer
    Dim iCol As Integer

    'Get multidimensional arrays of drop down database
    dropDwns = Pull.GetDropDwns
    myCst = Pull.GetCst(True)
    othCst = Pull.GetCst(False)

    'Clear old Dropdown values
    Sheets("DropDowns").Cells.Value = ""

    'Loop through columns in dropdown multidimensional array
    For iCol = 0 To UBound(dropDwns, 1)

        'Loop through rows in dropdown multidimensional array
        For iRow = 0 To UBound(dropDwns, 2)

            'Paste dropdowns in appropriate columns
            Sheets("DropDowns").Cells(iRow + 1, iCol + 1).Value = _
                dropDwns(iCol, iRow)
        Next
    Next

    'Loop through each element of assigned customer array
    For i = 0 To UBound(myCst)

        'Paste assigned customers in list format
        Sheets("DropDowns").Cells(i+1, 8).Value = myCst(i)
    Next

    'Loop through each element of unassigned customer array
    For i = 0 To UBound(othCst)

        'Paste unassigned customers in list format
        Sheets("DropDowns").Cells(i+1, 9).Value = othCst(i)
    Next
End Sub


'*******************************************************************************
'Update assigned/unassigned customer dropdown options
'*******************************************************************************
Sub ReviseDropDwns(varCst As Variant)

    'Focus on DropDowns sheet
    With Sheets("DropDowns")

        'Loop through all customers in passthrough array
        For i = 0 To Ubound(varCst)

            'Get last row (custom column)
            iLRow = .Cells(.Rows.Count,"H").End(xlUp).Row + 1

            'Add customers to assigned customer list
            .Cells(iLRow, "H").Value = varCst(i)

            'Remove customer from unassigned customer list
            .Cells(.Columns("I").Find(varCst(i)).Row, "I")).Delete
        Next
    End WIth
End Sub


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
    AddCondFormatting(2,3)

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


'*******************************************************************************
'Select all elements from multiuse listbox
'*******************************************************************************
Sub UpdateListboxAll(strSubject As String, strTxt As String, strFile As String)


End Sub


'*******************************************************************************
'Update multiuse listbox with list of customers. Boolean operator indicates
'if listbox should conatain assigned customers or unassigned customers.
'*******************************************************************************
Sub UpdateListboxCst(blMyCst As Boolean)

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

    'Highlight button
    ResetListToggle
    Sheets("Control Panel").Shapes("Listbox_Account_Tgl").Interior.Color = _
        RGB(64,64,64)

    'Correct listbox sizing
    ResizeListbox
End Sub


'*******************************************************************************
'Update multiuse listbox with list of account holders. Boolean operator
'indicates if listbox should conatain assigned/unassigned customers.
'*This function will only work for other user's account assignments
'*******************************************************************************
Sub UpdateListboxAss()

    'Focus on dropdowns sheet
    With Sheets("DropDowns")

        'Find last row of Dropowns sheet (custom column)
        iLRow = Cells(.Rows.Count, "J").End(xlUp).Row + 1

        'Update customer list be just unassigned customers
        Multiuse_Listbox.List = .Range("J1:J" & iLRow).Value
    End With

    'Update toggle color
    ResetListToggle
    Sheets("Control Panel").Shapes("Listbox_Holder_Tgl").Interior.Color = _
        RGB(64,64,64)

    'Correct listbox sizing
    ResizeListbox
End Sub


'*******************************************************************************
'Reset listbox toggle buttons to default color
'*******************************************************************************
Sub ResetListToggle()

    'Set all toggle
    With Sheets("Control Panel")
        .Shapes("Listbox_Holder_Tgl").Interior.Color = RGB(89,89,89)
        .Shapes("Listbox_Account_Tgl").Interior.Color = RGB(89,89,89)
        .Shapes("Listbox_All_Tgl").Interior.Color = RGB(89,89,89)
    End With
End Sub

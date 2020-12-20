Attribute VB_Name = "Format"

'Declare private module constants
Private Const shtPWD As String = "Dac123am"
Private Const shtProperties As String = _
    "UserInterFaceOnly:=True, " & _
    "AllowFormattingCells:=True, " & _
    "AllowDeletingRows:=True, " & _
    "AllowFormattingRows:=True, " & _
    "AllowInsertingRows:=True, " & _
    "AllowSorting:=False, " & _
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
    If obj.Sht = "Programs" Then AddCondFormatting
End Sub


'*******************************************************************************
'Add conditional formatting to the Programs tab. Weekly programs are highlighted
'green and programs expiring EOM are highlighted red.
'*******************************************************************************
Sub AddCondFormatting()

    'Declare sub variables
    Dim rngFrmt As Range
    Dim iStrt As Integer
    Dim iEnd As Integer
    Dim strDteRng As String

    'Get pertinant column indeces
    iStrt = oPrgms.ColIndex("START_DATE") + 1
    iEnd = oPrgms.ColIndex("END_DATE") + 1

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

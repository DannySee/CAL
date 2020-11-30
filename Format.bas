Attribute VB_Name = "Format"


'Declare private module constants
Private const shtPWD = "Dac123am"
Private const shtProperties = _
    "UserInterFaceOnly:=True, " & _
    "AllowFormattingCells:=True, " & _
    "AllowDeletingRows:=True, " & _
    "AllowFormattingRows:=True, " & _
    "AllowInsertingRows:=True, " & _
    "AllowSorting:=False, " & _
    "AllowFiltering:=True"


'Declare private module variables
Private iLRow As Long
Private iLCol As Integer


'*******************************************************************************
'Unlocks sheet using password constant. Parameter is sheet to unlock.
'*******************************************************************************
Sub ShtUnlock(strSheet As String)

    'unprotect sheet
    Sheets(strSheet).Unprotect shtPWD
End Sub


'*******************************************************************************
'Lock sheet using password constant. Parameter is sheet to lock.
'*******************************************************************************
Sub ShtLock(strSheet As String)

    'Protect sheet with constant variables
    Sheets(strSheet).Protect Password:=strPWD, shtProperties
End Sub


'*******************************************************************************
'Delete old sheet detail and paste new. Parameters are sheet name and open
'recordset.
'*******************************************************************************
Sub ShtRefresh(strSht As String, upd As ADODB.Recordset)

    'Unlock sheet
    ShtUnlock(strSht)

    'Clear previous sheet content and paste new
    With Sheets(strSht)
        .Rows(1).AutoFilter
        iLRow = .Range("A" & .Rows.Count).End(xlUp)
        .Range("A2:A" & iLRow + 1).EntireRow.Delete
        .Range("A2").CopyFromRecordset upd
        .Rows(1).AutoFilter
    End With

    'Format sheet
    AddFormat(strSht)

    'Lock sheet
    ShtLock(strSht)

    'Close recordset
    upd.Close
End Sub


'*******************************************************************************
'Format CAL sheet. Resize rows, correct borders, lock white space etc.
'Parameter is the sheet to be formatted.
'*******************************************************************************
Sub AddFormat(strSht As String)

    'Activate sheet
    With Sheets(strSht)

        'Get last row and last COLUMN_NUM
        iLRow = .Cells(.Rows.Count,1).End(xlUp).Row
        iLCol = .Cells(1,.Columns.Count).End(xlToLeft).Column

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

    'If active sheet is Programs tab then add end date condition formatting
    If strSht = "Programs" Then AddCondFormatting
End Sub


'*******************************************************************************
'Add conditional formatting to the Programs tab. Weekly programs are highlighted
'green and programs expiring EOM are highlighted red.
'*******************************************************************************
Sub AddCondFormatting()

    'Declare sub variables
    Dim rngFrmt As Range
    Dim iLRow As Long
    Dim iStrt As Integer
    Dim iEnd As Integer
    Dim strDteRng As String

    'Get pertinant column indeces
    iStrt = oPrgms.ColIndex("START_DATE") + 1
    iEnd = oPrgms.ColIndex("END_DATE") + 1

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
'Parameters include array of dropdowns & assigned/unassigned customers.
'*******************************************************************************
Sub AddDropDwns(dropDwns As Variant, myCst as Variant, othCst As Variant)

    'Declare sub variables
    Dim iRow As Integer
    Dim iCol As Integer
    Dim i As Integer

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

    'Add data validation to programs tab (create drop down lists)
    AddDataValidation
End Sub


'*******************************************************************************
'Add data validaiton to Programs tab. Include appropriate drop down list for
'all restricted fields. BB Format will always include a blank list (field is
'deactivated).
'*******************************************************************************
Sub AddDataValidation()

    'Declare sub variables
    Dim rngYN As Range
    Dim rngDrp As Range
    Dim rngPrgm As Range
    Dim varPrgmRng As Variant
    Dim iLRow As Long
    Dim iDropLR As Integer
    Dim iTier As Integer
    Dim iVAType As Integer
    Dim iCost As Integer
    Dim iBB As Integer
    Dim iCAType As Integer
    Dim iRebate As Integer
    Dim iApprop As Integer
    Dim iCst As Integer
    Dim i As Integer

    'Get indeces for pertinant Excel (Programs) indeces
    iTier = oPrgms.ColIndex("TIER") + 1, _
    iVAType = oPrgms.ColIndex("VEND_AGMT_TYPE") + 1, _
    iCost = oPrgms.ColIndex("COST_BASIS") + 1, _
    iBB = oPrgms.ColINdex("BILLBACK_FORMAT") + 1 _
    iCAType = oPrgms.ColIndex("CUST_AGMT_TYPE") + 1, _
    iRebate = oPrgms.ColIndex("REBATE_BASIS") + 1, _
    iApprop = oPrgms.ColIndex("APPROP_NAME") + 1 _
    iCst = oPrgms.ColIndex("CUSTOMER") + 1)

    'Activate programs tab
    With Sheets("Programs")

        'Remove data validation from sheet
        .Cells.Validation.Delete

        'Find last row
        iLRow = .Cells(.Rows.Count, 1).End(xlUp).Row

        'Set range for fields which will have Y/N drop down
        Set rngYN = .Range(.Cells(2, oPrgms.ColIndex("DAB")), _
            .Cells(iLRow, oPrgms.ColIndex("TIMELINESS")))

        'Create an array of pertinant Excel (Programs) ranges
        varPrgmRng = Array( _
            .Range(.Cells(2, iTier), .Cells(iLRow, iTier)), _
            .Range(.Cells(2, iVAType), .Cells(iLRow, iVAType)), _
            .Range(.Cells(2, iCost), .Cells(iLRow, iCost)), _
            .Range(.Cells(2, iBB), .Cells(iLRow, iBB)), _
            .Range(.Cells(2, iCAType), .Cells(iLRow, iCAType)), _
            .Range(.Cells(2, iRebate), .Cells(iLRow, iRebate)), _
            .Range(.Cells(2, iApprop), .Cells(iLRow, iApprop)), _
            .Range(.Cells(2, iCst), .Cells(iLRow, iCst)))
    End With

    'Add Y/N drop down list to first three editable Fields
    rngYN.Validation.Add xlValidateList, Formula1:="Y,N"

    'Activate DropDowns tab
    With Sheets("DropDowns")

        'Loop through each column of DropDowns tab
        For i = 0 To UBound(varPrgmRng)

            'Get last row of dropdown fields
            iDropLR = .Cells(.Rows.Count, i + 1).End(xlUp).Row 

            'Save dropdown Range
            Set rngDrp = .Range(.Cells(1, i + 1), .Cells(iDropLR, i + 1))

            'Add data validation to Excel (Programs) range
            varPrgmRng(i).Validation.Add xlValidateList, _
                Formula1:="=DropDowns!" & rngDrp.Address
        Next
    End With
End Sub

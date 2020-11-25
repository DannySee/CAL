Attribute VB_Name = "UL_Maintenance"


Sub UL_Refresh()

    'Declare global variable
    Dim cnn                 As New ADODB.Connection
    Dim rst                 As New ADODB.Recordset
    Dim strCnn              As String
    Dim strProgramFields    As String
    Dim i                   As Long

    'Freeze events while data is being pulled
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Establish connection to SSMS
    strCnn = "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"
    cnn.Open strCnn
    cnn.CommandTimeout = 900

    'Setup fields variable
    strProgramFields = "PRIMARY_KEY," _
        & "CUSTOMER_ID," _
        & "PROGRAM_ID," _
        & "DAB," _
        & "SCRIPT_ASSIST," _
        & "TIMELINESS," _
        & "TIER," _
        & "CUSTOMER," _
        & "PROGRAM_DESCRIPTION," _
        & "START_DATE," _
        & "END_DATE," _
        & "LEAD_VA," _
        & "LEAD_CA," _
        & "VEND_AGMT_TYPE," _
        & "VENDOR_NUM," _
        & "BILLBACK_FORMAT," _
        & "COST_BASIS," _
        & "CUST_AGMT_TYPE," _
        & "REBATE_BASIS," _
        & "PRE_APPROVAL," _
        & "APPROP_NAME," _
        & "PRN_GRP," _
        & "PACKET," _
        & "PACKET_DL," _
        & "COMMENTS"

    'query to pull in Programs sheet
    rst.Open "SELECT " & strProgramFields & " " _
        & "FROM UL_Programs " _
        & "LEFT JOIN (" _
            & "SELECT MAX(END_DATE) AS ED, PROGRAM_ID AS PID " _
            & "FROM UL_Programs " _
            & "GROUP BY PROGRAM_ID) AS O " _
        & "ON PROGRAM_ID = O.PID " _
        & "WHERE CUSTOMER_ID IN (" _
            & "SELECT CUSTOMER_ID " _
            & "FROM UL_Account_Ass " _
            & "WHERE TIER_1 = '" & Application.Username & "' " _
            & "OR TIER_2 = '" & Application.Username & "' " _
            & "OR T1_ID = '" & Environ("Username") & "' " _
            & "OR T2_ID = '" & Environ("Username") & "') " _
        & "AND O.ED = END_DATE " _
        & "ORDER BY CUSTOMER, PROGRAM_DESCRIPTION", cnn

    'Clear previous sheet contents and paste new (pause protection)
    With Sheets("Programs")
        .Unprotect "Dac123am"
        .Rows(1).AutoFilter
        .Range("A2:A" & .Range("A" & .Rows.Count).End(xlUp).Row + 1).EntireRow.Delete
        .Range("A2").CopyFromRecordset rst
        .Rows(1).AutoFilter
        .Tab.Color = RGB(38, 38, 38)
    End With

    'Close recordset for next sheet
    rst.Close

    'Setup query string to pull in Customer Profile sheet
    rst.Open "SELECT * " _
        & "FROM UL_Customer_Profile " _
        & "WHERE CUSTOMER_ID IN (" _
            & "SELECT CUSTOMER_ID " _
            & "FROM UL_Account_Ass " _
            & "WHERE TIER_1 = '" & Application.Username & "' " _
            & "OR  TIER_2 = '" & Application.Username & "' " _
            & "OR T1_ID = '" & Environ("Username") & "' " _
            & "OR T2_ID = '" & Environ("Username") & "') " _
        & "ORDER BY CUSTOMER", cnn

    'Clear previous sheet contents and paste new
    With Sheets("Customer Profile")
        .Unprotect "Dac123am"
        .Rows(1).AutoFilter
        .Range("A2:A" & .Range("A" & .Rows.Count).End(xlUp).Row + 1).EntireRow.Delete
        .Range("A2").CopyFromRecordset rst
        .Rows(1).AutoFilter
    End With

    'Close recordset for next sheet
    rst.Close

    'Setup query string to pull in Programs sheet
    rst.Open "SELECT * " _
        & "FROM UL_Deviation_Loads " _
        & "WHERE CUSTOMER_ID IN (" _
            & "SELECT CUSTOMER_ID " _
            & "FROM UL_Account_Ass " _
            & "WHERE TIER_1 = '" & Application.Username & "' " _
            & "OR TIER_2 = '" & Application.Username & "' " _
            & "OR T1_ID = '" & Environ("Username") & "' " _
            & "OR T2_ID = '" & Environ("Username") & "') " _
        & "ORDER BY CUSTOMER, PROGRAM", cnn

    'Clear previous sheet contents and paste new
    With Sheets("Deviation Loads")
        .Unprotect "Dac123am"
        .Rows(1).AutoFilter
        .Range("A2:A" & .Range("A" & .Rows.Count).End(xlUp).Row + 1).EntireRow.Delete
        .Range("A2").CopyFromRecordset rst
        .Rows(1).AutoFilter
    End With

    'Close recordset for next sheet
    rst.Close

    'Clear sheet
    Sheets("DropDowns").Cells.value = ""

    'Loop through all columns in drop downs
    For i = 1 To 7

        'Setup query string to pull in Customer Profile sheet
        rst.Open "SELECT DROP_DOWN " _
            & "FROM UL_List_Options " _
            & "WHERE COLUMN_NUM = " & i & " " _
            & "ORDER BY DROP_DOWN", cnn

        'Paste to sheet
        Sheets("DropDowns").Cells(1, i).CopyFromRecordset rst

        'Close recordset for next iteration
        rst.Close
    Next

    'Add new customer drop (user's customer)
    rst.Open "SELECT CUSTOMER_NAME " _
        & "FROM UL_ACCOUNT_ASS " _
        & "WHERE TIER_1 = '" & Application.Username & "' " _
        & "OR TIER_2 = '" & Application.Username & "' " _
        & "OR T1_ID = '" & Environ("Username") & "' " _
        & "OR T2_ID = '" & Environ("Username") & "' " _
        & "ORDER BY CUSTOMER_NAME"
    Sheets("DropDowns").Cells(1, 8).CopyFromRecordset rst

    'Close recordset for next iteration
    rst.Close

    'Add new customer drop (Other customers)
    rst.Open "SELECT CUSTOMER_NAME " _
        & "FROM UL_ACCOUNT_ASS " _
        & "WHERE TIER_1 <> '" & Application.Username & "' " _
        & "AND TIER_2 <> '" & Application.Username & "' " _
        & "AND T1_ID <> '" & Environ("Username") & "' " _
        & "AND T2_ID <> '" & Environ("Username") & "' " _
        & "ORDER BY CUSTOMER_NAME"
    With Sheets("DropDowns")
        .Cells(1, 9).CopyFromRecordset rst
        Sheets("Control Panel").Cust_Add_Listbox.List = _
        .Range("I1:I" & .Range("I" & .Rows.Count).End(xlUp).Row).value
    End With

    'Close recordset
    rst.Close

    'Free objects
    Set cnn = Nothing

    'Call format sub
    UL_Format

    'Ensure control panel is active sheet
    Sheets("Control Panel").Activate

    'Reset settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Sub UL_Format()

    'Declare global variable
    Dim iLRow               As Long
    Dim iLCol               As Integer
    Dim ws                  As Worksheet

    'Initiate loop through worksheets
    For Each ws In Worksheets
        With ws

            'Validate which sheet the loop is currently on
            If .Name = "Upload Sheet" Then

                'Reset upload sheet data
                .Cells.value = ""
            Else
                If .Name <> "DropDowns" And .Name <> "Control Panel" And Not .Name Like "*Recover*" And Not .Name Like "*Insert*" And Not .Name = "Upload" Then

                    'AutoFit
                    .Columns.ColumnWidth = 100
                    .Rows.RowHeight = 100
                    .Rows.AutoFit
                    .Columns.AutoFit

                    'Find last row and column
                    iLRow = .Range("A" & .Rows.Count).End(xlUp).Row
                    iLCol = .Cells(1, .Columns.Count).End(xlToLeft).Column

                    'Reset borders
                    .Cells.Borders.LineStyle = xlNone
                    .Range(.Cells(1, 1), .Cells(iLRow, iLCol)).Borders.LineStyle = xlContinuous

                    'Formatting tasks specific to Programs List sheet
                    If .Name = "Programs" Then

                        'Hide proper columns
                        .Columns("A:C").Hidden = True

                        'Add data validation to pertinant columns
                        .Cells.Validation.Delete
                        .Range("D2:F" & iLRow).Validation.Add xlValidateList, Formula1:="Y,N"
                        .Range("G2:G" & iLRow).Validation.Add xlValidateList, Formula1:="=DropDowns!$A$1:$A$" _
                            & Sheets("DropDowns").Range("A" & Sheets("DropDowns").Rows.Count).End(xlUp).Row
                        .Range("N2:N" & iLRow).Validation.Add xlValidateList, Formula1:="=DropDowns!$B$1:$B$" _
                            & Sheets("DropDowns").Range("B" & Sheets("DropDowns").Rows.Count).End(xlUp).Row
                        .Range("P2:P" & iLRow).Validation.Add xlValidateList, Formula1:="=DropDowns!$C$1:$C$" _
                            & Sheets("DropDowns").Range("C" & Sheets("DropDowns").Rows.Count).End(xlUp).Row
                        .Range("Q2:Q" & iLRow).Validation.Add xlValidateList, Formula1:="=DropDowns!$D$1:$D$" _
                            & Sheets("DropDowns").Range("D" & Sheets("DropDowns").Rows.Count).End(xlUp).Row
                        .Range("R2:R" & iLRow).Validation.Add xlValidateList, Formula1:="=DropDowns!$E$1:$E$" _
                            & Sheets("DropDowns").Range("E" & Sheets("DropDowns").Rows.Count).End(xlUp).Row
                        .Range("S2:S" & iLRow).Validation.Add xlValidateList, Formula1:="=DropDowns!$F$1:$F$" _
                            & Sheets("DropDowns").Range("F" & Sheets("DropDowns").Rows.Count).End(xlUp).Row
                        .Range("U2:U" & iLRow).Validation.Add xlValidateList, Formula1:="=DropDowns!$G$1:$G$" _
                            & Sheets("DropDowns").Range("G" & Sheets("DropDowns").Rows.Count).End(xlUp).Row

                        'Set range for conditionaly formatting
                        With .Range(.Cells(2, "K"), .Cells(iLRow, "K"))

                            'Delete existing conditional formatting
                            .FormatConditions.Delete

                            'Add conditional formatting
                            .FormatConditions.Add(xlExpression, xlEqual, Formula1:="=($K2-$J2)=6").Interior.Color = RGB(137, 191, 101)
                            .FormatConditions.Add(xlCellValue, xlLess, "=" & CLng(DateSerial(Year(Now), Month(Now) + 1, 11))).Interior.Color = RGB(250, 120, 120)
                        End With
                    End If

                    'Hide appropriate columns
                    .Columns("A:B").Hidden = True

                    'Protect Worsheet
                    .Cells.Locked = False
                    .Range(.Cells(1, .Columns.Count).End(xlToLeft).Offset(0, 1), .Cells(1, .Columns.Count)).Locked = True
                    .Range(.Cells(iLRow + 1, 1), .Cells(.Rows.Count, iLCol)).Locked = True
                    .Protect Password:="Dac123am", UserInterFaceOnly:=True, AllowFormattingCells:=True, AllowDeletingRows:=True, _
                        AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=False, AllowFiltering:=True
                End If
            End If
        End With
    Next

    MsgBox "Done"

End Sub

Sub Upload_Programs()

    'Declare global variable
    Dim dictUL              As New Scripting.Dictionary
    Dim cnn                 As New ADODB.Connection
    Dim rst                 As New ADODB.Recordset
    Dim strCnn              As String
    Dim strVal              As String
    Dim varUpd              As Variant
    Dim varUpdCol           As Variant
    Dim strIns              As String
    Dim i                   As Long
    Dim m                   As Long
    Dim iSheet              As Integer

    'Reset settings on error
    On Error GoTo errorStep:

    'Focus on upload sheet
    With Sheets("Upload Sheet")

        'Ensure there is data to parse
        If .Range("A2") = "" And .Range("C2") = "" And .Range("E2") = "" Then
            Exit Sub
        End If

        'Establish connection to SSMS
        strCnn = "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"
        cnn.Open strCnn
        cnn.CommandTimeout = 900

        'Create dictionary for SSMS fields
        With dictUL
            .Add 1, Array("MAX(LEN(PROGRAM_ID)) ", "CUSTOMER ", "CUSTOMER ", 8, 3, 3, "GROUP BY CUSTOMER_ID", "", "", "_programs", "_customer", "_deviation")
            .Add 2, Array("CUSTOMER_ID", "CUSTOMER_ID", "CUSTOMER_ID")
            .Add 3, Array("PROGRAM_ID", "CUSTOMER", "CUSTOMER")
            .Add 4, Array("DAB", "ALT_NAME", "PROGRAM")
            .Add 5, Array("SCRIPT_ASSIST", "PACKET", "OWNER")
            .Add 6, Array("TIMELINESS", "PRICE_RULE", "DATE")
            .Add 7, Array("TIER", "NID", "SR")
            .Add 8, Array("CUSTOMER", "MASTER_PRN", "GRP")
            .Add 9, Array("PROGRAM_DESCRIPTION", "PRICING_PRN", "")
            .Add 10, Array("START_DATE", "GROUP_NAME", "")
            .Add 11, Array("END_DATE", "VPNA", "")
            .Add 12, Array("LEAD_VA", "NAM", "")
            .Add 13, Array("LEAD_CA", "CUST_CONTACT", "")
            .Add 14, Array("VEND_AGMT_TYPE", "NOTES")
            .Add 15, Array("VENDOR_NUM", "")
            .Add 16, Array("BILLBACK_FORMAT", "")
            .Add 17, Array("COST_BASIS", "")
            .Add 18, Array("CUST_AGMT_TYPE", "")
            .Add 19, Array("REBATE_BASIS", "")
            .Add 20, Array("PRE_APPROVAL", "")
            .Add 21, Array("APPROP_NAME", "")
            .Add 22, Array("PRN_GRP", "")
            .Add 23, Array("PACKET", "")
            .Add 24, Array("PACKET_DL", "")
            .Add 25, Array("COMMENTS", "")
            .Add "A0", 1
            .Add "A1", 3
            .Add "A2", 5
        End With

        'Ensure there is data to parse on
        For iSheet = 0 To 2

            'Ensure there is data to parse
            If .Cells(2, dictUL("A" & iSheet)) <> "" Then

                'Remove duplicates and sort
                .Range(.Cells(2, dictUL("A" & iSheet)), .Cells(.Rows.Count, dictUL("A" & iSheet) + 1).End(xlUp)).RemoveDuplicates _
                    Columns:=Array(1, 2), Header:=xlNo
                .Range(.Cells(2, dictUL("A" & iSheet)), .Cells(.Rows.Count, dictUL("A" & iSheet) + 1).End(xlUp)).Sort _
                    Key1:=.Cells(2, dictUL("A" & iSheet) + 1), Order1:=xlAscending, Header:=xlNo

                'Loop through upload sheet to find address of fields that were changed
                For i = 2 To .Cells(.Rows.Count, dictUL("A" & iSheet)).End(xlUp).Row

                    'Ensure customer and program IDs are populated
                    If Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 1) = "" Then

                        'Pull cusomter ID from customer profile table
                        rst.Open "SELECT DISTINCT CUSTOMER_ID FROM UL_Account_Ass WHERE CUSTOMER_NAME = '" & Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), dictUL(1)(iSheet + 3)) & "' ", cnn

                        'Paste CID
                        Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 2).CopyFromRecordset rst
                        rst.Close

                        'lookup program ID
                        rst.Open "SELECT DISTINCT CUSTOMER_ID, MAX(CAST(right(PROGRAM_ID, charindex('-', reverse(PROGRAM_ID)) - 1) AS INT)) + 1 " _
                            & "FROM UL_Programs " _
                            & "WHERE CUSTOMER = '" & Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), dictUL(1)(iSheet + 3)) & "' GROUP BY CUSTOMER_ID", cnn
                        If iSheet = 0 And rst.EOF = False Then
                            Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 2).CopyFromRecordset rst
                            Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 3) = "'" & Cells(.Cells(i, _
                                dictUL("A" & iSheet) + 1), 2) & "-" & Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 3)
                        ElseIf iSheet = 0 Then
                            Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 3).value = Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 2) & "-1"
                        End If
                        rst.Close
                    End If

                    'Setup sub loop integer and update string
                    varUpd = ""
                    strVal = ""
                    strIns = ""
                    m = i

                    'Loop through like rows
                    Do While .Cells(i, dictUL("A" & iSheet) + 1) = .Cells(m, dictUL("A" & iSheet) + 1)

                        'Create update string and value string
                        If varUpd = "" Then
                            varUpdCol = dictUL(.Cells(m, dictUL("A" & iSheet)).value)(iSheet)
                            varUpd = "'" & Replace(Sheets(iSheet + 1).Cells(.Cells(m, dictUL("A" & iSheet) + 1), _
                                .Cells(m, dictUL("A" & iSheet))).value, "'", "") & "'"
                        Else
                            If .Cells(m, dictUL("A" & iSheet)) = 15 Then
                                varUpdCol = varUpdCol & "*" & dictUL(.Cells(m, dictUL("A" & iSheet)).value)(iSheet)
                                varUpd = varUpd & "*" & Sheets(iSheet + 1).Cells(.Cells(m, dictUL("A" _
                                    & iSheet) + 1), .Cells(m, dictUL("A" & iSheet))).value
                            Else
                                varUpdCol = varUpdCol & "*" & dictUL(.Cells(m, dictUL("A" & iSheet)).value)(iSheet)
                                varUpd = varUpd & "*'" & Replace(Sheets(iSheet + 1).Cells(.Cells(m, dictUL("A" _
                                    & iSheet) + 1), .Cells(m, dictUL("A" & iSheet))).value, "'", "") & "'"
                            End If
                        End If

                        'Iterate sub loop
                        m = m + 1
                    Loop

                    'Reset upload look row
                    i = m - 1

                    'loop through row to be inserted
                    For Each r In Sheets(iSheet + 1).Range(Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1).value, 2), _
                        Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1).value, Sheets(iSheet + 1).Cells(1, Sheets(iSheet + 1).Columns.Count).End(xlToLeft).Column))

                        'Create insert and value String
                        If strIns = "" Then
                            strIns = dictUL(r.Column)(iSheet)
                            strVal = r.value
                        Else
                            strIns = strIns & "," & dictUL(r.Column)(iSheet)
                            If r.Column = 15 And iSheet = 0 Then
                                If r = "" Then
                                    strVal = strVal & "," & 0
                                Else
                                    strVal = strVal & "," & r.value
                                End If
                            Else
                                strVal = strVal & ",'" & Replace(r.value, "'", "") & "'"
                            End If
                        End If
                    Next

                    'Determine which form of upload to use for proper data management
                    If InStr(varUpdCol, "START_DATE") > 0 Or Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 1) = "" Then

                        'Check if new line
                        If Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 1) <> "" Then

                            'Query old end date
                            rst.Open "SELECT MAX(END_DATE) AS ED FROM UL_Programs WHERE PRIMARY_KEY = " & Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 1), cnn

                            'Update previous agreement end date
                            If CLng(rst.Fields("ED")) >= CLng(Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 10)) Then cnn.Execute _
                                ("EXEC update_end_programs '" & CLng(DateAdd("d", -3, Sheets(iSheet + 1).Range("J" _
                                & .Cells(i, dictUL("A" & iSheet) + 1).value))) & "','" _
                                & Sheets(iSheet + 1).Range("A" & .Cells(i, dictUL("A" & iSheet) + 1).value).value & "'")

                            'Close recorset
                            rst.Close
                        End If

                        'Insert new agreement line
                        If Left(strVal, 1) <> "," Then cnn.Execute ("EXEC insert" & dictUL(1)(iSheet + 9) & " " & strVal)
                    Else

                        'Delete edited agreement line
                        cnn.Execute ("EXEC delete_update" & dictUL(1)(iSheet + 9) & " " & Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1).value, 1).value)

                        'Insert new agreement line
                        cnn.Execute ("EXEC insert" & dictUL(1)(iSheet + 9) & " " & strVal)
                    End If

                    'Insert Primary Key if not already populated
                    If iSheet = 0 Then
                        rst.Open "SELECT MAX(PRIMARY_KEY) AS PKEY FROM UL_Programs WHERE PROGRAM_ID = '" & Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 3) & "'", cnn
                        Sheets(iSheet + 1).Cells(.Cells(i, dictUL("A" & iSheet) + 1), 1).value = rst.Fields("PKEY").value
                        rst.Close
                    End If
                Next
            End If
        Next

        'Clear Upload Sheet
        .Cells.value = ""
    End With

    'Free Objects
    Set cnn = Nothing

    'Exit macro
    Exit Sub

'Reset settings on error
errorStep:
    MsgBox "You've run into an unexpected error. Please save a copy of this CAL before refreshing."

End Sub

Sub Delete_Programs(iRow As Long, strSht As String, strFields As String, strTable As String, strDel As String, iDel As Integer)

    'Declare global variable
    Dim cnn                 As New ADODB.Connection
    Dim rst                 As New ADODB.Recordset
    Dim strCnn              As String
    Dim strVal              As String

    'Ensure correct sheet is active
    With Sheets(strSht)

        'Ensure there is data to parse
        If .Cells(iRow, 1) = "" Then
            Exit Sub
        End If

        'Establish connection to SSMS
        strCnn = "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"
        cnn.Open strCnn
        cnn.CommandTimeout = 900

        'Loop through row to be deleted
        For Each r In .Range(.Cells(iRow, 2), .Cells(iRow, .Cells(1, Columns.Count).End(xlToLeft).Column))

            'Create value string for insert statement
            If strVal = "" Then
                strVal = r.value
            Else
                If r.Column = 15 And strSht = "Programs" Then
                    strVal = strVal & "," & r.value
                Else
                    strVal = strVal & ",'" & r.value & "'"
                End If
            End If
        Next

        'Insert deleted record into archive table
        cnn.Execute ("EXEC insert" & strTable & " " & strVal & ",'" & Application.Username & "'")

        'Delete record from main programs table
        cnn.Execute ("EXEC delete" & strTable & " '" & Cells(iRow, iDel).value & "'")
    End With

    'Free objects
    Set cnn = Nothing

End Sub

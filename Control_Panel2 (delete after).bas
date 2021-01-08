Attribute VB_Name = "Control_Panel"

Sub OpCo_List()

    'Clear anything already up
    Cancel_Cust_Add

    'Call Cust_Add Sub (Shared views)
    Cust_Add

    'Freeze processes
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Update list to be just active customers
    With Sheets("DropDowns")

        Sheets("Control Panel").Cust_Add_Listbox.List = _
        .Range("H1:H" & .Range("H" & .Rows.Count).End(xlUp).Row + 1).value
    End With

    'Add OpCo List select key
    Sheets("Control Panel").Shapes("OpCo_List_Select").Visible = True

    'Reset settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Sub OpCo_List_Select()

    'Declare global variable
    Dim cnn         As New ADODB.Connection
    Dim rst         As New ADODB.Recordset
    Dim strCust     As String
    Dim strPacket   As String
    Dim strUid      As String
    Dim strPwd      As String
    Dim iErr        As Integer
    Dim i           As Long

    'Establish connection to SSMS
    cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Loop through all dropdowns to create string
    With Sheets("Control Panel").Cust_Add_Listbox
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                If strCust = "" Then
                    strCust = "'" & .List(i) & "'"
                Else
                    strCust = strCust & ",'" & .List(i) & "'"
                End If
            End If
        Next
    End With

    'hide customer additions pane
    Cancel_Cust_Add

    'Freeze processes
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Ensure something was selected
    If strCust = "" Or strCust = "(Blank)" Then
        MsgBox "You must select at least one customer"
        GoTo ResetSettings
    End If

    'Pull in Programs sheet additions
    rst.Open "SELECT DISTINCT PACKET " _
        & "FROM UL_Customer_Profile " _
        & "WHERE CUSTOMER IN (" & strCust & ")", cnn

    'Create string of packets
    Do While rst.EOF = False
        If strPacket = "" Then
            strPacket = "'" & rst.Fields("PACKET").value & "'"
        Else
            strPacket = strPacket & ",'" & rst.Fields("PACKET").value & "'"
        End If

        'Iterate recordset
        rst.MoveNext
    Loop

    'Close connection to SQL server and reopen in sus (240)
    rst.Close
    cnn.Close

    'Get username and password
    strUid = get_uid
    strPwd = get_pwd

    'Connect to OpCo
    On Error GoTo OpErr
    cnn.Open "DSN=AS240A;UID=" & strUid & ";PASSWORD=" & strPwd & ";"
    On Error GoTo ResetSettings

    'Ensure proper packet information is found
    If strPacket = "" Then
        MsgBox Replace(strCust, "'", "") & " has a missing or invalid packet name. Please validate the PACKET column on the Customer Profile tab and try again."
        GoTo ResetSettings
    End If

    'Query OpCo list for each packet
    rst.Open "SELECT DISTINCT TRIM(DVPKGS) AS PCKT, TRIM(DVT500) AS OP " _
        & "FROM SCDBFP10.PMDPDVRF " _
        & "INNER JOIN (" _
            & "SELECT DVPKGS AS PACKET, MAX(LENGTH(TRIM(DVT500))) AS LEN " _
            & "FROM SCDBFP10.PMDPDVRF " _
            & "WHERE TRIM(DVPKGS) IN (" & strPacket & ") " _
            & "GROUP BY DVPKGS) " _
        & "ON DVPKGS = PACKET AND LENGTH(TRIM(DVT500)) = LEN ", cnn

    'Loop through recordset
    Do While rst.EOF = False

        'Add sheet/workbook
        If ActiveSheet.Name = "Control Panel" Then
            Workbooks.Add
        Else
            Sheets.Add
        End If

        'Format
        ActiveSheet.Name = rst.Fields("PCKT").value
        Cells.Interior.Color = vbWhite
        Range("A1").value = "Packet"
        Range("B1").value = "OpCo List"
        With Range("A1:B1")
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = vbWhite
            .Font.Bold = True
        End With

        'Paste OpCo into list
        For i = 1 To Len(rst.Fields("OP").value)
            Cells(Rows.Count, 1).End(xlUp).Offset(1, 0) = rst.Fields("PCKT").value
            Cells(Rows.Count, 2).End(xlUp).Offset(1, 0) = "'" & Mid(rst.Fields("OP").value, i, 3)
            i = i + 2
        Next

        'Add borders and fit columns
        Range("A1:B" & Range("B" & Rows.Count).End(xlUp).Row).Borders.LineStyle = xlContinuous
        Columns.AutoFit

        'Iterate loop
        rst.MoveNext
    Loop

    'Free objects
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing

    Exit Sub

    'Could not connect to OpCo
    OpErr:

    'If catastrophic error
    If InStr(Err.Description, "Catastrophic") > 0 Then
        MsgBox "OBDC overload. Please close all open instances of Excel and try again."

        'Free objecs & delete temp sheet
        Call free_obj(cnn, rst)

        'End macro
        Exit Sub

    'If invalid password
    ElseIf iErr < 2 Then
        MsgBox "SUS credentials missing/expired. Please validate your username/password and ensure you have access to OpCo as240a."
        UserLog.Show
        strUid = get_uid
        strPwd = get_pwd
        iErr = iErr + 1
        Resume

    'Skip OpCo
    Else
        MsgBox "Could not reach OpCo. Please validate you have access to as240a"
        GoTo ResetSettings
    End If


End Sub

Sub Recover_Records()

    'Declare variables
    Dim cnn                 As New ADODB.Connection
    Dim rst                 As New ADODB.Recordset
    Dim ws                  As Worksheet
    Dim strProgramFields    As String
    Dim strCnn              As String

    'Freeze processes
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Unhide recovery sheet
    Sheets("Recover Deviation Loads").Visible = True
    Sheets("Recover Cust Profile").Visible = True
    Sheets("Recover Programs").Visible = True

    'Loop through worksheets to hide
    For Each ws In Worksheets
        If Not ws.Name Like "*Recover*" Then
            ws.Visible = False
        End If
    Next

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

    'Query deleted records
    rst.Open "SELECT " & strProgramFields & " " _
        & "FROM UL_Deleted_Programs " _
        & "WHERE CUSTOMER_ID IN (" _
            & "SELECT CUSTOMER_ID " _
            & "FROM UL_Account_Ass " _
            & "WHERE TIER_1 = '" & Application.Username & "' " _
            & "OR TIER_2 = '" & Application.Username & "' " _
            & "OR T1_ID = '" & Environ("Username") & "' " _
            & "OR T2_ID = '" & Environ("Username") & "') " _
        & "OR DEL_USER = '" & Application.Username & "' " _
        & "ORDER BY CUSTOMER, PROGRAM_DESCRIPTION", cnn

    'Clear previous sheet contents and paste new (pause protection)
    With Sheets("Recover Programs")
        .Unprotect "Dac123am"
        .Rows(1).AutoFilter
        .Range("A3:A" & .Range("A" & .Rows.Count).End(xlUp).Row + 1).EntireRow.Delete
        .Range("A3").CopyFromRecordset rst
        .Range(.Cells(.Rows.Count, 3).End(xlUp), .Cells(2, .Columns.Count).End(xlToLeft)).Borders.LineStyle = xlContinuous
        .Rows(1).AutoFilter
        .Protect Password:="Dac123am", UserInterFaceOnly:=True
    End With

    'Close recordset for next sheet
    rst.Close

    'Setup query string to pull in Customer Profile sheet
    rst.Open "SELECT * " _
        & "FROM UL_Deleted_Customer_Profile " _
        & "WHERE CUSTOMER_ID IN (" _
            & "SELECT CUSTOMER_ID " _
            & "FROM UL_Account_Ass " _
            & "WHERE TIER_1 = '" & Application.Username & "' " _
            & "OR TIER_2 = '" & Application.Username & "' " _
            & "OR T1_ID = '" & Environ("Username") & "' " _
            & "OR T2_ID = '" & Environ("Username") & "') " _
        & "OR DEL_USER = '" & Application.Username & "' " _
        & "ORDER BY CUSTOMER", cnn

    'Clear previous sheet contents and paste new
    With Sheets("Recover Cust Profile")
        .Unprotect "Dac123am"
        .Rows(1).AutoFilter
        .Range("A3:A" & .Range("A" & .Rows.Count).End(xlUp).Row + 1).EntireRow.Delete
        .Range("A3").CopyFromRecordset rst
        .Range(.Cells(.Rows.Count, 3).End(xlUp), .Cells(2, .Columns.Count).End(xlToLeft)).Borders.LineStyle = xlContinuous
        .Rows(1).AutoFilter
        .Protect Password:="Dac123am", UserInterFaceOnly:=True
    End With

    'Close recordset for next sheet
    rst.Close

    'Setup query string to pull in deviation loads sheet
    rst.Open "SELECT * " _
        & "FROM UL_Deleted_Deviation_Loads " _
        & "WHERE CUSTOMER_ID IN (" _
            & "SELECT CUSTOMER_ID " _
            & "FROM UL_Account_Ass " _
            & "WHERE TIER_1 = '" & Application.Username & "' " _
            & "OR TIER_1 = '" & Application.Username & "' " _
            & "OR T1_ID = '" & Environ("Username") & "' " _
            & "OR T2_ID = '" & Environ("Username") & "') " _
        & "OR DEL_USER = '" & Application.Username & "' " _
        & "ORDER BY CUSTOMER, PROGRAM", cnn

    'Clear previous sheet contents and paste new
    With Sheets("Recover Deviation Loads")
        .Unprotect "Dac123am"
        .Rows(1).AutoFilter
        .Range("A3:A" & .Range("A" & .Rows.Count).End(xlUp).Row + 1).EntireRow.Delete
        .Range("A3").CopyFromRecordset rst
        .Range(.Cells(.Rows.Count, 3).End(xlUp), .Cells(2, .Columns.Count).End(xlToLeft)).Borders.LineStyle = xlContinuous
        .Rows(1).AutoFilter
        .Protect Password:="Dac123am", UserInterFaceOnly:=True
    End With

    'Ensure recover programs is active
    Sheets("Recover Programs").Activate

    'Close recordset and free object
    rst.Close
    Set cnn = Nothing

    'Reset settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Sub Cancel_Recover()

    'Freeze processes
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Hide and unhide sheets
    Sheets("Programs").Visible = True
    Sheets("Customer Profile").Visible = True
    Sheets("Deviation Loads").Visible = True
    Sheets("Control Panel").Visible = True
    Sheets("Recover Programs").Visible = False
    Sheets("Recover Cust Profile").Visible = False
    Sheets("Recover Deviation Loads").Visible = False

    'Ensure control panel is active window
    Sheets("Control Panel").Activate

    'Reset settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Sub Confirm_Recover()

    'Delcare variables
    Dim cnn                 As New ADODB.Connection
    Dim rst                 As New ADODB.Recordset
    Dim olA                 As New Outlook.Application
    Dim olUser              As Object
    Dim strFields           As String
    Dim strCnn              As String
    Dim strVal              As String
    Dim iRow                As Long
    Dim i                   As Long

    'Freeze processes
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Ensure only one row is selected
    If Selection.Rows.Count > 1 Then
        MsgBox "Please select one row at a time"
        GoTo ResetSettings
    Else
        iRow = Selection.Row
    End If

    'Ensure blank row was not inserted
    If Cells(iRow, 1) = "" Or iRow = 1 Or iRow = 2 Then
        GoTo ResetSettings
    End If

    'Establish connection to SSMS
    strCnn = "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"
    cnn.Open strCnn
    cnn.CommandTimeout = 900

    'Determine which sheet is active
    If ActiveSheet.Name = "Recover Programs" Then

        'Setup fields variable
        strFields = "CUSTOMER_ID," _
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

        'Loop through selected row to create insert value
        For i = 2 To Cells(2, Columns.Count).End(xlToLeft).Column

            'Create insert string
            If strVal = "" Then
                strVal = Cells(iRow, i)
            Else
                If i = 15 Then
                    strVal = strVal & "," & Cells(iRow, i)
                Else
                    strVal = strVal & ",'" & Cells(iRow, i) & "'"
                End If
            End If
        Next

        'Insert deleted record back into programs
        cnn.Execute ("EXEC insert_programs " & strVal)

        'Query acount assignments
        rst.Open "SELECT CUSTOMER_ID " _
            & "FROM UL_Account_Ass " _
            & "WHERE CUSTOMER_ID = " & Cells(iRow, 2), cnn

        'If no account assignment is present, create one
        If rst.EOF = True Then

            'Find user manager
            Set olUser = olA.GetNamespace("MAPI").CreateRecipient(Application.Username).AddressEntry.GetExchangeUser.Manager

            'Insert into account assignment
            cnn.Execute ("EXEC insert_new_id " & Cells(iRow, 2).value & ",'" & strName & "','" _
                & olUser & "','" & Application.Username & "','" & Application.Username & "','" _
                & olA.GetNamespace("MAPI").CreateRecipient(olUser).AddressEntry.GetExchangeUser.Alias & "','" _
                & Environ("Username") & "','" & Environ("Username") & "'")
        End If

        'Delete record from deleted database
        cnn.Execute ("EXEC delete_final_programs '" & Cells(iRow, 1).value & "'")

    'Determine which sheet is being updated
    ElseIf ActiveSheet.Name = "Recover Cust Profile" Then

        'Setup fields variable
        strFields = "CUSTOMER_ID," _
            & "CUSTOMER," _
            & "ALT_NAME," _
            & "PACKET," _
            & "PRICE_RULE," _
            & "NID," _
            & "MASTER_PRN," _
            & "PRICING_PRN," _
            & "GROUP_NAME," _
            & "VPNA," _
            & "NAM," _
            & "CUST_CONTACT"

        'Loop through selected row to create insert value
        For i = 2 To Cells(2, Columns.Count).End(xlToLeft).Column

            'Create insert string
            If strVal = "" Then
                strVal = Cells(iRow, i)
            Else
                strVal = strVal & ",'" & Cells(iRow, i) & "'"
            End If
        Next

        'Insert deleted record back into programs
        cnn.Execute ("EXEC insert_customer " & strVal)

        'Delete record from deleted database
        cnn.Execute ("EXEC delete_final_customer " & Cells(iRow, 1))

    'Determine which sheet is being updated
    ElseIf ActiveSheet.Name = "Recover Deviation Loads" Then

        'Setup fields variable
        strFields = "CUSTOMER_ID," _
            & "CUSTOMER," _
            & "PROGRAM," _
            & "OWNER," _
            & "DATE," _
            & "SR"

        'Loop through selected row to create insert value
        For i = 2 To Cells(2, Columns.Count).End(xlToLeft).Column

            'Create insert string
            If strVal = "" Then
                strVal = Cells(iRow, i)
            Else
                strVal = strVal & ",'" & Cells(iRow, i) & "'"
            End If
        Next

        'Insert deleted record back into programs
        cnn.Execute ("EXEC insert_deviation " & strVal)

        'Delete record from deleted database
        cnn.Execute ("EXEC delete_final_deviation " & Cells(iRow, 1))
    End If

    'Delete selected row and move cursor
    Rows(iRow).Delete
    Cells(2, 1).Activate

    ResetSettings:
    'Reset settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Sub Overlap_Validation()

    'Freeze screen updating
    Application.ScreenUpdating = False

    'Unhide appropriate objects
    With ActiveSheet
        If .Shapes("Help_Pane").Visible = True Then
            Help_Cancel
        End If
        If .Shapes("Cust_Add_Pane").Visible = True Then
            Cancel_Cust_Add
        End If
        If .Shapes("Item_Lookup_Pane").Visible = True Then
            Item_Lookup_Cancel
        End If
        .Shapes("Overlap_Pane").Visible = True
        .Shapes("Overlap_Opco").Visible = True
        .Shapes("Overlap_VA").Visible = True
        .Shapes("Overlap_CA").Visible = True
        .Shapes("Overlap_Textbox_1").Visible = True
        .Shapes("Overlap_Textbox_2").Visible = True
        .Shapes("Overlap_Cancel").Visible = True
        .Shapes("Overlap_Validate").Visible = True
    End With

    'Unfreeze screen updating
    Application.ScreenUpdating = True

End Sub

Sub Overlap_Validation_Cancel()

    'Freeze screen updating
    Application.ScreenUpdating = False

    'Hide appropriate objects
    With ActiveSheet
        .Shapes("Overlap_Pane").Visible = False
        .Shapes("Overlap_Opco").Visible = False
        .Shapes("Overlap_VA").Visible = False
        .Shapes("Overlap_CA").Visible = False
        .Shapes("Overlap_Textbox_1").Visible = False
        .Shapes("Overlap_Textbox_2").Visible = False
        .Shapes("Overlap_Cancel").Visible = False
        .Shapes("Overlap_Validate").Visible = False
        .TextBoxes("Overlap_Textbox_1").Text = ""
        .TextBoxes("Overlap_Textbox_2").Text = ""
    End With

    'Unfreeze screen updating
    Application.ScreenUpdating = True

End Sub

Sub Overlap_Validation_Run()

    'Declare variables
    Dim cnn         As New ADODB.Connection
    Dim rst         As New ADODB.Recordset
    Dim strOpCo     As String
    Dim strAgmt1    As String
    Dim strAgmt2    As String
    Dim strUid      As String
    Dim strPwd      As String
    Dim iErr        As Integer
    Dim blAgmt      As Boolean

    'Freeze screen updating
    Application.ScreenUpdating = False

    'Get inputs from input form
    strOpCo = Trim(ActiveSheet.TextBoxes("Overlap_Opco").Text)
    strAgmt1 = Trim(ActiveSheet.TextBoxes("Overlap_Textbox_1").Text)
    strAgmt2 = Trim(ActiveSheet.TextBoxes("Overlap_Textbox_2").Text)
    If ActiveSheet.Shapes("Overlap_VA").ControlFormat.value = xlOn Then blAgmt = True

    'Ensure there is data to parse
    If strOpCo = "" Or strAgmt1 = "" Or strAgmt2 = "" Then
        MsgBox "Missing Data:" & vbNewLine & "OpCo, and two agreements must be indicated on form."
        GoTo ResetSettings
    End If

    'Create new sheet and add headers
    Workbooks.Add
    Range("A1").value = "PRN/GRP"
    Range("B1").value = "SHP"
    Range("C1").value = "ITEM"

    'Get username and password
    strUid = get_uid
    strPwd = get_pwd

    'Connect to OpCo
    On Error GoTo OpErr
    cnn.Open "DSN=AS" & strOpCo & "A;UID=" & strUid & ";PASSWORD=" & strPwd & ";"
    On Error GoTo ResetSettings

    'Determine if entry was VA or CA number
    If blAgmt = True Then

        'VA cust query
        rst.Open "SELECT DISTINCT TRIM(T2.QWPCSC) || ' ' || TRIM(T2.QWPCSP), TRIM(T2.QWCUNO) " _
            & "FROM (" _
                & "SELECT QWCUNO, COUNT(QWCUNO) " _
                & "FROM SCDBFP10.PMPZQWPF " _
                & "WHERE QWVAGN IN (" & strAgmt1 & "," & strAgmt2 & ") " _
                & "GROUP BY QWCUNO " _
                & "HAVING COUNT(QWCUNO) > 1) AS T1 " _
            & "LEFT JOIN (" _
                & "SELECT QWPCSC, QWPCSP, QWCUNO " _
                & "FROM SCDBFP10.PMPZQWPF " _
                & "WHERE QWVAGN IN (" & strAgmt1 & "," & strAgmt2 & ")) AS T2 " _
            & "ON T1.QWCUNO = T2.QWCUNO", cnn

        'Paste query
        Range("A2").CopyFromRecordset rst
        Columns(1).RemoveDuplicates Columns:=Array(1)
        Columns(2).RemoveDuplicates Columns:=Array(1)

        'Close recordset in prep for item query
        rst.Close

        'Pull overlapping items
        rst.Open "SELECT TRIM(QBITEM) " _
            & "FROM SCDBFP10.PMPZQBPF " _
            & "WHERE QBVAGN IN (" & strAgmt1 & "," & strAgmt2 & ") " _
            & "GROUP BY QBITEM " _
            & "HAVING COUNT(QBITEM) > 1", cnn

        'Paste query
        Range("C2").CopyFromRecordset rst

    'CA query
    Else

        'CA cust query
        rst.Open "SELECT DISTINCT TRIM(T2.QYPCSC) || ' ' || TRIM(T2.QYPCSP), TRIM(T2.QYCUNO) " _
            & "FROM (" _
                & "SELECT QYCUNO, COUNT(QYCUNO) " _
                & "FROM SCDBFP10.PMPZQYPF " _
                & "WHERE QYCANO IN (" & strAgmt1 & "," & strAgmt2 & ") " _
                & "GROUP BY QYCUNO " _
                & "HAVING COUNT(QYCUNO) > 1) AS T1 " _
            & "LEFT JOIN (" _
                & "SELECT QYPCSC, QYPCSP, QYCUNO " _
                & "FROM SCDBFP10.PMPZQYPF " _
                & "WHERE QYCANO IN (" & strAgmt1 & "," & strAgmt2 & ")) AS T2 " _
            & "ON T1.QYCUNO = T2.QYCUNO", cnn

        'Paste query
        Range("A2").CopyFromRecordset rst
        Columns(1).RemoveDuplicates Columns:=Array(1)
        Columns(2).RemoveDuplicates Columns:=Array(1)

        'Close recordset in prep for item query
        rst.Close

        'Pull overlapping items
        rst.Open "SELECT TRIM(QXITEM) " _
            & "FROM SCDBFP10.PMPZQXPF " _
            & "WHERE QXCANO IN (" & strAgmt1 & "," & strAgmt2 & ") " _
            & "GROUP BY QXITEM " _
            & "HAVING COUNT(QXITEM) > 1", cnn

        'Paste Query
        Range("C2").CopyFromRecordset rst
    End If

    'Format
    Columns.AutoFit
    Cells.Interior.Color = vbWhite
    With Range("A1:C1")
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    Range("A1:C" & ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row).Borders.LineStyle = xlContinuous

    'Free Objects
    Set cnn = Nothing

    ResetSettings:
    'unfreeze screen updating
    Application.ScreenUpdating = True

    Exit Sub

    'Could not connect to OpCo
    OpErr:

    'If catastrophic error
    If InStr(Err.Description, "Catastrophic") > 0 Then
        MsgBox "OBDC overload. Please close all open instances of Excel and try again."

        'Free objecs & delete temp sheet
        Call free_obj(cnn, rst)

        'End macro
        Exit Sub

    'If invalid password
    ElseIf iErr < 2 Then
        MsgBox "SUS credentials missing/expired. Please validate your username/password and ensure you have access to OpCo " & strOpCo & "."
        UserLog.Show
        strUid = get_uid
        strPwd = get_pwd
        iErr = iErr + 1
        Resume

    'Skip OpCo
    Else
        MsgBox "Could not reach OpCo. Please validate you have access to OpCo " & strOpCo & "."
        GoTo ResetSettings
    End If

End Sub

Sub Item_Lookup()

    'Puase screeen updating
    Application.ScreenUpdating = False

    'Hide lookup objects
    With ActiveSheet
        If .Shapes("Help_Pane").Visible = True Then
            Help_Cancel
        End If
        If .Shapes("Cust_Add_Pane").Visible = True Then
            Cancel_Cust_Add
        End If
        If .Shapes("Overlap_pane").Visible = True Then
            Overlap_Validation_Cancel
        End If
        .Shapes("Item_Lookup_Pane").Visible = True
        .Shapes("Item_Lookup_MPC").Visible = True
        .Shapes("Item_Lookup_GTIN").Visible = True
        .Shapes("Item_Lookup_Search").Visible = True
        .Shapes("Item_Lookup_Cancel").Visible = True
        .Shapes("Item_Lookup_List").Visible = True
    End With

    'Unpause screen updating
    Application.ScreenUpdating = True

End Sub

Sub Item_Lookup_Cancel()

    'Puase screeen updating
    Application.ScreenUpdating = False

    'Hide lookup objects
    With ActiveSheet
        .Unprotect "Dac123am"
        .Shapes("Item_Lookup_Pane").Visible = False
        .Shapes("Item_Lookup_MPC").Visible = False
        .Shapes("Item_Lookup_GTIN").Visible = False
        .Shapes("Item_Lookup_Search").Visible = False
        .Shapes("Item_Lookup_Cancel").Visible = False
        .Shapes("Item_Lookup_List").Visible = False
        .TextBoxes("Item_Lookup_MPC").Text = ""
        .TextBoxes("Item_Lookup_GTIN").Text = ""
        .Columns("P:Q").Hidden = True
        .Columns("P:Q").Locked = True
        .Columns("P:Q").Font.Color = vbWhite
        .Range("P6:Q" & .Rows.Count).value = ""
        .Range("P6:Q" & .Rows.Count).Borders.LineStyle = xlNone
        .Range("P6:Q" & .Rows.Count).Interior.Color = RGB(128, 128, 128)
        .Protect "Dac123am"
    End With

    'Unpause screen updating
    Application.ScreenUpdating = True

End Sub

Sub Item_Lookup_Search()

    'Declare variables
    Dim cnn         As New ADODB.Connection
    Dim rst         As New ADODB.Recordset
    Dim strMPC      As String
    Dim strGTIN     As String
    Dim strUid      As String
    Dim strPwd      As String
    Dim iErr        As Integer

    'Freeze screen updating
    Application.ScreenUpdating = False

    'Get variables from user input
    If Range("P6").value = "" And Range("Q6").value = "" Then

        'Get values from single input form
        strMPC = "'" & Trim(ActiveSheet.TextBoxes("Item_Lookup_MPC").Text) & "'"
        strGTIN = Trim(ActiveSheet.TextBoxes("Item_Lookup_GTIN").Text)
    Else

        'Create string for MPCs
        If Range("P6").value <> "" Then
            For Each r In Range("P6:P" & Range("P" & Rows.Count).End(xlUp).Row)
                If strMPC = "" Then
                    strMPC = "'" & Trim(r.value) & "'"
                Else
                    strMPC = strMPC & ",'" & Trim(r.value) & "'"
                End If
            Next
        End If

        'Create string for GTINS
        If Range("Q6").value <> "" Then
            For Each r In Range("Q6:Q" & Range("Q" & Rows.Count).End(xlUp).Row)
                If strGTIN = "" Then
                    strGTIN = Trim(r.value)
                Else
                    strGTIN = strGTIN & "," & r.value
                End If
            Next
        End If
    End If

    'Ensure there is data to parse
    If strMPC = "''" And strGTIN = "" Then
        MsgBox "At least one search criteria must be indicated."
        GoTo ResetSettings
    End If

    'Create new sheet and add headers
    Workbooks.Add
    Range("A1").value = "SUPC"
    Range("B1").value = "PACK/SIZE"
    Range("C1").value = "BRAND"
    Range("D1").value = "DESCRIPTION"
    Range("E1").value = "MPC"
    Range("F1").value = "GTIN"

    'Get username and password
    strUid = get_uid
    strPwd = get_pwd

    'Connect to OpCo
    On Error GoTo OpErr
    cnn.Open "DSN=AS240A;UID=" & strUid & ";PASSWORD=" & strPwd & ";"
    On Error GoTo ResetSettings

    'If GTIN is supplied
    If strGTIN <> "" Then

        'Query for SUPC using GTIN
        rst.Open "SELECT DISTINCT TRIM(JFITEM), TRIM(JFPACK) || '/' || TRIM(JFITSZ), TRIM(JFBRND), TRIM(JFITDS), TRIM(JFMNPC), JFEUPC " _
            & "FROM SCDBFP10.USIAJFPF " _
            & "WHERE JFEUPC IN (" & strGTIN & ")", cnn

        'if query returned no data
        If rst.EOF Or InStr(strMPC, ",") > 0 Then

            'Ensure data is in MPC column
            If strMPC <> "''" And strMPC <> "" Then

                'Close recordset in prep for next run
                rst.Close

                'Query for SUPC usinc MPC
                rst.Open "SELECT DISTINCT TRIM(JFITEM), TRIM(JFPACK) || '/' || TRIM(JFITSZ), TRIM(JFBRND), TRIM(JFITDS), TRIM(JFMNPC), JFEUPC " _
                    & "FROM SCDBFP10.USIAJFPF " _
                    & "WHERE TRIM(JFMNPC) IN (" & strMPC & ")", cnn

                'If query returned no data, close workbook
                If rst.EOF Then
                    ActiveWorkbook.Close SaveChanges:=False
                    MsgBox "No items were found."
                    GoTo ResetSettings
                End If
            Else

                'If query returned no data, close workbook
                ActiveWorkbook.Close SaveChanges:=False
                MsgBox "No items were found."
                GoTo ResetSettings
            End If
        End If

        'Paste Recordset to sheet
        Range("A2").CopyFromRecordset rst

    'If not GTIN is supplied
    Else

        'Query for SUPC usinc MPC
        rst.Open "SELECT DISTINCT TRIM(JFITEM), TRIM(JFPACK) || '/' || TRIM(JFITSZ), TRIM(JFBRND), TRIM(JFITDS), TRIM(JFMNPC), JFEUPC " _
            & "FROM SCDBFP10.USIAJFPF " _
            & "WHERE TRIM(JFMNPC) IN (" & strMPC & ")", cnn

        'If query returned no data, close workbook
        If rst.EOF Then
        ActiveWorkbook.Close SaveChanges:=False
            MsgBox "No items were found."
            GoTo ResetSettings
        End If

        'Paste recordset to sheet
        Range("A2").CopyFromRecordset rst
    End If

    'Format
    Columns("F").NumberFormat = "0"
    Columns.AutoFit
    Cells.Interior.Color = vbWhite
    Range("A1:F" & ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row).Borders.LineStyle = xlContinuous
    With Range("A1:F1")
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With

    ResetSettings:
    Application.ScreenUpdating = True
    Set cnn = Nothing

    Exit Sub

    'Could not connect to OpCo
    OpErr:

    'If catastrophic error
    If InStr(Err.Description, "Catastrophic") > 0 Then
        MsgBox "OBDC overload. Please close all open instances of Excel and try again."

        'Free objecs & delete temp sheet
        Call free_obj(cnn, rst)

        'End macro
        Exit Sub

    'If invalid password
    ElseIf iErr < 2 Then
        MsgBox "SUS credentials missing/expired. Please validate your username/password and ensure you have access to OpCo as240a."
        UserLog.Show
        strUid = get_uid
        strPwd = get_pwd
        iErr = iErr + 1
        Resume

    'Skip OpCo
    Else
        MsgBox "Could not reach OpCo. Please validate you have access to OpCo as240a."
        GoTo ResetSettings
    End If


End Sub

Sub Item_Lookup_List()

    'Freeze screen updating
    Application.ScreenUpdating = False

    'Hide and unhid appropriate objects
    With ActiveSheet
        .Unprotect "Dac123am"
        .Shapes("Item_Lookup_List").Visible = False
        .Columns("P:Q").Hidden = False
        .Columns("P:Q").Locked = False
        .Protect "Dac123am"
    End With

    'Unfreeze screen updating
    Application.ScreenUpdating = True

End Sub

Sub Request_Automation()

    'Freeze events
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Unhide all objects
    With ActiveSheet
        If .Shapes("Cust_Add_Pane").Visible = True Then
            Cancel_Cust_Add
        End If
        If .Shapes("Overlap_Pane").Visible = True Then
            Overlap_Validation_Cancel
        End If
        If .Shapes("Item_Lookup_Pane").Visible = True Then
            Item_Lookup_Cancel
        End If
        .Shapes("Help_Pane").Visible = True
        .Shapes("Help_Label").Visible = True
        .Shapes("Help_Body").Visible = True
        .Shapes("Help_Send").Visible = True
        .Shapes("Help_Cancel").Visible = True
        .Shapes ("Help_Pane")
    End With

    'Reset settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Function get_pwd() As String

    'Declare variables
    Dim cnn     As New ADODB.Connection
    Dim rst     As New ADODB.Recordset
    Dim i       As Integer
    Dim StrDec  As String
    Dim strPwd  As String
    Dim StrKey  As String

    'Establish connection to SQL Server
    cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query uid & password
    rst.Open "SELECT SUS_PWD, CRED_ID FROM Login_Cred WHERE NET_ID = '" & Environ("Username") & "'", cnn

    'if no records
    If rst.EOF = True Then
        StrDec = "null"
        GoTo EOF
    End If

    'save value to string
    strPwd = rst.Fields("SUS_PWD").value
    StrKey = rst.Fields("CRED_ID").value

    'Loop through each character in password
    For i = 1 To Len(strPwd)

        'Decrypt character of password
        StrDec = StrDec & Chr(Asc(Mid(strPwd, i, 1)) - Mid(StrKey, i, 1))
    Next

    EOF:
    'Close and free objects
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing

    'Return decrypted password
    get_pwd = StrDec

End Function

Function get_uid() As String

    'Declare variables
    Dim cnn     As New ADODB.Connection
    Dim rst     As New ADODB.Recordset
    Dim strUid  As String

    'Establish connection to SQL Server
    cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'query uid
    rst.Open "SELECT SUS_ID FROM Login_Cred WHERE NET_ID = '" & Environ("Username") & "'", cnn

    'If no records then go to end
    If rst.EOF Then
        strUid = "null"
        GoTo EOF
    End If

    'Save password to memory
    strUid = rst.Fields("SUS_ID").value

    EOF:
    'Close and free objects
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing

    'Return uid
    get_uid = strUid

End Function

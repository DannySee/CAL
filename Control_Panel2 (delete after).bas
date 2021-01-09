Attribute VB_Name = "Control_Panel"

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

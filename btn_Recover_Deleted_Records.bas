Attribute VB_Name = "btn_Recover_Deleted_Records"

'Declare private module constants
Private Const varSht As Variant = _Array(oPrgms.Shtx, oCst.Shtx, oDev.Shtx)


'*******************************************************************************
'Show all utility elements and popluate data
'*******************************************************************************
Private Sub Recover_Deleted_Initialize()

    'Declare module variables
    Dim strCst As String

    'Get String of my customer names
    strCst = GetStr(Pull.GetCst(True), True)

    'Clear data from sheets. Parameters: utility sheets + header row
    Utility.ClearShts(varSht, 3)

    'Show utility sheets
    Utility.ShowSheets(varSht, True)

    'Hide all other Sheets (Sheets that do not contain key word Recover)
    Utility.SheetVisible("Recover", False)

    'Format sheets and insert updated server data
    Utility.ShtRefresh(oPrgms.Shtx, Pull.GetDelRecords(strCst, oPrgms.Dbx))
    Utility.ShtRefresh(oCst.Shtx, Pull.GetDelRecords(strCst, oCst.Dbx))
    Utility.ShtRefresh(oDev.Shtx, Pull.GetDelRecords(strCst, oDev.Dbx))

    'Ensure recover programs is active
    Sheets(oPrgms.Shtx).Activate
End Sub


'*******************************************************************************
'Recover deleted program records
'*******************************************************************************
Sub Recover_Deleted_Prgm_Confirm()

    'Declare sub variables
    Dim iRow As Long

    'Ensure only one row is selected
    If Selection.Rows.Count = 1 Then

        'Set row number to variable
        iRow = Selection.Row


    'If more than one row was selected
    Else

        'Alert user of incorrect process
        MsgBox "Please select one row at a time"
    End If
End Sub


'*******************************************************************************
'Recover deleted customer profile records
'*******************************************************************************
Sub Recover_Deleted_Cst_Confirm()


End Sub


'*******************************************************************************
'Recover deleted program records
'*******************************************************************************
Sub Recover_Deleted_Dev_Confirm()


End Sub


'*******************************************************************************
'Reset window view (hide utility sheets and show defaults).
'*******************************************************************************
Sub Recover_Deleted_Cancel()

    'Show all sheets
    Utility.SheetVisible("Recover", True)

    'hide utility sheets
    Utility.ShowSheets(varSht, False)

    'Make control panel the active screen upon completion
    Sheets("Control Panel").Activates
End Sub

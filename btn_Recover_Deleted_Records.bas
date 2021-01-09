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
    Dim iPKey As Long

    'Get primary key of selected row
    iPKey = Utility.GetArchiveKey

    'If a selection was made insert recovery line into main table (UL_Programs)
    If iPKey <> 0 Then Push.RecoverDeleted(oPrgms, iPKey)
End Sub


'*******************************************************************************
'Recover deleted Customer Profile records
'*******************************************************************************
Sub Recover_Deleted_Cst_Confirm()

    'Declare sub variables
    Dim iPKey As Long

    'Get primary key of selected row
    iPKey = Utility.GetArchiveKey

    'If a selection was made insert recovery line into main table (UL_Programs)
    If iPKey <> 0 Then Push.RecoverDeleted(oCst, iPKey)
End Sub


'*******************************************************************************
'Recover deleted Deviation Loads records
'*******************************************************************************
Sub Recover_Deleted_Dev_Confirm()

    'Declare sub variables
    Dim iPKey As Long

    'Get primary key of selected row
    iPKey = Utility.GetArchiveKey

    'If a selection was made insert recovery line into main table (UL_Programs)
    If iPKey <> 0 Then Push.RecoverDeleted(oDev, iPKey)
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

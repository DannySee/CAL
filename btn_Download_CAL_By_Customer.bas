Attribute VB_Name = "btn_Download_CAL_By_Customer"


'*******************************************************************************
'Download CAL account assignments by customer.
'*******************************************************************************
Sub Initialize()

    'Declare sub variables
    Dim varCst As Variant
    Dim strCst As String
    Dim wb As Workbook
    Dim rngHeaders As Range

    'Get header range
    Set rngHeaders = Sheets("Programs").Rows(1).UsedRange

    'Get folder path from user selection
    Utility.SelectFolder

    'Get array of my customer names
    varCst = Pull.GetCst(True)

    'Alert user and exit sub if no customers were selected
    If IsEmpty(varCst) Then Goto NoCst

    'Loop through all assigned customers
    For each cst In varCst

        'Create new workbook
        Set wb = Utility.CreateWorkbook(cst)

        'Setup customer string with quote delimiters
        strCst = "'" & cst & "'"

        'Format sheets and insert updated server date
        Format.ShtRefresh(oPrgms, Pull.GetPrograms(strCst))
    Next

'Jump to label to skip sub routine
NoCst:
End Sub

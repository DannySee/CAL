Attribute VB_Name = "btn_Download_CAL_By_Customer"


'*******************************************************************************
'Download CAL account assignments by customer. Testing
'*******************************************************************************
Sub Initialize()

    'Declare sub variables
    Dim varCst As Variant
    Dim varHeaders As Variant
    Dim strCst As String
    Dim strPth As String
    Dim wb As Workbook

    'Get folder path from user selection
    strPth = Utility.SelectFolder

    'Get array of my customer names
    varCst = Pull.GetCst(True)

    'Alert user and exit sub if no assigned customers/selected folder
    If IsEmpty(varCst) Or strPth = "" Then Goto NoCst

    'Get array of customer friendly headers
    varHeaders = oBtnDwn.Headers

    'Loop through all assigned customers
    For each cst In varCst

        'Create new workbook with formatting
        Set wb = Utility.CreateWorkbook(cst)

        'Setup customer string with quote delimiters
        strCst = "'" & cst & "'"

        'Format sheets and insert updated server date
        Cells(2,1).CopyFromRecordset Pull.GetPrograms(strCst)*******************

        'Add headers to sheet to sheet with formatting
        Utility.AddHeaders(varHeaders)
    Next

'Jump to label to skip sub routine
NoCst:
End Sub
